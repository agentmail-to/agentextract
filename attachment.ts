// attachment — text extraction for email attachments.
//
// In scope: text, HTML, PDF, Word (.docx + legacy .doc), and Excel (.xlsx). Out of scope: nested
// emails (.eml), images/OCR, legacy .xls/.ppt, and archives — those return a labeled skip.
// Ships as a separate subpath export ("agentextract/attachment") so the heavy
// parsers stay out of the body extractor's bundle. New formats = new registry entries.

import { isUtf8 } from 'node:buffer' // fast native check: are these bytes valid utf-8?
import zlib from 'node:zlib' // streaming raw-inflate for the OOXML decompression-budget guard

// Converts bytes into text using a specific character encoding
import iconv from 'iconv-lite'
// Guesses what encoding a chunk of bytes is in (eg. UTF-8, UTF-16, or windows-1252)
import jschardet from 'jschardet'

/////////////////////////////////////////////////////////////
// CONSTANTS (tunable)

// Max: 10 MB - Any attachment bigger than this gets rejected before any processing happens at all
export const MAX_INPUT_BYTES = 10 * 1024 * 1024

// Max extracted text length (JS chars). Input is byte-capped, but output isn't: a big
// spreadsheet or HTML table can balloon into megabytes that then hit S3 and the search index.
// Past this the text is truncated (bounded silently — there is no truncation flag on the result).
export const MAX_OUTPUT_CHARS = 250_000

// Max PDF pages to parse. A PDF can declare an enormous page count; we read text from at most this
// many and then stop. Bounds worst-case parse work when the per-page text is too sparse to trip the
// output cap (a content-bearing PDF hits MAX_OUTPUT_CHARS within a few dozen pages first). A typical
// email PDF is a handful of pages, so this is far above any real attachment while still capping abuse.
export const MAX_PDF_PAGES = 2000

// Max total ACTUAL uncompressed size of an OOXML (zip) attachment, MEASURED by streaming inflation
// (see the DECOMPRESSION BUDGET guard — the declared central-directory size is attacker-controlled
// and not trusted). A small in-cap .docx/.xlsx can still decompress to hundreds of MB and OOM the
// worker — MAX_INPUT_BYTES only bounds the COMPRESSED size. Parsers then build an in-memory model on
// top (~8x for a cell-dense sheet), so keep this well under the heavy-parser memory floor (1024 MB):
// 50 MB uncompressed → ~400 MB RSS worst case.
export const MAX_UNCOMPRESSED_BYTES = 50 * 1024 * 1024

// Guard: if a handler takes >= 10 seconds, we stop awaiting it. Note: this can't
// cancel synchronous CPU-bound work — it only stops us waiting on a slow async parse.
export const HANDLER_TIMEOUT_MS = 10_000

// "Sniff" the first 8KB of the buffer to see if it looks like text.
const SNIFF_BYTES = 8 * 1024
const SNIFF_TEXT_RATIO = 0.85

// Detecting which encoding a chunk of bytes is in (eg. UTF-8, UTF-16, or windows-1252) 
// Looking at the first 64KB of the buffer.
const DETECT_SAMPLE_BYTES = 64 * 1024
// If the confidence is less than 70%, we don't trust the detection.
const DETECT_MIN_CONFIDENCE = 0.7

/////////////////////////////////////////////////////////////
// TYPES

// The status of the extraction.
export type ExtractionStatus =
    | 'extracted' // handler ran (extraction holds the text, or is omitted when it produced none)
    | 'skipped' // over MAX_INPUT_BYTES, or unsupported/unrecognized type — never attempted
    | 'failed' // parser threw, timed out, or decode failed

// Records which signal ended up deciding the route.
export type RoutedBy = 'content-type' | 'extension' | 'sniff' | 'none'

// The formats we can route to. Each maps to exactly one handler.
// text: plain text, csv, calendar, vcard, json, xml, yaml
// html: html
// pdf: pdf
// docx: docx (modern OOXML Word)
// doc: doc (legacy OLE Word, Word 97–2003)
// xlsx: xlsx (modern OOXML Excel)
export type HandlerKind = 'text' | 'html' | 'pdf' | 'docx' | 'doc' | 'xlsx'

// The input to the extraction.
export interface AttachmentInput {
    content: Buffer // the raw bytes of the attachment
    filename?: string // the name of the attachment
    contentType?: string // the content type of the attachment (eg. text/plain, text/html, application/pdf, ...)
}

// The output of the extraction.
export interface ExtractionResult {
    status: ExtractionStatus // the status of the extraction
    extraction?: string // the extracted text; omitted entirely (never '') when the handler produced none
    reason?: string // set on skipped / failed
}

// The context for a handler.
interface HandlerContext {
    content: Buffer // the raw bytes of the attachment
    filename?: string // the name of the attachment
    charsetHint?: string // from the content-type charset= param
}

// The output of a handler.
interface HandlerOutput {
    text: string // the text of the attachment
    empty?: boolean // handler's own emptiness call; defaults to text.trim() === ''
}

// Handler takes context and returns an output. 
interface Handler {
    kind: HandlerKind // the type of the handler (eg. text, html, pdf, docx, xlsx)
    contentTypes: string[] // exact, lowercased, param-stripped
    extensions: string[] // with leading dot, lowercased
    extract: (ctx: HandlerContext) => Promise<HandlerOutput>
}

/////////////////////////////////////////////////////////////
// CHARSET-CORRECT DECODING (direct-text handlers)
// HAVE to ensure that the text is decoded correctly using the correct charset.

// A byte-order mark is an unambiguous encoding signal. 
// Sometimes the file tags itself, which is a very reliable signal. 
const bomCharset = (content: Buffer): string | undefined => {
    if (content.length >= 3 && content[0] === 0xef && content[1] === 0xbb && content[2] === 0xbf) return 'utf-8'
    if (content.length >= 2 && content[0] === 0xff && content[1] === 0xfe) return 'utf-16le'
    if (content.length >= 2 && content[0] === 0xfe && content[1] === 0xff) return 'utf-16be'
    return undefined
}

// Priority-ordered fallback chain (highest priority to lowest priority - try the most trustworthy thing, degrade gracefully).
const resolveCharset = (content: Buffer, hint?: string): string => {
    // 1. A BOM is definitive — the bytes declare their own encoding. Trust it above the
    //    Content-Type hint: a stale/wrong charset= param must not override an in-band BOM
    //    (e.g. a UTF-16 file mislabeled charset=windows-1252 would otherwise decode as garbage).
    const bom = bomCharset(content)
    if (bom) return bom
    // 2. Valid UTF-8 is self-evidencing — the multi-byte sequences are near-impossible to hit by
    //    accident, so trust the content over a hint that would mangle it. Pure-ASCII text is also
    //    valid utf-8 and decodes identically under latin1/windows-1252, so this never loses data;
    //    it only stops a wrong hint from corrupting genuine utf-8 multi-byte characters.
    //    Guard on NUL: BOM-less UTF-16 ASCII (h\0e\0l\0...) is technically valid utf-8 yet is really
    //    UTF-16 — decoding it as utf-8 would keep the interleaved NULs. Genuine utf-8 text has none.
    //    (The router already skips NUL-heavy text upstream, so for now this only backstops a
    //    direct decodeText caller, but it keeps resolveCharset correct without relying on that.)
    if (isUtf8(content) && !content.subarray(0, SNIFF_BYTES).includes(0)) return 'utf-8'
    // 3. Explicit charset from the Content-Type, if iconv-lite knows it.
    if (hint && iconv.encodingExists(hint)) return hint
    // 4. Statistical detection on the head, gated on confidence.
    const detected = jschardet.detect(content.subarray(0, DETECT_SAMPLE_BYTES))
    if (
        detected &&
        detected.encoding &&
        detected.confidence > DETECT_MIN_CONFIDENCE &&
        iconv.encodingExists(detected.encoding)
    ) {
        return detected.encoding
    }
    // 5. Floor — no signal was trusted. This isn't valid utf-8 (step 2 ruled that out), so it's
    //    single-byte legacy text, and utf-8 would turn every high byte into U+FFFD — irreversible
    //    loss. So degrade gracefully instead: take jschardet's guess even below the confidence gate
    //    (a plausible decode beats guaranteed U+FFFD), and failing that latin1, which maps every
    //    byte to a character — possibly wrong, but reversible.
    if (detected?.encoding && iconv.encodingExists(detected.encoding)) return detected.encoding
    return 'latin1'
}

// Decode the raw bytes to text using the correct charset.
const decodeText = (content: Buffer, hint?: string): string => {
    const charset = resolveCharset(content, hint) // pick the charset (see resolveCharset above)
    const text = iconv
        .decode(content, charset) // convert the raw bytes into a JS string using that encoding
        .replace(/^﻿/, '') // strip BOM (cleans the final output)
        .replace(/\r\n?/g, '\n') // normalizes the line endings
    // U+FFFD means the charset was wrong — the decoder hit bytes it couldn't map (jschardet can
    // confidently confuse e.g. big5 for GB2312). Those replacement chars are unrecoverable, so fall
    // back to latin1, which maps every byte to a character (reversible, never destroys data). Genuine
    // utf-8 text can legitimately contain U+FFFD, so leave that case — identified by isUtf8 — alone.
    // A BOM makes the charset DEFINITIVE (e.g. a UTF-16 file that legitimately contains U+FFFD): trust
    // it, never re-decode as latin1 — that would keep the interleaved NULs from the two-byte units.
    if (text.includes('�') && !isUtf8(content) && !bomCharset(content)) {
        return iconv.decode(content, 'latin1').replace(/\r\n?/g, '\n')
    }
    return text
}

/////////////////////////////////////////////////////////////
// HANDLER REGISTRY

// Text handler - txt, csv, calendar, vcard, json, xml, yaml
const textHandler: Handler = {
    kind: 'text',
    contentTypes: [
        'text/plain',
        'text/csv',
        'text/tab-separated-values',
        'text/markdown',
        'text/calendar',
        'application/ics',
        'text/vcard',
        'text/x-vcard',
        'text/enriched',
        // Header-only / report MIME types (DSNs, ARF spam reports) — text, not a full email.
        // NB: message/global is NOT here — it's a full internationalized email, which is out of
        // scope in this version (no eml handler), so it routes to nothing and lands on a skip.
        'message/global-headers',
        'message/delivery-status',
        'message/feedback-report',
        'text/rfc822-headers', 
        // Common structured-but-plain-text payloads.
        'application/json',
        'application/xml',
        'application/yaml',
        'application/x-yaml',
        'text/yaml',
    ],
    extensions: [
        '.txt',
        '.log',
        '.csv',
        '.tsv',
        '.md',
        '.markdown',
        '.ics',
        '.vcf',
        '.json',
        '.xml',
        '.yaml',
        '.yml',
    ],
    extract: async ({ content, charsetHint }) => ({ text: decodeText(content, charsetHint) }),
}

// NOTE: Lazy loading so that a Lambda that only ever sees text attachments never pays to load pdf.js / mammoth (heavier dependencies). 

// Helper function: Shared HTML -> visible-text flattening, 
// used by the html handler to turn an HTML document into visible text.
const flattenHtml = async (html: string): Promise<string> => {
    const { convert } = await import('html-to-text')
    return convert(html, {
        wordwrap: false,
        selectors: [
            { selector: 'img', format: 'skip' },
            { selector: 'a', options: { ignoreHref: true } },
            { selector: 'script', format: 'skip' },
            { selector: 'style', format: 'skip' },
        ],
    })
}

// HTML handler - html
const htmlHandler: Handler = {
    kind: 'html',
    contentTypes: ['text/html', 'application/xhtml+xml'],
    extensions: ['.html', '.htm', '.xhtml'],
    extract: async ({ content, charsetHint }) => {
        const decoded = decodeText(content, charsetHint)
        return { text: await flattenHtml(decoded) }
    },
}

// PDF handler - pdf
const pdfHandler: Handler = {
    kind: 'pdf',
    contentTypes: ['application/pdf', 'application/x-pdf', 'application/acrobat', 'application/vnd.pdf'],
    extensions: ['.pdf'],
    extract: async ({ content }) => {
        const { getDocumentProxy } = await import('unpdf') // Unpdf is a library that extracts text from PDF
        const pdf = await getDocumentProxy(new Uint8Array(content)) // Converts the raw bytes into format unpdf expects
        // Iterate pages ourselves (rather than unpdf's extractText, which parses EVERY page up front)
        // so a pathological page count or a huge text layer can't run unbounded: cap the page count
        // and stop once accumulated text passes MAX_OUTPUT_CHARS. NB: this bounds OUR accumulation and
        // the pages parsed — it does NOT bound pdf.js's internal per-page decompression (no hook
        // exists); that residual is closed only at the Lambda memory limit.
        const pageCount = Math.min(pdf.numPages, MAX_PDF_PAGES)
        const pages: string[] = []
        let length = 0
        for (let n = 1; n <= pageCount; n++) {
            const page = await pdf.getPage(n)
            const { items } = await page.getTextContent()
            // Replicate unpdf's per-page join (str + a newline on hasEOL), then trim the page — keeps
            // output byte-identical to the previous extractText path for the common in-cap case.
            const pageText = (items as Array<{ str?: string; hasEOL?: boolean }>)
                .filter((item) => item.str != null)
                .map((item) => (item.str ?? '') + (item.hasEOL ? '\n' : ''))
                .join('')
                .trim()
            if (pageText) {
                pages.push(pageText)
                length += pageText.length + 2 // + the '\n\n' page join
                if (length > MAX_OUTPUT_CHARS) break // one page of overshoot, trimmed centrally
            }
        }
        const joined = pages.join('\n\n').trim() // trim each page, join with a blank line
        return { text: joined, empty: joined.length === 0 }
    },
}

// DOCX handler - docx
const docxHandler: Handler = {
    kind: 'docx',
    contentTypes: ['application/vnd.openxmlformats-officedocument.wordprocessingml.document'],
    extensions: ['.docx'],
    extract: async ({ content }) => {
        const { default: mammoth } = await import('mammoth') // Mammoth is a library that extracts text from DOCX files
        const { value } = await mammoth.extractRawText({ buffer: content }) // extract the text from the DOCX file
        return { text: value }
    },
}

// DOC handler — legacy Word (.doc). mammoth only reads the modern .docx zip, so the old
// OLE binary needs word-extractor instead.
const docHandler: Handler = {
    kind: 'doc',
    contentTypes: ['application/msword'],
    extensions: ['.doc'],
    extract: async ({ content }) => {
        const { default: WordExtractor } = await import('word-extractor')
        const doc = await new WordExtractor().extract(content)
        return { text: doc.getBody() } // main body only; headers/footers/notes are separate streams
    },
}

// XLSX handler — modern Excel (.xlsx). Flatten each sheet to text for search/indexing.
const xlsxHandler: Handler = {
    kind: 'xlsx',
    contentTypes: ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
    extensions: ['.xlsx'],
    extract: async ({ content }) => {
        const { default: ExcelJS } = await import('exceljs') // not SheetJS: no known parse-time CVEs on untrusted bytes
        const workbook = new ExcelJS.Workbook()
        // Cast: the value is a real Buffer; exceljs's load() type clashes with @types/node's generic Buffer.
        await workbook.xlsx.load(content as unknown as Parameters<typeof workbook.xlsx.load>[0])
        const sheets: string[] = []
        // Stop flattening once we've built MAX_OUTPUT_CHARS of text: a dense workbook is already in
        // memory after load(), so building the FULL flattened string (which the central cap would
        // then trim) doubles peak memory for nothing. exceljs's eachRow/eachSheet can't break, so we
        // gate on a running length and skip further work once over. The central cap does the final
        // surrogate-safe slice; this just keeps the intermediate string near the cap, not the doc.
        let length = 0
        let capped = false
        workbook.eachSheet((sheet) => {
            if (capped) return
            const rows: string[] = []
            // Drop empty cells/rows so a sparse sheet doesn't flatten into runs of empty tabs.
            sheet.eachRow({ includeEmpty: false }, (row) => {
                if (capped) return
                const cells: string[] = []
                // cell.text = the shown value (formula result, formatted date), not the raw number/formula.
                row.eachCell({ includeEmpty: false }, (cell) => cells.push(cell.text ?? ''))
                if (cells.length > 0) {
                    const line = cells.join('\t')
                    rows.push(line)
                    length += line.length + 1 // + newline
                    if (length > MAX_OUTPUT_CHARS) capped = true // one line of overshoot, trimmed centrally
                }
            })
            // Header + rows per sheet; a text-less sheet contributes nothing → extracted_empty.
            if (rows.length > 0) sheets.push(`=== ${sheet.name} ===\n${rows.join('\n')}`)
        })
        return { text: sheets.join('\n\n') }
    },
}

// NB: no EML handler in this version — nested emails (.eml / message/rfc822 / message/global) are
// out of scope and route to nothing, so they land on a labeled skip.

const REGISTRY: Handler[] = [textHandler, htmlHandler, pdfHandler, docxHandler, docHandler, xlsxHandler]

const findHandler = (kind: HandlerKind): Handler | undefined => REGISTRY.find((h) => h.kind === kind) // Look up the handler in that array whose kind matches the one requested.

/////////////////////////////////////////////////////////////
// ROUTING — look at ALL the available clues about an attachment, and decide which HandlerKind (if any) it should be treated as.
// You can't just trust one signal - you need to look at all the available clues.

// Split "type/subtype; charset=..." into a normalized type + the charset param.
const parseContentType = (raw?: string): { type?: string; charset?: string } => {
    if (!raw) return {}
    const [head, ...params] = raw.split(';')
    const type = head.trim().toLowerCase() || undefined
    let charset: string | undefined
    for (const param of params) {
        const match = /^\s*charset\s*=\s*"?([^";]+)"?\s*$/i.exec(param)
        if (match) charset = match[1].trim()
    }
    return { type, charset }
}

// If there is a filename, pull out the clean file extension 
const extensionOf = (filename?: string): string | undefined => {
    if (!filename) return undefined
    const match = /(\.[a-z0-9]+)$/i.exec(filename.trim())
    return match ? match[1].toLowerCase() : undefined
}

// Look up the handler in the registry by content type.
const findByContentType = (type: string): Handler | undefined =>
    REGISTRY.find((h) => h.contentTypes.includes(type)) ??
    // RTF is text/* but is really control-word markup and we have no handler for it. Skip it here,
    // before the text fallback below would decode those control words as body text.
    (type.includes('rtf') ? undefined :
    // Any other text/* we didn't enumerate is plain text — except html-ish subtypes
    // (e.g. text/x-amp-html), which are markup and must go through the html handler.
    type.startsWith('text/') ? (type.includes('html') ? htmlHandler : textHandler) : undefined)

// Given a file extension (like .pdf), find the handler in the registry whose extensions list includes it (just a straight lookup). 
const findByExtension = (ext: string): Handler | undefined => REGISTRY.find((h) => h.extensions.includes(ext))

// Generic "unknown binary" MIME labels — not real type info, safe to ignore.
// Set = fast, order-doesn't-matter membership checks.
const OCTET_STREAM_TYPES = new Set([
    'application/octet-stream',
    'binary/octet-stream',
    'application/download',
    'application/unknown',
])
// Sniff bytes only when we have no real type info — not for real types we
// just don't support (e.g. image/png stays a deliberate skip). 
const shouldSniff = (type?: string): boolean => !type || OCTET_STREAM_TYPES.has(type)

// Guess whether raw bytes are plain text by sampling the head of the file (first 8KB). 
// A single NUL byte -> treat as binary immediately.
// Otherwise, count printable bytes; call it text if 85%+ of the sample is printable.
const looksLikeText = (content: Buffer): boolean => {
    const sample = content.subarray(0, SNIFF_BYTES)
    if (sample.length === 0) return false
    let textish = 0
    for (const byte of sample) {
        if (byte === 0) return false // NUL -> binary
        const printable = byte === 9 || byte === 10 || byte === 13 || (byte >= 32 && byte !== 127)
        if (printable) textish++
    }
    return textish / sample.length >= SNIFF_TEXT_RATIO
}

// PDFs always start with the literal bytes "%PDF-"
const PDF_MAGIC = Buffer.from('%PDF-')
// Zip files always start with the byte sequence 0x50 0x4b 0x03 0x04 ("PK..").
// A .docx is a zip, so this lets us recognize one from its content.
const ZIP_MAGIC = Buffer.from([0x50, 0x4b, 0x03, 0x04])
// RTF starts with "{\rtf". It's ASCII, so looksLikeText would wrongly take a mislabeled one as
// text and leak the control words — detect it here to skip instead (see findByContentType).
const RTF_MAGIC = Buffer.from('{\\rtf')
// Legacy OLE files (.doc) start with this 8-byte signature. Shared by .xls/.ppt/.msg too, so —
// like OOXML — we only claim doc when the extension confirms it.
const OLE_MAGIC = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])

// docx and xlsx share the zip magic, but their main part has a distinct path stored (uncompressed)
// as a literal filename in the zip: word/document.xml vs xl/workbook.xml. Scanning for it tells the
// two apart from content alone — no extension needed.
const DOCX_PART = Buffer.from('word/document.xml')
const XLSX_PART = Buffer.from('xl/workbook.xml')

// Check if the file starts with the given magic bytes.
const startsWith = (content: Buffer, magic: Buffer): boolean =>
    content.length >= magic.length && content.subarray(0, magic.length).equals(magic)

// Which OOXML kind a zip is, by its main root part (word/document.xml vs xl/workbook.xml). A docx can
// embed a workbook, so identify the package's OWN parts by exact zip entry name — not by whichever
// marker appears first in the raw bytes, which an embedded object can reorder. Returns undefined for a
// non-OOXML zip (jar, plain archive, pptx).
const ooxmlKind = (content: Buffer): HandlerKind | undefined => {
    // Prefer exact zip ENTRY names from the central directory: an embedded object (an xlsx inside a
    // docx) contributes its xl/workbook.xml as bytes within another entry, NOT as an entry of this
    // package, so a root-entry match is not fooled by storage order the way a raw byte scan is.
    const names = zipEntryNames(content)
    if (names) {
        const hasDocx = names.includes('word/document.xml')
        const hasXlsx = names.includes('xl/workbook.xml')
        if (hasDocx) return 'docx' // a real word/document.xml root part wins (docx may embed a workbook)
        if (hasXlsx) return 'xlsx'
        return undefined // OOXML zip with neither root part (pptx, jar, plain archive)
    }
    // Fallback (archive not walkable): raw-bytes scan, earlier main-part marker wins.
    const d = content.indexOf(DOCX_PART)
    const x = content.indexOf(XLSX_PART)
    if (d === -1 && x === -1) return undefined
    if (x === -1) return 'docx'
    if (d === -1) return 'xlsx'
    return d < x ? 'docx' : 'xlsx'
}

/////////////////////////////////////////////////////////////
// DECOMPRESSION BUDGET (OOXML zip-bomb guard)
//
// A small in-cap .docx/.xlsx can inflate to hundreds of MB and OOM the worker. The zip central
// directory carries a self-declared uncompressed size, but that field is attacker-controlled — a
// crafted archive can declare "tiny" and still expand far past the cap. So we don't trust it: we
// STREAM-inflate each entry and count the REAL output bytes, aborting the instant the archive-wide
// total crosses the cap. Streaming (not inflateRawSync) keeps peak memory at ~one zlib chunk, never
// the full expansion. Fail-closed: anything we can't measure — malformed metadata, ZIP64, a corrupt
// stream, an unsupported compression method — is a skip, never waved through to the parser.

const EOCD_MAGIC = Buffer.from([0x50, 0x4b, 0x05, 0x06]) // End-of-Central-Directory signature bytes
const CD_SIG = 0x02014b50 // Central-Directory file header
const LOCAL_SIG = 0x04034b50 // Local file header

// The pre-parse decompression verdict. `ok` = safe to hand to the parser; otherwise `reason` is the
// skip message (readable, never "Infinity bytes").
type DecompressionCheck = { ok: true } | { ok: false; reason: string }

// Inflate one raw-deflate region, adding its output onto `runningTotal`, aborting the moment the
// total would exceed `cap`. Returns the new total, or a sentinel: -1 = over budget (destroyed
// early, full expansion never materialized), -2 = corrupt/unreadable stream. Byte counts are >= 0,
// so the negatives are unambiguous.
const inflateCounting = (comp: Buffer, runningTotal: number, cap: number): Promise<number> =>
    new Promise((resolve) => {
        const inflate = zlib.createInflateRaw()
        let total = runningTotal
        let settled = false
        const settle = (value: number) => {
            if (settled) return
            settled = true
            resolve(value)
        }
        inflate.on('data', (chunk: Buffer) => {
            total += chunk.length
            if (total > cap) {
                inflate.destroy() // stop inflating — we never hold the full expansion
                settle(-1)
            }
        })
        inflate.on('end', () => settle(total))
        inflate.on('error', () => settle(-2)) // truncated / encrypted / garbage deflate stream
        inflate.end(comp)
    })

// Find the EOCD the SAME way the parser's zip reader (jszip, used by exceljs/mammoth) does: the LAST
// occurrence of the signature in the whole buffer, no comment-length check. Matching the parser's
// selection is the point — the preflight must measure the exact central directory the parser will
// read. A comment-to-EOF invariant would DIVERGE: a second EOCD planted after the real one is picked
// by jszip's last-match scan but rejected by the invariant, so the parser could follow it to a bomb CD
// the preflight never measured. Measuring (not trusting declared sizes) makes following whichever EOCD
// jszip picks safe. As a bonus, last-match doesn't false-skip zips with trailing bytes after the EOCD.
const findEocd = (buf: Buffer): number => {
    const eocd = buf.lastIndexOf(EOCD_MAGIC)
    return eocd >= 0 && eocd + 22 <= buf.length ? eocd : -1 // need room for the 22-byte fixed record
}

// Read every entry's stored name from the central directory. Unlike scanning raw bytes, this sees
// only the PACKAGE's own parts: an embedded object's internal paths (e.g. an xlsx embedded inside a
// docx) live in another entry's data, not as entries here, so they can't fool root detection.
// Returns undefined when the central directory can't be walked (caller falls back to a raw scan).
const zipEntryNames = (buf: Buffer): string[] | undefined => {
    const eocd = findEocd(buf)
    if (eocd < 0) return undefined
    const entries = buf.readUInt16LE(eocd + 10)
    if (entries === 0xffff) return undefined // ZIP64 entry count — not chased here
    const names: string[] = []
    let p = buf.readUInt32LE(eocd + 16)
    for (let i = 0; i < entries; i++) {
        if (p + 46 > buf.length || buf.readUInt32LE(p) !== CD_SIG) return undefined
        const nameLen = buf.readUInt16LE(p + 28)
        if (p + 46 + nameLen > buf.length) return undefined
        names.push(buf.toString('latin1', p + 46, p + 46 + nameLen))
        p += 46 + nameLen + buf.readUInt16LE(p + 30) + buf.readUInt16LE(p + 32)
    }
    return names
}

// Measure a zip's ACTUAL decompressed size, capped. Walks the central directory to each entry's real
// (structural, not self-declared) compressed region, inflates it, and counts true output bytes.
const checkDecompressionBudget = async (buf: Buffer, cap: number): Promise<DecompressionCheck> => {
    const eocd = findEocd(buf)
    if (eocd < 0) return { ok: false, reason: 'malformed zip: no end-of-central-directory record' }

    const entries = buf.readUInt16LE(eocd + 10)
    const cdOffset = buf.readUInt32LE(eocd + 16)
    // ZIP64 / out-of-range sentinels: the true values live in a ZIP64 record we deliberately don't
    // chase. Treat as over-budget rather than trust the classic field or crash on the sentinel.
    if (entries === 0xffff || buf.readUInt32LE(eocd + 12) === 0xffffffff || cdOffset === 0xffffffff)
        return { ok: false, reason: 'zip declares a ZIP64 / out-of-range size' }

    let total = 0
    let p = cdOffset
    for (let i = 0; i < entries; i++) {
        // A missing/misaligned central-directory header means the offset was a lie (a partial walk
        // must not silently return the accumulated total). Fail closed.
        if (p + 46 > buf.length || buf.readUInt32LE(p) !== CD_SIG)
            return { ok: false, reason: 'malformed zip: truncated or misaligned central directory' }
        const method = buf.readUInt16LE(p + 10)
        const compSize = buf.readUInt32LE(p + 20)
        const localOffset = buf.readUInt32LE(p + 42)
        if (compSize === 0xffffffff || localOffset === 0xffffffff)
            return { ok: false, reason: 'zip declares a ZIP64 / out-of-range size' }
        // The local header's name/extra lengths can differ from the central copy, so read them here
        // to find where this entry's compressed bytes actually begin.
        if (localOffset + 30 > buf.length || buf.readUInt32LE(localOffset) !== LOCAL_SIG)
            return { ok: false, reason: 'malformed zip: bad local header offset' }
        const dataStart = localOffset + 30 + buf.readUInt16LE(localOffset + 26) + buf.readUInt16LE(localOffset + 28)
        const comp = buf.subarray(dataStart, dataStart + compSize)
        if (comp.length < compSize)
            return { ok: false, reason: 'malformed zip: compressed data runs past end of file' }

        if (method === 0) {
            total += comp.length // stored (no compression): output === input
        } else if (method === 8) {
            total = await inflateCounting(comp, total, cap)
            if (total === -1) return { ok: false, reason: `decompresses to over ${cap} bytes` }
            if (total === -2) return { ok: false, reason: 'malformed zip: unreadable compressed data' }
        } else {
            return { ok: false, reason: `zip uses unsupported compression method ${method}` }
        }
        if (total > cap) return { ok: false, reason: `decompresses to over ${cap} bytes` }
        p += 46 + buf.readUInt16LE(p + 28) + buf.readUInt16LE(p + 30) + buf.readUInt16LE(p + 32)
    }
    return { ok: true }
}

// Does the content match a confident binary claim? Catches a wrong claim before we hand the bytes
// to the parser. pdf/doc are a magic-byte prefix; docx/xlsx go by their OOXML main part (zip magic
// alone can't tell them apart). text/html have no single signature, so they always pass here —
// lying text claims are handled separately by bytesContradictTextClaim.
const magicOk = (kind: HandlerKind, content: Buffer): boolean => {
    if (kind === 'pdf') return startsWith(content, PDF_MAGIC)
    if (kind === 'doc') return startsWith(content, OLE_MAGIC)
    if (kind === 'docx') return ooxmlKind(content) === 'docx'
    if (kind === 'xlsx') return ooxmlKind(content) === 'xlsx'
    return true
}

// Magic bytes are the strongest sniff signal, so they go first.
// A BOM is next: UTF-16 text is full of NUL bytes, so looksLikeText would reject it as binary.
// A zip could be docx/xlsx/pptx/jar; ooxmlKind reads the content to pick docx vs xlsx (and skips
// the rest), so we no longer need the extension to disambiguate.
const sniff = (content: Buffer, ext?: string): HandlerKind | undefined => {
    if (startsWith(content, PDF_MAGIC)) return 'pdf'
    if (startsWith(content, ZIP_MAGIC)) return ooxmlKind(content)
    if (startsWith(content, OLE_MAGIC) && ext === '.doc') return 'doc'
    // RTF is ASCII text but control-word markup, and we have no handler for it. Skip before the
    // looksLikeText check that would misclaim it as plain text and leak the control words.
    if (startsWith(content, RTF_MAGIC)) return undefined
    if (bomCharset(content) || looksLikeText(content)) return 'text'
    return undefined
}

// A text/html claim is only as good as its bytes: a PDF shipped as text/plain would
// latin1-decode into garbage and be reported as 'extracted' — a silent quality failure,
// worse than a labeled skip. Provably-binary bytes (no BOM, NULs / non-printable head)
// void the claim; the sniff then rescues what the magic bytes prove (%PDF, zip+.docx)
// and anything else is left unrouted. Empty content stays with the claimed handler so a
// zero-byte text attachment still lands on extracted_empty rather than a skip.
const trustsExplicitUnicodeCharset = (charset?: string): boolean => {
    const normalized = charset?.trim().toLowerCase().replace(/_/g, '-')
    return normalized === 'utf-16' || normalized === 'utf-16le' || normalized === 'utf-16be'
}

// Known binary-container magic — %PDF, zip (docx/xlsx), OLE (legacy .doc). A genuine text file never
// starts with these, so their presence contradicts a text claim even under an explicit charset hint.
const hasKnownBinaryMagic = (content: Buffer): boolean =>
    startsWith(content, PDF_MAGIC) || startsWith(content, ZIP_MAGIC) || startsWith(content, OLE_MAGIC)

const bytesContradictTextClaim = (content: Buffer, charsetHint?: string): boolean => {
    if (content.length === 0 || bomCharset(content)) return false
    // Binary magic beats even a charset=utf-16 hint: a real utf-16 text file never begins with %PDF /
    // PK\x03\x04 / the OLE signature, but a binary mislabeled text/…;charset=utf-16 would otherwise be
    // trusted because the utf-16 exemption below suppresses the printable-ratio check → gibberish
    // reported as 'extracted'. Re-sniffing then recovers the real format (%PDF → pdf, zip → docx/xlsx).
    if (hasKnownBinaryMagic(content)) return true
    // Otherwise an explicit utf-16 hint is trusted (utf-16 text is NUL-heavy and fails looksLikeText);
    // any other claim must actually look like text.
    return !trustsExplicitUnicodeCharset(charsetHint) && !looksLikeText(content)
}

// Takes an AttachmentInput and produces the final routing decision. 
// Tries each signal in priority order and stops as soon as one works.
export const detectRoute = (input: AttachmentInput): { kind?: HandlerKind; routedBy: RoutedBy } => {
    const { type, charset: charsetHint } = parseContentType(input.contentType)
    const ext = extensionOf(input.filename)

    const byType = type ? findByContentType(type) : undefined
    const byExt = ext ? findByExtension(ext) : undefined
    const claimed = byType
        ? ({ kind: byType.kind, routedBy: 'content-type' } as const)
        : byExt
          ? ({ kind: byExt.kind, routedBy: 'extension' } as const)
          : undefined

    if (claimed) {
        // Distrust a text-decoding claim contradicted by the bytes (the type/extension lied), OR one
        // whose bytes are RTF — RTF is printable ASCII so it slips past bytesContradictTextClaim, but
        // it's control-word markup with no handler and must not decode as body text. Re-sniffing hits
        // the RTF_MAGIC skip in sniff() and lands on a labeled skip instead of leaking control words.
        if (
            (claimed.kind === 'text' || claimed.kind === 'html') &&
            (bytesContradictTextClaim(input.content, charsetHint) || startsWith(input.content, RTF_MAGIC))
        ) {
            const sniffed = sniff(input.content, ext)
            return sniffed ? { kind: sniffed, routedBy: 'sniff' } : { routedBy: 'none' }
        }
        // Distrust a binary claim the content contradicts — a PDF mislabeled .docx, or a
        // spreadsheet mislabeled .docx (caught by its OOXML part, not just the shared zip magic).
        // Handing wrong bytes to the parser just fails and loses any real text, so re-sniff to find
        // the true format. Empty content has nothing to check — leave it to the parser.
        if (input.content.length > 0 && !magicOk(claimed.kind, input.content)) {
            const sniffed = sniff(input.content, ext)
            return sniffed ? { kind: sniffed, routedBy: 'sniff' } : { routedBy: 'none' }
        }
        return claimed
    }

    if (shouldSniff(type)) {
        const sniffed = sniff(input.content, ext)
        if (sniffed) return { kind: sniffed, routedBy: 'sniff' }
    }

    return { routedBy: 'none' }
}

/////////////////////////////////////////////////////////////
// SAFETY

// Error class for handler timeouts.
class HandlerTimeoutError extends Error {
    constructor(ms: number) {
        super(`handler exceeded ${ms}ms`)
        this.name = 'HandlerTimeoutError'
    }
}

// Enforce a time limit so a slow/hung handler doesn't block forever.
const withTimeout = <T>(promise: Promise<T>, ms: number): Promise<T> =>
    new Promise((resolve, reject) => {
        const timer = setTimeout(() => reject(new HandlerTimeoutError(ms)), ms)
        promise.then(
            (value) => {
                clearTimeout(timer)
                resolve(value)
            },
            (error) => {
                clearTimeout(timer)
                reject(error)
            }
        )
    })

// Safely turn any thrown thing into a readable string, since JS lets you throw non-Errors.
const errorMessage = (error: unknown): string => (error instanceof Error ? error.message : String(error))

/////////////////////////////////////////////////////////////
// ENTRY POINT 
// Calls everything, in the right order, with the right safety nets around each risky step, 
// and produces a complete, labeled ExtractionResult.

// Main entry point for extracting an attachment. Returns a labeled ExtractionResult:
// status, the extracted text (omitted when there is none — never ''), and a reason on skip/fail.
export const extractAttachment = async (input: AttachmentInput): Promise<ExtractionResult> => {
    const byteSize = input.content.length

    // Empty content has nothing to extract — resolve it here so the status is consistent regardless
    // of the declared type (otherwise an empty PDF/OOXML routes into a parser that throws on zero
    // bytes → 'failed', while empty text → 'extracted'). Treat all empties as ran-but-empty.
    if (byteSize === 0) {
        return { status: 'extracted' }
    }

    // Size gate BEFORE any decode/parse. If the attachment is too large, skip it.
    if (byteSize > MAX_INPUT_BYTES) {
        return { status: 'skipped', reason: `${byteSize} bytes exceeds ${MAX_INPUT_BYTES}` }
    }

    const { type, charset: charsetHint } = parseContentType(input.contentType) // declared type + charset, parsed once
    const { kind } = detectRoute(input) // figure out what kind of file this is
    const handler = kind ? findHandler(kind) : undefined // look up the actual handler for that kind, if we found one

    // No handler found for this kind — unsupported or unrecognized format, skip it.
    if (!handler) {
        return { status: 'skipped', reason: type ? `unsupported type ${type}` : 'unrecognized attachment' }
    }

    // Decompression preflight: an OOXML (zip) attachment can inflate far beyond its on-disk size and
    // OOM the worker, so MEASURE its actual decompressed size (streaming-inflate, capped) and skip
    // before exceljs/mammoth ever touch it. Fail-closed — malformed/ZIP64/corrupt metadata is a skip,
    // never a pass, since the zip's self-declared sizes are attacker-controlled. (pdf/doc/text aren't zips.)
    if (kind === 'docx' || kind === 'xlsx') {
        const check = await checkDecompressionBudget(input.content, MAX_UNCOMPRESSED_BYTES)
        if (!check.ok) return { status: 'skipped', reason: check.reason }
    }

    // Run the handler safely — catch errors/timeouts instead of crashing, and build the final result.
    try {
        const output = await withTimeout( // Enforce a time limit so a slow/hung handler doesn't block forever.
            handler.extract({ content: input.content, filename: input.filename, charsetHint }),
            HANDLER_TIMEOUT_MS
        )
        const isEmpty = output.empty ?? output.text.trim().length === 0
        // Cap output centrally so a pathological document can't dump megabytes into S3 + the search
        // index. NB: the xlsx and pdf handlers already cap INCREMENTALLY as they build, so this is
        // just the final precise trim for them; the docx (mammoth) and html (html-to-text) handlers
        // return a full string, so for those this is a POST-MATERIALIZATION cap (peak memory follows
        // the whole document — hard containment is the host memory limit, see README).
        // Don't split a surrogate pair at the boundary (a lone half serializes as U+FFFD).
        const overCap = output.text.length > MAX_OUTPUT_CHARS
        const capEnd =
            overCap && output.text.charCodeAt(MAX_OUTPUT_CHARS - 1) >= 0xd800 && output.text.charCodeAt(MAX_OUTPUT_CHARS - 1) <= 0xdbff
                ? MAX_OUTPUT_CHARS - 1
                : MAX_OUTPUT_CHARS
        const text = overCap ? output.text.slice(0, capEnd) : output.text
        // The extracted text — omitted entirely (never '') when the handler produced none. A
        // consumer reads a present `extraction` as "has text", its absence as "ran, but empty".
        return isEmpty ? { status: 'extracted' } : { status: 'extracted', extraction: text }
    } catch (error) {
        // Attacker-controlled bytes: a throw or timeout is an expected event, captured as
        // 'failed' with the reason — never propagated to crash the caller.
        return { status: 'failed', reason: errorMessage(error) }
    }
}