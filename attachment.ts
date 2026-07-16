// attachment — text extraction for email attachments: text, HTML, PDF, Word (.docx + .doc), Excel
// (.xlsx). Everything else — nested emails, images/OCR, .xls/.ppt, archives — returns a labeled skip.
//
// Peer module to body.ts; neither depends on the other. Its own subpath export
// ("agentextract/attachment") keeps the heavy parsers out of the body extractor's bundle.

import { isUtf8 } from 'node:buffer' // native check: are these bytes valid utf-8?
import zlib from 'node:zlib' // streaming raw-inflate, for the decompression budget

import iconv from 'iconv-lite' // bytes -> text, in a given encoding
import jschardet from 'jschardet' // guesses which encoding bytes are in

/////////////////////////////////////////////////////////////
// CONSTANTS (tunable)

// Rejected before any decode or parse.
export const MAX_INPUT_BYTES = 10 * 1024 * 1024

// Input is byte-capped; output isn't. A big sheet or HTML table can balloon into megabytes that then
// hit S3 and the search index. Truncated silently past this — the result carries no truncation flag.
export const MAX_OUTPUT_CHARS = 250_000

// A PDF can declare an enormous page count. Bounds parse work when per-page text is too sparse to
// trip the output cap; a content-bearing PDF hits MAX_OUTPUT_CHARS within a few dozen pages first.
export const MAX_PDF_PAGES = 2000

// Max ACTUAL uncompressed size of an OOXML zip, measured by inflating it — the declared size is
// attacker-controlled (see DECOMPRESSION BUDGET). MAX_INPUT_BYTES only bounds the COMPRESSED size.
// Parsers build a model on top (~8x for a dense sheet), so stay well under the 1024 MB memory floor.
export const MAX_UNCOMPRESSED_BYTES = 50 * 1024 * 1024

// Guard: stop awaiting a slow handler. Can't cancel synchronous CPU already running inside a parser.
export const HANDLER_TIMEOUT_MS = 10_000

// Sniff the first 8KB to decide whether bytes look like text.
const SNIFF_BYTES = 8 * 1024
const SNIFF_TEXT_RATIO = 0.85

// Charset detection samples the head; below this confidence the guess isn't trusted.
const DETECT_SAMPLE_BYTES = 64 * 1024
const DETECT_MIN_CONFIDENCE = 0.7

/////////////////////////////////////////////////////////////
// TYPES

export type ExtractionStatus =
    | 'extracted' // handler ran; `extraction` holds the text, or is omitted when there was none
    | 'skipped' // over MAX_INPUT_BYTES, or unsupported/unrecognized — never attempted
    | 'failed' // parser threw, timed out, or decode failed

// Which signal decided the route.
export type RoutedBy = 'content-type' | 'extension' | 'sniff' | 'none'

// Formats we route to; each maps to exactly one handler. text covers txt/csv/ics/vcard/json/xml/yaml.
export type HandlerKind = 'text' | 'html' | 'pdf' | 'docx' | 'doc' | 'xlsx'

export interface AttachmentInput {
    content: Buffer
    filename?: string
    contentType?: string
}

export interface ExtractionResult {
    status: ExtractionStatus
    extraction?: string // omitted entirely (never '') when the handler produced no text
    reason?: string // set on skipped / failed
}

interface HandlerContext {
    content: Buffer
    filename?: string
    charsetHint?: string // from the content-type charset= param
}

interface HandlerOutput {
    text: string
    empty?: boolean // handler's own emptiness call; defaults to text.trim() === ''
}

interface Handler {
    kind: HandlerKind
    contentTypes: string[] // exact, lowercased, param-stripped
    extensions: string[] // with leading dot, lowercased
    extract: (ctx: HandlerContext) => Promise<HandlerOutput>
}

/////////////////////////////////////////////////////////////
// CHARSET-CORRECT DECODING (direct-text handlers)

// A byte-order mark is the file tagging its own encoding — unambiguous when present.
const bomCharset = (content: Buffer): string | undefined => {
    if (content.length >= 3 && content[0] === 0xef && content[1] === 0xbb && content[2] === 0xbf) return 'utf-8'
    if (content.length >= 2 && content[0] === 0xff && content[1] === 0xfe) return 'utf-16le'
    if (content.length >= 2 && content[0] === 0xfe && content[1] === 0xff) return 'utf-16be'
    return undefined
}

// Priority chain: most trustworthy signal first, degrading gracefully.
const resolveCharset = (content: Buffer, hint?: string): string => {
    // #1. A BOM is definitive. Above the hint: a stale charset= must not override an in-band BOM.
    const bom = bomCharset(content)
    if (bom) return bom
    // #2. Valid utf-8 is self-evidencing — multi-byte sequences are near-impossible by accident, so
    //     trust the bytes over a hint that would mangle them. ASCII decodes the same either way, so
    //     this never loses data.
    //     Guard on NUL: BOM-less utf-16 ASCII (h\0e\0l\0) is technically valid utf-8 but isn't utf-8;
    //     decoding it as such would keep the interleaved NULs. Real utf-8 text has none.
    if (isUtf8(content) && !content.subarray(0, SNIFF_BYTES).includes(0)) return 'utf-8'
    // #3. Explicit charset from the Content-Type, if iconv knows it.
    if (hint && iconv.encodingExists(hint)) return hint
    // #4. Statistical detection on the head, gated on confidence.
    const detected = jschardet.detect(content.subarray(0, DETECT_SAMPLE_BYTES))
    if (
        detected &&
        detected.encoding &&
        detected.confidence > DETECT_MIN_CONFIDENCE &&
        iconv.encodingExists(detected.encoding)
    ) {
        return detected.encoding
    }
    // #5. Floor. Step 2 ruled out utf-8, so these are single-byte legacy bytes and decoding as utf-8
    //     would turn every high byte into an irreversible U+FFFD. Take jschardet's guess even below
    //     the gate (a plausible decode beats guaranteed U+FFFD), else latin1 — maybe wrong, but it
    //     maps every byte, so it's reversible.
    if (detected?.encoding && iconv.encodingExists(detected.encoding)) return detected.encoding
    return 'latin1'
}

const decodeText = (content: Buffer, hint?: string): string => {
    const charset = resolveCharset(content, hint)
    const text = iconv
        .decode(content, charset)
        .replace(/^﻿/, '') // strip BOM
        .replace(/\r\n?/g, '\n') // normalize line endings
    // U+FFFD means the charset was wrong (jschardet can confidently confuse big5 for GB2312) and is
    // unrecoverable, so fall back to latin1, which maps every byte. Two exceptions: real utf-8 may
    // legitimately contain U+FFFD, and a BOM makes the charset definitive — re-decoding either as
    // latin1 would corrupt genuine text.
    if (text.includes('�') && !isUtf8(content) && !bomCharset(content)) {
        return iconv.decode(content, 'latin1').replace(/\r\n?/g, '\n')
    }
    return text
}

/////////////////////////////////////////////////////////////
// HANDLER REGISTRY

// Handlers lazy-load their parsers: a Lambda that only ever sees text never pays to load pdf.js
// or mammoth.

// Text — txt, csv, calendar, vcard, json, xml, yaml
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
        // NB: message/global is deliberately absent — a full email, out of scope, so it skips.
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

// HTML -> visible text.
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

// HTML
const htmlHandler: Handler = {
    kind: 'html',
    contentTypes: ['text/html', 'application/xhtml+xml'],
    extensions: ['.html', '.htm', '.xhtml'],
    extract: async ({ content, charsetHint }) => {
        const decoded = decodeText(content, charsetHint)
        return { text: await flattenHtml(decoded) }
    },
}

// PDF
const pdfHandler: Handler = {
    kind: 'pdf',
    contentTypes: ['application/pdf', 'application/x-pdf', 'application/acrobat', 'application/vnd.pdf'],
    extensions: ['.pdf'],
    extract: async ({ content }) => {
        const { getDocumentProxy } = await import('unpdf')
        const pdf = await getDocumentProxy(new Uint8Array(content))
        // Iterate pages ourselves — unpdf's extractText parses EVERY page up front, so a pathological
        // page count runs unbounded. Bounds our accumulation and the pages parsed, NOT pdf.js's
        // per-page decompression (no hook exists); that residual is the host memory limit's job.
        const pageCount = Math.min(pdf.numPages, MAX_PDF_PAGES)
        const pages: string[] = []
        let length = 0
        for (let n = 1; n <= pageCount; n++) {
            const page = await pdf.getPage(n)
            const { items } = await page.getTextContent()
            // Replicates unpdf's per-page join: str, plus a newline on hasEOL.
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
        const joined = pages.join('\n\n').trim()
        return { text: joined, empty: joined.length === 0 }
    },
}

// DOCX — modern OOXML Word
const docxHandler: Handler = {
    kind: 'docx',
    contentTypes: ['application/vnd.openxmlformats-officedocument.wordprocessingml.document'],
    extensions: ['.docx'],
    extract: async ({ content }) => {
        const { default: mammoth } = await import('mammoth')
        const { value } = await mammoth.extractRawText({ buffer: content })
        return { text: value }
    },
}

// DOC — legacy Word 97–2003. mammoth only reads the modern .docx zip, so the OLE binary needs
// word-extractor instead.
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

// XLSX — modern Excel. Each sheet flattened to text for search/indexing.
const xlsxHandler: Handler = {
    kind: 'xlsx',
    contentTypes: ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
    extensions: ['.xlsx'],
    extract: async ({ content }) => {
        const { default: ExcelJS } = await import('exceljs') // not SheetJS: no known parse-time CVEs
        const workbook = new ExcelJS.Workbook()
        // Cast: the value is a real Buffer; exceljs's load() type clashes with @types/node's.
        await workbook.xlsx.load(content as unknown as Parameters<typeof workbook.xlsx.load>[0])
        const sheets: string[] = []
        // Cap incrementally. The workbook is already in memory after load(), so building the FULL
        // string for the central cap to trim doubles peak memory for nothing. eachRow/eachSheet can't
        // break, so gate on a running length instead.
        let length = 0
        let capped = false
        workbook.eachSheet((sheet) => {
            if (capped) return
            const rows: string[] = []
            // Drop empty cells/rows so a sparse sheet doesn't flatten into runs of empty tabs.
            sheet.eachRow({ includeEmpty: false }, (row) => {
                if (capped) return
                const cells: string[] = []
                // cell.text = the shown value (formula result, formatted date), not the raw formula.
                row.eachCell({ includeEmpty: false }, (cell) => cells.push(cell.text ?? ''))
                if (cells.length > 0) {
                    const line = cells.join('\t')
                    rows.push(line)
                    length += line.length + 1 // + newline
                    if (length > MAX_OUTPUT_CHARS) capped = true // one line of overshoot, trimmed centrally
                }
            })
            if (rows.length > 0) sheets.push(`=== ${sheet.name} ===\n${rows.join('\n')}`)
        })
        return { text: sheets.join('\n\n') }
    },
}

const REGISTRY: Handler[] = [textHandler, htmlHandler, pdfHandler, docxHandler, docHandler, xlsxHandler]

const findHandler = (kind: HandlerKind): Handler | undefined => REGISTRY.find((h) => h.kind === kind)

/////////////////////////////////////////////////////////////
// ROUTING — weigh every clue (type, extension, bytes) to pick a handler. No single signal is trusted.

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

// Pull the clean extension off a filename, if there is one.
const extensionOf = (filename?: string): string | undefined => {
    if (!filename) return undefined
    const match = /(\.[a-z0-9]+)$/i.exec(filename.trim())
    return match ? match[1].toLowerCase() : undefined
}

const findByContentType = (type: string): Handler | undefined =>
    REGISTRY.find((h) => h.contentTypes.includes(type)) ??
    // RTF is text/* but really control-word markup, with no handler. Skip before the text fallback
    // below decodes those control words as body text.
    (type.includes('rtf') ? undefined :
    // Any other text/* is plain text — except html-ish subtypes (e.g. text/x-amp-html), which are
    // markup and belong to the html handler.
    type.startsWith('text/') ? (type.includes('html') ? htmlHandler : textHandler) : undefined)

const findByExtension = (ext: string): Handler | undefined => REGISTRY.find((h) => h.extensions.includes(ext))

// Generic "unknown binary" labels — not real type info, safe to ignore.
const OCTET_STREAM_TYPES = new Set([
    'application/octet-stream',
    'binary/octet-stream',
    'application/download',
    'application/unknown',
])
// Sniff only when there's no real type info — a type we simply don't support (image/png) stays a
// deliberate skip.
const shouldSniff = (type?: string): boolean => !type || OCTET_STREAM_TYPES.has(type)

// Do the bytes look like plain text? One NUL means binary; otherwise 85%+ of the head must be
// printable.
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

const PDF_MAGIC = Buffer.from('%PDF-')
// A .docx/.xlsx is a zip, so this recognizes one from content alone.
const ZIP_MAGIC = Buffer.from([0x50, 0x4b, 0x03, 0x04])
// RTF is ASCII, so looksLikeText would take a mislabeled one as text and leak the control words.
// Detected here to skip instead (see findByContentType).
const RTF_MAGIC = Buffer.from('{\\rtf')
// Legacy OLE (.doc). Shared with .xls/.ppt/.msg, so we only claim doc when the extension confirms it.
const OLE_MAGIC = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])

// docx and xlsx share the zip magic; their main part is what tells them apart.
const DOCX_PART = Buffer.from('word/document.xml')
const XLSX_PART = Buffer.from('xl/workbook.xml')

const startsWith = (content: Buffer, magic: Buffer): boolean =>
    content.length >= magic.length && content.subarray(0, magic.length).equals(magic)

// Which OOXML kind a zip is, by its root part. Returns undefined for a non-OOXML zip (pptx, jar).
const ooxmlKind = (content: Buffer): HandlerKind | undefined => {
    // Match exact zip ENTRY names, not raw bytes: a docx can embed a workbook, and that workbook's
    // xl/workbook.xml lives INSIDE another entry rather than as an entry of this package — so a raw
    // scan gets fooled by storage order, while a root-entry match doesn't.
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
// A small in-cap .docx/.xlsx can inflate to hundreds of MB and OOM the worker, and the zip's own
// declared size is attacker-controlled. So measure it: stream-inflate each entry (peak stays ~one
// zlib chunk), count REAL bytes, abort once the total crosses the cap. Unmeasurable = fail closed.

const EOCD_MAGIC = Buffer.from([0x50, 0x4b, 0x05, 0x06]) // End-of-Central-Directory
const CD_SIG = 0x02014b50 // Central-Directory file header
const LOCAL_SIG = 0x04034b50 // Local file header

// `ok` = safe to hand to the parser; otherwise `reason` is the skip message.
type DecompressionCheck = { ok: true } | { ok: false; reason: string }

// Inflate one raw-deflate region onto `runningTotal`, aborting the moment it would exceed `cap`.
// Returns the new total, or a sentinel — -1 = over budget, -2 = corrupt stream. Byte counts are
// never negative, so the sentinels are unambiguous.
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

// Pick the EOCD the way the parser's zip reader (jszip, inside exceljs/mammoth) does: the LAST
// signature in the buffer, no comment-length check. Matching its choice is the point — measuring a
// different directory than the parser reads is a bomb-bypass. An invariant here would diverge: jszip
// follows a second EOCD planted after the real one, so rejecting that leaves the parser inflating a
// directory we never measured. Bonus: last-match doesn't false-skip zips with bytes after the EOCD.
const findEocd = (buf: Buffer): number => {
    const eocd = buf.lastIndexOf(EOCD_MAGIC)
    return eocd >= 0 && eocd + 22 <= buf.length ? eocd : -1 // need room for the 22-byte fixed record
}

// Every entry's stored name. Sees only the PACKAGE's own parts, so an embedded object's internal
// paths can't fool root detection. Returns undefined when the directory can't be walked — the caller
// falls back to a raw scan.
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

// Measure a zip's ACTUAL decompressed size, capped, from each entry's real (structural, not
// self-declared) compressed region. The two invariants below pin us to the records jszip will read;
// both assert the FILE is self-consistent rather than mirroring jszip, so neither rots if it
// changes. Every real archive satisfies them (73 measured, 0 failures).
const checkDecompressionBudget = async (buf: Buffer, cap: number): Promise<DecompressionCheck> => {
    const eocd = findEocd(buf)
    if (eocd < 0) return { ok: false, reason: 'malformed zip: no end-of-central-directory record' }

    const entries = buf.readUInt16LE(eocd + 10)
    const cdSize = buf.readUInt32LE(eocd + 12)
    const cdOffset = buf.readUInt32LE(eocd + 16)
    // ZIP64 / out-of-range sentinels: the true values live in a ZIP64 record we don't chase. Treat as
    // over-budget rather than trust the classic field or crash on the sentinel.
    if (entries === 0xffff || cdSize === 0xffffffff || cdOffset === 0xffffffff)
        return { ok: false, reason: 'zip declares a ZIP64 / out-of-range size' }

    // Invariant 1: the directory must END exactly where the EOCD begins.
    // jszip doesn't read from cdOffset unconditionally — when eocdPos - (cdOffset + cdSize) is
    // positive it rebases every offset by that gap (its support for data prepended ahead of the
    // archive). A nonzero gap therefore aims the two readers at two different directories, and a
    // bomb planted at the rebased one is a bomb we never measure.
    if (cdOffset + cdSize !== eocd)
        return { ok: false, reason: 'malformed zip: central directory does not end at the end-of-central-directory record' }

    let total = 0
    let p = cdOffset
    for (let i = 0; i < entries; i++) {
        // Guard: a missing/misaligned header means the offset lied. Fail closed — a partial walk must
        // never silently return the total it accumulated so far.
        if (p + 46 > buf.length || buf.readUInt32LE(p) !== CD_SIG)
            return { ok: false, reason: 'malformed zip: truncated or misaligned central directory' }
        const method = buf.readUInt16LE(p + 10)
        const compSize = buf.readUInt32LE(p + 20)
        const localOffset = buf.readUInt32LE(p + 42)
        if (compSize === 0xffffffff || localOffset === 0xffffffff)
            return { ok: false, reason: 'zip declares a ZIP64 / out-of-range size' }
        // Read the local header's own name/extra lengths — they can differ from the central copy, and
        // they're what fixes where this entry's compressed bytes actually begin.
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
    // Invariant 2: walking exactly `entries` records must land exactly on the EOCD.
    // Invariant 1 isn't enough. jszip ignores the declared count, reading headers until the signature
    // stops matching, and doesn't error when its tally disagrees — so an archive can declare one
    // entry, store two, and size the directory honestly, hiding a record from this counted walk.
    // Landing on the EOCD proves none hides: the next bytes are the EOCD signature, so jszip's
    // signature-driven loop stops exactly where this one did, on the same records.
    if (p !== eocd)
        return { ok: false, reason: 'malformed zip: central directory holds more records than it declares' }
    return { ok: true }
}

/////////////////////////////////////////////////////////////
// ROUTING (continued) — claim verification + the final decision

// Do the bytes back up a binary claim? Catches a wrong one before the parser sees it. text/html have
// no single signature, so they always pass here — lying text claims go to bytesContradictTextClaim.
const magicOk = (kind: HandlerKind, content: Buffer): boolean => {
    if (kind === 'pdf') return startsWith(content, PDF_MAGIC)
    if (kind === 'doc') return startsWith(content, OLE_MAGIC)
    if (kind === 'docx') return ooxmlKind(content) === 'docx'
    if (kind === 'xlsx') return ooxmlKind(content) === 'xlsx'
    return true
}

// Magic bytes are the strongest signal, so they go first. A BOM is next: utf-16 text is NUL-heavy,
// so looksLikeText would reject it as binary. ooxmlKind picks docx vs xlsx from content, so a zip
// needs no extension to disambiguate.
const sniff = (content: Buffer, ext?: string): HandlerKind | undefined => {
    if (startsWith(content, PDF_MAGIC)) return 'pdf'
    if (startsWith(content, ZIP_MAGIC)) return ooxmlKind(content)
    if (startsWith(content, OLE_MAGIC) && ext === '.doc') return 'doc'
    // RTF: skip before looksLikeText misclaims this ASCII markup as text and leaks the control words.
    if (startsWith(content, RTF_MAGIC)) return undefined
    if (bomCharset(content) || looksLikeText(content)) return 'text'
    return undefined
}

// Which utf-16 flavour an explicit charset= names, if any. BOM-less utf-16 is NUL-heavy and fails
// looksLikeText, so the hint is the only thing keeping it routable — but the hint alone must not be
// enough to earn that exemption. See isWellFormedUtf16.
const claimsUtf16 = (charset?: string): 'utf-16' | 'utf-16le' | 'utf-16be' | undefined => {
    const normalized = charset?.trim().toLowerCase().replace(/_/g, '-')
    return normalized === 'utf-16' || normalized === 'utf-16le' || normalized === 'utf-16be' ? normalized : undefined
}

// A genuine text file never starts with these, so they contradict a text claim even under a hint.
const hasKnownBinaryMagic = (content: Buffer): boolean =>
    startsWith(content, PDF_MAGIC) || startsWith(content, ZIP_MAGIC) || startsWith(content, OLE_MAGIC)

// Are the bytes structurally well-formed utf-16? A printable-ratio test can't tell: read as utf-16,
// arbitrary bytes land across the BMP and are nearly all "printable", so png/jpeg/gif sail through.
// Well-formedness can. The surrogate block is 1/32 of the BMP, so binary hits it constantly and
// essentially never as a correct high-then-low pair; real utf-16 pairs every one and never carries
// the U+FFFE/U+FFFF noncharacters. Either tell proves the bytes aren't the utf-16 they claim to be.
const isWellFormedUtf16 = (content: Buffer, bigEndian: boolean): boolean => {
    const sample = content.subarray(0, SNIFF_BYTES)
    const end = sample.length - (sample.length % 2) // whole code units only
    if (end === 0) return false
    for (let i = 0; i < end; i += 2) {
        const unit = bigEndian ? sample.readUInt16BE(i) : sample.readUInt16LE(i)
        if (unit === 0xfffe || unit === 0xffff) return false // noncharacter
        if (unit >= 0xdc00 && unit <= 0xdfff) return false // low surrogate with no high before it
        if (unit >= 0xd800 && unit <= 0xdbff) {
            // Nothing after a high surrogate means two different things. A truncated sample just puts
            // the low half out of view — no evidence, and real text must not be voided over a
            // sampling artifact. A file that ENDS here genuinely ends unpaired, which utf-16 never does.
            if (i + 2 >= end) return content.length > SNIFF_BYTES
            const low = bigEndian ? sample.readUInt16BE(i + 2) : sample.readUInt16LE(i + 2)
            if (low < 0xdc00 || low > 0xdfff) return false // high surrogate not followed by a low
            i += 2 // consume the pair
        }
    }
    return true
}

// A text/html claim is only as good as its bytes: a PDF sent as text/plain would decode into garbage
// and report 'extracted' — a silent quality failure, worse than a labeled skip. Binary bytes void the
// claim; the sniff then rescues whatever the magic proves, and anything else is left unrouted. Empty
// content keeps the claimed handler, so a zero-byte text attachment lands on 'extracted', not a skip.
const bytesContradictTextClaim = (content: Buffer, charsetHint?: string): boolean => {
    if (content.length === 0 || bomCharset(content)) return false
    // Magic beats even a charset=utf-16 hint — real utf-16 never starts with %PDF/PK/OLE. Cheap and
    // decisive, so it runs first; but it's an allowlist of three, which is why the branch below can't
    // lean on it.
    if (hasKnownBinaryMagic(content)) return true
    const utf16 = claimsUtf16(charsetHint)
    if (utf16) {
        // The exemption suppresses the printable-ratio check, so the bytes have to earn it rather than
        // the sender just asserting it. Without this, every binary outside the three magics above —
        // png, jpeg, gif, gzip — keeps the exemption and decodes to gibberish reported as 'extracted'.
        // A bare `utf-16` is BOM-less by here, so iconv picks an endianness heuristically: accept
        // either, since either is what it may choose.
        const le = utf16 !== 'utf-16be' && isWellFormedUtf16(content, false)
        const be = utf16 !== 'utf-16le' && isWellFormedUtf16(content, true)
        return !le && !be
    }
    // Any other claim must actually look like text.
    return !looksLikeText(content)
}

// The final routing decision: try each signal in priority order, stop at the first that works.
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
        // Distrust a text claim the bytes contradict, or one whose bytes are RTF — RTF is printable
        // ASCII so it slips past bytesContradictTextClaim, and re-sniffing hits sniff()'s RTF skip
        // instead of decoding control words as body text.
        if (
            (claimed.kind === 'text' || claimed.kind === 'html') &&
            (bytesContradictTextClaim(input.content, charsetHint) || startsWith(input.content, RTF_MAGIC))
        ) {
            const sniffed = sniff(input.content, ext)
            return sniffed ? { kind: sniffed, routedBy: 'sniff' } : { routedBy: 'none' }
        }
        // Distrust a binary claim the content contradicts — a PDF mislabeled .docx, or a spreadsheet
        // mislabeled .docx. Wrong bytes just fail the parser and lose any real text, so re-sniff for
        // the true format. Empty content has nothing to check: leave it to the parser.
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

// JS lets you throw non-Errors, so normalize whatever came out.
const errorMessage = (error: unknown): string => (error instanceof Error ? error.message : String(error))

/////////////////////////////////////////////////////////////
// ENTRY POINT — every step in order, each risky one inside its own safety net.

// Returns a labeled result: status, the text (omitted when there is none — never ''), and a reason
// on skip/fail. Never throws, whatever the bytes.
export const extractAttachment = async (input: AttachmentInput): Promise<ExtractionResult> => {
    const byteSize = input.content.length

    // Resolve empties here so the status doesn't depend on the declared type — otherwise an empty
    // PDF routes into a parser that throws on zero bytes ('failed') while empty text is 'extracted'.
    // All empties are ran-but-empty.
    if (byteSize === 0) {
        return { status: 'extracted' }
    }

    // Size gate, before any decode or parse.
    if (byteSize > MAX_INPUT_BYTES) {
        return { status: 'skipped', reason: `${byteSize} bytes exceeds ${MAX_INPUT_BYTES}` }
    }

    const { type, charset: charsetHint } = parseContentType(input.contentType)
    const { kind } = detectRoute(input)
    const handler = kind ? findHandler(kind) : undefined

    // Unsupported or unrecognized format.
    if (!handler) {
        return { status: 'skipped', reason: type ? `unsupported type ${type}` : 'unrecognized attachment' }
    }

    // Zip bombs: measure the real decompressed size and skip before exceljs/mammoth touch the bytes.
    // (pdf/doc/text aren't zips.)
    if (kind === 'docx' || kind === 'xlsx') {
        const check = await checkDecompressionBudget(input.content, MAX_UNCOMPRESSED_BYTES)
        if (!check.ok) return { status: 'skipped', reason: check.reason }
    }

    try {
        const output = await withTimeout(
            handler.extract({ content: input.content, filename: input.filename, charsetHint }),
            HANDLER_TIMEOUT_MS
        )
        const isEmpty = output.empty ?? output.text.trim().length === 0
        // Central cap, so a pathological document can't dump megabytes into S3 and the search index.
        // xlsx and pdf already capped incrementally, so for them this is just the final precise trim;
        // docx and html return a full string, so for those it's POST-materialization — peak memory
        // follows the whole document, and hard containment is the host memory limit (see README).
        // Don't split a surrogate pair at the boundary: a lone half serializes as U+FFFD.
        const overCap = output.text.length > MAX_OUTPUT_CHARS
        const capEnd =
            overCap && output.text.charCodeAt(MAX_OUTPUT_CHARS - 1) >= 0xd800 && output.text.charCodeAt(MAX_OUTPUT_CHARS - 1) <= 0xdbff
                ? MAX_OUTPUT_CHARS - 1
                : MAX_OUTPUT_CHARS
        const text = overCap ? output.text.slice(0, capEnd) : output.text
        // A present `extraction` reads as "has text"; its absence as "ran, but empty".
        return isEmpty ? { status: 'extracted' } : { status: 'extracted', extraction: text }
    } catch (error) {
        // The bytes are attacker-controlled, so a throw or timeout is an expected event, not a bug:
        // label it and move on rather than crashing the caller.
        return { status: 'failed', reason: errorMessage(error) }
    }
}