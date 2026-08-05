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
// hit S3 and the search index. The CEILING: ExtractOptions.maxOutputChars may only tighten it.
// Cutting sets `truncated` — a partial extraction that reads as complete is worse than a missing one.
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

// The two refusals differ by whose fault it is, so a caller can branch on them: `skipped` is a file
// we chose not to read, `failed` is one we couldn't. Never collapse them.
export type ExtractionStatus =
    | 'extracted' // handler ran; `extraction` holds the text, or is omitted when there was none
    | 'skipped' // intact, but declined: over MAX_INPUT_BYTES or the decompression budget, unsupported, unrecognized
    | 'failed' // broken or hostile bytes: parser threw, timed out, decode failed, or the zip preflight found it malformed

// Which signal decided the route.
export type RoutedBy = 'content-type' | 'extension' | 'sniff' | 'none'

// Formats we route to; each maps to exactly one handler. text covers txt/csv/ics/vcard/json/xml/yaml.
export type HandlerKind = 'text' | 'html' | 'pdf' | 'docx' | 'doc' | 'xlsx'

export interface AttachmentInput {
    content: Buffer
    filename?: string
    contentType?: string
}

// How to extract, as opposed to AttachmentInput's what. Omitting both reproduces 0.3.0 behaviour.
export interface ExtractOptions {
    // Clamped to MAX_OUTPUT_CHARS — may only tighten. A streaming handler reads until the cap, so a
    // caller-supplied ceiling above ours would turn a safety constant into a footgun.
    maxOutputChars?: number
    // Appended only when the text was cut. The library owns WHEN, the caller owns WHAT, so wording
    // is tunable without a publish. OUTSIDE cap accounting: output may exceed the cap by this length.
    trailer?: string
}

export interface ExtractionResult {
    status: ExtractionStatus
    extraction?: string // omitted entirely (never '') when the handler produced no text
    reason?: string // set on skipped / failed
    // Whether the document continues past `extraction`. Set on `extracted` only — skipped/failed
    // have no text to have cut. Independent of `trailer`, so a consumer never parses the text for it.
    truncated?: boolean
}

interface HandlerContext {
    content: Buffer
    filename?: string
    charsetHint?: string // from the content-type charset= param
    // Both resolved centrally in extractAttachment; no handler defaults either for itself.
    maxOutputChars: number // the effective cap; handlers that build incrementally stop here
    deadline: number // Date.now() ceiling; handlers that yield check it between units of work
}

interface HandlerOutput {
    text: string
    empty?: boolean // handler's own emptiness call; defaults to text.trim() === ''
    // Set by a handler that stopped early; ORed with the entry point's over-cap check. A handler
    // stopping ON the cap or on the deadline lands under it and would otherwise look complete.
    truncated?: boolean
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
// or saxes.

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
    extract: async ({ content, maxOutputChars, deadline }) => {
        const { getDocumentProxy } = await import('unpdf')
        const pdf = await getDocumentProxy(new Uint8Array(content))
        // Iterate pages ourselves — unpdf's extractText parses EVERY page up front, so a pathological
        // page count runs unbounded. Bounds our accumulation and the pages parsed, NOT pdf.js's
        // per-page decompression (no hook exists); that residual is the host memory limit's job.
        const pageCount = Math.min(pdf.numPages, MAX_PDF_PAGES)
        const pages: string[] = []
        let length = 0
        let truncated = false
        for (let n = 1; n <= pageCount; n++) {
            // This loop awaits per page, so the deadline is enforceable here in a way withTimeout's
            // race is not. Before the fetch: stopping is only useful if it precedes the work.
            if (Date.now() > deadline) {
                truncated = true
                break
            }
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
                if (length > maxOutputChars) {
                    truncated = true
                    break // one page of overshoot, trimmed centrally
                }
            }
        }
        // Pages past the ceiling are text we never read.
        if (pageCount < pdf.numPages) truncated = true
        const joined = pages.join('\n\n').trim()
        return { text: joined, empty: joined.length === 0, truncated }
    },
}

// DOCX — modern OOXML Word. Streams word/document.xml; see DOCX STREAMING READER for the machinery,
// and for why the output contract is mammoth's, reproduced rather than invented.
const docxHandler: Handler = {
    kind: 'docx',
    contentTypes: ['application/vnd.openxmlformats-officedocument.wordprocessingml.document'],
    extensions: ['.docx'],
    extract: async ({ content, maxOutputChars, deadline }) => {
        // Fail closed, as the xlsx handler does on !rewritten. Both bails are unreachable rather
        // than defensive: an archive whose central directory can't be walked never gets past
        // checkDecompressionBudget to reach this (Invariant 2 makes budget-ok imply zipEntries-ok
        // over the same records), and routing pins the second — ooxmlKind only answers 'docx' for an
        // archive whose directory NAMES word/document.xml. Costs nothing, and keeps neither of those
        // proofs load-bearing.
        const entries = zipEntries(content)
        if (!entries) throw new Error('docx central directory could not be read for streaming')
        const part = entries.find((entry) => entry.name === DOCX_MAIN_PART)
        if (!part) throw new Error(`docx archive has no ${DOCX_MAIN_PART} entry`)

        const reader = await createDocxReader()
        // Decode ACROSS inflate chunks, not per chunk: a 16 KB boundary lands mid-sequence in any
        // document with a non-ASCII character, and chunk.toString('utf8') would turn that one
        // character into two U+FFFD. StringDecoder carries the partial bytes forward. (saxes handles
        // a surrogate pair split across write() calls itself; this is the layer below that.)
        const { StringDecoder } = await import('node:string_decoder')
        const decoder = new StringDecoder('utf8')

        let truncated = false
        try {
            for await (const chunk of docxMainPartChunks(part)) {
                // Both guards here, ahead of the work, at one inflate chunk of granularity. Finer
                // than the pdf per-page and xlsx per-row checks, and the only place a stop is
                // possible: saxes has no abort, so the way to stop parsing is to stop feeding it.
                // Per CHUNK rather than per emitted character, for the same reason the xlsx deadline
                // sits above its empty-row skip — tens of MB of w:pPr/w:rPr markup produces no text
                // at all, so a cap-only check would never fire on precisely the cheapest loop to
                // spin. Breaking a `for await` destroys the inflate stream, so the rest of the
                // document is never decompressed either.
                if (Date.now() > deadline || reader.chars() > maxOutputChars) {
                    truncated = true
                    break
                }
                reader.write(decoder.write(chunk))
            }
            // Only a read that ran to the end may assert the document ended cleanly.
            if (!truncated) reader.end(decoder.end())
        } catch (error) {
            // saxes is conformant where mammoth's DOM parser recovered, so a document the old reader
            // read to the end can stop short here. Text already extracted is still text, and the
            // contract has a word for "the document continues past this point" — so keep it and say
            // so. Nothing read means nothing to label: that stays a failure, which a caller can see
            // and retry. What this deliberately does NOT do is install a saxes error handler and
            // parse on; measured, that emits close-tag text as content and descends into elements
            // mammoth drops — silent wrong output, the one outcome this file fails over everywhere else.
            if (reader.text().trim().length === 0) throw error
            truncated = true
        }

        // Emptiness is left to the entry point: an empty document is exactly '\n\n', which trims to ''.
        return { text: reader.text(), truncated }
    },
}

// DOC — legacy Word 97–2003. The .docx reader above is an OOXML zip reader, so the OLE binary needs
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

// The slice of exceljs's Row we touch. Structural because the streaming iterator's declared element
// type (Row) disagrees with what it yields at runtime (Row[]), putting the real element out of reach.
interface StreamedRow {
    eachCell: (options: { includeEmpty: boolean }, callback: (cell: { text?: string }) => void) => void
}

// XLSX — modern Excel. Each sheet flattened to text for search/indexing.
const xlsxHandler: Handler = {
    kind: 'xlsx',
    contentTypes: ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
    extensions: ['.xlsx'],
    extract: async ({ content, maxOutputChars, deadline }) => {
        const { Readable } = await import('node:stream')
        const { default: ExcelJS } = await import('exceljs') // not SheetJS: no known parse-time CVEs
        // Stream rather than workbook.xlsx.load(), which materializes every cell as a live object
        // before any cap can apply — a 4 MB in-cap .xlsx peaked at hundreds of MB and OOMed a 1024 MB
        // worker. Peak now follows the shared-string table, not the cell graph. See
        // reorderForStreaming: without it this reader drops worksheets.
        const rewritten = reorderForStreaming(content)
        // Fail closed: the budget measured the central directory, but unzipper inflates what its
        // LOCAL-header walk finds, so the budget binds this path only through the rewritten archive.
        if (!rewritten) throw new Error('xlsx central directory could not be read for streaming')
        const { content: ordered, worksheets: expected } = rewritten
        const reader = new ExcelJS.stream.xlsx.WorkbookReader(Readable.from(ordered), {
            worksheets: 'emit',
            sharedStrings: 'cache', // the only mode that resolves t="s" cells to their text
            styles: 'cache', // number formats; 'ignore' would render dates as raw serials
            hyperlinks: 'ignore',
            entries: 'ignore',
        })

        const sheets: string[] = []
        let length = 0
        let truncated = false
        let seen = 0
        for await (const worksheet of reader) {
            seen++
            const rows: string[] = []
            for await (const batch of worksheet) {
                // Documented quirk: this yields an ARRAY of rows ("worksheetReader returns an array
                // of rows ... for performance reasons"); older exceljs yielded one.
                const batchRows = (Array.isArray(batch) ? batch : [batch]) as StreamedRow[]
                for (const row of batchRows) {
                    // ABOVE the empty-row skip: a contentless row is `continue`d, and a sheet of them
                    // trips no cap either, so a check below would never run. This loop awaits, so
                    // breaking really ends the parse — unlike withTimeout's race, which frees the
                    // slot while the parse detaches.
                    if (Date.now() > deadline) {
                        truncated = true
                        break
                    }
                    const cells: string[] = []
                    // cell.text = the shown value (formula result, formatted date), not the raw formula.
                    row.eachCell({ includeEmpty: false }, (cell) => cells.push(cell.text ?? ''))
                    // Drop empty cells/rows so a sparse sheet doesn't flatten into runs of empty tabs.
                    if (cells.length === 0) continue
                    const line = cells.join('\t')
                    rows.push(line)
                    length += line.length + 1 // + newline
                    if (length > maxOutputChars) {
                        truncated = true
                        break // one line of overshoot, trimmed centrally
                    }
                }
                if (truncated) break
            }
            // Cast for the TYPE, not the value: the .d.ts omits `name`, but the constructor always
            // sets it (worksheet-reader.js:21) and _parseWorksheet overwrites it from the model
            // (workbook-reader.js:306). The fallback is unreachable at 4.4.0, kept because this
            // library has already moved undocumented behaviour under us twice.
            const name = (worksheet as unknown as { name?: string }).name ?? `Sheet${sheets.length + 1}`
            if (rows.length > 0) sheets.push(`=== ${name} ===\n${rows.join('\n')}`)
            // Safe to abandon mid-archive: the reorder keeps every worksheet on the inline path,
            // which spools nothing, so no temp files are stranded (#2147).
            if (truncated) break
        }

        // The reorder closes the two KNOWN paths to entry loss; this closes the class. Any third
        // surfaces as worksheets never yielded, and becomes a labeled failure rather than half a
        // spreadsheet reported as 'extracted' — a caller can retry a failure but cannot tell a
        // truncated document from a complete one. Gated on !truncated: a deliberate stop leaves
        // sheets unread by design.
        if (!truncated && seen < expected) {
            throw new Error(`xlsx reader yielded ${seen} of ${expected} worksheets`)
        }

        return { text: sheets.join('\n\n'), truncated }
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
//
// WHICH READER THIS BINDS — since the .docx handler stopped using mammoth, one answer for both
// formats: OUR OWN zipEntries walk. Neither handler hands the original bytes to a third-party zip
// reader any more, so there is no foreign directory choice left to second-guess.
//
//   .docx -> zipEntries locates word/document.xml and the handler inflates exactly that region of
//     exactly this buffer. Nothing else in the archive is decompressed at read time.
//
//   .xlsx -> exceljs -> unzipper dispatches on LOCAL file-header signatures front-to-back
//     (unzipper/lib/parse.js:51), only skipping PAST directory records — so a central-directory
//     measurement would not bind it at all. reorderForStreaming binds it instead: it re-emits one
//     local header per measured entry and nothing else, so both walks see the same set by
//     construction. Hence the xlsx handler FAILS when the rewrite can't be produced; the original
//     bytes would drop the budget on the floor.
//
// So the two invariants below no longer exist to keep a foreign reader honest. Invariant 2 has
// instead become the proof that this measurement and zipEntries can never disagree — now
// load-bearing for BOTH formats. Invariant 1 is a structural check that no longer prevents anything.
// See each.
//
// STALE, AND KNOWN TO BE: this measures EVERY entry, while the .docx handler now inflates one. On
// real documents word/document.xml is 2–28% of the archive (a Word template measured 66 KB of 3.1 MB),
// so a document with ≥50 MB of compressible non-document parts — EMF/WMF vector art, embedded OLE
// objects — is declined over bytes that would never be read. Narrow: ordinary media is PNG/JPEG,
// already compressed, so 50 MB of it needs >10 MB of archive and MAX_INPUT_BYTES turns it away
// first. Left alone deliberately — loosening a zip-bomb gate is a permissiveness change that belongs
// with the input caps, not smuggled in behind a parser swap.

const EOCD_MAGIC = Buffer.from([0x50, 0x4b, 0x05, 0x06]) // End-of-Central-Directory
const CD_SIG = 0x02014b50 // Central-Directory file header
const LOCAL_SIG = 0x04034b50 // Local file header

// `ok` = safe to hand to the parser. Otherwise `status` is which kind of no, since the two differ to
// a caller: `failed` = the bytes are broken (the parser would have thrown anyway), `skipped` = intact
// but we decline — over budget, or a variant we don't chase. Mirrors the MAX_INPUT_BYTES precedent.
type DecompressionCheck = { ok: true } | { ok: false; status: 'failed' | 'skipped'; reason: string }

const corrupt = (reason: string): DecompressionCheck => ({ ok: false, status: 'failed', reason })
const declined = (reason: string): DecompressionCheck => ({ ok: false, status: 'skipped', reason })

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

// The LAST signature in the buffer, no comment-length check. Chosen originally to match jszip, which
// read .docx through mammoth: measuring a different directory than the parser reads is a bomb-bypass,
// and jszip follows a second EOCD planted after the real one. jszip is out of the read path now, so
// the choice is purely internal — this function is shared by BOTH walks (zipEntries and the budget),
// so any deterministic pick is self-consistent. Kept as last-match because it still has to agree with
// unzipper on the .xlsx path, and because it doesn't false-skip zips with bytes after the EOCD.
const findEocd = (buf: Buffer): number => {
    const eocd = buf.lastIndexOf(EOCD_MAGIC)
    return eocd >= 0 && eocd + 22 <= buf.length ? eocd : -1 // need room for the 22-byte fixed record
}

// One entry as the central directory describes it, plus its stored (still-compressed) bytes.
interface ZipEntry {
    name: string
    method: number
    crc: number
    compSize: number
    uncompSize: number
    data: Buffer // the raw stored bytes — deflated unless method is 0
}

// Walk the central directory. Sees only the PACKAGE's own parts, so an embedded object's internal
// paths can't fool root detection. Returns undefined when the directory can't be walked; what the
// two callers do with that differs, because they ask for different things:
//   ooxmlKind  — identification, so it degrades to a raw byte scan.
//   reorderForStreaming — the entry set unzipper will be given, so it refuses and its caller fails.
//   docxHandler — the entry it will inflate, so it refuses and the handler fails.
//
// Those last two make this walk load-bearing for the zip-bomb guard on BOTH formats, not merely a
// lenient identifier: on .xlsx the archive it emits is what makes the budget's measurement bind, and
// on .docx it is what picks the one region the handler decompresses (see DECOMPRESSION BUDGET).
// Still deliberately NOT merged with checkDecompressionBudget — that one is the measurement
// itself, and its two invariants must reject archives this one accepts. Merging would force a single
// contract onto both. What keeps the split safe is the one-way relation pinned at Invariant 2: this
// walk can never be the more permissive of the two.
const zipEntries = (buf: Buffer): ZipEntry[] | undefined => {
    const eocd = findEocd(buf)
    if (eocd < 0) return undefined
    const count = buf.readUInt16LE(eocd + 10)
    if (count === 0xffff) return undefined // ZIP64 entry count — not chased here
    const entries: ZipEntry[] = []
    let p = buf.readUInt32LE(eocd + 16)
    for (let i = 0; i < count; i++) {
        if (p + 46 > buf.length || buf.readUInt32LE(p) !== CD_SIG) return undefined
        const nameLen = buf.readUInt16LE(p + 28)
        if (p + 46 + nameLen > buf.length) return undefined
        const compSize = buf.readUInt32LE(p + 20)
        const localOffset = buf.readUInt32LE(p + 42)
        if (compSize === 0xffffffff || localOffset === 0xffffffff) return undefined // ZIP64
        // The local header carries its own name/extra lengths, which can differ from the central
        // copy — they are what fixes where this entry's bytes actually start.
        if (localOffset + 30 > buf.length || buf.readUInt32LE(localOffset) !== LOCAL_SIG) return undefined
        const dataStart = localOffset + 30 + buf.readUInt16LE(localOffset + 26) + buf.readUInt16LE(localOffset + 28)
        const data = buf.subarray(dataStart, dataStart + compSize)
        if (data.length < compSize) return undefined
        entries.push({
            name: buf.toString('latin1', p + 46, p + 46 + nameLen),
            method: buf.readUInt16LE(p + 10),
            crc: buf.readUInt32LE(p + 16),
            compSize,
            uncompSize: buf.readUInt32LE(p + 24),
            data,
        })
        p += 46 + nameLen + buf.readUInt16LE(p + 30) + buf.readUInt16LE(p + 32)
    }
    return entries
}

const zipEntryNames = (buf: Buffer): string[] | undefined => zipEntries(buf)?.map((entry) => entry.name)

// Measure a zip's ACTUAL decompressed size, capped, from each entry's real (structural, not
// self-declared) compressed region. The two invariants below were written to pin us to the records
// jszip would read on the old .docx path; jszip is gone, and what they now do is stated at each.
// Both assert the FILE is self-consistent rather than mirroring any particular reader, which is why
// neither rotted when the reader changed. Every real archive satisfies them (73 measured, 0 failures).
const checkDecompressionBudget = async (buf: Buffer, cap: number): Promise<DecompressionCheck> => {
    const eocd = findEocd(buf)
    if (eocd < 0) return corrupt('malformed zip: no end-of-central-directory record')

    const entries = buf.readUInt16LE(eocd + 10)
    const cdSize = buf.readUInt32LE(eocd + 12)
    const cdOffset = buf.readUInt32LE(eocd + 16)
    // ZIP64 / out-of-range sentinels: the true values live in a ZIP64 record we don't chase. Treat as
    // over-budget rather than trust the classic field or crash on the sentinel.
    if (entries === 0xffff || cdSize === 0xffffffff || cdOffset === 0xffffffff)
        return declined('zip declares a ZIP64 / out-of-range size')

    // Invariant 1: the directory must END exactly where the EOCD begins.
    // Written when .docx went through jszip, which rebases every offset by a positive
    // `eocdPos - (cdOffset + cdSize)` gap — its support for data prepended ahead of the archive — so
    // a nonzero gap aimed the two readers at two different directories and a bomb planted at the
    // rebased one went unmeasured. Nothing rebases now: zipEntries reads cdOffset raw, exactly as
    // this does, so a gap can no longer split the walks. Demoted from a bypass guard to a structural
    // check, and kept because it costs nothing and fails such a file HERE with an accurate message
    // rather than a few lines down on a local header that doesn't match.
    // Deliberate: that gap is also how a self-extracting archive legitimately carries its stub, so
    // this calls real files malformed. Accepted — an email attachment has no business being one.
    if (cdOffset + cdSize !== eocd)
        return corrupt('malformed zip: central directory does not end at the end-of-central-directory record')

    let total = 0
    let p = cdOffset
    for (let i = 0; i < entries; i++) {
        // Guard: a missing/misaligned header means the offset lied. Fail closed — a partial walk must
        // never silently return the total it accumulated so far.
        if (p + 46 > buf.length || buf.readUInt32LE(p) !== CD_SIG)
            return corrupt('malformed zip: truncated or misaligned central directory')
        const method = buf.readUInt16LE(p + 10)
        const compSize = buf.readUInt32LE(p + 20)
        const localOffset = buf.readUInt32LE(p + 42)
        if (compSize === 0xffffffff || localOffset === 0xffffffff)
            return declined('zip declares a ZIP64 / out-of-range size')
        // Read the local header's own name/extra lengths — they can differ from the central copy, and
        // they're what fixes where this entry's compressed bytes actually begin.
        if (localOffset + 30 > buf.length || buf.readUInt32LE(localOffset) !== LOCAL_SIG)
            return corrupt('malformed zip: bad local header offset')
        const dataStart = localOffset + 30 + buf.readUInt16LE(localOffset + 26) + buf.readUInt16LE(localOffset + 28)
        const comp = buf.subarray(dataStart, dataStart + compSize)
        if (comp.length < compSize)
            return corrupt('malformed zip: compressed data runs past end of file')

        if (method === 0) {
            total += comp.length // stored (no compression): output === input
        } else if (method === 8) {
            total = await inflateCounting(comp, total, cap)
            if (total === -1) return declined(`decompresses to over ${cap} bytes`)
            if (total === -2) return corrupt('malformed zip: unreadable compressed data')
        } else {
            return declined(`zip uses unsupported compression method ${method}`)
        }
        if (total > cap) return declined(`decompresses to over ${cap} bytes`)
        p += 46 + buf.readUInt16LE(p + 28) + buf.readUInt16LE(p + 30) + buf.readUInt16LE(p + 32)
    }
    // Invariant 2: walking exactly `entries` records must land exactly on the EOCD.
    // This one GAINED a job. Its original reason was jszip, which ignores the declared count and
    // reads headers until the signature stops matching: an archive could declare one entry, store
    // two, size the directory honestly, and hide a record from this counted walk. jszip is out of the
    // read path entirely now — but zipEntries is IN it, for both formats, and this is what pins the
    // two walks together. They read the same count from the same findEocd and advance by the same
    // expression, so they visit the same records; zipEntries has exactly ONE bail this function lacks
    // — a name length running past the buffer — and that bail pushes `p` past the EOCD, which lands
    // here. That is the whole proof that budget-ok implies zipEntries-ok over the same set. The .docx
    // handler inflates one of those entries and the .xlsx preflight re-emits all of them; neither can
    // reach bytes this function did not measure.
    if (p !== eocd)
        return corrupt('malformed zip: central directory holds more records than it declares')
    return { ok: true }
}

/////////////////////////////////////////////////////////////
// XLSX STREAMING PREFLIGHT (exceljs entry-order workaround)
//
// A worksheet entry reaching exceljs's reader before `this.sharedStrings && this.workbookRels` are
// both set takes a spool branch (workbook-reader.js:110) that pipes it to a temp file and awaits
// that with the zip stream paused — violating unzipper's contract ("If you do not intend to consume
// an entry stream's raw data, call autodrain() ... Otherwise the stream will halt"). It halts, and
// later entries are never emitted: measured on a 2-sheet workbook, entries dropped in ~34% of reads,
// and where xl/workbook.xml is also lost the reader throws "Cannot read properties of undefined
// (reading 'sheets')". Upstream exceljs #2790, #3064, #2147 — all open, none since 4.4.0 (Oct 2023).
//
// We can't fix their reader, only choose what bytes it reads: ordering both parts ahead of every
// worksheet keeps it off that branch. Entry bytes are copied still-compressed, so nothing is
// inflated or recompressed and peak stays ~2x the (already capped) compressed size.
//
// Verified: 300 reads across 6 workbooks, byte-identical to workbook.xlsx.load(), zero drops and
// zero throws — against 0/50 clean on the worst of them before the reorder.

// Ordered first, so every flag the worksheet branch tests is set before a worksheet is reached.
// xl/workbook.xml is not one of those flags but leads anyway: it sets this.model, dereferenced by
// the same function for the sheet NAME, and the source of the 'sheets' throw when missing.
const XLSX_LEADING_ENTRIES = [
    '[Content_Types].xml',
    '_rels/.rels',
    'xl/workbook.xml', // -> this.model  (sheet names; the 'sheets' TypeError without it)
    'xl/_rels/workbook.xml.rels', // -> this.workbookRels
    'xl/sharedStrings.xml', // -> this.sharedStrings
    'xl/styles.xml', // -> this.styles      (number formats)
]

// A workbook with no strings has no xl/sharedStrings.xml, so this.sharedStrings is never set and
// ordering alone cannot lift it out of the spool branch (measured: 35 of 50 reads dropped sheets).
// Injecting an empty table sets the flag and changes no cell — a workbook omitting the part has no
// `t="s"` cell to resolve through it. Handed to the reader only, never written back.
const EMPTY_SHARED_STRINGS = Buffer.from(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>',
    'latin1'
)
// CRC32 of the literal above, precomputed — the payload is fixed, so computing it per extraction
// would be work with one possible answer. A test recomputes it, so the two can't drift.
const EMPTY_SHARED_STRINGS_CRC = 0x2949bd0b

const LOCAL_HEADER_BYTES = 30
const CENTRAL_HEADER_BYTES = 46
const EOCD_BYTES = 22
const ZIP_VERSION = 20 // 2.0 — the floor for deflate, which is all we re-emit

// One xl/worksheets/sheetN.xml per worksheet the reader should yield. Counted from the archive, not
// xl/workbook.xml's <sheet> list: that list also names chartsheets, which have no worksheet part and
// would make a correct read look like it had lost one.
const WORKSHEET_ENTRY = /^xl\/worksheets\/sheet\d+\.xml$/

// Rebuild with XLSX_LEADING_ENTRIES first, everything else after in its original order, reporting
// the worksheet count so the caller can verify it got them all. Returns undefined when the directory
// can't be walked, and the caller MUST fail rather than fall back to the original bytes (see
// DECOMPRESSION BUDGET). No known archive reaches this — every bail below is also a budget
// rejection, and the lone divergence (a name length overrunning the buffer) is caught by Invariant
// 2 — so failing closed only keeps that from being load-bearing.
const reorderForStreaming = (buf: Buffer): { content: Buffer; worksheets: number } | undefined => {
    const entries = zipEntries(buf)
    if (!entries) return undefined

    const complete = entries.some((entry) => entry.name === 'xl/sharedStrings.xml')
        ? entries
        : [
              ...entries,
              {
                  name: 'xl/sharedStrings.xml',
                  method: 0, // stored: the injected table is 153 bytes, so compressing it is noise
                  crc: EMPTY_SHARED_STRINGS_CRC,
                  compSize: EMPTY_SHARED_STRINGS.length,
                  uncompSize: EMPTY_SHARED_STRINGS.length,
                  data: EMPTY_SHARED_STRINGS,
              },
          ]

    // The EOCD's entry count is 16-bit, where 0xffff means "ZIP64, the real count is elsewhere".
    // zipEntries refuses an archive already declaring it, but the injection above can carry a
    // 65534-entry archive onto it, writing a sentinel where a count belongs. Reachable inside
    // MAX_INPUT_BYTES: ~76 bytes of headers per entry, so 65534 fit in ~5 MB.
    if (complete.length >= 0xffff) return undefined

    const rank = (name: string) => {
        const index = XLSX_LEADING_ENTRIES.indexOf(name)
        return index === -1 ? XLSX_LEADING_ENTRIES.length : index
    }
    // Stable by construction: equal ranks keep their original relative order, so worksheets stay in
    // sheet1, sheet2, ... order and the output's sheet sequence is unchanged.
    const ordered = [...complete].sort((a, b) => rank(a.name) - rank(b.name))

    const locals: Buffer[] = []
    const centrals: Buffer[] = []
    let offset = 0
    for (const entry of ordered) {
        const name = Buffer.from(entry.name, 'latin1')

        const local = Buffer.alloc(LOCAL_HEADER_BYTES)
        local.writeUInt32LE(LOCAL_SIG, 0)
        local.writeUInt16LE(ZIP_VERSION, 4)
        local.writeUInt16LE(entry.method, 8)
        local.writeUInt32LE(entry.crc, 14)
        local.writeUInt32LE(entry.compSize, 18)
        local.writeUInt32LE(entry.uncompSize, 22)
        local.writeUInt16LE(name.length, 26)
        locals.push(local, name, entry.data)

        const central = Buffer.alloc(CENTRAL_HEADER_BYTES)
        central.writeUInt32LE(CD_SIG, 0)
        central.writeUInt16LE(ZIP_VERSION, 4)
        central.writeUInt16LE(ZIP_VERSION, 6)
        central.writeUInt16LE(entry.method, 10)
        central.writeUInt32LE(entry.crc, 16)
        central.writeUInt32LE(entry.compSize, 20)
        central.writeUInt32LE(entry.uncompSize, 24)
        central.writeUInt16LE(name.length, 28)
        central.writeUInt32LE(offset, 42)
        centrals.push(central, name)
        offset += LOCAL_HEADER_BYTES + name.length + entry.data.length
    }

    const directory = Buffer.concat(centrals)
    const end = Buffer.alloc(EOCD_BYTES)
    end.writeUInt32LE(0x06054b50, 0)
    end.writeUInt16LE(ordered.length, 8)
    end.writeUInt16LE(ordered.length, 10)
    end.writeUInt32LE(directory.length, 12)
    end.writeUInt32LE(offset, 16)
    return {
        content: Buffer.concat([...locals, directory, end]),
        worksheets: complete.filter((entry) => WORKSHEET_ENTRY.test(entry.name)).length,
    }
}

/////////////////////////////////////////////////////////////
// DOCX STREAMING READER (WordprocessingML -> raw text)
//
// Reads word/document.xml straight out of the archive with a SAX parser, so the largest live object
// is one inflate chunk plus the text kept so far — never a DOM. mammoth, which this replaces, built
// an xmldom tree AND a document model on top of it, then let the central cap throw almost all of it
// away. Measured through dist/ at a 1024 MB heap on a 45 MB document.xml inside a 3.65 MB archive —
// every gate cleared, nothing skipped:
//
//                       mammoth                     this
//     concurrency 1     812 ms /  607 MB            68 ms / 175 MB
//     concurrency 2    1552 ms /  970 MB            59 ms / 164 MB
//     concurrency 4    3512 ms / 1217 MB            58 ms / 168 MB
//
// mammoth built 43,420,044 characters every time and kept 250,000 — 0.6%. Its peak tracks the
// document and multiplies by concurrency; this one is flat, because the cap now stops the READ
// rather than trimming the result. The sharpest case isn't even the big archive: a 0.41 MB
// attachment holding one 43 MB w:t peaked mammoth at 1270 MB, versus 199 MB here.
//
// Same failure mode as the .xlsx reader, one format over, and reachable from half a megabyte.
//
// THE OUTPUT CONTRACT is mammoth's extractRawText, reproduced deliberately rather than invented
// (raw-text.js is 13 lines): w:t text verbatim, w:tab -> '\t', a paragraph's children then TWO
// newlines, everything else its children only. Nothing is trimmed, so a document ends with '\n\n'.
//
// THE SHAPE is a WHITELIST, mirroring body-reader.js:46-57 — an element mammoth has no handler for
// has its ENTIRE SUBTREE dropped rather than recursed into, and bare text outside a w:t is dropped
// too. So the tables below are what we KEEP; anything absent disappears with its children. That is
// the load-bearing decision: "emit every w:t, newline on </w:p>" over-extracts on any document with
// a text box or a field. It also DELETES work rather than adding it —
//   - mammoth's 20-entry ignore list (w:pPr, w:rPr, w:sectPr, w:proofErr, w:tblGrid, ...) has no
//     counterpart here. It exists only to suppress a warning; for raw text "deliberately ignored"
//     and "unrecognised" produce the identical empty result, and the default already IS that result.
//   - every emit-nothing leaf mammoth spells out (w:br, w:cr, w:fldChar, w:instrText,
//     w:footnoteReference, w:bookmarkStart) is likewise just the default.
//   - w:sdt keeping only w:sdtContent, and mc:AlternateContent keeping only mc:Fallback, fall out
//     for free: w:sdtPr and mc:Choice simply aren't on the list.
//
// Fidelity is pinned by tests/docx-fidelity.test.ts against mammoth itself over a real Word corpus.
// The deliberate divergences are listed at DOCX_CONTAINERS.

const DOCX_MAIN_PART = 'word/document.xml'

// Namespace URI -> the prefix mammoth's element names carry (office-xml-reader.js:10-36). Matching a
// literal `w:` prefix instead would be a shippable bug three times over: the ISO-strict format binds
// w to a different URI (mammoth's own corpus carries strict-format.docx), a producer may bind
// wordprocessingml as the DEFAULT namespace — giving <document>/<body>/<p> with no prefix at all —
// and a producer may pick any prefix it likes. Only these four URIs are ever matched; mammoth's
// other eleven mappings feed its HTML converter and its ignore list, neither of which survives here.
const OOXML_PREFIXES: Record<string, string> = {
    'http://schemas.openxmlformats.org/wordprocessingml/2006/main': 'w', // transitional
    'http://purl.oclc.org/ooxml/wordprocessingml/main': 'w', // ISO strict
    'http://schemas.openxmlformats.org/markup-compatibility/2006': 'mc',
    'urn:schemas-microsoft-com:vml': 'v',
}

// Pre-bound so an undeclared prefix doesn't end the parse. saxes is stricter than the DOM parser
// this replaces and fails with `unbound namespace prefix` where xmldom shrugs, and real producers do
// emit stray o:/w10:/wne:/wps: markup without declaring it. Binding them up front costs nothing —
// the elements are dropped either way, only the throw is avoided — and it converts the likeliest
// strictness regression into a non-event. Verified: an in-document xmlns still shadows these, so a
// strict-format file resolves to the strict URI and routes correctly.
const OOXML_ASSUMED_PREFIXES: Record<string, string> = {
    w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    v: 'urn:schemas-microsoft-com:vml',
    o: 'urn:schemas-microsoft-com:office:office',
    w10: 'urn:schemas-microsoft-com:office:word',
    r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
}

// Recurse into these; emit nothing of their own. Absent from the list = subtree dropped.
const DOCX_CONTAINERS = new Set([
    'w:document',
    'w:body',
    'w:r',
    'w:hyperlink',
    'w:ins',
    'w:smartTag',
    'w:tbl',
    'w:tr',
    'w:tc',
    'w:sdt', // w:sdtPr isn't here, which is the whole of mammoth's firstOrEmpty("w:sdtContent")
    'w:sdtContent',
    'mc:AlternateContent',
    'mc:Fallback', // likewise mc:Choice: dropped by omission
    'w:object',
    'w:drawing',
    'w:txbxContent',
    'v:group',
    'v:rect',
    'v:roundrect',
    'v:shape',
    'v:textbox',
])
// Deliberately NOT here, each matching mammoth:
//   wp:inline / wp:anchor — mammoth's readDrawingElement digs for a:blip images and never reads
//     arbitrary children, so the DrawingML text-box path (wps:txbx) yields nothing. Text boxes reach
//     us through the VML fallback above, which is the branch mc:Fallback selects anyway.
//   v:imagedata, v:shapetype, v:shadow — images and shape defs; no text either way.
//   w:fldSimple, w:delText, w:ruby — mammoth DROPS these with their children. w:delText is right
//     (deleted text does not belong in an extract); w:fldSimple and w:ruby are real text loss in
//     mammoth, replicated on purpose so this change stays a pure port. Recovering them is one array
//     entry each, deliberately left to its own change once the corpus says how often they matter.
//
// Deviations from mammoth, all four in tests/docx-fidelity.test.ts's ACCEPTED_DIVERGENCES if they
// ever show up on a real document: a w:sdt carrying wordml:checkbox has its first text character
// REPLACED by a checkbox node in mammoth (we keep the character — strictly better, and costs nothing
// here); w:sym needs mammoth's dingbat-to-unicode table to map (we drop it, and it is the only
// element that would make this reader read an attribute at all); a w:t holding a comment or a nested
// element makes mammoth's text() throw "Not implemented" (we extract); and the main part is read at
// its conventional path rather than resolved through _rels/.rels, which routing already requires.

// Emit a literal, then drop any children — mammoth's handlers for these ignore children entirely.
const DOCX_LITERALS: Record<string, string> = {
    'w:tab': '\t',
    'w:noBreakHyphen': '\u2011',
    'w:softHyphen': '\u00ad',
}

// xmldom normalised these to '\n' before mammoth ever saw the markup (dom-parser.js:34-38). saxes
// does the XML-standard \r\n and \r itself — correctly, across chunk boundaries — but leaves NEL and
// LINE SEPARATOR alone. Both are single characters, so matching xmldom on the way out needs no
// cross-chunk state, unlike normalising the raw markup would.
const XML_EXTRA_SEPARATORS = /[\u0085\u2028]/g

// The two <w:del> markers that change what an ANCESTOR emits, so they have to be spotted even though
// they sit inside an already-dropped w:pPr / w:trPr subtree. ECMA-376 17.13.5.15 (deleted paragraph
// mark: the mark is gone, so the paragraph merges into the next one) and 17.13.5.12 (deleted table
// row: the whole row goes).
const PARAGRAPH_DELETED = ['w:p', 'w:pPr', 'w:rPr', 'w:del']
const ROW_DELETED = ['w:tr', 'w:trPr', 'w:del']

const endsWith = (stack: string[], path: string[]): boolean =>
    path.every((name, i) => stack[stack.length - path.length + i] === name)

// A paragraph's own text, and the text hoisted out of any w:pict inside it. mammoth returns a
// (value, extra) pair from every element it reads: w:pict moves its whole result into `extra`
// (body-reader.js:439 .toExtra()), extras bubble up through every container, and only w:p reinserts
// them — as a SIBLING AFTER the paragraph (:296 .insertExtra()). So VML text-box text lands after
// the paragraph's '\n\n' rather than glued into the middle of it, and an extra that never reaches a
// w:p is silently lost. Two string fields per open w:p / w:pict reproduce all of that.
interface DocxFrame {
    value: string
    extra: string
}

interface DocxReader {
    write: (chunk: string) => void // one decoded slice of document.xml; THROWS on malformed XML
    chars: () => number // characters emitted so far — what the handler's cap check reads
    end: (tail: string) => void // flush, close, and run the end-of-document checks
    text: () => string // the text, complete or partial
}

const createDocxReader = async (): Promise<DocxReader> => {
    // Lazy like every other parser here: a Lambda that only ever sees text never loads saxes.
    const { SaxesParser } = await import('saxes')

    const stack: string[] = [] // qualified names of every open element, innermost last
    const frames: DocxFrame[] = [{ value: '', extra: '' }] // index 0 is the document; extras reaching it are lost
    const deleted: boolean[] = [false] // one flag per open w:p; index 0 pairs with the root frame
    let top = frames[frames.length - 1]
    let skip = -1 // stack index where the dropped subtree began, or -1 when we're reading
    let rowDeleted = false // a w:trPr said its row is deleted; act on it once that w:trPr closes
    let sawBody = false
    let chars = 0

    const emit = (text: string): void => {
        top.value += text
        chars += text.length
    }

    const parser = new SaxesParser({ xmlns: true, additionalNamespaces: OOXML_ASSUMED_PREFIXES })

    parser.on('opentag', (tag) => {
        // A direct port of mammoth's convertName (xml/reader.js:53-66): mapped URI -> `w:t`,
        // unmapped -> `{uri}local`, no namespace -> the bare local name. Only the first form can
        // match a table, so the other two exist to guarantee a miss — but they keep the stack
        // readable in a debugger, and they are what makes an unmapped namespace fail closed.
        const uri = tag.uri ?? ''
        const prefix = OOXML_PREFIXES[uri]
        const name = uri === '' ? tag.local : prefix === undefined ? `{${uri}}${tag.local}` : `${prefix}:${tag.local}`
        stack.push(name)

        if (skip >= 0) {
            // Inside a dropped subtree nothing is emitted, and only the two ancestor-affecting
            // markers are still looked for. The `skip ===` tests prove the marker really sits in the
            // properties element of a LIVE w:p / w:tr — i.e. that the drop began at that w:pPr /
            // w:trPr — rather than somewhere deeper in unrelated debris.
            if (name === 'w:del') {
                if (skip === stack.length - 3 && endsWith(stack, PARAGRAPH_DELETED)) deleted[deleted.length - 1] = true
                else if (skip === stack.length - 2 && endsWith(stack, ROW_DELETED)) rowDeleted = true
            }
            return
        }

        // w:t is the ONE text-bearing element; the text handler recognizes it off the stack top, so
        // it needs no flag of its own. Deliberately before the container test, so its children —
        // illegal anyway — fall through to the drop below exactly as they do in mammoth.
        if (name === 'w:t') return

        if (name === 'w:p' || name === 'w:pict') {
            top = { value: '', extra: '' }
            frames.push(top)
            if (name === 'w:p') deleted.push(false)
            return
        }
        if (DOCX_CONTAINERS.has(name)) {
            if (name === 'w:body') sawBody = true
            return
        }
        const literal = DOCX_LITERALS[name]
        if (literal !== undefined) emit(literal)
        // The whitelist default: this element and everything under it is gone.
        skip = stack.length - 1
    })

    parser.on('text', (text) => {
        if (skip < 0 && stack[stack.length - 1] === 'w:t') emit(text.replace(XML_EXTRA_SEPARATORS, '\n'))
    })

    parser.on('closetag', () => {
        const name = stack.pop()

        if (skip >= 0) {
            if (stack.length !== skip) return // still inside the dropped subtree
            skip = -1
            // A deleted row's properties have just closed; drop the rest of the row with them.
            if (rowDeleted && name === 'w:trPr') {
                rowDeleted = false
                skip = stack.length - 1
            }
            return
        }

        if (name !== 'w:p' && name !== 'w:pict') return
        const frame = top
        frames.pop()
        top = frames[frames.length - 1]
        if (name === 'w:pict') {
            // Hoist: the picture's own text becomes the parent's `extra`, behind any extra it
            // already carried. It reaches the output only if some ancestor w:p reinserts it.
            top.extra += frame.extra + frame.value
        } else if (deleted.pop()) {
            // The paragraph mark was deleted, so there is no paragraph break here: this text runs
            // straight into the next paragraph. mammoth re-reads the stashed children in the next
            // paragraph's context and lands on the same string; suppressing the tail is the
            // streaming form of that, and is equivalent because this reader is context-free.
            top.value += frame.value
            top.extra += frame.extra
        } else {
            top.value += `${frame.value}\n\n${frame.extra}`
            chars += 2
        }
    })

    return {
        write: (chunk) => void parser.write(chunk),
        chars: () => chars,
        end: (tail) => {
            parser.write(tail).close() // close() is the well-formedness check: it throws on an unclosed element
            // mammoth throws "Could not find the body element: are you sure this is a docx file?"
            // here — for a foreign root AND for a w:document with no w:body, both verified. Keeping
            // it means a zip whose main part is not WordprocessingML stays a labeled failure instead
            // of becoming an 'extracted' with no text.
            if (!sawBody) throw new Error('docx main part has no w:body element')
        },
        // Frames above the root are only still open on the truncated path; concatenating every
        // frame's `value` in order is exactly the partial text in document order, so a document that
        // is one enormous paragraph still returns what was read instead of ''. On the complete path
        // there is only the root frame and this is just frames[0].value. Pending `extra` is dropped
        // either way — mammoth's own behaviour for an extra that never reaches a w:p.
        text: () => frames.reduce((text, frame) => text + frame.value, ''),
    }
}

// One shape for both storage methods, so the handler's loop has a single form. Method 0 is its own
// single chunk (bounded by MAX_INPUT_BYTES, and no real producer stores document.xml); method 8
// yields a chunk at a time under backpressure, which is what keeps peak memory flat.
const docxMainPartChunks = (part: ZipEntry): AsyncIterable<Buffer> | Iterable<Buffer> => {
    if (part.method === 0) return [part.data]
    const inflate = zlib.createInflateRaw()
    inflate.end(part.data) // pushes the compressed bytes; the readable side inflates only on demand
    return inflate
}

/////////////////////////////////////////////////////////////
// ROUTING (continued) — claim verification + the final decision

// Extensions naming a NON-doc OLE format. The OLE magic is shared, so it alone can't confirm a doc
// claim — but an extension naming a sibling format contradicts it. A missing extension contradicts
// nothing, so a real .doc sent without a filename still routes.
const NON_DOC_OLE_EXTENSIONS = new Set(['.xls', '.ppt', '.msg'])

// Do the bytes back up a binary claim? Catches a wrong one before the parser sees it. text/html have
// no single signature, so they always pass here — lying text claims go to bytesContradictTextClaim.
const magicOk = (kind: HandlerKind, content: Buffer, ext?: string): boolean => {
    if (kind === 'pdf') return startsWith(content, PDF_MAGIC)
    if (kind === 'doc') return startsWith(content, OLE_MAGIC) && !(ext && NON_DOC_OLE_EXTENSIONS.has(ext))
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
        if (input.content.length > 0 && !magicOk(claimed.kind, input.content, ext)) {
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

// A caller's cap may only tighten: absent or non-finite falls back to the ceiling, negative clamps
// to 0, fractional floors (a cap is a whole number of chars).
const resolveCap = (requested?: number): number =>
    requested === undefined || !Number.isFinite(requested)
        ? MAX_OUTPUT_CHARS
        : Math.min(MAX_OUTPUT_CHARS, Math.max(0, Math.floor(requested)))

// Returns a labeled result: status, the text (omitted when there is none — never ''), whether it was
// truncated, and a reason on skip/fail. Never throws, whatever the bytes.
export const extractAttachment = async (
    input: AttachmentInput,
    options: ExtractOptions = {}
): Promise<ExtractionResult> => {
    const byteSize = input.content.length

    // Resolve empties here so the status doesn't depend on the declared type — otherwise an empty
    // PDF routes into a parser that throws on zero bytes ('failed') while empty text is 'extracted'.
    // All empties are ran-but-empty.
    // `truncated: false` is stated because this is the one 'extracted' return that never reaches the
    // cap logic below, and omitting it would hand `undefined` to a consumer testing `=== false`.
    if (byteSize === 0) {
        return { status: 'extracted', truncated: false }
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

    // Zip bombs: measure the real decompressed size before either OOXML handler inflates anything.
    // The preflight also decides which refusal this is — over budget skips, malformed fails.
    // (pdf/doc/text aren't zips.)
    if (kind === 'docx' || kind === 'xlsx') {
        const check = await checkDecompressionBudget(input.content, MAX_UNCOMPRESSED_BYTES)
        if (!check.ok) return { status: check.status, reason: check.reason }
    }

    const maxOutputChars = resolveCap(options.maxOutputChars)

    try {
        const output = await withTimeout(
            handler.extract({
                content: input.content,
                filename: input.filename,
                charsetHint,
                maxOutputChars,
                // The same budget withTimeout races on, as a value a handler can act on: the race
                // can't cancel work already inside a parser; a handler checking this stops itself.
                deadline: Date.now() + HANDLER_TIMEOUT_MS,
            }),
            HANDLER_TIMEOUT_MS
        )
        // Central cap, so a pathological document can't dump megabytes into S3 and the search index.
        // Handlers that build incrementally overshoot by one unit, so for them this is the final
        // precise trim; html alone still returns a full string, so for that one it's
        // POST-materialization — peak memory follows the whole document, and hard containment is the
        // host memory limit (see README). Don't split a surrogate pair at the boundary: a lone half
        // serializes as U+FFFD.
        const overCap = output.text.length > maxOutputChars
        const capEnd =
            overCap && output.text.charCodeAt(maxOutputChars - 1) >= 0xd800 && output.text.charCodeAt(maxOutputChars - 1) <= 0xdbff
                ? maxOutputChars - 1
                : maxOutputChars
        const text = overCap ? output.text.slice(0, capEnd) : output.text
        // Decided on the FINAL text: a tight enough cap slices a non-empty extraction to '', and the
        // contract is that `extraction` is omitted rather than ever being ''. `||`, not `??`, because
        // the handler's call can only ADD emptiness — pdf computes `empty` from its pre-cap page join
        // and answers `false` for text this then slices away, which `??` would emit as '' (or, with a
        // trailer, as a bare trailer and no document text).
        const isEmpty = text.trim().length === 0 || (output.empty ?? false)
        // Either signal means the document continues past `text`: the handler stopped itself, or it
        // handed back more than the cap and we cut it.
        const truncated = (output.truncated ?? false) || overCap
        // A present `extraction` reads as "has text"; its absence as "ran, but empty". The trailer
        // goes on AFTER the cap slice, so it never displaces extracted text.
        return isEmpty
            ? { status: 'extracted', truncated }
            : { status: 'extracted', extraction: truncated && options.trailer ? text + options.trailer : text, truncated }
    } catch (error) {
        // The bytes are attacker-controlled, so a throw or timeout is an expected event, not a bug:
        // label it and move on rather than crashing the caller.
        return { status: 'failed', reason: errorMessage(error) }
    }
}