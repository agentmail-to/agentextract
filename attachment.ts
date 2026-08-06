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

// How far AHEAD of that timer a handler's own deadline sits. Without the gap the two are the same
// instant and a handler can never report its own stop: the checks are `Date.now() > deadline`, so the
// earliest one fires is deadline + 1ms, by which point withTimeout has already rejected — and a
// partial-but-real extraction comes back as `failed`, discarding exactly the work those checks exist
// to keep. This buys the turn a handler needs to finish the unit in flight and return. A margin, not
// a guarantee: one page or row slower than this still loses the race, and `failed` is the right
// answer when a single unit overruns.
export const HANDLER_DEADLINE_MARGIN_MS = 1_000

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
    // Date.now() ceiling, sitting a margin INSIDE withTimeout's; handlers that yield check it between
    // units of work. See HANDLER_DEADLINE_MARGIN_MS for why the two instants must not coincide.
    deadline: number
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
        const rewritten = await reorderForStreaming(content)
        // Fail closed: the budget measured the central directory, but unzipper inflates what its
        // LOCAL-header walk finds, so the budget binds this path only through the rewritten archive.
        if (!rewritten.ok) throw new Error(rewritten.reason)
        // The workbook's worksheets, in tab order — see SHEET IDENTITY. Both the names below and the
        // backstop's count read off this rather than off the archive or the reader.
        const { content: ordered, sheets: resolved } = rewritten
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
                // 4.4.0 yields ONE Row per iteration (worksheet-reader.js:275 pushes
                // `{eventType: 'row', value: row}`, and :104-112 yields each `value` through) — the
                // other way round from how this once read. The normalization stays anyway: exceljs
                // documents the batched shape, which is why the wrong version of this was believable,
                // and it costs one predicate to be right under either.
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
            // Positional, and sound because the rebuild laid out exactly `resolved` and nothing else
            // this reader dispatches as a worksheet, in this order. Reading worksheet.name instead is
            // what produced "Sheet1" for a legal absolute rel Target: exceljs matches rel.Target
            // against one exact spelling (workbook-reader.js:302) and gives up on every other, and
            // its .d.ts does not admit the field either way. The fallback is unreachable — an
            // out-of-range index means the reader emitted a part we never wrote.
            const name = resolved[seen - 1]?.name ?? `Sheet${seen}`
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
        //
        // `!==`, not `<`: over-yielding is now just as wrong, because it means the reader dispatched
        // a worksheet the rebuild did not lay out — which would have silently shifted every name
        // above. Unreachable by construction, and cheap enough not to leave that proof load-bearing.
        if (!truncated && seen !== resolved.length) {
            throw new Error(`xlsx reader yielded ${seen} of ${resolved.length} worksheets`)
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
// WHICH READER THIS BINDS — the two formats no longer share one:
//
//   .docx -> mammoth -> jszip reads the CENTRAL DIRECTORY, so measuring it measures what jszip will
//     inflate. The invariants below pin both walks to the same records; their reasoning is this case.
//
//   .xlsx -> exceljs -> unzipper dispatches on LOCAL file-header signatures front-to-back
//     (unzipper/lib/parse.js:51), only skipping PAST directory records — so the invariants do not
//     bind it, and a local entry the directory omits would still be inflated unmeasured.
//     reorderForStreaming binds it instead. What it writes is bounded in BOTH directions, and only
//     stating the first is how a 21x amplification once read as impossible from this comment:
//       - NO MORE. Every local header it writes comes from a MEASURED central-directory record, and
//         each such record is written AT MOST ONCE. Sheet layout carries resolved ZipEntry objects
//         rather than looking parts up by name — the load-bearing half, because a zip name is not a
//         key: 21 entries all named xl/worksheets/sheet1.xml once collapsed onto the first of them
//         and re-emitted its bytes 21 times, turning a measured 1.5 MB into 31 MB.
//       - NO FEWER IS FINE. Some measured records are deliberately not written at all (orphan
//         worksheet parts — see SHEET IDENTITY). Dropping only removes bytes from the reader's
//         reach, so it cannot loosen the bound.
//       - ONE EXCEPTION, ours. The injected empty xl/sharedStrings.xml matches no central record. It
//         is a fixed 153-byte literal in this file, not anything the input controls.
//     So what unzipper can inflate is at most what was measured, plus those 153 bytes. Hence the
//     xlsx handler FAILS when the rewrite can't be produced; the original bytes would drop the
//     budget on the floor.

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

// Pick the EOCD the way jszip (inside mammoth, and so the .docx reader) does: the LAST signature in
// the buffer, no comment-length check. Matching its choice is the point — measuring a different
// directory than the parser reads is a bomb-bypass. An invariant here would diverge: jszip
// follows a second EOCD planted after the real one, so rejecting that leaves the parser inflating a
// directory we never measured. Bonus: last-match doesn't false-skip zips with bytes after the EOCD.
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

// Walk the central directory for entry NAMES. Sees only the PACKAGE's own parts, so an embedded
// object's internal paths can't fool root detection — which is the whole reason ooxmlKind prefers
// this to a raw byte scan.
//
// Deliberately more lenient than zipEntries below, and split from it for exactly that reason: naming
// an entry needs no local header, no compressed region and no ZIP64 size fields, so every extra bail
// zipEntries takes for the REWRITE would here only push identification back onto the byte scan this
// exists to beat. Nothing is inflated on the strength of this answer, so leniency costs nothing.
const zipEntryNames = (buf: Buffer): string[] | undefined => {
    const eocd = findEocd(buf)
    if (eocd < 0) return undefined
    const count = buf.readUInt16LE(eocd + 10)
    if (count === 0xffff) return undefined // ZIP64 entry count — not chased here
    const names: string[] = []
    let p = buf.readUInt32LE(eocd + 16)
    for (let i = 0; i < count; i++) {
        if (p + 46 > buf.length || buf.readUInt32LE(p) !== CD_SIG) return undefined
        const nameLen = buf.readUInt16LE(p + 28)
        if (p + 46 + nameLen > buf.length) return undefined
        names.push(buf.toString('latin1', p + 46, p + 46 + nameLen))
        p += 46 + nameLen + buf.readUInt16LE(p + 30) + buf.readUInt16LE(p + 32)
    }
    return names
}

// The same walk, plus everything needed to RE-EMIT each entry: its stored bytes, located through the
// local header, and the fields a rebuilt header has to carry. Returns undefined when the directory
// can't be walked — and its one caller, reorderForStreaming, refuses rather than degrading, because
// the entry set it produces is what unzipper will be given.
//
// That makes this walk load-bearing for the .xlsx zip-bomb guard, not merely an identifier: the
// archive it feeds is what makes the budget's measurement bind (see DECOMPRESSION BUDGET). Still
// deliberately NOT merged with checkDecompressionBudget — that one is the measurement itself, and
// its two invariants must reject archives this one accepts. Merging would force a single contract
// onto both. What keeps the split safe is the one-way relation pinned at Invariant 2: this walk can
// never be the more permissive of the two.
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

// Measure a zip's ACTUAL decompressed size, capped, from each entry's real (structural, not
// self-declared) compressed region. The two invariants below pin us to the records jszip will read
// — the .docx path; the .xlsx path is bound differently, see the section header. Both assert the
// FILE is self-consistent rather than mirroring jszip, so neither rots if it changes. Every real
// archive satisfies them (73 measured, 0 failures).
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
    // jszip rebases every offset by a positive `eocdPos - (cdOffset + cdSize)` gap — its support for
    // data prepended ahead of the archive. A nonzero gap aims the two readers at two different
    // directories, and a bomb planted at the rebased one is one we never measure.
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
    // Invariant 1 isn't enough. jszip ignores the declared count, reading headers until the signature
    // stops matching, and doesn't error when its tally disagrees — so an archive can declare one
    // entry, store two, and size the directory honestly, hiding a record from this counted walk.
    // Landing on the EOCD proves none hides: the next bytes are the EOCD signature, so jszip's
    // signature-driven loop stops exactly where this one did, on the same records.
    // Doubles as the reason zipEntries can never be the weaker of the two walks: the one bail it has
    // that this function lacks — a name length running past the buffer — pushes `p` past the EOCD,
    // which lands here.
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
// Verified: 300 reads across 6 workbooks, zero drops and zero throws — against 0/50 clean on the
// worst of them before the reorder. Cell values, dates, number formats, booleans, formula results,
// rich text and unicode all match workbook.xlsx.load(). Two shapes deliberately do NOT, both pinned
// by tests: load() proxied a merged cell's master value into every slave, so a horizontal merge
// repeated the label once per column and a vertical one emitted trailing rows carrying nothing else,
// and an error-valued formula stringified to "[object Object]". Streaming emits the merge once and
// the error cell as empty. Better output for a search index either way, but merged workbooks do
// change row and column shape against main, so it is stated rather than filed under "identical".

// Ordered first, so every flag the worksheet branch tests is set before a worksheet is reached.
// xl/workbook.xml is not one of those flags but leads anyway: it sets this.model, whose absence is
// the 'sheets' TypeError above. Sheet NAMES no longer come from it — we resolve those ourselves, so
// the reader's own copy is only what keeps it from throwing.
const XLSX_LEADING_ENTRIES = [
    '[Content_Types].xml',
    '_rels/.rels',
    'xl/workbook.xml', // -> this.model  (the 'sheets' TypeError without it)
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

// One xl/worksheets/sheetN.xml per worksheet, anchored.
const WORKSHEET_PART = /^xl\/worksheets\/sheet\d+\.xml$/

// What EXCELJS treats as a worksheet: the same shape, UNANCHORED (workbook-reader.js:311). So it
// dispatches xl/worksheets/sheet1.xml.bak, or a copy nested under any prefix, where WORKSHEET_PART
// does not. That divergence was harmless while names came off the reader — it only ever made the
// backstop's expected count a subset of what was emitted, which cannot false-fire. It is not
// harmless now that names are positional, since one unresolved emission shifts every later one. The
// rebuild drops everything this matches that is not a resolved part, making emitted === laid out.
const EXCELJS_WORKSHEET_DISPATCH = /xl\/worksheets\/sheet\d+[.]xml/

/////////////////////////////////////////////////////////////
// SHEET IDENTITY
//
// Which worksheets exist, in what order, under what names — resolved from the WORKBOOK, not from the
// archive. exceljs takes all three from zip layout: it emits in entry order, and it names a sheet by
// matching rel.Target against one exact spelling (workbook-reader.js:302,
// `worksheets/sheet${n}.xml`), so a legal absolute Target ("/xl/worksheets/sheet1.xml") resolves to
// no name at all. workbook.xlsx.load() went through xl/workbook.xml's <sheets> list and the rels
// instead, which is why streaming diverged from it three ways at once: a workbook whose tabs had been
// dragged came back in file order, absolute targets came back as "Sheet1, Sheet2, ...", and orphan
// sheetN.xml parts that no <sheet> references were emitted as sheets of their own.
//
// Reading the two parts here puts the authority where the format puts it and makes
// reorderForStreaming the single place that decides — which is also what lets the reader be named
// positionally, and what keeps the lost-worksheet backstop counting something real.
//
// The workbook is the ORDERING and NAMING authority only. MEMBERSHIP stays the archive's: see the
// tail of workbookWorksheets for why handing it that third job silently deletes text.

const WORKBOOK_PART = 'xl/workbook.xml'
const WORKBOOK_RELS_PART = 'xl/_rels/workbook.xml.rels'

// The relationship types naming a sheet that has NO xl/worksheets part, so a declaration carrying
// one is accounted for rather than lost. Matched on the type, which the format defines, and never on
// the target's path, which the producer writes: a path whitelist is walked into by pointing a
// worksheet relationship at a planted xl/chartsheets/sheet1.xml. Anchored at the end so a type
// merely ending in these words cannot pass. Macrosheets are deliberately absent — treating one as
// unplaced costs a rescued orphan in the output, treating it as accounted-for could cost rows.
const NON_WORKSHEET_REL = /\/relationships\/(chartsheet|dialogsheet)$/

// One worksheet as the workbook describes it: the archive entry holding it, and its tab name.
//
// The ENTRY, deliberately, and not its name. Zip permits duplicate entry names, so a name is not a
// key: resolving layout through a name -> entry lookup mapped every reference to the FIRST entry
// carrying that name and re-emitted its bytes once per reference. An archive of 21 entries all named
// xl/worksheets/sheet1.xml turned a measured 1.5 MB into a 31 MB rebuild that way — bytes the
// decompression budget counted once and unzipper would then inflate 21 times. Carrying the entry
// makes that unrepresentable rather than merely guarded against.
interface WorkbookSheet {
    entry: ZipEntry
    name: string
}

const worksheetEntries = (entries: ZipEntry[]): ZipEntry[] => entries.filter((entry) => WORKSHEET_PART.test(entry.name))

// A worksheet part the workbook does not name. sheetN.xml's own number is the most stable label
// available, and is what the reader's own fallback produced for these.
const partFallbackName = (part: string): string => `Sheet${/(\d+)\.xml$/.exec(part)?.[1] ?? ''}`

// A ceiling on the two metadata parts, which is NOT the decompression budget. The budget bounds the
// whole archive at 50 MB, and letting one part spend all of it here costs ~2x that in peak: the
// Buffer, then the UTF-16 string the parse needs. Measured on a 46 KB attachment whose workbook.xml
// inflates to 42 MB — heap +89 MB and RSS 319 MB, against +49 MB and 205 MB before this resolver
// existed, on the path whose whole point was holding 24 concurrent parses inside 662 MB.
//
// exceljs is no argument for spending it: it streams both parts through saxes
// (workbook-reader.js:157-166 `parseStream(iterateStream(entry))`) and never holds either whole, so
// this is a memory class the read did not previously have, not one it already paid.
//
// 4 MB because a real workbook.xml is kilobytes — thousands of sheets plus their defined names still
// land far under it — so the cap can only be reached by a part padded to reach it. Over the cap is
// treated as unreadable, which degrades to archiveWorksheets rather than failing.
const MAX_METADATA_BYTES = 4 * 1024 * 1024

// Inflate one entry, bounded. Only ever called on the two parts above. undefined = unreadable or
// over the cap, which the caller treats as "the workbook did not tell us" rather than as an error.
const inflateEntry = (entry: ZipEntry): Promise<Buffer | undefined> => {
    // Stored: output === input, and the subarray is a view on bytes already resident.
    if (entry.method === 0) {
        return Promise.resolve(entry.data.length <= MAX_METADATA_BYTES ? entry.data : undefined)
    }
    if (entry.method !== 8) return Promise.resolve(undefined) // the budget already refuses these
    // maxOutputLength, not a post-hoc length check: it errors on the chunk that would cross the cap,
    // so the allocation never happens. uncompSize is self-declared and cannot be the guard.
    return new Promise((resolve) =>
        zlib.inflateRaw(entry.data, { maxOutputLength: MAX_METADATA_BYTES }, (error, out) =>
            resolve(error ? undefined : out)
        )
    )
}

// Namespace-aware, because a local name alone is not an identity. Matching `sheet` anywhere accepted
// a foreign <foo:sheet> planted in an <extLst> extension block: it claimed the real sheet's part
// first, so the sheet came back under the attacker's name and in the attacker's tab position, and
// the genuine declaration was skipped as already claimed. Element identity is (namespace, local
// name), and the parent has to be <sheets> — the only place ECMA-376 12.3.2 puts a tab declaration.
//
// The cost is that saxes with xmlns on treats an UNDECLARED prefix as fatal. That is contained: the
// caller catches the throw and degrades to archiveWorksheets, the same answer it already gives for a
// workbook it cannot read, and exceljs's own parse of those bytes fails too.
const SPREADSHEETML_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const PACKAGE_RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const OFFICE_RELS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

// One relationship as the workbook points at it. The TYPE is carried, not just the target: a Target
// is a string the producer chose, so classifying a declaration by the shape of its path let a
// worksheet relationship aimed at xl/chartsheets/anything read as a legitimate chart sheet — and the
// worksheet part that really held those rows was then dropped as an orphan.
interface WorkbookRel {
    target: string
    type: string
}

// <sheets> is tab order (ECMA-376 12.3.2); each <sheet> carries its name and an r:id into the rels.
// A parse and not a scan, because sheet names carry XML entities and either quote style.
const parseWorkbookParts = async (workbookXml: string, relsXml: string) => {
    const { SaxesParser } = await import('saxes')

    // EVERY <sheet> is recorded, including one missing its name or its r:id. Those cannot be
    // resolved, but the caller has to know the workbook DECLARED something it could not place —
    // dropping them here would make an unresolvable sheet indistinguishable from an orphan part.
    const declared: { name: string; rId?: string }[] = []
    const workbook = new SaxesParser<{ xmlns: true }>({ xmlns: true })
    // Parent tracking, so <sheet> counts only as a child of <sheets>.
    const open: { uri: string; local: string }[] = []
    workbook.on('opentag', (tag) => {
        const parent = open[open.length - 1]
        open.push({ uri: tag.uri, local: tag.local })
        if (tag.uri !== SPREADSHEETML_NS || tag.local !== 'sheet') return
        if (parent?.uri !== SPREADSHEETML_NS || parent.local !== 'sheets') return
        // r:id is namespace-qualified; `name` and the sibling `sheetId` are not, so neither can be
        // mistaken for it however the producer bound its prefixes.
        const rId = Object.values(tag.attributes).find((a) => a.uri === OFFICE_RELS_NS && a.local === 'id')?.value
        declared.push({ name: tag.attributes.name?.value ?? '', rId })
    })
    workbook.on('closetag', () => void open.pop())
    workbook.write(workbookXml).close()

    const targets = new Map<string, WorkbookRel>()
    const rels = new SaxesParser<{ xmlns: true }>({ xmlns: true })
    rels.on('opentag', (tag) => {
        if (tag.uri !== PACKAGE_RELS_NS || tag.local !== 'Relationship') return
        const value = (name: string) => tag.attributes[name]?.value
        const [Id, Target, Type, TargetMode] = [value('Id'), value('Target'), value('Type'), value('TargetMode')]
        // An external target points outside the package, so it is never an entry we hold.
        if (Id && Target && TargetMode !== 'External') targets.set(Id, { target: Target, type: Type ?? '' })
    })
    rels.write(relsXml).close()

    return { declared, targets }
}

// A Target is relative to the rels part's base — xl/ — but may legally be an absolute package path
// or climb out with '..'. exceljs compares the raw string against one form and misses every other;
// this maps all of them onto the archive's entry name. undefined = it escapes the package.
const resolveRelTarget = (target: string): string | undefined => {
    const trimmed = target.trim()
    if (trimmed === '') return undefined
    // OPC targets are URI references, so any character in one may legally be percent-encoded —
    // "worksheets/sheet%32.xml" names sheet2.xml. Decoding is what matches those to the entry rather
    // than missing it. A malformed escape throws; that target is simply unresolvable.
    let decoded: string
    try {
        decoded = decodeURIComponent(trimmed)
    } catch {
        return undefined
    }
    const path: string[] = []
    for (const segment of (decoded.startsWith('/') ? decoded.slice(1) : `xl/${decoded}`).split('/')) {
        if (segment === '' || segment === '.') continue
        if (segment === '..') {
            if (path.pop() === undefined) return undefined
            continue
        }
        path.push(segment)
    }
    return path.length > 0 ? path.join('/') : undefined
}

// The workbook's own worksheets, in tab order. undefined when it cannot say — a missing or unreadable
// workbook.xml / rels part. Deliberately not an error: the streaming reader degrades on exactly that
// input (it never sets this.workbookRels, so every sheet takes the spool path) yet still reads to
// completion, so failing here would turn a workbook that works today into a `failed`. archiveWorksheets
// is the fallback, and it is what main did for every workbook.
const workbookWorksheets = async (entries: ZipEntry[]): Promise<WorkbookSheet[] | undefined> => {
    const workbookEntry = entries.find((entry) => entry.name === WORKBOOK_PART)
    const relsEntry = entries.find((entry) => entry.name === WORKBOOK_RELS_PART)
    if (!workbookEntry || !relsEntry) return undefined

    const [workbookXml, relsXml] = await Promise.all([inflateEntry(workbookEntry), inflateEntry(relsEntry)])
    if (!workbookXml || !relsXml) return undefined

    let parsed: Awaited<ReturnType<typeof parseWorkbookParts>>
    try {
        parsed = await parseWorkbookParts(workbookXml.toString('utf8'), relsXml.toString('utf8'))
    } catch {
        return undefined // malformed — exceljs's own parse of the same bytes fails too
    }

    const worksheets = worksheetEntries(entries)
    const byPart = new Map<string, ZipEntry>()
    for (const entry of worksheets) if (!byPart.has(entry.name)) byPart.set(entry.name, entry)

    // Two sets, because a zip name is not a key here either. `claimedNames` answers "did any
    // declaration reference this part?", which is what makes an entry an ORPHAN. `placed` answers
    // "did we lay this entry out?", which is what makes it redundant. The second of two entries
    // sharing a claimed name is neither: referenced, never placed, and not an orphan.
    const claimedNames = new Set<string>()
    const placed = new Set<ZipEntry>()
    const ordered: WorkbookSheet[] = []
    // Declarations we could not place at all. Distinct from a declaration we placed OUTSIDE the
    // worksheets, which is a positive answer and costs nothing.
    let unplaced = 0
    for (const { name, rId } of parsed.declared) {
        const rel = rId === undefined ? undefined : parsed.targets.get(rId)
        if (rel === undefined || name === '') {
            unplaced += 1 // no relationship, an external one, or nothing to call it
            continue
        }
        // A chartsheet or dialogsheet is a real declaration with no xl/worksheets part for the reader
        // to yield, so it is accounted for and must NOT count as unplaced — otherwise every workbook
        // holding one flips into the rescue below and resurrects genuine orphans.
        //
        // Decided on the relationship TYPE, never on the target's path. The path is a string the
        // producer chose: classifying by it let a worksheet relationship aimed at a planted
        // xl/chartsheets/sheet1.xml read as a chart sheet, leaving the worksheet part that held the
        // rows unclaimed for the orphan rule to drop. Type is what the format states the target IS.
        if (NON_WORKSHEET_REL.test(rel.type)) continue
        // Anything else is expected to be a worksheet, whatever its declared type — an unknown or
        // absent type resolves here too, and unplaced is the safe answer for it.
        const part = resolveRelTarget(rel.target)
        if (part === undefined || !WORKSHEET_PART.test(part)) {
            unplaced += 1
            continue
        }
        claimedNames.add(part)
        const entry = byPart.get(part)
        if (!entry) {
            unplaced += 1 // declares a worksheet this archive does not hold
            continue
        }
        if (placed.has(entry)) continue // a second <sheet> on one part: another tab name, no new rows
        placed.add(entry)
        ordered.push({ entry, name })
    }

    // MEMBERSHIP is the workbook's only while the workbook accounted for everything it declared.
    //
    // An unplaced worksheet part is normally an ORPHAN — no <sheet> references it, load() ignored it,
    // and the rebuild drops it (that is the point). Two shapes are NOT orphans and must ship, because
    // dropping either deletes text invisibly: the second of two entries sharing a claimed name, and
    // — once ANY declaration went unplaced — every remaining part, since orphan and victim stop being
    // distinguishable from here. Invisibly, because the backstop compares `seen` against this list's
    // own length: both sides move together and it can never fire. That is silent partial output, the
    // one outcome this file fails over. A rescued part loses its tab position, which is unknowable,
    // never its rows.
    const rescued = worksheets.filter((entry) => !placed.has(entry) && (unplaced > 0 || claimedNames.has(entry.name)))
    // Zero worksheets is a real answer for a chartsheet-only workbook — but only from a workbook that
    // told us something. A <sheets> we could not read one declaration out of is not that answer, and
    // honoring it as one drops every worksheet part the archive holds: the whole document, silently,
    // as an `extracted` with no text. undefined hands the question to archiveWorksheets instead,
    // which is what main did for every workbook.
    return parsed.declared.length > 0
        ? [...ordered, ...rescued.map((entry) => ({ entry, name: partFallbackName(entry.name) }))]
        : undefined
}

// The fallback, which is what main did for every workbook: the archive's own worksheet parts, in
// entry order, named after their part number. Distinct ENTRIES, so entries sharing a name stay
// distinct here rather than collapsing onto whichever came first.
const archiveWorksheets = (entries: ZipEntry[]): WorkbookSheet[] =>
    worksheetEntries(entries).map((entry) => ({ entry, name: partFallbackName(entry.name) }))

/////////////////////////////////////////////////////////////

// Rebuild with XLSX_LEADING_ENTRIES first, then the resolved worksheets in TAB order, then everything
// else in its original order — minus any further part this reader would dispatch as a worksheet.
// Reports the sheets it laid out so the caller can name and count them positionally. Refuses rather
// than falling back to the original bytes (see DECOMPRESSION BUDGET), and says WHICH refusal: a
// directory we could not read and a rewrite we declined to produce are different facts, and only the
// first is unreachable — the entry-count bail below has a test driving it.
type Reorder = { ok: true; content: Buffer; sheets: WorkbookSheet[] } | { ok: false; reason: string }

const reorderForStreaming = async (buf: Buffer): Promise<Reorder> => {
    const entries = zipEntries(buf)
    // Unreachable: every zipEntries bail is also a budget rejection except a name length overrunning
    // the buffer, which Invariant 2 catches. Failing closed keeps that from being load-bearing.
    if (!entries) return { ok: false, reason: 'xlsx central directory could not be read for streaming' }

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
    if (complete.length >= 0xffff) {
        return { ok: false, reason: `xlsx would rewrite to ${complete.length} entries, at the 0xffff count sentinel` }
    }

    const sheets = (await workbookWorksheets(entries)) ?? archiveWorksheets(entries)

    const leadingRank = XLSX_LEADING_ENTRIES.length
    const rank = (name: string) => {
        const index = XLSX_LEADING_ENTRIES.indexOf(name)
        return index === -1 ? leadingRank : index
    }
    // Three groups rather than one sort, because the middle is ordered by the WORKBOOK and the last
    // DROPS entries rather than placing them.
    const ordered = [
        // Every flag the reader's worksheet branch tests, set before a worksheet is reached.
        ...complete.filter((entry) => rank(entry.name) < leadingRank).sort((a, b) => rank(a.name) - rank(b.name)),
        // Tab order. The resolved ENTRIES, carried, never re-looked-up by name — a name is not a key
        // in a zip, and a lookup here is what re-emitted one entry's bytes once per duplicate name.
        // Both resolvers draw from this archive's own worksheet entries and each returns any given
        // entry at most once, so this writes each of them exactly once.
        ...sheets.map((sheet) => sheet.entry),
        // Everything else — minus every remaining part the reader would dispatch as a worksheet.
        // Orphan sheetN.xml parts no <sheet> references die here, which is both what stops them being
        // emitted as sheets of their own and what makes the Nth emission exactly sheets[N - 1].
        ...complete.filter(
            (entry) => rank(entry.name) === leadingRank && !EXCELJS_WORKSHEET_DISPATCH.test(entry.name)
        ),
    ]

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
    return { ok: true, content: Buffer.concat([...locals, directory, end]), sheets }
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

// A caller's cap may only tighten: absent or NaN falls back to the ceiling, negative clamps to 0,
// fractional floors (a cap is a whole number of chars). The infinities go through the clamp rather
// than the fallback, so the function is monotonic across its whole domain — -Infinity used to mean
// "no cap at all" while -5 meant 0, which is a seam nothing benefits from. Both routes are safe
// either way: this can only ever tighten.
const resolveCap = (requested?: number): number =>
    requested === undefined || Number.isNaN(requested)
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

    // Zip bombs: measure the real decompressed size before exceljs/mammoth touch the bytes. The
    // preflight also decides which refusal this is — over budget skips, malformed fails.
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
                // The same budget withTimeout races on, less a margin, as a value a handler can act
                // on: the race can't cancel work already inside a parser; a handler checking this
                // stops itself. The margin is not decoration — it is what lets that stop be
                // REPORTED. See HANDLER_DEADLINE_MARGIN_MS.
                deadline: Date.now() + HANDLER_TIMEOUT_MS - HANDLER_DEADLINE_MARGIN_MS,
            }),
            HANDLER_TIMEOUT_MS
        )
        // Central cap, so a pathological document can't dump megabytes into S3 and the search index.
        // Handlers that build incrementally overshoot by one unit, so for them this is the final
        // precise trim; docx and html return a full string, so for those it's POST-materialization —
        // peak memory follows the whole document, and hard containment is the host memory limit
        // (see README). Don't split a surrogate pair at the boundary: a lone half serializes as U+FFFD.
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