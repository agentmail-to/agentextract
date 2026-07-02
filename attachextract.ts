// attachextract — text extraction for email attachments.
//
// Sibling to agentextract (which strips quoted history from an email BODY). This
// entry takes a text-bearing attachment — plain text, CSV, calendar (.ics), the
// message/*-header types, and (in later steps) HTML, PDF, Word, and nested email —
// and returns its plain text. Because the heavier handlers rely on real parsers
// (unpdf, mammoth, html-to-text, mailparser), this ships as a SEPARATE subpath
// export ("agentextract/attachments") so those deps stay out of the body
// extractor's bundle.
//
// Design: a type-keyed handler REGISTRY. Handlers are pure `bytes -> text`; ALL
// safety (size cap, timeout, fail-open) and result assembly live in the entry
// point, so a new format is a new registry entry, never a rewrite, and every
// handler is trivially testable on its own.
//
// Scope (from a ~7.7k-attachment corpus): images (~61%) are out of scope — no OCR;
// the text-extractable ~37% is dominated by PDF, with docx, nested email, csv, and
// calendar next. Spreadsheets (.xlsx), legacy binary (.doc/.xls), presentations,
// and archives are out of scope and return a labeled skip.

import iconv from 'iconv-lite'
import jschardet from 'jschardet'

/////////////////////////////////////////////////////////////
// CONSTANTS (tunable)

// Bump when extraction logic changes so a downstream cache can invalidate on it.
export const EXTRACTION_VERSION = '1'

// Enforced BEFORE any decode/parse. Corpus p100 for text-extractable types is a
// 20 MB PDF; 10 MB skips ~0.4% (a dozen monster PDFs) and bounds Lambda memory.
export const MAX_INPUT_BYTES = 10 * 1024 * 1024

// Soft guard around library handlers. NOTE: Promise.race cannot cancel synchronous
// CPU-bound work — it stops us *awaiting* a slow parse (unpdf/mammoth are mostly
// async, so it mostly holds); true isolation is a worker-thread concern, out of
// scope here. Used by the library handlers in later steps.
export const HANDLER_TIMEOUT_MS = 20_000

// A nested email (.eml) is descended one level; an eml inside an eml is not.
export const MAX_NESTING_DEPTH = 1

// "Looks like text" sniff reads the head of the buffer only.
const SNIFF_BYTES = 8 * 1024
const SNIFF_TEXT_RATIO = 0.85

// Charset detection is bounded to the head — jschardet on a full 10 MB buffer is
// wasteful, and encoding is uniform across a file in practice.
const DETECT_SAMPLE_BYTES = 64 * 1024
const DETECT_MIN_CONFIDENCE = 0.7

// Below this average characters/page, a PDF that still has *some* text is flagged
// lowTextDensity (likely image-heavy) — not discarded. A PDF with zero text is
// extracted_empty instead. Both are OCR candidates downstream.
const PDF_MIN_CHARS_PER_PAGE = 50

// A .docx is a zip; embedded images live under word/media/. extractRawText drops them,
// so if a docx carries media but little text, its content is probably in the images.
const DOCX_MEDIA_MARKER = Buffer.from('word/media/')
const DOCX_MIN_TEXT_CHARS = 200

/////////////////////////////////////////////////////////////
// TYPES

export type ExtractionStatus =
    | 'extracted' // real text obtained
    | 'extracted_empty' // handler succeeded but produced no text (e.g. scanned PDF) — terminal, valid
    | 'skipped_oversize' // over MAX_INPUT_BYTES, never attempted
    | 'skipped_unsupported_type' // image, spreadsheet, archive, binary, etc.
    | 'failed' // parser threw, timed out, or decode failed 

// Which signal decided the route — recorded so downstream can see how confident a
// route was ('sniff' is the weakest).
export type RoutedBy = 'content-type' | 'extension' | 'sniff' | 'none'

// The formats we can route to. Each maps to exactly one handler.
export type HandlerKind = 'text' | 'html' | 'pdf' | 'docx' | 'eml'

export interface AttachmentInput {
    content: Buffer
    filename?: string
    contentType?: string
}

export interface ExtractionResult {
    filename?: string
    detectedType?: HandlerKind // the handler that ran; undefined when nothing matched
    routedBy: RoutedBy
    charset?: string // set when a text handler decoded
    byteSize: number
    status: ExtractionStatus
    reason?: string // set on failed / skipped_*
    extractedText?: string // set on 'extracted'
    extractionVersion: string
    children?: ExtractionResult[] // .eml only — one result per inner attachment (later step)
    // Image-awareness signals. We never read the images themselves (no OCR), but we flag
    // when a document likely holds text we could not reach, so a future OCR pass can target
    // exactly these attachments/pages. OCR-candidate rule downstream: status ===
    // 'extracted_empty' (image-only) OR lowTextDensity === true (mixed/sparse).
    lowTextDensity?: boolean // real text obtained, but sparse / some pages had none — likely image-heavy
    pageCount?: number // PDF only — total pages
    emptyPageCount?: number // PDF only — pages with no extractable text (likely image pages)
}

// A handler does ONLY bytes -> text. Safety and assembly are the entry point's job.
interface HandlerContext {
    content: Buffer
    filename?: string
    charsetHint?: string // from the content-type charset= param
    depth: number // nesting guard for the eml handler
}

interface HandlerOutput {
    text: string
    charset?: string // text handlers report the charset they decoded with
    empty?: boolean // handler's own emptiness call; defaults to text.trim() === ''
    children?: ExtractionResult[] // eml only (later step)
    lowTextDensity?: boolean // image-awareness signal (pdf/docx)
    pageCount?: number // pdf only
    emptyPageCount?: number // pdf only
}

interface Handler {
    kind: HandlerKind
    contentTypes: string[] // exact, lowercased, param-stripped
    extensions: string[] // with leading dot, lowercased
    extract: (ctx: HandlerContext) => Promise<HandlerOutput>
}

/////////////////////////////////////////////////////////////
// CHARSET-CORRECT DECODING (direct-text handlers)
//
// Buffer.toString('utf8') silently turns windows-1252 / ISO-8859-x / Shift_JIS into
// mojibake. Resolve the charset first, then decode via iconv-lite (which never
// throws — invalid bytes become the replacement char, not an exception).

// A byte-order mark is an unambiguous encoding signal. Reading it directly (rather
// than leaving it to jschardet) makes UTF-16 / UTF-8-BOM files decode correctly even
// when the statistical detector is unsure — or if jschardet is ever dropped.
const bomCharset = (content: Buffer): string | undefined => {
    if (content.length >= 3 && content[0] === 0xef && content[1] === 0xbb && content[2] === 0xbf) return 'utf-8'
    if (content.length >= 2 && content[0] === 0xff && content[1] === 0xfe) return 'utf-16le'
    if (content.length >= 2 && content[0] === 0xfe && content[1] === 0xff) return 'utf-16be'
    return undefined
}

const resolveCharset = (content: Buffer, hint?: string): string => {
    // 1. Explicit charset from the Content-Type, if iconv-lite knows it.
    if (hint && iconv.encodingExists(hint)) return hint
    // 2. A BOM is definitive — trust it over statistical detection.
    const bom = bomCharset(content)
    if (bom) return bom
    // 3. Statistical detection on the head, gated on confidence.
    const detected = jschardet.detect(content.subarray(0, DETECT_SAMPLE_BYTES))
    if (
        detected &&
        detected.encoding &&
        detected.confidence > DETECT_MIN_CONFIDENCE &&
        iconv.encodingExists(detected.encoding)
    ) {
        return detected.encoding
    }
    // 4. utf-8, then 5. latin1 as the guaranteed-decodable floor.
    return iconv.encodingExists('utf-8') ? 'utf-8' : 'latin1'
}

const decodeText = (content: Buffer, hint?: string): { text: string; charset: string } => {
    const charset = resolveCharset(content, hint)
    const text = iconv
        .decode(content, charset)
        .replace(/^﻿/, '') // strip BOM
        .replace(/\r\n?/g, '\n') // normalize CRLF / CR -> LF
    return { text, charset }
}

/////////////////////////////////////////////////////////////
// HANDLER REGISTRY
//
// Step 1 registers only the direct-text handler. HTML / PDF / DOCX / EML handlers
// (and their content-types, extensions, and magic-byte sniffing) are added in the
// following steps — additively, never rewriting the router below.

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
        // Header-only / partial-message MIME types (DSNs, forwards) — text, not full email.
        'message/global-headers',
        'message/global',
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
    extract: async ({ content, charsetHint }) => decodeText(content, charsetHint),
}

// The library handlers below lazy-load their parser via dynamic import() so a Lambda
// that only ever sees text attachments never pays to load pdf.js / mammoth, and so
// the ESM-only `unpdf` loads cleanly from this CommonJS module.

// Shared HTML -> visible-text flattening, used by the html handler and reused by the
// eml handler for html-only forwarded bodies. Never regex-strips tags — html-to-text
// walks the DOM; script/style/img are dropped and hrefs ignored.
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

const htmlHandler: Handler = {
    kind: 'html',
    contentTypes: ['text/html', 'application/xhtml+xml'],
    extensions: ['.html', '.htm', '.xhtml'],
    // HTML bytes can be non-utf8 too, so decode with the same charset logic first.
    extract: async ({ content, charsetHint }) => {
        const { text: decoded, charset } = decodeText(content, charsetHint)
        return { text: await flattenHtml(decoded), charset }
    },
}

const pdfHandler: Handler = {
    kind: 'pdf',
    contentTypes: ['application/pdf'],
    extensions: ['.pdf'],
    // Per-page (mergePages: false) so we can distinguish a truly text-less PDF
    // (extracted_empty) from one with real text plus image pages we didn't read.
    extract: async ({ content }) => {
        const { getDocumentProxy, extractText } = await import('unpdf')
        const pdf = await getDocumentProxy(new Uint8Array(content))
        const { totalPages, text } = await extractText(pdf, { mergePages: false })
        const pages = text.map((page) => page.trim())
        const joined = pages.join('\n\n').trim()
        const emptyPageCount = pages.filter((page) => page.length === 0).length
        const chars = pages.reduce((sum, page) => sum + page.length, 0)
        // Zero text -> image-only/scanned -> extracted_empty. Otherwise keep the real
        // text, but flag when some pages had no text or the overall density is low —
        // those are likely images a future OCR pass should target.
        const empty = joined.length === 0
        const lowTextDensity = !empty && (emptyPageCount > 0 || (totalPages > 0 && chars / totalPages < PDF_MIN_CHARS_PER_PAGE))
        return { text: joined, empty, lowTextDensity, pageCount: totalPages, emptyPageCount }
    },
}

const docxHandler: Handler = {
    kind: 'docx',
    contentTypes: ['application/vnd.openxmlformats-officedocument.wordprocessingml.document'],
    extensions: ['.docx'],
    // extractRawText only (never the HTML path) — its blank-line paragraph separation
    // is the structure we keep. Empty/whitespace-only falls to the default trim rule.
    extract: async ({ content }) => {
        const { default: mammoth } = await import('mammoth')
        const { value } = await mammoth.extractRawText({ buffer: content })
        // Has embedded images but little text -> content is probably in the images.
        const lowTextDensity = content.includes(DOCX_MEDIA_MARKER) && value.trim().length < DOCX_MIN_TEXT_CHARS
        return { text: value, lowTextDensity }
    },
}

const emlHandler: Handler = {
    kind: 'eml',
    contentTypes: ['message/rfc822'],
    extensions: ['.eml'],
    // A forwarded/attached email. Take its subject + text body (html body flattened via
    // the shared HTML path when there's no text part), then recurse each inner attachment
    // back through the pipeline. MAX_NESTING_DEPTH bounds this: an email nested inside an
    // email is body-read, but its own attachments are not descended into (email-bomb guard).
    extract: async ({ content, depth }) => {
        const { simpleParser } = await import('mailparser')
        // Skip mailparser's own html<->text synthesis so our HTML handler owns that path.
        const parsed = await simpleParser(content, {
            skipHtmlToText: true,
            skipTextToHtml: true,
            skipImageLinks: true,
            skipTextLinks: true,
        })

        const body = parsed.text?.trim() ? parsed.text : parsed.html ? await flattenHtml(parsed.html) : ''
        const text = [parsed.subject, body].filter((part) => part && part.trim().length > 0).join('\n\n')

        // Each inner attachment is a first-class attachment: its own size cap, routing, and
        // fail-open handling, via the same entry point at depth + 1.
        const children =
            depth < MAX_NESTING_DEPTH && parsed.attachments.length > 0
                ? await Promise.all(
                      parsed.attachments.map((attachment) =>
                          extractAt(
                              {
                                  content: attachment.content,
                                  filename: attachment.filename,
                                  contentType: attachment.contentType,
                              },
                              depth + 1
                          )
                      )
                  )
                : undefined

        return { text, children }
    },
}

const REGISTRY: Handler[] = [textHandler, htmlHandler, pdfHandler, docxHandler, emlHandler]

const findHandler = (kind: HandlerKind): Handler | undefined => REGISTRY.find((h) => h.kind === kind)

/////////////////////////////////////////////////////////////
// ROUTING — never trust one signal

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

const extensionOf = (filename?: string): string | undefined => {
    if (!filename) return undefined
    const match = /(\.[a-z0-9]+)$/i.exec(filename.trim())
    return match ? match[1].toLowerCase() : undefined
}

const findByContentType = (type: string): Handler | undefined =>
    REGISTRY.find((h) => h.contentTypes.includes(type)) ??
    // Any other text/* subtype we didn't enumerate is still plain text — except
    // text/html, which has its own handler (added later) and must not decode as raw text.
    (type.startsWith('text/') && type !== 'text/html' ? textHandler : undefined)

const findByExtension = (ext: string): Handler | undefined => REGISTRY.find((h) => h.extensions.includes(ext))

// Only sniff when the declared type is absent or a generic "unknown binary" label —
// a real media type we don't support (image/png, application/zip) is a deliberate
// skip, not a sniff candidate.
const OCTET_STREAM_TYPES = new Set([
    'application/octet-stream',
    'binary/octet-stream',
    'application/download',
    'application/unknown',
])
const shouldSniff = (type?: string): boolean => !type || OCTET_STREAM_TYPES.has(type)

// "Looks like text": no NUL byte, and >=85% of the head is printable/whitespace.
// High bytes (>=128) count as text-ish (utf-8 / latin1 content); C0 control chars
// other than tab/LF/CR do not.
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
const ZIP_MAGIC = Buffer.from([0x50, 0x4b, 0x03, 0x04]) // "PK\x03\x04"
const startsWith = (content: Buffer, magic: Buffer): boolean =>
    content.length >= magic.length && content.subarray(0, magic.length).equals(magic)

// Magic bytes are the strongest sniff signal, so they go first. A BOM is next: UTF-16
// text is full of NUL bytes, so looksLikeText would reject it as binary — the BOM
// overrides that. A .docx is a zip, but a bare zip could be xlsx/pptx/jar, so we only
// claim docx when the extension confirms it (already caught by extension routing before
// we sniff — kept as a defensive backstop).
const sniff = (content: Buffer, ext?: string): HandlerKind | undefined => {
    if (startsWith(content, PDF_MAGIC)) return 'pdf'
    if (startsWith(content, ZIP_MAGIC) && ext === '.docx') return 'docx'
    if (bomCharset(content) || looksLikeText(content)) return 'text'
    return undefined
}

export const detectRoute = (input: AttachmentInput): { kind?: HandlerKind; routedBy: RoutedBy } => {
    const { type } = parseContentType(input.contentType)
    const ext = extensionOf(input.filename)

    const byType = type ? findByContentType(type) : undefined
    if (byType) return { kind: byType.kind, routedBy: 'content-type' }

    const byExt = ext ? findByExtension(ext) : undefined
    if (byExt) return { kind: byExt.kind, routedBy: 'extension' }

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

// Reject once ms elapses. Guards slow *async* handlers (see HANDLER_TIMEOUT_MS note).
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

const errorMessage = (error: unknown): string => (error instanceof Error ? error.message : String(error))

/////////////////////////////////////////////////////////////
// ENTRY POINT
//
// bytes + metadata -> ExtractionResult. Pure and fail-open: no network, no
// filesystem, no throws. Every non-ideal outcome is a labeled status, never a gap.

const extractAt = async (input: AttachmentInput, depth: number): Promise<ExtractionResult> => {
    const byteSize = input.content.length
    const base = {
        filename: input.filename,
        byteSize,
        extractionVersion: EXTRACTION_VERSION,
    }

    // Size gate BEFORE any decode/parse.
    if (byteSize > MAX_INPUT_BYTES) {
        return {
            ...base,
            routedBy: 'none',
            status: 'skipped_oversize',
            reason: `${byteSize} bytes exceeds ${MAX_INPUT_BYTES}`,
        }
    }

    const { charset: charsetHint } = parseContentType(input.contentType)
    const { kind, routedBy } = detectRoute(input)
    const handler = kind ? findHandler(kind) : undefined

    if (!handler) {
        return {
            ...base,
            routedBy,
            status: 'skipped_unsupported_type',
            reason: parseContentType(input.contentType).type
                ? `unsupported type ${parseContentType(input.contentType).type}`
                : 'unrecognized attachment',
        }
    }

    try {
        const output = await withTimeout(
            handler.extract({ content: input.content, filename: input.filename, charsetHint, depth }),
            HANDLER_TIMEOUT_MS
        )
        const isEmpty = output.empty ?? output.text.trim().length === 0
        return {
            ...base,
            detectedType: handler.kind,
            routedBy,
            charset: output.charset,
            status: isEmpty ? 'extracted_empty' : 'extracted',
            extractedText: isEmpty ? undefined : output.text,
            children: output.children,
            // Image-awareness signals pass through unchanged; page counts are reported
            // even on extracted_empty so an OCR pass knows how many pages to render.
            lowTextDensity: output.lowTextDensity,
            pageCount: output.pageCount,
            emptyPageCount: output.emptyPageCount,
        }
    } catch (error) {
        // Attacker-controlled bytes: a throw or timeout is an expected event, captured as
        // 'failed' with the reason — never propagated to crash the caller.
        return {
            ...base,
            detectedType: handler.kind,
            routedBy,
            status: 'failed',
            reason: errorMessage(error),
        }
    }
}

export const extractAttachment = (input: AttachmentInput): Promise<ExtractionResult> => extractAt(input, 0)
