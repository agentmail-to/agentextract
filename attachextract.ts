// attachextract — text extraction for email attachments.
//
// In scope: text, HTML, PDF, Word (.docx + legacy .doc), and nested emails (.eml). Out of
// scope: images/OCR, spreadsheets (.xlsx), and archives — those return a labeled skip.
// Ships as a separate subpath export ("agentextract/attachments") so the heavy
// parsers stay out of the body extractor's bundle. New formats = new registry entries.

// Converts bytes into text using a specific character encoding 
import iconv from 'iconv-lite'
// Guesses what encoding a chunk of bytes is in (eg. UTF-8, UTF-16, or windows-1252)
import jschardet from 'jschardet'

/////////////////////////////////////////////////////////////
// CONSTANTS (tunable)

// Max: 10 MB - Any attachment bigger than this gets rejected before any processing happens at all
export const MAX_INPUT_BYTES = 10 * 1024 * 1024

// Guard: if a handler takes >= 20 seconds, we stop awaiting it. Note: this can't
// cancel synchronous CPU-bound work — it only stops us waiting on a slow async parse.
export const HANDLER_TIMEOUT_MS = 20_000

// For nested emails (.emls), we only descend one level (eg. a forward of a forward).
export const MAX_NESTING_DEPTH = 1

// "Sniff" the first 8KB of the buffer to see if it looks like text.
const SNIFF_BYTES = 8 * 1024
const SNIFF_TEXT_RATIO = 0.85

// Detecting which encoding a chunk of bytes is in (eg. UTF-8, UTF-16, or windows-1252) 
// Looking at the first 64KB of the buffer.
const DETECT_SAMPLE_BYTES = 64 * 1024
// If the confidence is less than 70%, we don't trust the detection.
const DETECT_MIN_CONFIDENCE = 0.7

// below ~50 avg chars/page, a PDF with some text gets flagged as likely image-heavy 
const PDF_MIN_CHARS_PER_PAGE = 50

// Setup: converts folder-name string into raw bytes, so the code can search for that literal byte sequence inside the zip's raw contents. 
const DOCX_MEDIA_MARKER = Buffer.from('word/media/')
// below 200 chars AND the doc has embedded media -> flagged as likely image-heavy 
const DOCX_MIN_TEXT_CHARS = 200

/////////////////////////////////////////////////////////////
// TYPES

// The status of the extraction.
export type ExtractionStatus =
    | 'extracted' // real text obtained
    | 'extracted_empty' // handler succeeded but produced no text (e.g. scanned PDF) — terminal, valid
    | 'skipped_oversize' // over MAX_INPUT_BYTES, never attempted
    | 'skipped_unsupported_type' // image, spreadsheet, archive, binary, etc.
    | 'failed' // parser threw, timed out, or decode failed 

// Records which signal ended up deciding the route.
export type RoutedBy = 'content-type' | 'extension' | 'sniff' | 'none'

// The formats we can route to. Each maps to exactly one handler.
// text: plain text, csv, calendar, vcard, json, xml, yaml
// html: html
// pdf: pdf
// docx: docx (modern OOXML Word)
// doc: doc (legacy OLE Word, Word 97–2003)
// eml: eml
export type HandlerKind = 'text' | 'html' | 'pdf' | 'docx' | 'doc' | 'eml'

// The input to the extraction.
export interface AttachmentInput {
    content: Buffer // the raw bytes of the attachment
    filename?: string // the name of the attachment
    contentType?: string // the content type of the attachment (eg. text/plain, text/html, application/pdf, ...)
}

// The output of the extraction.
export interface ExtractionResult {
    filename?: string // the name of the attachment
    detectedType?: HandlerKind // the handler that ran; undefined when nothing matched (eg. image, spreadsheet, archive, binary, etc.)
    routedBy: RoutedBy // how we decided what handler to use 
    charset?: string // set when a text handler decoded 
    byteSize: number // the size of the attachment in bytes 
    status: ExtractionStatus // the status of the extraction 
    reason?: string // set on failed / skipped_* 
    extractedText?: string // present string on 'extracted'/'extracted_empty' ('' when empty); undefined on skip/failed — see the field-presence note at the entry point
    children?: ExtractionResult[] // .eml only — one result per inner attachment

    // Image-awareness signals. No OCR, but we flag when a document likely holds text we could not reach (eg. image-heavy). 
    // For future OCR passes.
    lowTextDensity?: boolean // real text obtained, but sparse / some pages had none — likely image-heavy
    pageCount?: number // PDF only — total pages 
    emptyPageCount?: number // PDF only — pages with no extractable text (likely image pages)
}

// The context for a handler.
interface HandlerContext {
    content: Buffer // the raw bytes of the attachment
    filename?: string // the name of the attachment
    charsetHint?: string // from the content-type charset= param
    depth: number // nesting guard for the eml handler
}

// The output of a handler.
interface HandlerOutput {
    text: string // the text of the attachment
    charset?: string // text handlers report the charset they decoded with
    empty?: boolean // handler's own emptiness call; defaults to text.trim() === ''
    children?: ExtractionResult[] // eml only
    lowTextDensity?: boolean // image-awareness signal (pdf/docx)
    pageCount?: number // pdf only
    emptyPageCount?: number // pdf only
}

// Handler takes context and returns an output. 
interface Handler {
    kind: HandlerKind // the type of the handler (eg. text, html, pdf, docx, eml)
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

// Decode the text using the correct charset.
const decodeText = (content: Buffer, hint?: string): { text: string; charset: string } => {
    const charset = resolveCharset(content, hint) // get the correct charset (done above)
    const text = iconv 
        .decode(content, charset) // convert the raw bytes into a JS string using that encoding
        .replace(/^﻿/, '') // strip BOM (cleans the final output)
        .replace(/\r\n?/g, '\n') // normalizes the line endings
    return { text, charset } // return the decoded text and the charset
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
        // NB: message/global is NOT here — it's a full internationalized email (eml handler).
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
    extract: async ({ content, charsetHint }) => decodeText(content, charsetHint),
}

// NOTE: Lazy loading so that a Lambda that only ever sees text attachments never pays to load pdf.js / mammoth (heavier dependencies). 

// Helper function: Shared HTML -> visible-text flattening, 
// used by the html handler and reused by the eml handler for html-only forwarded bodies. 
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
        const { text: decoded, charset } = decodeText(content, charsetHint) 
        return { text: await flattenHtml(decoded), charset } 
    },
}

// PDF handler - pdf
const pdfHandler: Handler = {
    kind: 'pdf',
    contentTypes: ['application/pdf'],
    extensions: ['.pdf'],
    extract: async ({ content }) => {
        const { getDocumentProxy, extractText } = await import('unpdf') // Unpdf is a library that extracts text from PDF
        const pdf = await getDocumentProxy(new Uint8Array(content)) // Converts the raw bytes into format unpdf expects
        const { totalPages, text } = await extractText(pdf, { mergePages: false }) // extract text per page and return the total number of pages and the text
        const pages = text.map((page) => page.trim()) // trim the whitespace 
        const joined = pages.join('\n\n').trim() // join the pages with a blank line 
        const emptyPageCount = pages.filter((page) => page.length === 0).length // count the number of empty pages (no text)
        const chars = pages.reduce((sum, page) => sum + page.length, 0) // count the number of characters in the text (across all pages)
        const empty = joined.length === 0 // true if no text at all 
        const lowTextDensity = !empty && (emptyPageCount > 0 || (totalPages > 0 && chars / totalPages < PDF_MIN_CHARS_PER_PAGE)) // flag as image-heavy if any page has no text, or overall density is low
        return { text: joined, empty, lowTextDensity, pageCount: totalPages, emptyPageCount }
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
        const lowTextDensity = content.includes(DOCX_MEDIA_MARKER) && value.trim().length < DOCX_MIN_TEXT_CHARS // flag as image-heavy if the document has embedded images but little text
        return { text: value, lowTextDensity } // return the text and the low text density
    },
}

// DOC handler. 
// Separate from docx: mammoth only reads the modern OOXML zip, so pre-2007 .doc needs word-extractor, which parses the OLE compound binary.
const docHandler: Handler = {
    kind: 'doc',
    contentTypes: ['application/msword'],
    extensions: ['.doc'],
    extract: async ({ content }) => {
        const { default: WordExtractor } = await import('word-extractor') // parses the legacy OLE .doc binary
        const doc = await new WordExtractor().extract(content) // accepts the raw buffer directly
        return { text: doc.getBody() } // getBody() is the main document text (headers/footers/notes are separate)
    },
}

// EML handler - eml
const emlHandler: Handler = {
    kind: 'eml',
    // message/global is the internationalized (SMTPUTF8) equivalent of message/rfc822 — a
    // full email, not a headers blob, so it parses through mailparser like any other .eml.
    contentTypes: ['message/rfc822', 'message/global'],
    extensions: ['.eml'],
    extract: async ({ content, depth }) => {
        const { simpleParser } = await import('mailparser') // mailparser is a library that parses emails
        const parsed = await simpleParser(content, {
            skipHtmlToText: true, // skip the html to text conversion so our HTML handler owns that path
            skipTextToHtml: true, // skip the text to html conversion so our HTML handler owns that path
            skipImageLinks: true, // skip the image links so we don't include them in the text
            skipTextLinks: true, // skip the text links so we don't include them in the text
        })
        const body = parsed.text?.trim() ? parsed.text : parsed.html ? await flattenHtml(parsed.html) : '' // get the text from the email or the html body
        const text = [parsed.subject, body].filter((part) => part && part.trim().length > 0).join('\n\n') // join the subject and body with a blank line
        // Start building the (optional) array of results for any nested attachments. 
        // Recursively extract the attachments.
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

const REGISTRY: Handler[] = [textHandler, htmlHandler, pdfHandler, docxHandler, docHandler, emlHandler]

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
    // Guard: If RTF returns undefined - text/rtf would originally return the text handler
    // Checked before the text handler on purpose to avoid misclaiming it as plain text and leaking the control words
    (type.includes('rtf') ? undefined :
    // If text/ or text/html, return the html handler if it's text/html, otherwise return the text handler (if none of the above, return undefined)
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
// RTF documents always start with the literal "{\rtf". RTF is ASCII, so looksLikeText would
// otherwise grab a mislabeled/extensionless one as text and leak its control words — detect it
// here to skip instead (see the text/rtf note on findByContentType).
const RTF_MAGIC = Buffer.from('{\\rtf')
// Legacy OLE compound files (Word 97–2003 .doc) start with this fixed 8-byte signature. It is
// NOT unique to Word — legacy .xls / .ppt / .msg share it — so, like the zip→docx case, we only
// claim doc when the extension confirms it; other OLE payloads stay unrouted.
const OLE_MAGIC = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])

// Check if the file starts with the given magic bytes.
const startsWith = (content: Buffer, magic: Buffer): boolean =>
    content.length >= magic.length && content.subarray(0, magic.length).equals(magic)

// Magic bytes are the strongest sniff signal, so they go first. 
// A BOM is next: UTF-16 text is full of NUL bytes, so looksLikeText would reject it as binary 
// A .docx is a zip, but a bare zip could be xlsx/pptx/jar, 
// so we only claim docx when the extension confirms it. 
const sniff = (content: Buffer, ext?: string): HandlerKind | undefined => {
    if (startsWith(content, PDF_MAGIC)) return 'pdf'
    if (startsWith(content, ZIP_MAGIC) && ext === '.docx') return 'docx'
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
const bytesContradictTextClaim = (content: Buffer): boolean =>
    content.length > 0 && !bomCharset(content) && !looksLikeText(content)

// Takes an AttachmentInput and produces the final routing decision. 
// Tries each signal in priority order and stops as soon as one works.
export const detectRoute = (input: AttachmentInput): { kind?: HandlerKind; routedBy: RoutedBy } => {
    const { type } = parseContentType(input.contentType)
    const ext = extensionOf(input.filename)

    const byType = type ? findByContentType(type) : undefined
    const byExt = ext ? findByExtension(ext) : undefined
    const claimed = byType
        ? ({ kind: byType.kind, routedBy: 'content-type' } as const)
        : byExt
          ? ({ kind: byExt.kind, routedBy: 'extension' } as const)
          : undefined

    // Distrust a text-decoding claim contradicted by the bytes (the type/extension lied).
    if (claimed && (claimed.kind === 'text' || claimed.kind === 'html') && bytesContradictTextClaim(input.content)) {
        const sniffed = sniff(input.content, ext)
        return sniffed ? { kind: sniffed, routedBy: 'sniff' } : { routedBy: 'none' }
    }
    if (claimed) return claimed

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

// Main entry point for extracting an attachment. 
// Takes an AttachmentInput and a depth (for nested emails) and returns an ExtractionResult.
const extractAt = async (input: AttachmentInput, depth: number): Promise<ExtractionResult> => {
    const byteSize = input.content.length
    const base = {
        filename: input.filename,
        byteSize,
    }

    // Size gate BEFORE any decode/parse. If the attachment is too large, skip it.
    if (byteSize > MAX_INPUT_BYTES) {
        return {
            ...base,
            routedBy: 'none',
            status: 'skipped_oversize',
            reason: `${byteSize} bytes exceeds ${MAX_INPUT_BYTES}`,
        }
    }

    const { type, charset: charsetHint } = parseContentType(input.contentType) // declared type + charset, parsed once
    const { kind, routedBy } = detectRoute(input) // figure out what kind of file this is, and confidence level
    const handler = kind ? findHandler(kind) : undefined // look up the actual handler for that kind, if we found one

    // No handler found for this kind — unsupported or unrecognized format, skip it.
    if (!handler) { 
        return {
            ...base,
            routedBy,
            status: 'skipped_unsupported_type',
            reason: type ? `unsupported type ${type}` : 'unrecognized attachment',
        }
    }

    // Run the handler safely — catch errors/timeouts instead of crashing, and build the final result.
    try {
        const output = await withTimeout( // Enforce a time limit so a slow/hung handler doesn't block forever.
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
            // A handler that ran is a completed extraction: emit a present string even when empty
            // ('' on extracted_empty), never undefined. This mirrors the body extractor's contract
            // (extractEmailBody returns '' for an empty result) so a downstream field-presence
            // cache — "extracted field present ⇒ cached", no version stamp — treats a legitimately
            // empty attachment as cached instead of re-extracting it forever. Skip/failed statuses
            // below leave extractedText undefined, which correctly reads as "not cached, re-attempt".
            extractedText: isEmpty ? '' : output.text,
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
// Public entry point — starts the recursion at depth 0.
export const extractAttachment = (input: AttachmentInput): Promise<ExtractionResult> => extractAt(input, 0)