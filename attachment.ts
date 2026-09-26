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

// Blocks a handler concatenates into one extraction — PDF pages, XLSX sheets — are joined by this.
// The incremental length accounting has to charge for it BEFORE the join, so both must read it from
// here: a separator that disagrees with its own charged width makes the cap off by that difference.
const BLOCK_SEPARATOR = '\n\n'

// A PDF can declare an enormous page count. What this buys is DETERMINISM, not work-bounding: the
// page loop already checks the deadline before every page, so a pathological page count is bounded
// in time with or without this. Without it, though, where a huge sparse PDF stops would depend on
// how fast the host is, and the same attachment would extract differently on a re-index. A declared
// page count is a fixed property of the file, so cutting on it cuts at the same place every time.
// Only the sparse case can reach it: at even 125 characters per page MAX_OUTPUT_CHARS binds first,
// so a document stopped here is one carrying almost no text at all. 2000 is a backstop rather than
// a tuned threshold — an email attachment past a few hundred pages is already an outlier, and the
// suite is indifferent to raising it (it passes with this at 1e9), so it is sized for headroom.
export const MAX_PDF_PAGES = 2000

// Max ACTUAL uncompressed size of an OOXML zip, measured by inflating it — the declared size is
// attacker-controlled (see DECOMPRESSION BUDGET). MAX_INPUT_BYTES only bounds the COMPRESSED size.
// Parsers build a model on top (~8x for a dense sheet), so stay well under the 1024 MB memory floor.
export const MAX_UNCOMPRESSED_BYTES = 50 * 1024 * 1024

// Ceiling on a <v> shared-string index before the scan stops reading digits. Derived, not picked: the
// smallest legal <si> record ('<si><t></t></si>') is 16 bytes, so MAX_UNCOMPRESSED_BYTES bounds a
// workbook at ~3.3M entries even if the shared-string table were the only part in it. Any index past
// this cannot resolve, so the digits after it carry no information and only serve to make the scan
// retain an attacker-sized numeric string. Rounded up to 10M to stay clearly above the real bound
// rather than tracking it exactly — the point is to stop unbounded accumulation, not to be tight.
const MAX_SHARED_STRING_INDEX = 10_000_000

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

// Granularity at which an already-resident buffer is fed to a parser — XLSX control XML and a STORED
// .docx part both come this way. Only the cap and deadline checks BETWEEN slices can stop a parse,
// so this is how long a stop can be deferred: a whole buffer handed over at once cannot be
// interrupted at all, and one character at a time pays per-slice overhead on every byte. 16 KB is
// the compromise, and is chosen to match what zlib pushes out of createInflateRaw so a stored part
// and a deflated one arrive in comparable pieces rather than on two different response latencies.
// Pinned from BOTH sides by 'checks the deadline between chunks of a stored document part', which
// sizes its fixture so that the first slice holds the prefix and 5,000 paragraphs do not fit in one:
// the suite fails at 64 and again at 1 GB. That test states 16 KiB in prose rather than importing
// this, so retuning it here surfaces as an unrelated-looking deadline assertion — change both.
const STREAM_SLICE_UNITS = 16 * 1024

// saxes resolves namespace prefixes by scanning the open-tag stack, so attacker-controlled nesting
// makes its xmlns mode quadratic even when the XML is tiny. Twenty nested Word tables legitimately
// reach depth 65, so the ceiling must leave real structure room while still bounding a synchronous
// parser call the between-chunk deadline cannot interrupt.
const MAX_XML_NESTING_DEPTH = 256

const assertXmlDepth = (depth: number): void => {
    if (depth > MAX_XML_NESTING_DEPTH) {
        throw new ExtractionFailure('malformed')
    }
}

// How much of the head the three byte-shape tests read: the printable-ratio sniff (looksLikeText),
// the utf-8 NUL scan, and the utf-16 well-formedness check. Shared deliberately — each one is
// deciding the same question, "do the opening bytes agree with a text claim", and a file that
// disagrees does so immediately. 8 KB because every format these must separate is distinguishable
// far inside it: a binary file puts a NUL, an unpaired surrogate, or a non-printable byte in its
// first few hundred, so raising this buys no accuracy and lowering it starts missing headers that
// pad with ASCII. The one place it is more than a budget is the utf-16 check's truncation branch
// (see isWellFormedUtf16): a trailing high surrogate at the sample edge means "out of view" for a
// longer file and "genuinely unpaired" for a shorter one, so the two readings are split at exactly
// this boundary. Changing this value moves that line — it is not only a performance knob.
const SNIFF_BYTES = 8 * 1024
// Share of the sampled bytes that must be printable for the bytes to pass as text. Not a tuned
// number and does not need to be: the distributions are nowhere near it. Real text is ~100%
// printable (the exceptions are stray control bytes, a handful per file), and binary that gets this
// far is ~30-50% by the structure of byte values, since only 95 of 256 are printable ASCII. Anything
// from roughly 0.6 to 0.95 separates them identically; 0.85 leaves room for a text file carrying
// some control noise without letting through anything with real binary density. Worth knowing when
// changing it: the NUL check above is doing the discrimination the suite actually exercises, and no
// test distinguishes this value — it passes at 0, 0.5, 0.7 and 0.95, and fails only when set above 1
// (unreachable, so every input is binary). A file that is, say, 60% printable is untested territory.
const SNIFF_TEXT_RATIO = 0.85

// Charset detection samples the head. 64 KB rather than SNIFF_BYTES because this asks a harder
// question than "is it text": telling the legacy single-byte encodings apart is a statistical
// judgement over character frequencies, and jschardet sharpens with more of them. Not currently
// pinned from either side — the suite passes with this at 4 bytes and at 1 GB — so treat it as a
// sampling budget rather than a tuned threshold.
const DETECT_SAMPLE_BYTES = 64 * 1024

/////////////////////////////////////////////////////////////
// TYPES

// The two refusals differ by whose fault it is, so a caller can branch on them: `skipped` is a file
// we chose not to read, `failed` is one we couldn't. Never collapse them.
// WHY a file was refused, as a value rather than a sentence. `reason` used to be free text that
// interpolated numbers the caller already had — its own byte count, its own content type, constants
// this module exports — so it read like a diagnostic while carrying nothing a consumer could branch
// on without a regex. These are what a consumer actually needs to tell apart, and each one maps to a
// different thing to DO about it.
export type ExtractionReason =
    // --- skipped: intact, and declined ---
    | 'too-large' // over MAX_INPUT_BYTES, before any decode or parse
    | 'expands-too-large' // over MAX_UNCOMPRESSED_BYTES once actually inflated; also the zip-bomb signal
    | 'unsupported-format' // a type we recognize and do not handle
    | 'unrecognized' // nothing — type, extension or bytes — identified it
    | 'password-protected' // readable bytes, locked content: ask the sender, do not retry
    | 'unsupported-zip-feature' // ZIP64, or a compression method this reader does not implement
    // --- failed: we could not read it ---
    | 'malformed' // the bytes are broken and a parser rejected them. Never retryable
    | 'wrong-document-shape' // parses, but is not the document it claims to be (no w:body, no main part)
    | 'timed-out' // ran out of time. The ONLY retryable failure here, which is the point of naming it
    // --- failed: our fault, not the file's ---
    // An invariant of ours tripped, or a pinned dependency moved under us. Worth separating because
    // a file-shaped failure is routine and this is not: a spike of these should page someone rather
    // than feed a retry loop, and no amount of retrying or re-sending the attachment will help.
    | 'internal'

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

// Why an `extracted` result carries no text. Set ONLY when `status` is 'extracted' and `extraction`
// is absent, because that is the one outcome a caller could not otherwise interpret: before this, a
// scanned page, a zero-byte file and a whitespace-only file were the same result object, so "we
// found nothing" and "there is nothing" were indistinguishable, and no consumer could decide
// whether OCR was worth trying.
export type EmptyReason =
    // Read, and genuinely holds no text. Nothing else will get more out of it.
    | 'no-text-content'
    // HAS content, but none of it is text — a scan, a photographed page, a deck of images. Text
    // extraction is the wrong tool here and OCR is the next step. Reported only where the format
    // lets us PROVE the difference, never guessed.
    | 'no-text-layer'

export interface ExtractionResult {
    status: ExtractionStatus
    extraction?: string // omitted entirely (never '') when the handler produced no text
    reason?: ExtractionReason // set on skipped / failed, and never on extracted
    // Whether the document continues past `extraction`. Set on `extracted` only — skipped/failed
    // have no text to have cut. Independent of `trailer`, so a consumer never parses the text for it.
    truncated?: boolean
    // Set on `extracted` with no `extraction`, and never otherwise — so exactly one of `extraction`
    // and `emptyReason` is present on every successful result.
    emptyReason?: EmptyReason
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
    // Set by a handler that can PROVE the document holds non-text content; the entry point turns it
    // into `emptyReason: 'no-text-layer'` when the extraction came out empty. Proof, not inference —
    // a handler that cannot tell an image-only document from an empty one leaves this unset, and the
    // result says 'no-text-content', which is the honest answer for a format that cannot tell.
    hasNonTextContent?: boolean
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

// Which utf-16 flavour an explicit or resolved charset names, if any.
const claimsUtf16 = (charset?: string): 'utf-16' | 'utf-16le' | 'utf-16be' | undefined => {
    const normalized = charset?.trim().toLowerCase().replace(/_/g, '-')
    if (normalized === 'utf-16' || normalized === 'utf16') return 'utf-16'
    if (normalized === 'utf-16le' || normalized === 'utf16le' || normalized === 'ucs-2' || normalized === 'ucs2') {
        return 'utf-16le'
    }
    if (normalized === 'utf-16be' || normalized === 'utf16be') return 'utf-16be'
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
    // #4. Statistical detection on the head, DELIBERATELY UNGATED on jschardet's confidence score.
    //     Step 2 ruled out utf-8, so these are single-byte legacy bytes: decoding them as utf-8 would
    //     turn every high byte into an irreversible U+FFFD, and a low-confidence guess still beats
    //     that. A confidence gate used to sit here and could not ever change the answer — the branch
    //     below it returned the same `detected.encoding` the gate had just rejected, so every value
    //     of the threshold produced identical output. Re-adding one is not a tightening: it means
    //     choosing latin1 over a plausible guess, which is a behaviour change and needs its own case.
    const detected = jschardet.detect(content.subarray(0, DETECT_SAMPLE_BYTES))
    if (detected?.encoding && iconv.encodingExists(detected.encoding)) return detected.encoding
    // #5. Floor. latin1 maps every byte, so it is wrong-but-reversible rather than lossy.
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
    // legitimately contain U+FFFD, and a BOM or explicitly declared UTF-16 charset makes the decode
    // definitive — re-decoding any of them as latin1 would corrupt genuine text. A detector guess is
    // intentionally not exempt: if guessed UTF-16 produces U+FFFD, latin1 remains the lossless floor.
    if (text.includes('�') && !isUtf8(content) && !bomCharset(content) && !claimsUtf16(hint)) {
        return iconv.decode(content, 'latin1').replace(/\r\n?/g, '\n')
    }
    return text
}

// A w:-namespaced attribute, resolved by URI rather than by literal prefix, for the same reason
// OOXML_PREFIXES exists. Module scope because the part shapes above read it too.
type SaxesAttributes = Record<string, string | { uri?: string; local?: string; value: string }>
const wordAttribute = (attributes: SaxesAttributes, local: string): string | undefined => {
    for (const attribute of Object.values(attributes)) {
        if (
            typeof attribute !== 'string' &&
            attribute.local === local &&
            attribute.uri !== undefined &&
            OOXML_PREFIXES[attribute.uri] === 'w'
        ) {
            return attribute.value
        }
    }
    return undefined
}

// WHERE a part keeps its block-level content. The main document buries it one level down
// (w:document > w:body); a header or footer IS its own container; footnotes, endnotes and comments
// hold a repeating wrapper per item. Everything below that container — paragraphs, runs, tables —
// is identical vocabulary across all five, which is why one reader serves them with only this
// varying. The main-document shape is expressed here exactly as it was hard-coded before, so the
// path this file has always taken is unchanged rather than re-derived.
interface DocxPartShape {
    root: string // the document element
    content: string // element that opens a content region; === root when the root holds it directly
    contentDepth: number // stack.length at which `content` is expected, so a nested namesake cannot open one
    repeats: boolean // may open more than once (one w:footnote per note), rather than exactly once
    // A footnotes/endnotes part always carries the separator and continuation-separator notes Word
    // uses to draw the rule above the note area. They are structure, not content, and emitting them
    // puts a stray empty paragraph in the output of every document that has a single footnote.
    skipItem?: (attributes: SaxesAttributes) => boolean
}

const DOCX_SEPARATOR_NOTE_TYPES = new Set(['separator', 'continuationSeparator'])
const isSeparatorNote = (attributes: SaxesAttributes): boolean => {
    const type = wordAttribute(attributes, 'type')
    return type !== undefined && DOCX_SEPARATOR_NOTE_TYPES.has(type)
}

const DOCX_MAIN_SHAPE: DocxPartShape = { root: 'w:document', content: 'w:body', contentDepth: 2, repeats: false }
const DOCX_HEADER_SHAPE: DocxPartShape = { root: 'w:hdr', content: 'w:hdr', contentDepth: 1, repeats: false }
const DOCX_FOOTER_SHAPE: DocxPartShape = { root: 'w:ftr', content: 'w:ftr', contentDepth: 1, repeats: false }
const DOCX_FOOTNOTES_SHAPE: DocxPartShape = {
    root: 'w:footnotes',
    content: 'w:footnote',
    contentDepth: 2,
    repeats: true,
    skipItem: isSeparatorNote,
}
const DOCX_ENDNOTES_SHAPE: DocxPartShape = {
    root: 'w:endnotes',
    content: 'w:endnote',
    contentDepth: 2,
    repeats: true,
    skipItem: isSeparatorNote,
}
const DOCX_COMMENTS_SHAPE: DocxPartShape = { root: 'w:comments', content: 'w:comment', contentDepth: 2, repeats: true }

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
        try {
            // Iterate pages ourselves — unpdf's extractText parses EVERY page up front, so a pathological
            // page count runs unbounded. Bounds our accumulation and the pages parsed, NOT pdf.js's
            // per-page decompression (no hook exists); that residual is the host memory limit's job.
            const pageCount = Math.min(pdf.numPages, MAX_PDF_PAGES)
            const pages: string[] = []
            let length = 0
            let truncated = false
            let pagesRead = 0 // pages whose text content we actually asked for, deadline stops excluded
            for (let n = 1; n <= pageCount; n++) {
                // This loop awaits per page, so the deadline is enforceable here in a way withTimeout's
                // race is not. Before the fetch: stopping is only useful if it precedes the work.
                if (Date.now() > deadline) {
                    truncated = true
                    break
                }
                const page = await pdf.getPage(n)
                const { items } = await page.getTextContent()
                pagesRead++
                // Replicates unpdf's per-page join: str, plus a newline on hasEOL.
                const pageText = (items as Array<{ str?: string; hasEOL?: boolean }>)
                    .filter((item) => item.str != null)
                    .map((item) => (item.str ?? '') + (item.hasEOL ? '\n' : ''))
                    .join('')
                    .trim()
                if (pageText) {
                    length += pageText.length + (pages.length === 0 ? 0 : BLOCK_SEPARATOR.length) // only BETWEEN pages
                    pages.push(pageText)
                    if (length > maxOutputChars) {
                        truncated = true
                        break // one page of overshoot, trimmed centrally
                    }
                }
            }
            // Pages past the ceiling are text we never read.
            if (pageCount < pdf.numPages) truncated = true
            const joined = pages.join(BLOCK_SEPARATOR).trim()
            // A PDF that HAS pages and yielded no text from any of them is the canonical scan: the
            // pages carry images, and pdf.js found no text operators to report. That is provable
            // here in a way it is not in most formats, and it is the difference between "send this
            // to OCR" and "there was nothing to read". Deliberately requires a page we actually
            // visited — a zero-page PDF, or one whose pages were all skipped by the deadline, has
            // shown us nothing to draw a conclusion from.
            return { text: joined, empty: joined.length === 0, truncated, hasNonTextContent: pagesRead > 0 }
        } finally {
            // getDocumentProxy exposes pdf.js's loading task; release its worker handler, document
            // bytes, and page cache on success, early stop, and errors alike. Cleanup is best-effort:
            // its rejection must not discard extracted text or replace the primary parse error.
            try {
                await pdf.loadingTask?.destroy?.()
            } catch {
                // There is no diagnostics channel in the extraction result; preserve the real result.
            }
        }
    },
}

// DOCX — modern OOXML Word. Streams word/document.xml; see DOCX STREAMING READER for the machinery,
// and for why the output contract is mammoth's, reproduced rather than invented.
// The parts a .docx keeps its text in, beyond word/document.xml. All five carry the same
// paragraph/run vocabulary as the body, differing only in the container the shape names — which is
// why one reader reads all of them and this table is the entire addition.
//
// ORDER IS THE CAP'S PRIORITY ORDER, not the document's reading order, which cannot be
// reconstructed anyway: a footnote's body lives here while its reference sits inline in the body,
// and a header is repeated per section rather than positioned once. Body first, then the notes a
// reader would follow, then the margins — so a document that runs into MAX_OUTPUT_CHARS loses its
// page furniture before it loses a footnote, and a footnote before it loses a paragraph.
//
// Matched by CONVENTIONAL PATH, as the main part already is, rather than resolved through
// word/_rels/document.xml.rels. Consistent with the existing reader, and the failure mode is benign
// in a way the main part's is not: a part missed here costs its own text, nothing else.
const DOCX_AUXILIARY_PARTS: { pattern: RegExp; shape: DocxPartShape; dedupe: boolean }[] = [
    { pattern: /^word\/footnotes\.xml$/, shape: DOCX_FOOTNOTES_SHAPE, dedupe: false },
    { pattern: /^word\/endnotes\.xml$/, shape: DOCX_ENDNOTES_SHAPE, dedupe: false },
    { pattern: /^word\/comments\.xml$/, shape: DOCX_COMMENTS_SHAPE, dedupe: false },
    // A section can declare up to three headers (default, first page, even pages) and Word writes
    // one part each, usually with identical text. Emitted once per distinct text, because repeating
    // a letterhead once per section is noise that also eats the cap.
    { pattern: /^word\/header\d*\.xml$/, shape: DOCX_HEADER_SHAPE, dedupe: true },
    { pattern: /^word\/footer\d*\.xml$/, shape: DOCX_FOOTER_SHAPE, dedupe: true },
]

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
        if (!entries) throw new ExtractionFailure('malformed')
        // JSZip (and therefore Mammoth's previous reader) resolves duplicate names to the final
        // central-directory record. Preserve that compatibility rather than silently switching the
        // extracted document when an ambiguous archive reaches this lower-level ZIP reader.
        const part = entries.findLast((entry) => opcKey(entry.name) === DOCX_MAIN_PART)
        if (!part) throw new ExtractionFailure('wrong-document-shape')

        // Decode ACROSS inflate chunks, not per chunk: a 16 KB boundary lands mid-sequence in any
        // document with a non-ASCII character, and chunk.toString('utf8') would turn that one
        // character into two U+FFFD. StringDecoder carries the partial bytes forward. (saxes handles
        // a surrogate pair split across write() calls itself; this is the layer below that.)
        const { StringDecoder } = await import('node:string_decoder')

        // One part, read to its end or to the budget it was given. Returns the text and whether it
        // stopped early, and throws only the way the main part always has — so the caller can treat
        // the main part as load-bearing and every other part as best-effort.
        const readPart = async (entry: ZipEntry, shape: DocxPartShape, budget: number) => {
            const reader = await createDocxReader(budget, shape)
            const decoder = new StringDecoder('utf8')
            let partTruncated = false
            let sawContent: boolean | undefined
            try {
                for await (const chunk of docxMainPartChunks(entry)) {
                // Both guards here, ahead of the work, at one inflate chunk of granularity. Finer
                // than the pdf per-page and xlsx per-row checks, and the only place a stop is
                // possible: saxes has no abort, so the way to stop parsing is to stop feeding it.
                // Per CHUNK rather than per emitted character, for the same reason the xlsx deadline
                // sits above its empty-row skip — tens of MB of w:pPr/w:rPr markup produces no text
                // at all, so a cap-only check would never fire on precisely the cheapest loop to
                // spin. Breaking a `for await` destroys the inflate stream, so the rest of the
                // document is never decompressed either.
                    if (Date.now() > deadline || reader.shouldStop()) {
                        partTruncated = true
                        break
                    }
                    reader.write(decoder.write(chunk))
                }
                // Only a read that ran to the end may assert the part ended cleanly.
                if (!partTruncated) sawContent = reader.end(decoder.end())
                if (reader.overCap()) partTruncated = true
            } catch (error) {
            // saxes is conformant where mammoth's DOM parser recovered, so a document the old reader
            // read to the end can stop short here. Text already extracted is still text, and the
            // contract has a word for "the document continues past this point" — so keep it and say
            // so. Nothing read means nothing to label: that stays a failure, which a caller can see
            // and retry. What this deliberately does NOT do is install a saxes error handler and
            // parse on; measured, that emits close-tag text as content and descends into elements
            // mammoth drops — silent wrong output, the one outcome this file fails over everywhere else.
                if (reader.text().trim().length === 0) throw error
                partTruncated = true
            }
            return { text: reader.text(), truncated: partTruncated, sawContent }
        }

        const body = await readPart(part, DOCX_MAIN_SHAPE, maxOutputChars)
        // Keep this semantic assertion outside the malformed-XML recovery above. A bodyless main
        // part can contain parseable paragraph text, but it is still not a Word document; catching
        // this assertion as though parsing stopped midway would mislabel the foreign content as a
        // useful truncated prefix.
        if (body.sawContent === false) throw new ExtractionFailure('wrong-document-shape')

        // Concatenated, NOT joined with a separator: the reader already terminates every paragraph
        // with one, so a part's text ends where the next can start. Adding another here is what put
        // a pair of blank paragraphs between the body and its own footnotes.
        const sections = [body.text]
        let truncated = body.truncated
        const used = () => sections.reduce((total, part) => total + part.length, 0)

        for (const { pattern, shape, dedupe } of DOCX_AUXILIARY_PARTS) {
            const seen = new Set<string>()
            for (const entry of entries) {
                if (truncated || Date.now() > deadline) {
                    // Ran out of room or time. The body is already read, so this is a truncation of
                    // the document, exactly as stopping mid-body would be — not a failure.
                    truncated = true
                    break
                }
                const key = opcKey(entry.name)
                if (key === undefined || !pattern.test(key)) continue
                const remaining = maxOutputChars - used()
                if (remaining <= 0) {
                    truncated = true
                    break
                }
                try {
                    const read = await readPart(entry, shape, remaining)
                    // Trimmed for the DECISIONS (is there anything here, have we already emitted
                    // it), raw for the OUTPUT, so a part keeps the paragraph breaks it read.
                    const key = read.text.trim()
                    if (key === '' || (dedupe && seen.has(key))) continue
                    seen.add(key)
                    sections.push(read.text)
                    if (read.truncated) truncated = true
                } catch {
                    // BEST-EFFORT, and the reason the main part is read separately above: a broken
                    // footnotes part costs its own text and nothing else. Turning a readable document
                    // into a failure over its margins would trade a whole extraction for a fragment.
                    truncated = true
                }
            }
            if (truncated) break
        }

        // Emptiness is left to the entry point: an empty document is exactly '\n\n', which trims to ''.
        return { text: sections.join(''), truncated }
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
interface StreamedCell {
    address: string
    value?: unknown
    text?: string
    formula?: string
    result?: unknown
    numFmt?: string
    workbook?: { properties?: { model?: { date1904?: boolean } } }
}

interface StreamedRow {
    number?: number
    eachCell: (options: { includeEmpty: boolean }, callback: (cell: StreamedCell) => void) => void
}

type ExcelXmlChunks = AsyncIterable<Buffer | string>
interface ExcelXmlEntry {
    setEncoding: (encoding: BufferEncoding) => unknown
    pipe: (destination: ExcelXmlEntry) => ExcelXmlEntry
}
interface ExcelJsInternalReader {
    sharedStrings?: unknown[]
    _parseRels?: (entry: ExcelXmlEntry) => Promise<void>
    _parseWorkbook?: (entry: ExcelXmlEntry) => Promise<void>
    _parseSharedStrings?: (entry: ExcelXmlEntry) => AsyncIterable<unknown>
    _parseStyles?: (entry: ExcelXmlEntry) => Promise<void>
    _parseWorksheet?: (chunks: ExcelXmlChunks, sheetNo: string) => Iterable<unknown>
    _parseHyperlinks?: (chunks: ExcelXmlChunks, sheetNo: string) => Iterable<unknown>
}

// ExcelJS's worksheet parser already rejects malformed XML as it receives it, but never calls
// saxes.close() at natural EOF. This lexical companion enforces the common depth ceiling and records
// whether the root actually closed. It understands XML comments/CDATA/PIs so a closing-tag-shaped
// string cannot spoof completion. Worksheet value metadata is collected separately by the
// namespace-aware companion below.
interface ExcelXmlScan {
    booleanCells: Set<string>
    inlineCells?: Map<string, string>
    emptySharedCells?: Map<number, Set<number>>
    write: (chunk: string) => void
    complete: (allowRootless?: boolean) => boolean
}

const createXmlScan = (): ExcelXmlScan => {
    type Mode = 'text' | 'tag' | 'comment' | 'cdata' | 'pi' | 'declaration'
    let mode: Mode = 'text'
    let tag = ''
    let quote = ''
    let suffix = ''
    let lastUnquoted = ''
    let declarationDepth = 0
    let rootClosed = false
    let depth = 0
    const booleanCells = new Set<string>()
    // This scanner buffers a tag's text to read its name when the tag closes, so an unterminated '<'
    // in a hostile file would otherwise accumulate the rest of the document into one string. Past
    // this point the characters are dropped and only the name-bearing head is kept — which is all
    // that is read. 8 KB because the name is the first token: no legal element name approaches it,
    // and the slack covers a tag carrying a long attribute list before its '>'. A tag longer than
    // this is truncated for NAMING purposes only; depth tracking still follows its real '>'.
    const TAG_CAPTURE_LIMIT = 8 * 1024

    const finishTag = (): void => {
        const body = tag.slice(1, tag.endsWith('>') ? -1 : undefined).trim()
        const closing = body.startsWith('/')
        if (closing) {
            depth = Math.max(0, depth - 1)
            if (depth === 0) rootClosed = true
        } else if (body.endsWith('/') || lastUnquoted === '/') {
            if (depth === 0) rootClosed = true
        } else {
            assertXmlDepth(++depth)
        }
        tag = ''
        quote = ''
        lastUnquoted = ''
        mode = 'text'
    }

    const write = (chunk: string): void => {
        for (let i = 0; i < chunk.length; i++) {
            const char = chunk[i]
            if (mode === 'text') {
                if (char === '<') {
                    tag = '<'
                    lastUnquoted = '<'
                    mode = 'tag'
                }
                continue
            }
            if (mode === 'comment' || mode === 'cdata' || mode === 'pi') {
                suffix = (suffix + char).slice(-3)
                if (
                    (mode === 'comment' && suffix === '-->') ||
                    (mode === 'cdata' && suffix === ']]>') ||
                    (mode === 'pi' && suffix.endsWith('?>'))
                ) {
                    suffix = ''
                    mode = 'text'
                }
                continue
            }
            if (mode === 'declaration') {
                if (quote) {
                    if (char === quote) quote = ''
                } else if (char === '"' || char === "'") quote = char
                else if (char === '[') declarationDepth++
                else if (char === ']') declarationDepth = Math.max(0, declarationDepth - 1)
                else if (char === '>' && declarationDepth === 0) mode = 'text'
                continue
            }

            if (tag.length < TAG_CAPTURE_LIMIT) tag += char
            if (tag === '<!--') {
                tag = ''
                suffix = ''
                mode = 'comment'
                continue
            }
            if (tag === '<![CDATA[') {
                tag = ''
                suffix = ''
                mode = 'cdata'
                continue
            }
            if (tag === '<?') {
                tag = ''
                suffix = ''
                mode = 'pi'
                continue
            }
            if (/^<!DOCTYPE$/i.test(tag)) {
                tag = ''
                quote = ''
                declarationDepth = 0
                mode = 'declaration'
                continue
            }
            if (quote) {
                if (char === quote) quote = ''
            } else if (char === '"' || char === "'") quote = char
            else if (char === '>') finishTag()
            else if (char !== ' ' && char !== '\t' && char !== '\n' && char !== '\r') lastUnquoted = char
        }
    }

    return {
        booleanCells,
        write,
        complete: (allowRootless = false) => mode === 'text' && depth === 0 && (rootClosed || allowRootless),
    }
}

const isExcelDateFormat = (format?: string): boolean =>
    Boolean(format?.replace(/\[[^\]]*]/g, '').replace(/"[^"]*"/g, '').match(/[ymdhMsb]+/))

const streamedFormula = (cell: StreamedCell): { result: unknown; text: string } | undefined => {
    if (cell.formula !== undefined) return { result: cell.result, text: cell.text ?? '' }
    if (typeof cell.value !== 'object' || cell.value === null) return undefined
    const shared = cell.value as { formula?: unknown; result?: unknown }
    if (shared.formula !== '') return undefined
    return { result: shared.result, text: shared.result == null ? '' : String(shared.result) }
}

const excelFormulaText = (cell: StreamedCell, booleanCells: Set<string>): string => {
    const formula = streamedFormula(cell)
    if (formula !== undefined && typeof formula.result === 'number' && Number.isFinite(formula.result)) {
        if (booleanCells.delete(cell.address)) return String(formula.result !== 0)
        if (isExcelDateFormat(cell.numFmt)) {
            const date1904 = cell.workbook?.properties?.model?.date1904 ?? false
            const milliseconds = Math.round((formula.result - 25569 + (date1904 ? 1462 : 0)) * 86_400_000)
            return new Date(milliseconds).toString()
        }
    } else {
        booleanCells.delete(cell.address)
    }
    return formula?.text ?? cell.text ?? ''
}

const excelCellPosition = (address: string): { row: number; column: number } | undefined => {
    const match = /^([A-Za-z]+)(\d+)$/.exec(address)
    if (!match) return undefined
    let column = 0
    for (const char of match[1].toUpperCase()) column = column * 26 + char.charCodeAt(0) - 64
    const row = Number(match[2])
    return Number.isSafeInteger(row) && row > 0 ? { row, column } : undefined
}

// XLSX — modern Excel. Each sheet flattened to text for search/indexing.
const xlsxHandler: Handler = {
    kind: 'xlsx',
    contentTypes: ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
    extensions: ['.xlsx'],
    extract: async ({ content, maxOutputChars, deadline }) => {
        const { Readable, Transform } = await import('node:stream')
        const { StringDecoder } = await import('node:string_decoder')
        const [{ default: ExcelJS }, { SaxesParser }] = await Promise.all([
            import('exceljs'), // not SheetJS: no known parse-time CVEs
            import('saxes'),
        ])
        // Stream rather than workbook.xlsx.load(), which materializes every cell as a live object
        // before any cap can apply — a 4 MB in-cap .xlsx peaked at hundreds of MB and OOMed a 1024 MB
        // worker. Peak now follows the shared-string table, not the cell graph. See
        // reorderForStreaming: without it this reader drops worksheets.
        const rewritten = await reorderForStreaming(content)
        // Fail closed: the budget measured the central directory, but unzipper inflates what its
        // LOCAL-header walk finds, so the budget binds this path only through the rewritten archive.
        if (!rewritten.ok) throw new ExtractionFailure(rewritten.reason)
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

        // ExcelJS 4.4 decodes each unzipper output chunk independently before feeding saxes. A
        // multi-byte code point split by inflate therefore becomes U+FFFD without any parse error.
        // Its entry parsers accept streams and its worksheet parsers accept async iterables, so
        // normalize both boundaries to decoded strings with state carried across chunks.
        const internal = reader as unknown as ExcelJsInternalReader
        const hook = <K extends keyof ExcelJsInternalReader>(name: K): NonNullable<ExcelJsInternalReader[K]> => {
            const method = internal[name]
            if (typeof method !== 'function') {
                throw new ExtractionFailure('internal')
            }
            return method.bind(internal) as NonNullable<ExcelJsInternalReader[K]>
        }
        let truncated = false
        const booleanCellLimit = 2 * maxOutputChars + 1_024
        const decodeEntry = (entry: ExcelXmlEntry, guardDepth = true, allowRootless = false): ExcelXmlEntry => {
            entry.setEncoding('utf8') // Node streams use StringDecoder internally across chunks.
            if (!guardDepth) return entry
            const scan = createXmlScan()
            const guard = new Transform({
                decodeStrings: false,
                encoding: 'utf8',
                transform(chunk, _encoding, done) {
                    try {
                        scan.write(String(chunk))
                    } catch (error) {
                        done(error as Error)
                        return
                    }
                    done(null, chunk)
                },
                flush(done) {
                    if (!scan.complete(allowRootless)) {
                        done(new ExtractionFailure('malformed'))
                        return
                    }
                    done()
                },
            })
            return entry.pipe(guard as unknown as ExcelXmlEntry)
        }
        const decodeChunks = async function* (
            chunks: ExcelXmlChunks,
            scan?: ExcelXmlScan
        ): AsyncIterable<string> {
            const decoder = new StringDecoder('utf8')
            let stopped = false
            const slices = function* (text: string): Iterable<string> {
                for (let offset = 0; offset < text.length; offset += STREAM_SLICE_UNITS) {
                    if (Date.now() > deadline) {
                        truncated = true
                        stopped = true
                        return
                    }
                    const slice = text.slice(offset, offset + STREAM_SLICE_UNITS)
                    scan?.write(slice)
                    yield slice
                }
            }
            for await (const chunk of chunks) {
                yield* slices(typeof chunk === 'string' ? chunk : decoder.write(chunk))
                if (stopped) return
            }
            yield* slices(decoder.end())
            if (stopped) return
            // Reached only at NATURAL EOF. Abandonment for the cap/deadline skips this assertion,
            // keeping intentional partial reads labeled `truncated` rather than parse failures.
            if (scan && !scan.complete()) throw new ExtractionFailure('malformed')
        }
        const parseRels = hook('_parseRels')
        internal._parseRels = (entry) => parseRels(decodeEntry(entry, false))
        const parseWorkbook = hook('_parseWorkbook')
        internal._parseWorkbook = (entry) => parseWorkbook(decodeEntry(entry, false))
        const parseSharedStrings = hook('_parseSharedStrings')
        internal._parseSharedStrings = (entry) => parseSharedStrings(decodeEntry(entry, true, true))
        const parseStyles = hook('_parseStyles')
        internal._parseStyles = (entry) => parseStyles(decodeEntry(entry, true, true))
        const worksheetScans: ExcelXmlScan[] = []
        const createWorksheetScan = (): ExcelXmlScan => {
            const base = createXmlScan()
            const booleanCells = new Set<string>()
            const inlineCells = new Map<string, string>()
            const emptySharedCells = new Map<number, Set<number>>()
            let emptySharedCount = 0
            // Keep only enough scan-order metadata to produce the capped output. The extra 1,024
            // cells in booleanCellLimit cover sparse/empty inline strings without making this map
            // grow with worksheet size.
            const inlineTextLimit = booleanCellLimit
            let inlineChars = 0
            const open: { uri: string; local: string }[] = []
            let cell:
                | {
                      depth: number
                      address: string
                      inline: boolean
                      shared: boolean
                      text: string
                      textDepth?: number
                      sharedIndexState: 'leading' | 'sign' | 'digits' | 'done' | 'invalid' | 'overflow'
                      sharedIndexValue: number
                      sharedIndexSign: 1 | -1
                      valueDepth?: number
                  }
                | undefined
            const inline = new SaxesParser({
                xmlns: true,
                // ExcelJS's non-namespace parser tolerated undeclared extension prefixes. Treat
                // them as foreign markup instead of failing every worksheet in the attachment.
                resolvePrefix: () => UNBOUND_NAMESPACE,
            })
            const isSpreadsheet = (element: { uri: string; local: string } | undefined, local: string): boolean =>
                element !== undefined && SPREADSHEETML_NS.has(element.uri) && element.local === local
            inline.on('opentag', (tag) => {
                assertXmlDepth(open.length + 1)
                const isCell =
                    SPREADSHEETML_NS.has(tag.uri) &&
                    tag.local === 'c' &&
                    open.length === 3 &&
                    isSpreadsheet(open[0], 'worksheet') &&
                    isSpreadsheet(open[1], 'sheetData') &&
                    isSpreadsheet(open[2], 'row')
                open.push({ uri: tag.uri, local: tag.local })
                if (isCell && cell === undefined) {
                    const address = tag.attributes.r?.value ?? ''
                    const type = tag.attributes.t?.value
                    cell = {
                        depth: open.length,
                        address,
                        inline: type === 'inlineStr',
                        shared: type === 's',
                        text: '',
                        sharedIndexState: 'leading',
                        sharedIndexValue: 0,
                        sharedIndexSign: 1,
                    }
                    if (type === 'b' && address !== '' && booleanCells.size < booleanCellLimit) {
                        booleanCells.add(address)
                    }
                } else if (cell?.shared && SPREADSHEETML_NS.has(tag.uri) && tag.local === 'v') {
                    const cellElement = open[cell.depth - 1]
                    if (open.length === cell.depth + 1 && isSpreadsheet(cellElement, 'c')) {
                        cell.valueDepth = open.length
                    }
                } else if (cell?.inline && SPREADSHEETML_NS.has(tag.uri) && tag.local === 't') {
                    const cellElement = open[cell.depth - 1]
                    const direct =
                        open.length === cell.depth + 2 &&
                        isSpreadsheet(cellElement, 'c') &&
                        isSpreadsheet(open[open.length - 2], 'is')
                    const rich =
                        open.length === cell.depth + 3 &&
                        isSpreadsheet(cellElement, 'c') &&
                        isSpreadsheet(open[open.length - 3], 'is') &&
                        isSpreadsheet(open[open.length - 2], 'r')
                    if (direct || rich) cell.textDepth = open.length
                }
            })
            const appendInline = (text: string): void => {
                if (!cell?.inline || cell.textDepth !== open.length) return
                const room = inlineTextLimit - inlineChars - cell.text.length
                if (room > 0) cell.text += text.slice(0, room)
            }
            const appendSharedIndex = (text: string): void => {
                if (!cell?.shared || cell.valueDepth !== open.length) return
                for (const char of text) {
                    if (
                        cell.sharedIndexState === 'done' ||
                        cell.sharedIndexState === 'invalid' ||
                        cell.sharedIndexState === 'overflow'
                    ) {
                        return
                    }
                    if (cell.sharedIndexState === 'leading') {
                        if (/\s/u.test(char)) continue
                        if (char === '+' || char === '-') {
                            cell.sharedIndexSign = char === '-' ? -1 : 1
                            cell.sharedIndexState = 'sign'
                            continue
                        }
                        if (char < '0' || char > '9') {
                            cell.sharedIndexState = 'invalid'
                            return
                        }
                        cell.sharedIndexState = 'digits'
                    } else if (cell.sharedIndexState === 'sign') {
                        if (char < '0' || char > '9') {
                            cell.sharedIndexState = 'invalid'
                            return
                        }
                        cell.sharedIndexState = 'digits'
                    } else if (char < '0' || char > '9') {
                        // parseInt stops at the first non-decimal character: 1e3, 1.5 and 1abc
                        // all mean index 1, while 0x10 means index 0 under radix 10.
                        cell.sharedIndexState = 'done'
                        return
                    }
                    cell.sharedIndexValue = cell.sharedIndexValue * 10 + char.charCodeAt(0) - 48
                    if (cell.sharedIndexValue > MAX_SHARED_STRING_INDEX) {
                        // The decompression ceiling cannot hold this many <si> records. Stop reading
                        // digits now rather than retaining an attacker-sized value string.
                        cell.sharedIndexState = 'overflow'
                        return
                    }
                }
            }
            inline.on('text', (text) => {
                appendInline(text)
                appendSharedIndex(text)
            })
            inline.on('cdata', (text) => {
                appendInline(text)
                appendSharedIndex(text)
            })
            inline.on('closetag', (tag) => {
                if (
                    cell?.textDepth === open.length &&
                    SPREADSHEETML_NS.has(tag.uri) &&
                    tag.local === 't'
                ) {
                    cell.textDepth = undefined
                }
                if (
                    cell?.valueDepth === open.length &&
                    SPREADSHEETML_NS.has(tag.uri) &&
                    tag.local === 'v'
                ) {
                    cell.valueDepth = undefined
                }
                if (
                    cell?.depth === open.length &&
                    SPREADSHEETML_NS.has(tag.uri) &&
                    tag.local === 'c'
                ) {
                    if (
                        cell.inline &&
                        cell.address !== '' &&
                        inlineCells.size < inlineTextLimit &&
                        inlineChars < inlineTextLimit
                    ) {
                        inlineCells.set(cell.address, cell.text)
                        inlineChars += cell.text.length
                    }
                    if (cell.shared) {
                        const hasIndex = cell.sharedIndexState === 'digits' || cell.sharedIndexState === 'done'
                        const index = cell.sharedIndexSign * cell.sharedIndexValue
                        const sharedStrings = internal.sharedStrings ?? []
                        if (cell.sharedIndexState === 'overflow') {
                            throw new ExtractionFailure('malformed')
                        }
                        if (hasIndex && (index < 0 || index >= sharedStrings.length)) {
                            throw new ExtractionFailure('malformed')
                        }
                        if (hasIndex && sharedStrings[index] == null) {
                            const position = excelCellPosition(cell.address)
                            if (position && emptySharedCount < inlineTextLimit) {
                                const columns = emptySharedCells.get(position.row) ?? new Set<number>()
                                if (!columns.has(position.column)) emptySharedCount++
                                columns.add(position.column)
                                emptySharedCells.set(position.row, columns)
                            }
                        }
                    }
                    cell = undefined
                }
                open.pop()
            })
            return {
                booleanCells,
                inlineCells,
                emptySharedCells,
                write: (chunk) => {
                    base.write(chunk)
                    inline.write(chunk)
                },
                complete: () => {
                    inline.close()
                    return base.complete()
                },
            }
        }
        const parseWorksheet = hook('_parseWorksheet')
        internal._parseWorksheet = (chunks, sheetNo) => {
            const scan = createWorksheetScan()
            worksheetScans.push(scan)
            return parseWorksheet(decodeChunks(chunks, scan), sheetNo)
        }
        const parseHyperlinks = hook('_parseHyperlinks')
        internal._parseHyperlinks = (chunks, sheetNo) => parseHyperlinks(decodeChunks(chunks), sheetNo)

        const sheets: string[] = []
        let length = 0
        let seen = 0
        for await (const worksheet of reader) {
            // Row-less sheets never enter the inner loop. Check here as well so thousands of them
            // cannot run to withTimeout and discard text already extracted from earlier sheets.
            if (Date.now() > deadline) {
                truncated = true
                break
            }
            seen++
            const worksheetScan = worksheetScans[seen - 1]
            const booleanCells = worksheetScan?.booleanCells ?? new Set<string>()
            const name = resolved[seen - 1]?.name ?? `Sheet${seen}`
            const header = `=== ${name} ===\n`
            const rows: string[] = []
            for await (const batch of worksheet) {
                // 4.4.0 yields ONE Row per iteration (worksheet-reader.js:275 pushes
                // `{eventType: 'row', value: row}`, and :104-112 yields each `value` through) — the
                // other way round from how this once read. The normalization stays anyway: exceljs
                // documents the batched shape, which is why the wrong version of this was believable,
                // and it costs one predicate to be right under either.
                const batchRows = (Array.isArray(batch) ? batch : [batch]) as Array<StreamedRow | null | undefined>
                for (const row of batchRows) {
                    // ABOVE the empty-row skip: a contentless row is `continue`d, and a sheet of them
                    // trips no cap either, so a check below would never run. This loop awaits, so
                    // breaking really ends the parse — unlike withTimeout's race, which frees the
                    // slot while the parse detaches.
                    if (Date.now() > deadline) {
                        truncated = true
                        break
                    }
                    // ExcelJS emits a row event at every </row>, even outside sheetData, but its
                    // current row is null there. The full-workbook reader ignored those events;
                    // preserve that behavior instead of turning one stray extension element into
                    // total extraction failure.
                    if (!row) continue
                    const cells: { column: number; text: string }[] = []
                    const seenColumns = new Set<number>()
                    // cell.text = the shown value (formula result, formatted date), not the raw formula.
                    row.eachCell({ includeEmpty: false }, (cell) => {
                        const inline = worksheetScan?.inlineCells?.get(cell.address)
                        if (inline !== undefined) worksheetScan?.inlineCells?.delete(cell.address)
                        const position = excelCellPosition(cell.address)
                        const column = position?.column ?? cells.length + 1
                        seenColumns.add(column)
                        cells.push({ column, text: inline ?? excelFormulaText(cell, booleanCells) })
                    })
                    const rowNumber = row.number
                    const emptyShared = rowNumber === undefined ? undefined : worksheetScan?.emptySharedCells?.get(rowNumber)
                    // An explicit empty shared string matters only as a separator between real
                    // cells. Blank-only rows stay absent, and leading/trailing placeholders do not
                    // manufacture whitespace-only output or consume the cap.
                    if (emptyShared && rowNumber !== undefined && cells.length > 1) {
                        const columns = cells.map((cell) => cell.column)
                        const first = Math.min(...columns)
                        const last = Math.max(...columns)
                        for (const column of emptyShared) {
                            if (column > first && column < last && !seenColumns.has(column)) {
                                cells.push({ column, text: '' })
                            }
                        }
                        worksheetScan?.emptySharedCells?.delete(rowNumber)
                    }
                    // Drop empty cells/rows so a sparse sheet doesn't flatten into runs of empty tabs.
                    if (cells.length === 0) continue
                    cells.sort((a, b) => a.column - b.column)
                    const line = cells.map((cell) => cell.text).join('\t')
                    length +=
                        rows.length === 0
                            ? (sheets.length === 0 ? 0 : BLOCK_SEPARATOR.length) + header.length + line.length
                            : 1 + line.length
                    rows.push(line)
                    if (length > maxOutputChars) {
                        truncated = true
                        break // one line of overshoot, trimmed centrally
                    }
                }
                if (truncated) break
            }
            worksheetScan?.booleanCells.clear()
            worksheetScan?.inlineCells?.clear()
            worksheetScan?.emptySharedCells?.clear()
            // Positional, and sound because the rebuild laid out exactly `resolved` and nothing else
            // this reader dispatches as a worksheet, in this order. Reading worksheet.name instead is
            // what produced "Sheet1" for a legal absolute rel Target: exceljs matches rel.Target
            // against one exact spelling (workbook-reader.js:302) and gives up on every other, and
            // its .d.ts does not admit the field either way. The fallback is unreachable — an
            // out-of-range index means the reader emitted a part we never wrote.
            if (rows.length > 0) sheets.push(`${header}${rows.join('\n')}`)
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
            throw new ExtractionFailure('internal')
        }

        return { text: sheets.join(BLOCK_SEPARATOR), truncated }
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

// docx and xlsx share the zip magic; their main part is what tells them apart. The names are needed
// in two shapes — as entry-name strings for the exact-match path, and as bytes for the raw scan the
// fallback uses — so the string is canonical here and the Buffer is derived from it. Both shapes
// have to name the SAME part or the two paths disagree about what a package is.
const DOCX_MAIN_PART = 'word/document.xml'
const WORKBOOK_PART = 'xl/workbook.xml'
const DOCX_PART = Buffer.from(DOCX_MAIN_PART)
const XLSX_PART = Buffer.from(WORKBOOK_PART)
// The other OPC part names this module names by hand, kept with the two above so there is a single
// home for them: routing, the streaming rebuild and the sheet-identity resolver all spell the same
// parts, and a rebuild that ordered a part the resolver reads under a different spelling would put
// exceljs back on the branch the rebuild exists to avoid.
const WORKBOOK_RELS_PART = 'xl/_rels/workbook.xml.rels'
const CONTENT_TYPES_PART = '[Content_Types].xml'

const startsWith = (content: Buffer, magic: Buffer): boolean =>
    content.length >= magic.length && content.subarray(0, magic.length).equals(magic)

// Which OOXML kind a zip is, by its root part. Returns undefined for a non-OOXML zip (pptx, jar).
const ooxmlKind = (content: Buffer): HandlerKind | undefined => {
    // Match exact zip ENTRY names, not raw bytes: a docx can embed a workbook, and that workbook's
    // xl/workbook.xml lives INSIDE another entry rather than as an entry of this package — so a raw
    // scan gets fooled by storage order, while a root-entry match doesn't.
    const names = zipEntryNames(content)
    if (names) {
        const folded = new Set(names.flatMap((name) => opcKey(name) ?? []))
        const hasDocx = folded.has(DOCX_MAIN_PART)
        const hasXlsx = folded.has(WORKBOOK_PART)
        if (hasDocx) return 'docx' // a real word/document.xml root part wins (docx may embed a workbook)
        if (hasXlsx) return 'xlsx'
        return undefined // OOXML zip with neither root part (pptx, jar, plain archive)
    }
    // Fallback (archive not walkable): raw-bytes scan, earlier main-part marker wins.
    // Fold one bounded copy rather than comparing every byte against both needles in nested loops.
    const folded = Buffer.from(content)
    for (let i = 0; i < folded.length; i++) {
        if (folded[i] >= 0x41 && folded[i] <= 0x5a) folded[i] += 0x20
    }
    const d = folded.indexOf(DOCX_PART)
    const x = folded.indexOf(XLSX_PART)
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
//     (unzipper/lib/parse.js:51), only skipping PAST directory records — so the invariants do not
//     bind it, and a local entry the directory omits would still be inflated unmeasured.
//     reorderForStreaming binds it instead. What it writes must be bounded in BOTH directions —
//     stating only the first once let a 21x amplification read as impossible from here:
//       - NO MORE. Every local header it writes comes from a MEASURED central-directory record, and
//         each record is written AT MOST ONCE. Sheet layout carries resolved ZipEntry objects rather
//         than looking parts up by name, because a zip name is not a key: entries sharing one name
//         collapse onto whichever came first, and re-emit its bytes once per reference.
//       - NO FEWER IS FINE. Some measured records are deliberately not written at all (orphan
//         worksheet parts — see SHEET IDENTITY). Dropping only removes bytes from the reader's
//         reach, so it cannot loosen the bound.
//       - TWO EXCEPTIONS, ours. The injected empty xl/sharedStrings.xml and workbook.xml.rels match no
//         central records. They are fixed 153- and 140-byte literals, not anything the input controls.
//     So what unzipper can inflate is at most what was measured, plus those 293 bytes. Hence the
//     xlsx handler FAILS when the rewrite can't be produced; the original bytes would drop the
//     budget on the floor.
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
// Fixed-record sizes from APPNOTE 4.3.7/4.3.12/4.3.16 — the bytes before each record's variable-length
// tail (name, extra, comment). Used by both the walks below and the rebuild in reorderForStreaming.
const LOCAL_HEADER_BYTES = 30
const CENTRAL_HEADER_BYTES = 46
const EOCD_BYTES = 22

// `ok` = safe to hand to the parser. Otherwise `status` is which kind of no, since the two differ to
// a caller: `failed` = the bytes are broken (the parser would have thrown anyway), `skipped` = intact
// but we decline — over budget, or a variant we don't chase. Mirrors the MAX_INPUT_BYTES precedent.
type DecompressionCheck = { ok: true } | { ok: false; status: 'failed' | 'skipped'; reason: ExtractionReason }

const corrupt = (reason: ExtractionReason = 'malformed'): DecompressionCheck => ({ ok: false, status: 'failed', reason })
const declined = (reason: ExtractionReason): DecompressionCheck => ({ ok: false, status: 'skipped', reason })

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
    return eocd >= 0 && eocd + EOCD_BYTES <= buf.length ? eocd : -1 // need room for the fixed record
}

// One entry as the central directory describes it, plus its stored (still-compressed) bytes.
interface ZipEntry {
    name: string
    utf8Name: boolean
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
        if (p + CENTRAL_HEADER_BYTES > buf.length || buf.readUInt32LE(p) !== CD_SIG) return undefined
        const nameLen = buf.readUInt16LE(p + 28)
        if (p + CENTRAL_HEADER_BYTES + nameLen > buf.length) return undefined
        const encoding = (buf.readUInt16LE(p + 8) & 0x0800) !== 0 ? 'utf8' : 'latin1'
        names.push(buf.toString(encoding, p + CENTRAL_HEADER_BYTES, p + CENTRAL_HEADER_BYTES + nameLen))
        p += CENTRAL_HEADER_BYTES + nameLen + buf.readUInt16LE(p + 30) + buf.readUInt16LE(p + 32)
    }
    return names
}

// The same walk, plus everything needed to RE-EMIT each entry: its stored bytes, located through the
// local header, and the fields a rebuilt header has to carry. Returns undefined when the directory
// can't be walked. Both callers refuse rather than degrade: reorderForStreaming because the entry
// set it produces is what unzipper will be given, and docxHandler because this walk chooses the exact
// compressed region it will inflate.
//
// That makes this walk load-bearing for the zip-bomb guard on BOTH formats, not merely an identifier:
// the archive it feeds makes the budget's measurement bind for .xlsx, while the entry it selects is
// the only region inflated for .docx (see DECOMPRESSION BUDGET). Still
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
        if (p + CENTRAL_HEADER_BYTES > buf.length || buf.readUInt32LE(p) !== CD_SIG) return undefined
        const nameLen = buf.readUInt16LE(p + 28)
        if (p + CENTRAL_HEADER_BYTES + nameLen > buf.length) return undefined
        const compSize = buf.readUInt32LE(p + 20)
        const localOffset = buf.readUInt32LE(p + 42)
        if (compSize === 0xffffffff || localOffset === 0xffffffff) return undefined // ZIP64
        // The local header carries its own name/extra lengths, which can differ from the central
        // copy — they are what fixes where this entry's bytes actually start.
        if (localOffset + LOCAL_HEADER_BYTES > buf.length || buf.readUInt32LE(localOffset) !== LOCAL_SIG) return undefined
        const dataStart = localOffset + LOCAL_HEADER_BYTES + buf.readUInt16LE(localOffset + 26) + buf.readUInt16LE(localOffset + 28)
        const data = buf.subarray(dataStart, dataStart + compSize)
        if (data.length < compSize) return undefined
        entries.push({
            name: buf.toString((buf.readUInt16LE(p + 8) & 0x0800) !== 0 ? 'utf8' : 'latin1', p + CENTRAL_HEADER_BYTES, p + CENTRAL_HEADER_BYTES + nameLen),
            utf8Name: (buf.readUInt16LE(p + 8) & 0x0800) !== 0,
            method: buf.readUInt16LE(p + 10),
            crc: buf.readUInt32LE(p + 16),
            compSize,
            uncompSize: buf.readUInt32LE(p + 24),
            data,
        })
        p += CENTRAL_HEADER_BYTES + nameLen + buf.readUInt16LE(p + 30) + buf.readUInt16LE(p + 32)
    }
    return entries
}

// Measure a zip's ACTUAL decompressed size, capped, from each entry's real (structural, not
// self-declared) compressed region. The two invariants below were written to pin us to the records
// jszip would read on the old .docx path; jszip is gone, and what they now do is stated at each.
// Both assert the FILE is self-consistent rather than mirroring any particular reader, which is why
// neither rotted when the reader changed. Every real archive satisfies them (73 measured, 0 failures).
const checkDecompressionBudget = async (buf: Buffer, cap: number): Promise<DecompressionCheck> => {
    const eocd = findEocd(buf)
    if (eocd < 0) return corrupt()

    const entries = buf.readUInt16LE(eocd + 10)
    const cdSize = buf.readUInt32LE(eocd + 12)
    const cdOffset = buf.readUInt32LE(eocd + 16)
    // ZIP64 / out-of-range sentinels: the true values live in a ZIP64 record we don't chase. Treat as
    // over-budget rather than trust the classic field or crash on the sentinel.
    if (entries === 0xffff || cdSize === 0xffffffff || cdOffset === 0xffffffff)
        return declined('unsupported-zip-feature')

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
        return corrupt()

    let total = 0
    let p = cdOffset
    for (let i = 0; i < entries; i++) {
        // Guard: a missing/misaligned header means the offset lied. Fail closed — a partial walk must
        // never silently return the total it accumulated so far.
        if (p + CENTRAL_HEADER_BYTES > buf.length || buf.readUInt32LE(p) !== CD_SIG)
            return corrupt()
        const method = buf.readUInt16LE(p + 10)
        const compSize = buf.readUInt32LE(p + 20)
        const localOffset = buf.readUInt32LE(p + 42)
        if (compSize === 0xffffffff || localOffset === 0xffffffff)
            return declined('unsupported-zip-feature')
        // Read the local header's own name/extra lengths — they can differ from the central copy, and
        // they're what fixes where this entry's compressed bytes actually begin.
        if (localOffset + LOCAL_HEADER_BYTES > buf.length || buf.readUInt32LE(localOffset) !== LOCAL_SIG)
            return corrupt()
        const dataStart = localOffset + LOCAL_HEADER_BYTES + buf.readUInt16LE(localOffset + 26) + buf.readUInt16LE(localOffset + 28)
        const comp = buf.subarray(dataStart, dataStart + compSize)
        if (comp.length < compSize)
            return corrupt()

        if (method === 0) {
            total += comp.length // stored (no compression): output === input
        } else if (method === 8) {
            total = await inflateCounting(comp, total, cap)
            if (total === -1) return declined('expands-too-large')
            if (total === -2) return corrupt()
        } else {
            return declined('unsupported-zip-feature')
        }
        if (total > cap) return declined('expands-too-large')
        p += CENTRAL_HEADER_BYTES + buf.readUInt16LE(p + 28) + buf.readUInt16LE(p + 30) + buf.readUInt16LE(p + 32)
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
        return corrupt()
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
    CONTENT_TYPES_PART,
    '_rels/.rels',
    WORKBOOK_PART, // -> this.model  (the 'sheets' TypeError without it)
    WORKBOOK_RELS_PART, // -> this.workbookRels
    'xl/sharedStrings.xml', // -> this.sharedStrings
    'xl/styles.xml', // -> this.styles      (number formats)
]

// ECMA-376 defines part-name equivalence by folding A-Z only. JavaScript's Unicode lowercasing also
// merges distinct legal names such as Ä.xml and ä.xml, so it cannot be used for any OPC identity.
const asciiFold = (value: string): string => value.replace(/[A-Z]/g, (char) => char.toLowerCase())

// OPC part names are URI paths, and the ZIP item name is that canonical escaped path without its
// leading slash. Decode only percent-encoded UNRESERVED ASCII (the URI-equivalent spellings); keep
// escapes such as %20 and %2F as part of the item name. URL supplies dot-segment resolution and
// escapes raw spaces/non-ASCII consistently for relationships, content types, and ZIP entries.
const canonicalOpcPartName = (value: string, base = ''): string | undefined => {
    const trimmed = value.trim()
    if (trimmed === '' || trimmed.includes('\\')) return undefined
    try {
        const root = new URL('https://agentextract.invalid/')
        const baseUrl = new URL(base, root)
        const resolved = new URL(trimmed, baseUrl)
        if (resolved.origin !== root.origin || resolved.search !== '' || resolved.hash !== '') return undefined
        const path = resolved.pathname
            .slice(1)
            .replace(/%([0-9a-f]{2})/gi, (escape, hex: string) => {
                const char = String.fromCharCode(Number.parseInt(hex, 16))
                return /^[A-Za-z0-9._~-]$/.test(char) ? char : escape.toUpperCase()
            })
        return path === '' ? undefined : path
    } catch {
        return undefined
    }
}

const opcKey = (value: string, base = ''): string | undefined => {
    const canonical = canonicalOpcPartName(value, base)
    return canonical === undefined ? undefined : asciiFold(canonical)
}

const XLSX_CANONICAL_CONTROL_PARTS = new Map(XLSX_LEADING_ENTRIES.map((name) => [asciiFold(name), name]))
const XLSX_READER_CONTROL_PARTS = new Set(XLSX_CANONICAL_CONTROL_PARTS.keys())

// A workbook with no strings has no xl/sharedStrings.xml, so this.sharedStrings is never set and
// ordering alone cannot lift it out of the spool branch (measured: 35 of 50 reads dropped sheets).
// Injecting an empty table sets the flag and changes no valid cell. The worksheet companion rejects
// any `t="s"` reference when the original archive omitted this part; otherwise ExcelJS can silently
// erase values and even whole row-only sheets. Handed to the reader only, never written back.
const EMPTY_SHARED_STRINGS = Buffer.from(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>',
    'latin1'
)
// CRC32 of the literal above, precomputed — the payload is fixed, so computing it per extraction
// would be work with one possible answer. A test recomputes it, so the two can't drift.
const EMPTY_SHARED_STRINGS_CRC = 0x2949bd0b

// A missing, empty, or malformed workbook relationship part leaves exceljs's workbookRels unset,
// sending every sheet through its temp-file spool branch. Early break/throw then strands those files.
// Sheet identity is resolved from the ORIGINAL part before the rebuild; ExcelJS needs only a truthy
// relationship model after that, so its private copy always receives this known-valid empty set and
// keeps every sheet on the inline reader, where abandoning the generator has nothing to clean up.
const EMPTY_WORKBOOK_RELS = Buffer.from(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
    'latin1'
)
const EMPTY_WORKBOOK_RELS_CRC = 0x9f1f1b86

const ZIP_VERSION = 20 // 2.0 — the floor for deflate, which is all we re-emit

// The conventional worksheet path, anchored. OPC part URIs are ASCII-case-insensitive, so this also
// recognizes a stored xl/Worksheets/Sheet2.xml. The rebuild canonicalizes every resolved sheet before
// exceljs sees it because exceljs's streaming dispatch below is case-sensitive.
const WORKSHEET_PART = /^xl\/worksheets\/sheet\d+\.xml$/
const isWorksheetPart = (name: string): boolean => {
    const key = opcKey(name)
    return key !== undefined && WORKSHEET_PART.test(key)
}

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

// WORKBOOK_PART, WORKBOOK_RELS_PART and CONTENT_TYPES_PART are declared with the format magic
// above, where routing needs them first.

// The relationship types naming a sheet that has NO xl/worksheets part, mapped to the part family
// each one must live in, in both flavours: Transitional (schemas.openxmlformats.org) and Strict
// (purl.oclc.org). A declaration carrying one is accounted for rather than counted lost.
//
// EXACT URIs, not a suffix match. A relationship type is an identifier the format defines, so
// anything else bearing the same tail is a stranger: https://invalid.example/relationships/chartsheet
// passed a tail match and, with a planted xl/chartsheets/fake.xml to point at, took a real sheet's
// place. Membership of this map is the whole authority — an unrecognized type is expected to be a
// worksheet, so it reaches the rescue instead of bypassing it.
//
// Macrosheets are deliberately absent: treating one as unplaced costs a rescued orphan in the
// output, treating it as accounted-for could cost rows.
const NON_WORKSHEET_REL = new Map([
    ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet', 'chartsheets'],
    ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet', 'dialogsheets'],
    ['http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet', 'chartsheets'],
    ['http://purl.oclc.org/ooxml/officeDocument/relationships/dialogsheet', 'dialogsheets'],
])

// Unlike the non-worksheet families above, a worksheet part may live anywhere in the package. An
// explicit relationship type is therefore the authority that lets a non-conventional target enter
// the worksheet set; absent or unknown types retain the conservative conventional-path fallback.
const WORKSHEET_REL = new Set([
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
    'http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet',
])
const WORKSHEET_CONTENT_TYPE = new Set([
    'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
    'application/vnd.ms-excel.worksheet+xml', // ISO Strict
])

// One worksheet as the workbook describes it: the archive entry holding it, and its tab name.
//
// The ENTRY, deliberately, and not its name: zip permits duplicate entry names, so a name is not a
// key, and carrying the entry makes re-emitting one record several times unrepresentable rather than
// merely guarded against. That is the "NO MORE" half of the bound — see DECOMPRESSION BUDGET.
interface WorkbookSheet {
    entry: ZipEntry
    name: string
}

const worksheetEntries = (entries: ZipEntry[]): ZipEntry[] => entries.filter((entry) => isWorksheetPart(entry.name))

// A worksheet part the workbook does not name. sheetN.xml's own number is the most stable label
// available, and is what the reader's own fallback produced for these.
const partFallbackName = (part: string): string =>
    `Sheet${/(\d+)\.xml$/.exec(asciiFold(canonicalOpcPartName(part) ?? part))?.[1] ?? ''}`

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
// land far under it — so the cap can only be reached by a part padded to reach it. An unreadable
// workbook/rels pair degrades to archive order augmented by whichever relationship/content-type
// declarations remain readable. Unreadable content types become an unknown authorization signal.
const MAX_METADATA_BYTES = 4 * 1024 * 1024

// Inflate one entry, bounded. Only called on the three metadata parts above. undefined = unreadable or
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
// Namespace mode would normally make an undeclared extension prefix fatal. The parsers below map
// only such unclaimed prefixes to UNBOUND_NAMESPACE, so foreign extension markup remains foreign
// and cannot impersonate a SpreadsheetML declaration while the rest of the workbook stays readable.
// Both OOXML flavours. ECMA-376 Transitional is what Excel writes by default; ISO/IEC 29500 Strict
// re-homes the same vocabulary under purl.oclc.org, and "Excel Workbook (Strict Open XML)" is a
// documented save-as target — so a Strict workbook is a legal .xlsx, not a curiosity. Recognizing
// only Transitional cost it nothing but its identity: no <sheet> matched, so tab order and names
// fell back to the archive's. The Open XML SDK maps the two the same way when reading.
const SPREADSHEETML_NS = new Set([
    'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'http://purl.oclc.org/ooxml/spreadsheetml/main',
])
const OFFICE_RELS_NS = new Set([
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'http://purl.oclc.org/ooxml/officeDocument/relationships',
])
// The .rels grammar is OPC (29500-2), which Strict does not re-home — the flavours differ in the
// part markup, not in the package. Relationship TYPE values do move, so NON_WORKSHEET_REL lists both
// spellings of each one explicitly.
const PACKAGE_RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
// The content-types part has its own OPC namespace, distinct from the relationships one above. Read
// by two separate scans of [Content_Types].xml — one resolving worksheet parts, one classifying
// them — which each used to declare it locally; they must agree, so there is one of it.
const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'

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
const parseWorkbookParts = async (workbookXml: string, relsXml: string, contentTypesXml?: string) => {
    const { SaxesParser } = await import('saxes')

    // EVERY <sheet> is recorded, including one missing its name or its r:id. Those cannot be
    // resolved, but the caller has to know the workbook DECLARED something it could not place —
    // dropping them here would make an unresolvable sheet indistinguishable from an orphan part.
    const declared: { name: string; rId?: string }[] = []
    const workbook = new SaxesParser({ xmlns: true, resolvePrefix: () => UNBOUND_NAMESPACE })
    // Parent tracking, so <sheet> counts only as a child of <sheets>.
    const open: { uri: string; local: string }[] = []
    workbook.on('opentag', (tag) => {
        assertXmlDepth(open.length + 1)
        const parent = open[open.length - 1]
        open.push({ uri: tag.uri, local: tag.local })
        if (!SPREADSHEETML_NS.has(tag.uri) || tag.local !== 'sheet') return
        if (parent === undefined || !SPREADSHEETML_NS.has(parent.uri) || parent.local !== 'sheets') return
        // r:id is namespace-qualified; `name` and the sibling `sheetId` are not, so neither can be
        // mistaken for it however the producer bound its prefixes.
        const rId = Object.values(tag.attributes).find((a) => OFFICE_RELS_NS.has(a.uri) && a.local === 'id')?.value
        declared.push({ name: tag.attributes.name?.value ?? '', rId })
    })
    workbook.on('closetag', () => void open.pop())
    workbook.write(workbookXml).close()

    const targets = new Map<string, WorkbookRel>()
    const rels = new SaxesParser({ xmlns: true, resolvePrefix: () => UNBOUND_NAMESPACE })
    let relsDepth = 0
    rels.on('opentag', (tag) => {
        assertXmlDepth(++relsDepth)
        if (tag.uri !== PACKAGE_RELS_NS || tag.local !== 'Relationship') return
        const value = (name: string) => tag.attributes[name]?.value
        const [Id, Target, Type, TargetMode] = [value('Id'), value('Target'), value('Type'), value('TargetMode')]
        // An external target points outside the package, so it is never an entry we hold.
        if (Id && Target && TargetMode !== 'External') targets.set(Id, { target: Target, type: Type ?? '' })
    })
    rels.on('closetag', () => void relsDepth--)
    rels.write(relsXml).close()

    // When readable, package content types corroborate a worksheet relationship before a custom
    // target is accepted. Missing/malformed/over-cap metadata is unknown rather than a negative:
    // the relationship can retain its present target, and the rescue below keeps conventional parts
    // too, so an untrusted declaration cannot silently replace their rows.
    const worksheetParts = new Set<string>()
    const worksheetExtensions = new Set<string>()
    const overriddenParts = new Set<string>()
    let contentTypesKnown = false
    if (contentTypesXml !== undefined) {
        const contentTypes = new SaxesParser({
            xmlns: true,
            resolvePrefix: () => UNBOUND_NAMESPACE,
        })
        let contentTypesDepth = 0
        contentTypes.on('opentag', (tag) => {
            assertXmlDepth(++contentTypesDepth)
            if (tag.uri !== CONTENT_TYPES_NS) return
            if (tag.local === 'Override') {
                const partName = tag.attributes.PartName?.value
                if (!partName) return
                const part = opcKey(partName)
                if (part === undefined) return
                overriddenParts.add(part)
                const contentType = tag.attributes.ContentType?.value?.trim().toLowerCase()
                if (contentType && WORKSHEET_CONTENT_TYPE.has(contentType)) worksheetParts.add(part)
            } else if (tag.local === 'Default') {
                const contentType = tag.attributes.ContentType?.value?.trim().toLowerCase()
                if (!contentType || !WORKSHEET_CONTENT_TYPE.has(contentType)) return
                const extension = tag.attributes.Extension?.value?.replace(/^\./, '')
                const foldedExtension = extension === undefined ? undefined : asciiFold(extension)
                if (foldedExtension) worksheetExtensions.add(foldedExtension)
            }
        })
        contentTypes.on('closetag', () => void contentTypesDepth--)
        try {
            contentTypes.write(contentTypesXml).close()
            contentTypesKnown = true
        } catch {
            // Workbook relationships remain usable when this optional authorization source is
            // malformed. Discard any declarations observed before the parse error so partial
            // metadata cannot be mistaken for a complete negative answer.
            worksheetParts.clear()
            worksheetExtensions.clear()
            overriddenParts.clear()
        }
    }

    return { declared, targets, worksheetParts, worksheetExtensions, overriddenParts, contentTypesKnown }
}

// A Target is relative to the rels part's base — xl/ — but may legally be an absolute package path
// or climb out with '..'. exceljs compares the raw string against one form and misses every other;
// this maps all of them onto the archive's entry name. undefined = it escapes the package.
const resolveRelTarget = (target: string): string | undefined => {
    return canonicalOpcPartName(target, 'xl/')
}

// When workbook.xml or its relationships exceed the resolver's materialization ceiling, retain the
// archive-order fallback but augment it with custom worksheet parts identified by whichever OPC
// metadata remains readable. A valid custom part is declared by an exact worksheet relationship,
// an exact worksheet content type, or both; conventional xl/worksheets parts remain the baseline.
const metadataFallbackWorksheets = async (
    entries: ZipEntry[],
    relsXml?: Buffer,
    contentTypesXml?: Buffer
): Promise<WorkbookSheet[]> => {
    const { SaxesParser } = await import('saxes')
    const candidateNames = new Set(
        worksheetEntries(entries).flatMap((entry) => {
            const key = opcKey(entry.name)
            return key === undefined ? [] : [key]
        })
    )
    const relationshipCandidates = new Set<string>()
    const contentTypeCandidates = new Set<string>()
    const worksheetExtensions = new Set<string>()
    const overriddenParts = new Set<string>()
    let contentTypesKnown = false

    if (relsXml !== undefined) {
        const rels = new SaxesParser({ xmlns: true, resolvePrefix: () => UNBOUND_NAMESPACE })
        let depth = 0
        rels.on('opentag', (tag) => {
            assertXmlDepth(++depth)
            if (tag.uri !== PACKAGE_RELS_NS || tag.local !== 'Relationship') return
            const type = tag.attributes.Type?.value
            const target = tag.attributes.Target?.value
            if (!type || !target || tag.attributes.TargetMode?.value === 'External' || !WORKSHEET_REL.has(type)) return
            const resolved = resolveRelTarget(target)
            if (resolved !== undefined) relationshipCandidates.add(asciiFold(resolved))
        })
        rels.on('closetag', () => void depth--)
        try {
            rels.write(relsXml.toString('utf8')).close()
        } catch {
            relationshipCandidates.clear()
        }
    }

    if (contentTypesXml !== undefined) {
        const contentTypes = new SaxesParser({
            xmlns: true,
            resolvePrefix: () => UNBOUND_NAMESPACE,
        })
        let depth = 0
        contentTypes.on('opentag', (tag) => {
            assertXmlDepth(++depth)
            if (tag.uri !== CONTENT_TYPES_NS) return
            const contentType = tag.attributes.ContentType?.value?.trim().toLowerCase()
            if (tag.local === 'Override') {
                const name = tag.attributes.PartName?.value
                if (!name) return
                const part = opcKey(name)
                if (part === undefined) return
                overriddenParts.add(part)
                if (contentType && WORKSHEET_CONTENT_TYPE.has(contentType)) contentTypeCandidates.add(part)
            } else if (tag.local === 'Default') {
                if (!contentType || !WORKSHEET_CONTENT_TYPE.has(contentType)) return
                const extension = tag.attributes.Extension?.value?.replace(/^\./, '')
                if (extension) worksheetExtensions.add(asciiFold(extension))
            }
        })
        contentTypes.on('closetag', () => void depth--)
        try {
            contentTypes.write(contentTypesXml.toString('utf8')).close()
            contentTypesKnown = true
        } catch {
            contentTypeCandidates.clear()
            worksheetExtensions.clear()
            overriddenParts.clear()
        }
    }

    for (const entry of entries) {
        const name = opcKey(entry.name)
        if (name === undefined) continue
        if (XLSX_READER_CONTROL_PARTS.has(name)) continue
        const dot = name.lastIndexOf('.')
        const slash = name.lastIndexOf('/')
        const extension = dot > slash ? name.slice(dot + 1) : ''
        const contentTyped =
            contentTypeCandidates.has(name) || (!overriddenParts.has(name) && worksheetExtensions.has(extension))
        if (
            contentTypeCandidates.has(name) ||
            (relationshipCandidates.has(name) && (!contentTypesKnown || contentTyped))
        ) {
            candidateNames.add(name)
        }
    }

    return entries
        .filter((entry) => {
            const key = opcKey(entry.name)
            return key !== undefined && candidateNames.has(key)
        })
        .map((entry) => ({ entry, name: partFallbackName(entry.name) }))
}

// The workbook's own worksheets, in tab order. undefined when it cannot say — a missing or unreadable
// workbook.xml / rels part. Deliberately not an error: the streaming reader degrades on exactly that
// input (it never sets this.workbookRels, so every sheet takes the spool path) yet still reads to
// completion, so failing here would turn a workbook that works today into a `failed`. archiveWorksheets
// is the fallback, and it is what main did for every workbook.
const workbookWorksheets = async (entries: ZipEntry[]): Promise<WorkbookSheet[] | undefined> => {
    const metadata = new Map<string, ZipEntry>()
    for (const entry of entries) {
        const key = opcKey(entry.name)
        if (key !== undefined && !metadata.has(key)) metadata.set(key, entry)
    }
    const workbookEntry = entries.find((entry) => entry.name === WORKBOOK_PART) ?? metadata.get(asciiFold(WORKBOOK_PART))
    const relsEntry = entries.find((entry) => entry.name === WORKBOOK_RELS_PART) ?? metadata.get(asciiFold(WORKBOOK_RELS_PART))
    const contentTypesEntry =
        entries.find((entry) => entry.name === CONTENT_TYPES_PART) ?? metadata.get(asciiFold(CONTENT_TYPES_PART))
    if (!workbookEntry || !relsEntry) return undefined

    const [workbookXml, relsXml, contentTypesXml] = await Promise.all([
        inflateEntry(workbookEntry),
        inflateEntry(relsEntry),
        contentTypesEntry ? inflateEntry(contentTypesEntry) : undefined,
    ])
    if (!workbookXml || !relsXml) {
        const fallback = await metadataFallbackWorksheets(entries, relsXml, contentTypesXml)
        if (fallback.length === 0) {
            throw new ExtractionFailure('wrong-document-shape')
        }
        return fallback
    }

    let parsed: Awaited<ReturnType<typeof parseWorkbookParts>>
    try {
        parsed = await parseWorkbookParts(
            workbookXml.toString('utf8'),
            relsXml.toString('utf8'),
            contentTypesXml?.toString('utf8')
        )
    } catch {
        return undefined // malformed — exceljs's own parse of the same bytes fails too
    }

    // OPC compares part URIs ASCII-case-insensitively. Keep the actual entry as the value: the target
    // may differ only in case, or an explicit worksheet relationship may point outside the usual
    // xl/worksheets/sheetN.xml convention. reorderForStreaming gives those entries canonical names
    // in its private copy so exceljs's narrower, case-sensitive dispatcher can yield them.
    const present = new Set(metadata.keys())
    const worksheets = worksheetEntries(entries)
    const byPart = metadata

    // Two sets, because a zip name is not a key here either. `claimedNames` answers "did any
    // declaration reference this part?", which is what makes an entry an ORPHAN. `placed` answers
    // "did we lay this entry out?", which is what makes it redundant. The second of two entries
    // sharing a claimed name is neither: referenced, never placed, and not an orphan.
    const claimedNames = new Set<string>()
    const placed = new Set<ZipEntry>()
    const ordered: WorkbookSheet[] = []
    // Declarations we could not place at all. Distinct from a declaration we placed OUTSIDE the
    // worksheets, which is a positive answer and costs nothing.
    let hasUnplaced = false
    for (const { name, rId } of parsed.declared) {
        const rel = rId === undefined ? undefined : parsed.targets.get(rId)
        if (rel === undefined || name === '') {
            hasUnplaced = true // no relationship, an external one, or nothing to call it
            continue
        }
        const resolvedPart = resolveRelTarget(rel.target)
        const part = resolvedPart === undefined ? undefined : asciiFold(resolvedPart)
        // A chartsheet or dialogsheet is a real declaration with no xl/worksheets part for the reader
        // to yield, so it is accounted for and must NOT count as unplaced — otherwise every workbook
        // holding one flips into the rescue below and resurrects genuine orphans.
        //
        // Type and target must name the SAME KIND. Neither signal alone is sound: both the type and
        // the target path are producer-controlled, so either can be made to contradict the other,
        // and every contradiction accepted here leaves the real worksheet unclaimed for the orphan
        // rule to drop. The regression tests hold one fixture per way they can disagree.
        //
        // A path convention is load-bearing here where it could not be for the worksheet case,
        // because the asymmetry inverted: a chart sheet stored somewhere unconventional now reads as
        // unplaced, which costs a rescued orphan in the output. Guessing the other way costs rows.
        const otherFamily = NON_WORKSHEET_REL.get(rel.type)
        if (otherFamily !== undefined) {
            if (part === undefined || !present.has(part) || !part.startsWith(`xl/${otherFamily}/`)) hasUnplaced = true
            continue
        }
        const dot = part?.lastIndexOf('.') ?? -1
        const slash = part?.lastIndexOf('/') ?? -1
        const extension = dot > slash ? asciiFold(part?.slice(dot + 1) ?? '') : undefined
        const authorizedByContentType =
            part !== undefined &&
            (parsed.worksheetParts.has(part) ||
                (!parsed.overriddenParts.has(part) &&
                    extension !== undefined &&
                    parsed.worksheetExtensions.has(extension)))
        // A conventional worksheet path remains recoverable under a missing/unknown type. Outside
        // that convention, require the exact worksheet relationship identity and, when the content-
        // type table was readable, its authorization.
        // Reader-control parts are never eligible even when both declarations are forged together:
        // if styles.xml became the resolved sheet, the real sheetN.xml would otherwise be dropped as
        // an orphan and this same central record would be written twice by the rebuild.
        if (
            part === undefined ||
            (!isWorksheetPart(part) &&
                (!WORKSHEET_REL.has(rel.type) ||
                    (parsed.contentTypesKnown && !authorizedByContentType) ||
                    XLSX_READER_CONTROL_PARTS.has(part)))
        ) {
            hasUnplaced = true
            continue
        }
        claimedNames.add(part)
        const entry = byPart.get(part)
        if (!entry) {
            hasUnplaced = true // declares a worksheet this archive does not hold
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
    // The two metadata declarations that authorize a custom path live in the same untrusted package.
    // If one places a sheet outside the conventional path while conventional parts remain, those
    // parts are no longer provable orphans — a forged relationship + content-type pair could be
    // pointing at theme XML while the real rows sit in sheetN.xml. Rescue them all. A legitimate
    // custom sheet may therefore keep a stale conventional orphan, preferring possible duplicate
    // text over the silent row loss this feature exists to prevent.
    const placedNonConventional = ordered.some((sheet) => !isWorksheetPart(sheet.entry.name))
    const rescued = worksheets.filter(
        (entry) =>
            !placed.has(entry) &&
            (hasUnplaced ||
                placedNonConventional ||
                claimedNames.has(opcKey(entry.name) ?? ''))
    )
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
type Reorder = { ok: true; content: Buffer; sheets: WorkbookSheet[] } | { ok: false; reason: ExtractionReason }

const reorderForStreaming = async (buf: Buffer): Promise<Reorder> => {
    const entries = zipEntries(buf)
    // Unreachable: every zipEntries bail is also a budget rejection except a name length overrunning
    // the buffer, which Invariant 2 catches. Failing closed keeps that from being load-bearing.
    if (!entries) return { ok: false, reason: 'malformed' }

    // A ZIP may carry the same OPC part more than once, but there is no well-defined winner. Passing
    // duplicate controls through is worse than refusing them: ExcelJS consumes every occurrence, so
    // a later empty shared-string table silently turns text cells into missing values, while a later
    // relationship part can make the emitted sheet set disagree with the workbook. Compare with OPC
    // identity (ASCII case-folding), catching exact duplicates and case aliases alike.
    const seenControls = new Set<string>()
    for (const entry of entries) {
        const folded = opcKey(entry.name)
        if (folded === undefined) continue
        if (!XLSX_READER_CONTROL_PARTS.has(folded)) continue
        // Two records claiming the same control part: the reader would take one and the resolver the
        // other, so which sheet names come out would depend on zip order.
        if (seenControls.has(folded)) return { ok: false, reason: 'malformed' }
        seenControls.add(folded)
    }
    // OPC part names compare ASCII-case-insensitively; exceljs's streaming dispatch does not. Give
    // the first case-variant control part its canonical name in the private copy, just as worksheets
    // below are canonicalized. If an exact spelling already exists, leave any invalid case-duplicate
    // alone rather than creating a second entry with the same name.
    const exactControls = new Set(entries.map((entry) => entry.name))
    const canonicalized = new Set<string>()
    const complete = entries.map((entry) => {
        const canonical = XLSX_CANONICAL_CONTROL_PARTS.get(opcKey(entry.name) ?? '')
        if (!canonical || entry.name === canonical || exactControls.has(canonical) || canonicalized.has(canonical)) {
            return entry
        }
        canonicalized.add(canonical)
        return { ...entry, name: canonical }
    })
    if (!complete.some((entry) => entry.name === 'xl/sharedStrings.xml')) {
        complete.push({
            name: 'xl/sharedStrings.xml',
            utf8Name: false,
            method: 0, // stored: the injected table is 153 bytes, so compressing it is noise
            crc: EMPTY_SHARED_STRINGS_CRC,
            compSize: EMPTY_SHARED_STRINGS.length,
            uncompSize: EMPTY_SHARED_STRINGS.length,
            data: EMPTY_SHARED_STRINGS,
        })
    }
    const safeWorkbookRels: ZipEntry = {
        name: WORKBOOK_RELS_PART,
        utf8Name: false,
        method: 0,
        crc: EMPTY_WORKBOOK_RELS_CRC,
        compSize: EMPTY_WORKBOOK_RELS.length,
        uncompSize: EMPTY_WORKBOOK_RELS.length,
        data: EMPTY_WORKBOOK_RELS,
    }
    const workbookRelsIndex = complete.findIndex((entry) => entry.name === WORKBOOK_RELS_PART)
    if (workbookRelsIndex === -1) complete.push(safeWorkbookRels)
    else complete[workbookRelsIndex] = safeWorkbookRels

    const sheets = (await workbookWorksheets(entries)) ?? archiveWorksheets(entries)
    const sheetEntries = new Set(sheets.map((sheet) => sheet.entry))
    // ExcelJS 4.4.0 recognizes only this spelling. Renaming the private streamed copy closes two
    // valid-OPC gaps at once: case-variant part names and relationship-declared custom part paths.
    // Each source ZipEntry is still carried once and its compressed bytes are untouched.
    const streamedSheets = sheets.map((sheet, i) => ({ ...sheet.entry, name: `xl/worksheets/sheet${i + 1}.xml` }))

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
        ...streamedSheets,
        // Everything else — minus every remaining part the reader would dispatch as a worksheet.
        // Orphan sheetN.xml parts no <sheet> references die here, which is both what stops them being
        // emitted as sheets of their own and what makes the Nth emission exactly sheets[N - 1].
        ...complete.filter(
            (entry) =>
                rank(entry.name) === leadingRank &&
                !sheetEntries.has(entry) &&
                !EXCELJS_WORKSHEET_DISPATCH.test(entry.name)
        ),
    ]

    // The EOCD's entry count is 16-bit, where 0xffff means "ZIP64, the real count is elsewhere".
    // zipEntries refuses an archive already declaring it, but the injection above can carry a
    // 65534-entry archive onto it, writing a sentinel where a count belongs. Reachable inside
    // MAX_INPUT_BYTES: ~76 bytes of headers per entry, so 65534 fit in ~5 MB.
    //
    // Counted on `ordered`, which is what the EOCD below is actually written from. Counting `complete`
    // instead refused archives whose rewrite lands well under the sentinel, since the third group
    // drops entries — the count that mattered was never the one being checked.
    if (ordered.length >= 0xffff) {
        return { ok: false, reason: 'unsupported-zip-feature' }
    }

    const locals: Buffer[] = []
    const centrals: Buffer[] = []
    let offset = 0
    for (const entry of ordered) {
        const name = Buffer.from(entry.name, entry.utf8Name ? 'utf8' : 'latin1')
        const flags = entry.utf8Name ? 0x0800 : 0

        const local = Buffer.alloc(LOCAL_HEADER_BYTES)
        local.writeUInt32LE(LOCAL_SIG, 0)
        local.writeUInt16LE(ZIP_VERSION, 4)
        local.writeUInt16LE(flags, 6)
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
        central.writeUInt16LE(flags, 8)
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
//   - w:sdt keeps its FIRST w:sdtContent, and mc:AlternateContent its FIRST mc:Fallback. Their
//     property/choice siblings are absent from the whitelist, while selection state below prevents
//     malformed packages from concatenating a second eligible child.
//
// Fidelity is pinned by tests/docx-fidelity.test.ts against mammoth itself over a real Word corpus.
// The deliberate divergences are listed at DOCX_CONTAINERS.

// DOCX_MAIN_PART is declared with the format magic above, where routing needs it first.

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

// Any prefix that reaches saxes undeclared resolves HERE, and this URI is deliberately absent from
// OOXML_PREFIXES — so the element names `{urn:agentextract:unbound}local`, misses the whitelist, and
// its subtree is dropped. Which is precisely what mammoth does with an unmapped namespace
// (xml/reader.js:53-66 emits the same `{uri}local` shape, and no handler matches it).
//
// It exists because saxes fails the whole parse on an unbound prefix — for ATTRIBUTES too
// (saxes.js:1920-1925) — where xmldom shrugged. That is not a rare shape: Word 2013 and later stamp
// w15:paraId on EVERY w:p, and w16cid:durableId alongside it, so a document that carries those
// without their xmlns (a fragment assembled by a templating tool, a repaired file) used to lose
// everything after its first paragraph. An allowlist cannot close that — there is always another
// prefix — so this resolves the open class instead, and the map below shrinks to the only entries
// that do something an allowlist has to do.
const UNBOUND_NAMESPACE = 'urn:agentextract:unbound'

// The three prefixes where GUESSING is better than dropping, because these are the only URIs
// OOXML_PREFIXES maps: an undeclared `w:` resolved to the sentinel above would take the whole
// document body with it. Every other prefix a real document leaves undeclared — o:, w10:, wne:, r:,
// wp:, a:, pic:, wps:, w14: — resolves to a URI this reader does not map either way, so guessing it
// and sentinelling it are the same outcome, and the sentinel needs no maintenance.
// An in-document xmlns still shadows these (saxes.resolve checks scope first, then this map, then
// the sentinel), so a strict-format file resolves to the strict URI and routes correctly.
const OOXML_ASSUMED_PREFIXES: Record<string, string> = {
    w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    v: 'urn:schemas-microsoft-com:vml',
}

// Recurse into these; emit nothing of their own. Absent from the list = subtree dropped.
const DOCX_CONTAINERS = new Set([
    'w:r',
    'w:hyperlink',
    'w:ins',
    'w:smartTag',
    'w:tbl',
    'w:tr',
    'w:tc',
    'w:sdt',
    'mc:AlternateContent',
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
// Deviations from mammoth, added to tests/docx-fidelity.test.ts's ACCEPTED_DIVERGENCES if they
// ever show up on a real document: a w:sdt carrying wordml:checkbox has its first text character
// REPLACED by a checkbox node in mammoth (we keep the character — strictly better, and costs nothing
// here); a w:t holding a comment or a nested element makes mammoth's text() throw "Not implemented"
// (we extract); the main part is read at
// its conventional path rather than resolved through _rels/.rels, which routing already requires;
// and a deleted-mark paragraph whose stash never reaches a surviving paragraph is retained here
// while mammoth drops it. That last direction is deliberately availability-preserving.

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
// w:p is silently lost. Paragraph/picture frames reproduce that ordering; table/cell frames add a
// bounded choice between Mammoth's normal vertical-merge filtering and its malformed-table bail-out.
interface DocxFrame {
    kind: 'root' | 'paragraph' | 'picture' | 'table' | 'cell'
    value: string
    extra: string
    // Extras from deleted paragraphs are claimed by the next paragraph when it OPENS. Keeping the
    // ownership on that frame prevents a nested text-box paragraph from stealing them on close.
    claimedExtra: string
    allValue?: string
    orphanPicture?: boolean
    discardedValue?: boolean
    capCrossedBeforeDeferredExtra?: boolean
}

interface DocxTableCell {
    depth: number
    column: number
    gridSpan: number
    vMerge: boolean | null
    finalized: boolean
    suppressed: boolean
    frame: DocxFrame
    chars: number
    tcPrSeen: boolean
    tcPrDepth?: number
    gridSpanSeen: boolean
    vMergeSeen: boolean
}

interface DocxTable {
    columns: Set<number>
    column: number
    frame: DocxFrame
    mergeBailout: boolean
    deferredChars: number
    rowDepth?: number
    cell?: DocxTableCell
}

interface PendingDeletedParagraph {
    value: string
    extra: string
    chars: number
}


interface DocxReader {
    write: (chunk: string) => void // one decoded slice of document.xml; THROWS on malformed XML
    chars: () => number // characters emitted so far — what the handler's cap check reads
    overCap: () => boolean
    shouldStop: () => boolean // cap crossed and any output-order-sensitive frame has closed
    end: (tail: string) => boolean // flush/close; whether the part contained its content container
    text: () => string // the text, complete or partial
}

const createDocxReader = async (maxOutputChars: number, shape: DocxPartShape = DOCX_MAIN_SHAPE): Promise<DocxReader> => {
    // Lazy like every other parser here: a Lambda that only ever sees text never loads saxes.
    const { SaxesParser } = await import('saxes')
    const { hex: dingbatHex } = await import('dingbat-to-unicode')

    const stack: string[] = [] // qualified names of every open element, innermost last
    const frames: DocxFrame[] = [{ kind: 'root', value: '', extra: '', claimedExtra: '' }] // root extras are lost
    const deleted: boolean[] = [false] // one flag per open w:p; index 0 pairs with the root frame
    const selectors: { depth: number; parent: string; target: string; selected: boolean }[] = []
    const tables: DocxTable[] = []
    let top = frames[frames.length - 1]
    let skip = -1 // stack index where the dropped subtree began, or -1 when we're reading
    let rowDeleted = false // a w:trPr said its row is deleted; act on it once that w:trPr closes
    let sawBody = false
    let inBody = false
    let chars = 0
    // Mammoth holds a deleted-mark paragraph until the following paragraph is read. Its value
    // merges into that paragraph, while any text-box `extra` still follows the completed merged
    // paragraph. Both stay bounded here because the next paragraph may cross a suppressed cell.
    let pendingDeleted: PendingDeletedParagraph = { value: '', extra: '', chars: 0 }
    // Descendant deleted paragraphs are parsed eagerly here, while Mammoth sees them only after an
    // outer deleted paragraph is claimed. Preserve that delayed order: closing the outer paragraph
    // puts its stash first and queues the descendant state for the paragraph after the claimant.
    const pendingDeletedQueue: PendingDeletedParagraph[] = []
    let capExceeded = false
    let drainUntil: DocxFrame | undefined
    let drainAncestors: Set<DocxFrame> | undefined
    let drainPendingDeletedContent = false
    let discardDeferredText = false
    let deferredFreshFrames: Set<DocxFrame> | undefined
    let discardTableText = false
    // Exactly the cap: nothing is stored past what can be returned. The entry point strips a
    // trailing surrogate half unconditionally, so stopping flush against the cap cannot split a
    // character, and truncation is reported from capExceeded rather than inferred from overshoot.
    const storageLimit = maxOutputChars

    const appendRaw = (frame: DocxFrame, field: 'value' | 'extra' | 'claimedExtra' | 'allValue', text: string): void => {
        const value = frame[field] ?? ''
        const room = storageLimit - value.length
        if (room > 0) frame[field] = value + text.slice(0, room)
    }

    const append = (
        frame: DocxFrame,
        field: 'value' | 'extra' | 'claimedExtra',
        text: string,
        options: { source?: DocxFrame; deferred?: boolean } = {}
    ): void => {
        // A frame close must still move text retained before the cap into its parent. Carrying the
        // source separately from fresh appends prevents body-final stashes and newly generated
        // separators from using that privilege. Deferred commits are retained only on the picture
        // drain path; a table drain drops them because missing table text sorts before every extra.
        // Once the drain boundary closes, only frames that were already open at that boundary may
        // finish propagating their retained prefix through those same ancestors. Elements parsed
        // later in the same SAX chunk cannot append either fresh text or frame-close separators.
        if (
            drainAncestors &&
            (options.source === undefined || !drainAncestors.has(options.source) || !drainAncestors.has(frame))
        ) {
            return
        }
        if (discardTableText && (options.source === undefined || options.deferred)) return
        if (discardDeferredText && options.source === undefined) {
            // A nested picture may cross the cap before direct text that appears ahead of it after
            // Mammoth hoists the picture. Retain fresh value on frames that were ancestors of that
            // exact picture, plus direct paragraph text parsed after the picture closes but before
            // its enclosing paragraph closes. That direct text also sorts before the picture.
            const directParagraphPrefix =
                frame.kind === 'paragraph' &&
                drainUntil?.kind === 'paragraph' &&
                frames.includes(drainUntil) &&
                !frames.some((candidate) => candidate.kind === 'picture')
            if (options.deferred || field !== 'value' || (!deferredFreshFrames?.has(frame) && !directParagraphPrefix)) {
                if (field === 'value' && text.length > 0) frame.discardedValue = true
                return
            }
        }
        appendRaw(frame, field, text)
        if (frame.kind === 'table' && field === 'value') appendRaw(frame, 'allValue', text)
    }

    const charge = (length: number): void => {
        // Mammoth still walks an orphan picture, so a deleted paragraph inside it can affect the
        // next body paragraph, but the picture's own output is unreachable until that happens.
        if (frames.some((frame) => frame.orphanPicture)) return
        const activeCell = [...tables].reverse().find((table) => table.cell !== undefined)?.cell
        if (activeCell) {
            activeCell.chars += length
            // A cell is speculative until it closes: a deleted paragraph can uncharge its text,
            // and vertical-merge suppression can remove the entire cell. Latching truncation here
            // loses later document text even when neither candidate survives. The cell-close path
            // charges only the committed variant; tables are cap-clipped while that decision waits.
            return
        }
        chars += length
        if (capExceeded || chars <= maxOutputChars) return
        capExceeded = true
        // Text inside a picture is not output where it is read: the picture is hoisted behind the
        // enclosing paragraph. Continue only until that paragraph closes, so the cap returns the
        // box text in its real output position without reading the rest of the document.
        const outerPicture = frames.findIndex((frame) => frame.kind === 'picture')
        if (outerPicture > 0) {
            const innerPicture = frames.findLastIndex((frame) => frame.kind === 'picture')
            frames[innerPicture].capCrossedBeforeDeferredExtra = true
            deferredFreshFrames = new Set(frames.slice(0, innerPicture))
            drainUntil = [...frames.slice(0, outerPicture)].reverse().find((frame) => frame.kind === 'paragraph')
            // Picture text is deferred until after the enclosing paragraph. Once it crosses the
            // output cap, stop accepting more deferred text while the paragraph drains for direct
            // text that sorts before it. Already-bounded frame content can still propagate at close
            // and supply the useful prefix of the first text box.
            discardDeferredText = true
        } else {
            // Table output is withheld until the table closes because a later non-row/non-cell
            // child can make Mammoth cancel vertical-merge suppression retroactively. Drain the
            // outermost table so the returned text uses the table's final, known variant.
            drainUntil = frames.find((frame) => frame.kind === 'table')
            discardTableText = drainUntil !== undefined
        }
    }

    const uncharge = (length: number): void => {
        if (frames.some((frame) => frame.orphanPicture)) return
        const activeCell = [...tables].reverse().find((table) => table.cell !== undefined)?.cell
        if (activeCell) activeCell.chars = Math.max(0, activeCell.chars - length)
        else chars = Math.max(0, chars - length)
    }

    const emit = (text: string): void => {
        append(top, 'value', text)
        charge(text.length)
    }

    const emitSymbol = (attributes: Record<string, string | { uri?: string; local?: string; value: string }>): void => {
        const font = wordAttribute(attributes, 'font')
        const char = wordAttribute(attributes, 'char')
        if (!font || !char) return
        const mapped = dingbatHex(font, char) ?? (/^F0..$/.test(char) ? dingbatHex(font, char.slice(2)) : undefined)
        if (mapped) emit(mapped.string)
    }

    const currentTable = (): DocxTable | undefined => tables[tables.length - 1]

    const finalizeTableCell = (): DocxTableCell | undefined => {
        const table = currentTable()
        const cell = table?.cell
        if (!table || !cell || cell.finalized) return cell
        cell.finalized = true
        if (cell.vMerge === true && table.columns.has(cell.column)) cell.suppressed = true
        else table.columns.add(cell.column)
        return cell
    }

    // calculateRowSpans() bails out when the flattened children of a table are not all rows, or a
    // row's are not all cells. These are the handlers that directly create a document node rather
    // than an ignored property/range end. Recursive wrappers are decided by the children they
    // expose, so they are deliberately absent here.
    const TABLE_STRUCTURAL_NODES = new Set([
        'w:p',
        'w:r',
        'w:t',
        'w:tab',
        'w:noBreakHyphen',
        'w:softHyphen',
        'w:sym',
        'w:hyperlink',
        'w:tbl',
        'w:tr',
        'w:tc',
        'w:footnoteReference',
        'w:endnoteReference',
        'w:commentReference',
        'w:br',
        'w:bookmarkStart',
    ])

    const openSelector = (name: string): void => {
        if (name === 'w:sdt') selectors.push({ depth: stack.length, parent: name, target: 'w:sdtContent', selected: false })
        else if (name === 'mc:AlternateContent')
            selectors.push({ depth: stack.length, parent: name, target: 'mc:Fallback', selected: false })
    }

    const selectorAllows = (name: string): boolean => {
        const selector = selectors[selectors.length - 1]
        if (!selector || stack.length !== selector.depth + 1) return true
        if (name !== selector.target || selector.selected) return false
        selector.selected = true
        return true
    }

    const closeSelector = (name: string): void => {
        const selector = selectors[selectors.length - 1]
        if (selector?.parent === name && selector.depth === stack.length + 1) selectors.pop()
    }

    const parser = new SaxesParser({
        xmlns: true,
        additionalNamespaces: OOXML_ASSUMED_PREFIXES,
        // Consulted only after scope and the map above (saxes.js:1845-1862), so this catches exactly
        // the prefixes nothing else claimed, and turns "unbound prefix" from a fatal parse error into
        // an unmapped namespace — which this reader already knows how to drop.
        resolvePrefix: () => UNBOUND_NAMESPACE,
    })

    parser.on('opentag', (tag) => {
        // A direct port of mammoth's convertName (xml/reader.js:53-66): mapped URI -> `w:t`,
        // unmapped -> `{uri}local`, an explicitly empty default namespace -> the bare local name.
        // With no default declaration saxes asks resolvePrefix(''), which returns the unbound URI
        // above. Only a mapped name can match a table; every other form deliberately misses.
        const uri = tag.uri ?? ''
        const prefix = Object.prototype.hasOwnProperty.call(OOXML_PREFIXES, uri) ? OOXML_PREFIXES[uri] : undefined
        const name = uri === '' ? tag.local : prefix === undefined ? `{${uri}}${tag.local}` : `${prefix}:${tag.local}`
        stack.push(name)
        assertXmlDepth(stack.length)

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

        // mammoth reads the w:body child, not every paragraph anywhere under w:document. Preserve
        // the root only as a path to that child; any sibling before or after the first body is a
        // malformed-document subtree and is dropped without affecting the body itself.
        if (!inBody) {
            // The root is kept only as a path to the container below it; for a header or footer the
            // two are the same element and there is no such step. `repeats` is what separates the
            // one-body parts from a footnotes part, where every w:footnote opens a fresh region.
            if (name === shape.root && stack.length === 1 && shape.content !== shape.root) return
            if (
                name === shape.content &&
                stack.length === shape.contentDepth &&
                (shape.contentDepth === 1 || stack[0] === shape.root) &&
                (shape.repeats || !sawBody)
            ) {
                sawBody = true
                // A separator note is structure — counted as seen, so the part still reads as
                // well-formed, but not entered, so it contributes no paragraph.
                if (shape.skipItem?.(tag.attributes) ?? false) {
                    skip = stack.length - 1
                    return
                }
                inBody = true
                return
            }
            skip = stack.length - 1
            return
        }

        // firstOrEmpty semantics are selection, not merely recursion. A malformed producer may put
        // two sdtContent/Fallback children under one parent; Mammoth reads the first and ignores the
        // rest. Non-target siblings (sdtPr/Choice included) are dropped by the same gate.
        const selector = selectors[selectors.length - 1]
        const selectedContainer =
            selector !== undefined &&
            selector.depth + 1 === stack.length &&
            selector.target === name &&
            !selector.selected
        if (!selectorAllows(name)) {
            skip = stack.length - 1
            return
        }

        const parent = stack[stack.length - 2]
        const openTable = currentTable()
        if (openTable && openTable.cell === undefined && TABLE_STRUCTURAL_NODES.has(name)) {
            const goBack = name === 'w:bookmarkStart' && wordAttribute(tag.attributes, 'name') === '_GoBack'
            const expected = openTable.rowDepth === undefined ? 'w:tr' : 'w:tc'
            if (!goBack && name !== expected) {
                openTable.mergeBailout = true
            }
        }

        const enclosingCell = openTable?.cell
        if (
            enclosingCell?.tcPrDepth !== undefined &&
            stack.length === enclosingCell.tcPrDepth + 1 &&
            parent === 'w:tcPr' &&
            name !== 'w:gridSpan' &&
            name !== 'w:vMerge'
        ) {
            // Test against the enclosing table captured before any structural branch can push a
            // nested table and replace currentTable(). tcPr is metadata-only; every direct child
            // except the two properties above is dropped with its entire subtree.
            skip = stack.length - 1
            return
        }

        if (name === 'w:tbl') {
            const frame: DocxFrame = { kind: 'table', value: '', allValue: '', extra: '', claimedExtra: '' }
            top = frame
            frames.push(frame)
            tables.push({
                columns: new Set(),
                column: 0,
                frame,
                mergeBailout: false,
                deferredChars: 0,
            })
        } else if (name === 'w:tr') {
            const table = currentTable()
            // A malformed row nested inside a cell is content of that cell, not a new row in the
            // enclosing table. Let wrapped rows recurse, but only while this table owns no cell.
            if (table && table.cell === undefined) {
                table.column = 0
                table.rowDepth = stack.length
            }
        } else if (name === 'w:tc') {
            const table = currentTable()
            if (table && table.cell === undefined) {
                const frame: DocxFrame = { kind: 'cell', value: '', extra: '', claimedExtra: '' }
                top = frame
                frames.push(frame)
                table.cell = {
                    depth: stack.length,
                    column: table.column,
                    gridSpan: 1,
                    vMerge: null,
                    finalized: false,
                    suppressed: false,
                    frame,
                    chars: 0,
                    tcPrSeen: false,
                    gridSpanSeen: false,
                    vMergeSeen: false,
                }
            }
        } else if (name === 'w:tcPr') {
            const cell = currentTable()?.cell
            if (parent !== 'w:tc' || !cell || stack.length !== cell.depth + 1 || cell.tcPrSeen) {
                skip = stack.length - 1
                return
            }
            cell.tcPrSeen = true
            cell.tcPrDepth = stack.length
        } else if (name === 'w:gridSpan') {
            const cell = currentTable()?.cell
            if (!cell || parent !== 'w:tcPr' || cell.tcPrDepth !== stack.length - 1 || cell.gridSpanSeen) {
                skip = stack.length - 1
                return
            }
            cell.gridSpanSeen = true
            const span = Number.parseInt(wordAttribute(tag.attributes, 'val') ?? '', 10)
            if (cell && Number.isFinite(span) && span > 0) cell.gridSpan = span
            skip = stack.length - 1
            return
        } else if (name === 'w:vMerge') {
            const cell = currentTable()?.cell
            if (!cell || parent !== 'w:tcPr' || cell.tcPrDepth !== stack.length - 1 || cell.vMergeSeen) {
                skip = stack.length - 1
                return
            }
            cell.vMergeSeen = true
            const value = wordAttribute(tag.attributes, 'val')
            if (cell) cell.vMerge = value === undefined || value === '' || value === 'continue'
            skip = stack.length - 1
            return
        }

        // w:t is the ONE text-bearing element; the text handler recognizes it off the stack top, so
        // it needs no flag of its own. Deliberately before the container test, so its children —
        // illegal anyway — fall through to the drop below exactly as they do in mammoth.
        if (name === 'w:t') return
        // These elements recurse only as the selected direct child of their owning selector.
        // Orphan forms are unknown elements in Mammoth and their entire subtree is dropped.
        if (selectedContainer) return
        if (name === 'w:sym') {
            emitSymbol(tag.attributes)
            skip = stack.length - 1 // the symbol is attribute-only; illegal children stay ignored
            return
        }

        if (name === 'w:p' || name === 'w:pict') {
            if (name === 'w:pict') {
                // A picture opened after an enclosing picture crossed the cap will be hoisted
                // ahead of that enclosing picture's retained value. Since fresh deferred text is
                // no longer accepted, emitting the retained value alone would create a hole.
                for (const [index, frame] of frames.entries()) {
                    const hasOpenParagraph = frames
                        .slice(index + 1)
                        .some((candidate) => candidate.kind === 'paragraph')
                    if (frame.kind === 'picture' && frame.capCrossedBeforeDeferredExtra && !hasOpenParagraph) {
                        frame.discardedValue = true
                    }
                }
            }
            const claimedValue = name === 'w:p' ? pendingDeleted.value : ''
            const claimedExtra = name === 'w:p' ? pendingDeleted.extra : ''
            const claimedChars = name === 'w:p' ? pendingDeleted.chars : 0
            if (name === 'w:p') {
                pendingDeleted = pendingDeletedQueue.shift() ?? { value: '', extra: '', chars: 0 }
            }
            const orphanPicture =
                name === 'w:pict' &&
                !frames.some((frame) => frame.kind === 'paragraph' || frame.orphanPicture)
            top = {
                kind: name === 'w:p' ? 'paragraph' : 'picture',
                value: claimedValue,
                extra: '',
                claimedExtra,
                orphanPicture,
                // Fresh paragraph output opened after a picture cap is later than retained output
                // whose propagation is still draining. Mark even an empty paragraph now; otherwise
                // its close-only separator can survive text that was already discarded.
                discardedValue:
                    name === 'w:p' && discardDeferredText && frames.some((frame) => frame.kind === 'picture'),
            }
            frames.push(top)
            if (name === 'w:p') {
                deleted.push(false)
                // The deleted paragraph's contents were uncharged from their old container when
                // stashed. Charge them to the next paragraph's actual cell/global owner now.
                charge(claimedChars)
                if (drainPendingDeletedContent) {
                    drainUntil = top
                    drainPendingDeletedContent = false
                }
            }
            return
        }
        if (DOCX_CONTAINERS.has(name) || name === 'w:tcPr') {
            openSelector(name)
            return
        }
        const literal = Object.prototype.hasOwnProperty.call(DOCX_LITERALS, name) ? DOCX_LITERALS[name] : undefined
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
            if (name === 'w:tr') {
                const table = currentTable()
                if (table?.rowDepth === stack.length + 1) table.rowDepth = undefined
            }
            return
        }
        if (name === shape.content && stack.length === shape.contentDepth - 1 && inBody) {
            // Availability-preserving divergence retained from the previous reader: Mammoth drops
            // a final deleted-mark paragraph because nothing claims its stash; attachment
            // extraction keeps its plain text. Extras remain deferred and therefore absent.
            if (pendingDeleted.value !== '') {
                append(top, 'value', pendingDeleted.value)
                charge(pendingDeleted.value.length)
                pendingDeleted = { value: '', extra: '', chars: 0 }
                pendingDeletedQueue.length = 0
            }
            inBody = false
            return
        }

        if (name === 'w:tcPr') {
            const cell = currentTable()?.cell
            if (cell?.tcPrDepth === stack.length + 1) {
                cell.tcPrDepth = undefined
                finalizeTableCell()
            }
            return
        }
        if (name === 'w:tc') {
            const table = currentTable()
            if (table?.cell && stack.length !== table.cell.depth - 1) return
            const cell = finalizeTableCell()
            if (table && cell) {
                frames.pop()
                top = frames[frames.length - 1]
                if (cell.frame.discardedValue) table.frame.discardedValue = true
                appendRaw(table.frame, 'allValue', cell.frame.value)
                if (cell.suppressed) table.deferredChars += cell.chars
                else {
                    appendRaw(table.frame, 'value', cell.frame.value)
                    // Clear before charging so a nested table commits into its enclosing cell,
                    // while an outermost table commits into the extraction's global count.
                    table.cell = undefined
                    charge(cell.chars)
                }
                table.column += cell.gridSpan
                table.cell = undefined
            }
            return
        }
        if (name === 'w:tr') {
            const table = currentTable()
            if (table?.rowDepth === stack.length + 1) table.rowDepth = undefined
            return
        }
        if (name === 'w:tbl') {
            const table = tables.pop()
            if (table) {
                frames.pop()
                top = frames[frames.length - 1]
                if (table.frame.discardedValue) {
                    if (top.kind !== 'root') top.discardedValue = true
                    if (drainUntil === table.frame) {
                        drainUntil = undefined
                        drainAncestors = new Set(frames)
                    }
                    return
                }
                const value = table.mergeBailout ? table.frame.allValue ?? '' : table.frame.value
                append(top, 'value', value, { source: table.frame })
                append(top, 'extra', table.frame.extra, { source: table.frame, deferred: true })
                if (table.mergeBailout) charge(table.deferredChars)
                if (drainUntil === table.frame) {
                    drainUntil = undefined
                    drainAncestors = new Set(frames)
                }
            }
            return
        }

        if (name !== 'w:p' && name !== 'w:pict') {
            if (name !== undefined) closeSelector(name)
            return
        }
        const frame = top
        frames.pop()
        top = frames[frames.length - 1]
        let closedDeleted = false
        if (name === 'w:pict') {
            // Hoist: the picture's own text becomes the parent's `extra`, behind any extra it
            // already carried. It reaches the output only if some ancestor w:p reinserts it.
            // These strings were retained before the cap crossed; moving them upward does not add
            // new deferred content, and keeps the useful prefix of the first text box.
            if (!frame.orphanPicture && !frame.discardedValue) {
                append(top, 'extra', frame.extra, { source: frame, deferred: true })
                append(top, 'extra', frame.value, { source: frame, deferred: true })
            }
        } else {
            const wasDeleted = deleted.pop() ?? false
            closedDeleted = wasDeleted
            if (wasDeleted) {
                // Mammoth stashes the paragraph's XML children globally, then prepends them to the
                // next paragraph even when it crosses a table-cell boundary. Holding the rendered
                // pair does the same for raw text and prevents a vMerge continuation from deleting
                // content that belongs to the following visible paragraph.
                const deferredChars = frame.value.length + frame.claimedExtra.length + frame.extra.length
                uncharge(deferredChars)
                if (pendingDeleted.value !== '' || pendingDeleted.extra !== '' || pendingDeleted.chars > 0) {
                    pendingDeletedQueue.unshift(pendingDeleted)
                }
                pendingDeleted = {
                    value: frame.value,
                    extra: `${frame.claimedExtra}${frame.extra}`.slice(0, storageLimit),
                    chars: deferredChars,
                }
            } else if (frame.discardedValue) {
                // A missing child is a hole before this paragraph's own separator. Carry that fact
                // through same-order containers. A picture is the ordering boundary: it drops its
                // output without invalidating direct text that precedes it in the parent paragraph.
                if (top.kind !== 'root' && top.kind !== 'picture') top.discardedValue = true
            } else if (
                frames.some((candidate) => candidate.orphanPicture) &&
                !deleted.some((candidate) => candidate)
            ) {
                // This paragraph is reachable only through an orphan picture. Keep parsing for the
                // deleted-paragraph side effect above, but do not copy ordinary text into every
                // open picture frame — none of those values can reach document output.
            } else {
                append(top, 'value', frame.value, { source: frame })
                append(top, 'value', '\n\n', discardDeferredText ? { source: frame } : undefined)
                append(top, 'value', frame.claimedExtra, { source: frame, deferred: true })
                append(top, 'value', frame.extra, { source: frame, deferred: true })
                charge(2)
            }
        }

        if (drainUntil === frame) {
            drainUntil = undefined
            drainAncestors = new Set(frames)
            if (closedDeleted) drainPendingDeletedContent = true
        }
    })

    return {
        write: (chunk) => void parser.write(chunk),
        chars: () => chars,
        overCap: () => capExceeded,
        shouldStop: () => capExceeded && drainUntil === undefined && !drainPendingDeletedContent,
        end: (tail) => {
            parser.write(tail).close() // close() is the well-formedness check: it throws on an unclosed element
            return sawBody
        },
        // An open picture frame is NOT inline text: once complete, it is hoisted into an enclosing
        // paragraph's `extra` and appears only after that paragraph's break. On truncation, exclude
        // it and every nested frame so the returned text remains a prefix of complete output.
        // Without a picture, concatenating open values is the partial document order and lets one
        // enormous paragraph return what was read. Pending `extra` stays deferred either way.
        text: () => {
            const picture = frames.findIndex((frame) => frame.kind === 'picture')
            const cell = frames.findIndex((frame) => frame.kind === 'cell')
            const boundary = [picture, cell].filter((index) => index >= 0).reduce((a, b) => Math.min(a, b), frames.length)
            const visible = frames.slice(0, boundary)
            let text = ''
            for (const frame of visible) {
                let value = frame.value
                if (frame.kind === 'table') {
                    const all = frame.allValue ?? ''
                    let common = 0
                    while (common < value.length && common < all.length && value[common] === all[common]) common++
                    value = value.slice(0, common)
                }
                text += value.slice(0, storageLimit - text.length)
                if (text.length >= storageLimit) break
            }
            return text
        },
    }
}


const storedDocxChunks = function* (data: Buffer): Iterable<Buffer> {
    for (let offset = 0; offset < data.length; offset += STREAM_SLICE_UNITS) {
        yield data.subarray(offset, offset + STREAM_SLICE_UNITS)
    }
}

// One shape for both storage methods, so the handler's loop has a single form. Stored parts are
// sliced rather than handed over whole, keeping cap/deadline checks enforceable at the same bounded
// granularity as deflate's output. Method 8 yields under backpressure from zlib.
const docxMainPartChunks = (part: ZipEntry): AsyncIterable<Buffer> | Iterable<Buffer> => {
    if (part.method === 0) return storedDocxChunks(part.data)
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

// A password-protected OOXML file is not a zip at all: Office wraps the encrypted package in an OLE
// container, so routing sees OLE magic where it expected PK and declines the file as unrecognized.
// True, and useless to a caller — the file was recognized fine, it is locked. OLE directory entry
// names are UTF-16LE, so the marker stream's name appears verbatim in the bytes. Bounded to the head
// like every other sniff: the directory of a real encrypted package sits well inside it.
const ENCRYPTED_PACKAGE_MARKER = Buffer.from('EncryptedPackage', 'utf16le')
const looksEncryptedOffice = (content: Buffer): boolean =>
    startsWith(content, OLE_MAGIC) && content.subarray(0, DETECT_SAMPLE_BYTES).includes(ENCRYPTED_PACKAGE_MARKER)

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

// A genuine text file never starts with these, so they contradict a text claim even under a hint.
const hasKnownBinaryMagic = (content: Buffer): boolean =>
    startsWith(content, PDF_MAGIC) || startsWith(content, ZIP_MAGIC) || startsWith(content, OLE_MAGIC)

// Are the bytes structurally well-formed utf-16? A printable-ratio test can't tell: read as utf-16,
// arbitrary bytes land across the BMP and are nearly all "printable", so png/jpeg/gif sail through.
// Well-formedness can. The surrogate block is 1/32 of the BMP, so binary hits it constantly and
// essentially never as a correct high-then-low pair; attachment text pairs every one and carries
// neither embedded NULs nor the U+FFFE/U+FFFF noncharacters. Any tell disproves the charset claim.
const isWellFormedUtf16 = (content: Buffer, bigEndian: boolean): boolean => {
    const sample = content.subarray(0, SNIFF_BYTES)
    const end = sample.length - (sample.length % 2) // whole code units only
    if (end === 0) return false
    for (let i = 0; i < end; i += 2) {
        const unit = bigEndian ? sample.readUInt16BE(i) : sample.readUInt16LE(i)
        if (unit === 0 || unit === 0xfffe || unit === 0xffff) return false // embedded NUL or noncharacter
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

// Handlers signal WHY by throwing this; the entry point reads `code` off it instead of turning a
// message into a reason. Anything else that escapes a parser is the file's fault by default, which
// is the same assumption the old free-text path made — it just says so now.
class ExtractionFailure extends Error {
    constructor(readonly code: ExtractionReason) {
        super(code)
        this.name = 'ExtractionFailure'
    }
}

// KNOWN COST of reasons being codes: an unrecognized throw is attributed to the bytes, so a genuine
// bug of ours that escapes here is reported as 'malformed' and blames the file. The free-text reason
// this replaced carried the message, which named it. Accepted rather than hidden — our own invariant
// sites throw ExtractionFailure('internal') explicitly, and the only losses are the throws we did
// not anticipate, which were never something a caller could branch on either way.
const failureReason = (error: unknown): ExtractionReason =>
    error instanceof ExtractionFailure ? error.code : error instanceof HandlerTimeoutError ? 'timed-out' : 'malformed'

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

/////////////////////////////////////////////////////////////
// ENTRY POINT — every step in order, each risky one inside its own safety net.

// A caller's cap may only tighten: absent or NaN falls back to the ceiling, negative clamps to 0,
// fractional floors (a cap is a whole number of chars). The infinities go through the clamp rather
// than the fallback, so the function is monotonic across its whole domain — -Infinity used to mean
// "no cap at all" while -5 meant 0, which is a seam nothing benefits from. Both routes are safe
// either way: this can only ever tighten.
const resolveCap = (requested?: number): number =>
    typeof requested !== 'number' || Number.isNaN(requested)
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
        return { status: 'extracted', truncated: false, emptyReason: 'no-text-content' }
    }

    // Size gate, before any decode or parse.
    if (byteSize > MAX_INPUT_BYTES) {
        return { status: 'skipped', reason: 'too-large' }
    }

    const { type, charset: charsetHint } = parseContentType(input.contentType)
    const { kind } = detectRoute(input)
    const handler = kind ? findHandler(kind) : undefined

    // Unsupported or unrecognized format.
    if (!handler) {
        // Say WHY before falling back to the generic refusals: "unrecognized" is true of the bytes
        // and unhelpful about the file, and a caller who knows a document is merely locked can ask
        // its sender for it rather than treating the attachment as unreadable junk.
        if (looksEncryptedOffice(input.content)) {
            return { status: 'skipped', reason: 'password-protected' }
        }
        return { status: 'skipped', reason: type ? 'unsupported-format' : 'unrecognized' }
    }

    const maxOutputChars = resolveCap(options.maxOutputChars)

    try {
        // Zip bombs: measure the real decompressed size before either OOXML handler inflates
        // anything. This intentionally precedes (and is not charged to) the handler timeout, but it
        // must remain inside the never-throws boundary: an unexpected zlib rejection is a labeled
        // failure, just like a parser rejection.
        if (kind === 'docx' || kind === 'xlsx') {
            const check = await checkDecompressionBudget(input.content, MAX_UNCOMPRESSED_BYTES)
            if (!check.ok) return { status: check.status, reason: check.reason }
        }

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
        // precise trim; html alone still returns a full string, so for that one it's
        // POST-materialization — peak memory follows the whole document, and hard containment is the
        // host memory limit (see README). Don't split a surrogate pair at the boundary: a lone half
        // serializes as U+FFFD.
        const overCap = output.text.length > maxOutputChars
        const capped = overCap ? output.text.slice(0, maxOutputChars) : output.text
        // A trailing high surrogate is half a character however it got there — our cut, or a handler
        // that stopped exactly on the cap. Checked unconditionally rather than only when we cut,
        // because the alternative is requiring every incremental handler to store past its own cap
        // purely so this check can notice, which is what the CAP_STORAGE_SLACK constant used to buy.
        const lastUnit = capped.charCodeAt(capped.length - 1)
        const text = lastUnit >= 0xd800 && lastUnit <= 0xdbff ? capped.slice(0, -1) : capped
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
            ? {
                  status: 'extracted',
                  truncated,
                  // The whole point of the field: a scan and an empty file stop looking alike here.
                  emptyReason: output.hasNonTextContent ? 'no-text-layer' : 'no-text-content',
              }
            : { status: 'extracted', extraction: truncated && options.trailer ? text + options.trailer : text, truncated }
    } catch (error) {
        // The bytes are attacker-controlled, so a throw or timeout is an expected event, not a bug:
        // label it and move on rather than crashing the caller.
        return { status: 'failed', reason: failureReason(error) }
    }
}
