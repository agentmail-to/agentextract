import { readFileSync } from 'node:fs'
import { join } from 'node:path'

import { describe, it, expect } from 'vitest'

import { extractAttachment, detectRoute, MAX_INPUT_BYTES, MAX_OUTPUT_CHARS } from '../attachextract'

// Real fixtures generated once with macOS textutil (.docx) and cupsfilter (.pdf), and exceljs (.xlsx).
// vitest runs from the repo root, so resolve against cwd.
const fixture = (name: string) => readFileSync(join(process.cwd(), 'tests', 'fixtures', name))
const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

// ---------------------------------------------------------------------------
// Synthetic tests for attachextract.ts — Step 1 (baseline + direct-text)
// ---------------------------------------------------------------------------
// Library handlers (html/pdf/docx/eml) and their failure/timeout paths land in
// later steps; here we cover routing, charset decoding, and the safety gates.

const buf = (s: string) => Buffer.from(s, 'utf8')

// Routing --------------------------------------------------------------------

describe('attachextract — routing (never trust one signal)', () => {
    // Honest content-type: the common case. Routes on the type, extracts, reports it.
    it('routes text/csv by content-type', async () => {
        const r = await extractAttachment({ content: buf('a,b\n1,2'), contentType: 'text/csv', filename: 'data.csv' })
        expect(r.status).toBe('extracted')
        expect(r.detectedType).toBe('text')
        expect(r.routedBy).toBe('content-type')
        expect(r.extractedText).toBe('a,b\n1,2')
    })

    // The real-world liar: providers ship CSVs as application/octet-stream. The
    // extension has to rescue the route the content-type failed.
    it('routes a mislabeled octet-stream CSV by its .csv extension', async () => {
        const r = await extractAttachment({
            content: buf('id,name\n1,ada'),
            contentType: 'application/octet-stream',
            filename: 'report.csv',
        })
        expect(r.status).toBe('extracted')
        expect(r.routedBy).toBe('extension')
        expect(r.extractedText).toBe('id,name\n1,ada')
    })

    // No usable type AND no extension — only the bytes are left. The "looks like
    // text" sniff carries it.
    it('routes an octet-stream blob with no filename by sniffing the bytes', async () => {
        const r = await extractAttachment({ content: buf('just some plain text'), contentType: 'application/octet-stream' })
        expect(r.status).toBe('extracted')
        expect(r.routedBy).toBe('sniff')
        expect(r.extractedText).toBe('just some plain text')
    })

    // Extensionless text file, no content-type at all — sniff again.
    it('routes an extensionless, typeless text file by sniffing', async () => {
        const r = await extractAttachment({ content: buf('line one\nline two'), filename: 'notes' })
        expect(r.status).toBe('extracted')
        expect(r.routedBy).toBe('sniff')
    })

    // Any text/* subtype we didn't enumerate is still plain text.
    it('treats an unenumerated text/* subtype as text', async () => {
        const r = await extractAttachment({ content: buf('BEGIN:VCARD'), contentType: 'text/x-unknown' })
        expect(r.detectedType).toBe('text')
        expect(r.status).toBe('extracted')
    })

    // Header-only MIME types (DSNs / forwards) are text, not a full email.
    it('routes message/global-headers as text', async () => {
        const r = await extractAttachment({ content: buf('From: a@b.com\nTo: c@d.com'), contentType: 'message/global-headers' })
        expect(r.detectedType).toBe('text')
        expect(r.status).toBe('extracted')
    })

    // DSN status blobs and ARF spam reports are header-style text payloads.
    it('routes message/delivery-status and message/feedback-report as text', async () => {
        const dsn = await extractAttachment({
            content: buf('Reporting-MTA: dns; mx.example.com\nAction: failed\nStatus: 5.1.1'),
            contentType: 'message/delivery-status',
        })
        expect(dsn.detectedType).toBe('text')
        expect(dsn.status).toBe('extracted')
        const arf = await extractAttachment({
            content: buf('Feedback-Type: abuse\nUser-Agent: SomeReporter/1.0'),
            contentType: 'message/feedback-report',
        })
        expect(arf.detectedType).toBe('text')
        expect(arf.status).toBe('extracted')
    })

    // message/global is a FULL internationalized email (SMTPUTF8 rfc822), not a headers
    // blob — it must parse through the eml handler, not decode as raw text.
    it('routes message/global to the eml handler', async () => {
        const r = await extractAttachment({ content: fixture('plain.eml'), contentType: 'message/global' })
        expect(r.detectedType).toBe('eml')
        expect(r.status).toBe('extracted')
        expect(r.extractedText).toContain('Meeting notes')
    })

    // AMP HTML (text/x-amp-html in real traffic) is markup: it must flatten through the
    // html handler, not decode as raw text and surface tag soup.
    it('routes html-ish text/* subtypes through the html handler', async () => {
        const r = await extractAttachment({
            content: buf('<div>AMP body text</div><style amp-custom>.a{color:red}</style>'),
            contentType: 'text/x-amp-html',
        })
        expect(r.detectedType).toBe('html')
        expect(r.status).toBe('extracted')
        expect(r.extractedText).toContain('AMP body text')
        expect(r.extractedText).not.toContain('color:red')
        expect(r.extractedText).not.toContain('<div>')
    })

    // A PDF shipped as text/plain must NOT be latin1-decoded into garbage and reported
    // as extracted — the magic bytes void the false text claim and rescue the route.
    it('overrides a lying text/plain claim when the bytes are a PDF', async () => {
        const r = await extractAttachment({ content: fixture('sample.pdf'), contentType: 'text/plain', filename: 'report.txt' })
        expect(r.detectedType).toBe('pdf')
        expect(r.routedBy).toBe('sniff')
        expect(r.status).toBe('extracted')
        expect(r.extractedText).toContain('extractable text content')
    })

    // Same lie, docx flavor: zip magic + .docx extension rescue the route.
    it('overrides a lying text claim when the bytes are a docx', async () => {
        const r = await extractAttachment({ content: fixture('sample.docx'), contentType: 'text/plain', filename: 'sample.docx' })
        expect(r.detectedType).toBe('docx')
        expect(r.routedBy).toBe('sniff')
        expect(r.status).toBe('extracted')
    })

    // Unrescuable binary behind a text claim becomes a labeled skip, not mojibake.
    it('skips provably-binary bytes behind a text/plain claim', async () => {
        const r = await extractAttachment({ content: Buffer.from([0x00, 0x01, 0x02, 0x03]), contentType: 'text/plain' })
        expect(r.status).toBe('skipped_unsupported_type')
        expect(r.extractedText).toBeUndefined()
    })

    // The byte-verification guard must not break the zero-byte edge: an empty text
    // attachment stays with the text handler and lands on extracted_empty.
    it('keeps a zero-byte text attachment on the text handler (extracted_empty)', async () => {
        const r = await extractAttachment({ content: Buffer.alloc(0), contentType: 'text/plain', filename: 'empty.txt' })
        expect(r.detectedType).toBe('text')
        expect(r.status).toBe('extracted_empty')
        // A completed-but-empty extraction is present-empty, not undefined, so a field-presence
        // cache counts it as cached (mirrors the body extractor's '' contract).
        expect(r.extractedText).toBe('')
    })

    // A BOM'd UTF-16 file (NUL-heavy bytes) declared as text must survive the guard.
    it('keeps a declared UTF-16 text file with a BOM on the text handler', async () => {
        const r = await extractAttachment({ content: Buffer.from('\ufeffhello', 'utf16le'), contentType: 'text/plain' })
        expect(r.detectedType).toBe('text')
        expect(r.status).toBe('extracted')
        expect(r.extractedText).toBe('hello')
    })

    // A declared, recognized non-text media type is a deliberate skip — we do NOT
    // sniff it (sniff is only for missing/unknown/octet-stream).
    it('skips a declared image type without sniffing', async () => {
        const r = await extractAttachment({ content: buf('...'), contentType: 'image/png', filename: 'logo.png' })
        expect(r.status).toBe('skipped_unsupported_type')
        expect(r.routedBy).toBe('none')
        expect(r.extractedText).toBeUndefined()
    })

    // Binary blob (contains a NUL) mislabeled octet-stream: the sniff must reject it
    // rather than decode garbage.
    it('skips a binary octet-stream blob the sniff rejects', async () => {
        const r = await extractAttachment({ content: Buffer.from([0x00, 0x01, 0x02, 0x03]), contentType: 'application/octet-stream' })
        expect(r.status).toBe('skipped_unsupported_type')
        expect(r.routedBy).toBe('none')
    })

    // UTF-16 text is full of NUL bytes, so the plain "looks like text" heuristic would
    // wrongly reject an undeclared UTF-16 file as binary. The BOM check rescues it.
    it('sniffs an undeclared UTF-16 file as text via its BOM', async () => {
        const content = Buffer.from('﻿hello', 'utf16le')
        const r = await extractAttachment({ content })
        expect(r.routedBy).toBe('sniff')
        expect(r.status).toBe('extracted')
        expect(r.extractedText).toBe('hello')
    })

    // RTF is text/* by MIME but its body is control-word markup. With no RTF handler,
    // decoding it as raw text would leak \rtf1/\pard/\fonttbl into extractedText — worse
    // than a labeled skip. Both the content-type path and the sniff path must skip it.
    it('skips RTF instead of leaking control words as text', async () => {
        const rtf = buf('{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Times;}}\\pard Dear team, the meeting is at noon.\\par}')
        // 1. Declared text/rtf — must not fall through the text/* fallback.
        const byType = await extractAttachment({ content: rtf, contentType: 'text/rtf', filename: 'memo.rtf' })
        expect(byType.status).toBe('skipped_unsupported_type')
        expect(byType.extractedText).toBeUndefined()
        // 2. Mislabeled octet-stream — the {\rtf magic must skip it before looksLikeText grabs it.
        const bySniff = await extractAttachment({ content: rtf, contentType: 'application/octet-stream' })
        expect(bySniff.status).toBe('skipped_unsupported_type')
        expect(bySniff.routedBy).toBe('none')
        expect(bySniff.extractedText).toBeUndefined()
    })

    // detectRoute is the pure routing decision, independent of extraction.
    it('detectRoute records the deciding signal', () => {
        expect(detectRoute({ content: buf('x'), contentType: 'text/plain' }).routedBy).toBe('content-type')
        expect(detectRoute({ content: buf('x'), contentType: 'application/octet-stream', filename: 'a.txt' }).routedBy).toBe('extension')
        expect(detectRoute({ content: buf('x'), contentType: 'application/octet-stream' }).routedBy).toBe('sniff')
        expect(detectRoute({ content: Buffer.from([0x00]), contentType: 'application/octet-stream' }).routedBy).toBe('none')
    })
})

// Charset decoding -----------------------------------------------------------

describe('attachextract — charset-correct decoding', () => {
    // windows-1252 'é' (0xE9) is invalid UTF-8 on its own. Naive Buffer.toString('utf8')
    // mangles it to the replacement char; we must decode with the declared charset.
    it('decodes a windows-1252 file that would mojibake under naive utf-8', async () => {
        const content = Buffer.from([0x63, 0x61, 0x66, 0xe9]) // "café" in windows-1252
        expect(content.toString('utf8')).not.toBe('café') // proves the naive path is broken
        expect(content.toString('utf8')).toContain('�')

        const r = await extractAttachment({ content, contentType: 'text/csv; charset=windows-1252' })
        expect(r.status).toBe('extracted')
        expect(r.extractedText).toBe('café')
        expect(r.charset).toBe('windows-1252')
    })

    // BOM stripped, CRLF/CR normalized to LF.
    it('strips the BOM and normalizes newlines', async () => {
        const content = Buffer.concat([Buffer.from('﻿'), buf('a\r\nb\rc')])
        const r = await extractAttachment({ content, contentType: 'text/plain; charset=utf-8' })
        expect(r.extractedText).toBe('a\nb\nc')
    })

    // A BOM is a definitive charset signal, ahead of jschardet.
    it('decodes a UTF-16 file by its BOM', async () => {
        const content = Buffer.from('﻿hello world', 'utf16le') // FF FE + LE-encoded text
        const r = await extractAttachment({ content, contentType: 'text/plain' })
        expect(r.status).toBe('extracted')
        expect(r.extractedText).toBe('hello world')
        expect(r.charset).toBe('utf-16le')
    })
})

// Safety ---------------------------------------------------------------------

describe('attachextract — safety gates', () => {
    // Over the cap: skipped BEFORE any decode, no text produced.
    it('skips oversize input before any work', async () => {
        const content = Buffer.alloc(MAX_INPUT_BYTES + 1, 0x41)
        const r = await extractAttachment({ content, contentType: 'text/plain', filename: 'big.txt' })
        expect(r.status).toBe('skipped_oversize')
        expect(r.routedBy).toBe('none')
        expect(r.extractedText).toBeUndefined()
        expect(r.byteSize).toBe(MAX_INPUT_BYTES + 1)
    })

    // A handler that succeeds but yields no text is a terminal, valid outcome.
    it('reports empty/whitespace-only text as extracted_empty, not extracted', async () => {
        const r = await extractAttachment({ content: buf('   \n\t  '), contentType: 'text/plain' })
        expect(r.status).toBe('extracted_empty')
        expect(r.extractedText).toBe('') // present-empty (field-presence cache treats it as cached), not undefined
        expect(r.charset).toBeDefined()
    })
})

// Result contract ------------------------------------------------------------

describe('attachextract — result contract', () => {
    it('echoes filename + byteSize on the result', async () => {
        const content = buf('hello')
        const r = await extractAttachment({ content, contentType: 'text/plain', filename: 'hi.txt' })
        expect(r.filename).toBe('hi.txt')
        expect(r.byteSize).toBe(content.length)
    })
})

// HTML handler ---------------------------------------------------------------

describe('attachextract — html handler', () => {
    it('flattens HTML to visible text', async () => {
        const r = await extractAttachment({
            content: buf('<html><body><h1>Title</h1><p>Body text here.</p></body></html>'),
            contentType: 'text/html',
        })
        expect(r.status).toBe('extracted')
        expect(r.detectedType).toBe('html')
        expect(r.routedBy).toBe('content-type')
        // html-to-text uppercases <h1> headings by default, so match case-insensitively.
        expect(r.extractedText?.toLowerCase()).toContain('title')
        expect(r.extractedText).toContain('Body text here.')
    })

    // Non-rendered content must be dropped; real table text must survive.
    it('drops script/style and keeps table cell text', async () => {
        const html =
            '<style>.a{color:red}</style><script>alert("x")</script>' +
            '<table><tr><td>Cell A</td><td>Cell B</td></tr></table>'
        const r = await extractAttachment({ content: buf(html), contentType: 'text/html', filename: 'page.html' })
        expect(r.extractedText).toContain('Cell A')
        expect(r.extractedText).toContain('Cell B')
        expect(r.extractedText).not.toContain('color:red')
        expect(r.extractedText).not.toContain('alert')
    })
})

// PDF handler ----------------------------------------------------------------

describe('attachextract — pdf handler', () => {
    it('extracts text from a real PDF', async () => {
        const r = await extractAttachment({ content: fixture('sample.pdf'), contentType: 'application/pdf', filename: 'sample.pdf' })
        expect(r.status).toBe('extracted')
        expect(r.detectedType).toBe('pdf')
        expect(r.extractedText).toContain('extractable text content')
    })

    // A PDF mislabeled as octet-stream with no .pdf name is rescued by the %PDF magic bytes.
    it('routes a mislabeled PDF by its magic bytes', async () => {
        const r = await extractAttachment({ content: fixture('sample.pdf'), contentType: 'application/octet-stream' })
        expect(r.routedBy).toBe('sniff')
        expect(r.detectedType).toBe('pdf')
        expect(r.status).toBe('extracted')
    })

    // Attacker-controlled bytes that pass the %PDF gate but don't parse must fail cleanly,
    // never throw — this is also the reachability case for the 'failed' status.
    it('returns failed (not a throw) on a corrupt PDF', async () => {
        const r = await extractAttachment({ content: buf('%PDF-1.4\nthis is not a real pdf body'), contentType: 'application/pdf' })
        expect(r.status).toBe('failed')
        expect(r.reason).toBeDefined()
        expect(r.detectedType).toBe('pdf')
    })
})

// DOCX handler ---------------------------------------------------------------

describe('attachextract — docx handler', () => {
    it('extracts a real .docx and preserves paragraph breaks', async () => {
        const r = await extractAttachment({ content: fixture('sample.docx'), contentType: DOCX_TYPE, filename: 'sample.docx' })
        expect(r.status).toBe('extracted')
        expect(r.detectedType).toBe('docx')
        expect(r.extractedText).toContain('First paragraph')
        expect(r.extractedText).toContain('Second paragraph')
        // mammoth separates paragraphs with a blank line — the structure we keep.
        expect(r.extractedText).toMatch(/First paragraph[\s\S]*\n\n[\s\S]*Second paragraph/)
    })

    // An empty document parses successfully but yields no text.
    it('reports an empty .docx as extracted_empty', async () => {
        const r = await extractAttachment({ content: fixture('empty.docx'), contentType: DOCX_TYPE })
        expect(r.status).toBe('extracted_empty')
        expect(r.extractedText).toBe('') // present-empty, not undefined (field-presence cache contract)
    })

    // A docx mislabeled octet-stream is caught by its extension.
    it('routes a docx by extension when the type lies', async () => {
        const r = await extractAttachment({ content: fixture('sample.docx'), contentType: 'application/octet-stream', filename: 'sample.docx' })
        expect(r.routedBy).toBe('extension')
        expect(r.detectedType).toBe('docx')
        expect(r.status).toBe('extracted')
    })
})

// DOC handler (legacy OLE Word) ----------------------------------------------
// Real fixture generated with macOS textutil (-convert doc). Word 97–2003 is an OLE
// compound binary mammoth can't read, so it routes to word-extractor instead.

describe('attachextract — doc handler', () => {
    it('extracts a real legacy .doc by content-type', async () => {
        const r = await extractAttachment({ content: fixture('sample.doc'), contentType: 'application/msword', filename: 'report.doc' })
        expect(r.status).toBe('extracted')
        expect(r.detectedType).toBe('doc')
        expect(r.routedBy).toBe('content-type')
        expect(r.extractedText).toContain('First paragraph about revenue')
        expect(r.extractedText).toContain('Second paragraph about costs')
    })

    // A .doc mislabeled octet-stream is rescued by the shared OLE magic + .doc extension —
    // the extension gate matters because .xls/.ppt/.msg carry the identical OLE signature.
    it('routes a mislabeled .doc by its OLE magic + extension', async () => {
        const r = await extractAttachment({ content: fixture('sample.doc'), contentType: 'application/octet-stream', filename: 'report.doc' })
        expect(r.detectedType).toBe('doc')
        expect(r.status).toBe('extracted')
        expect(r.extractedText).toContain('First paragraph about revenue')
    })

    // OLE bytes with a non-.doc name must NOT be claimed as doc (could be xls/ppt/msg).
    it('does not claim OLE bytes as doc without a .doc extension', async () => {
        const r = await extractAttachment({ content: fixture('sample.doc'), contentType: 'application/octet-stream', filename: 'book.xls' })
        expect(r.detectedType).toBeUndefined()
        expect(r.status).toBe('skipped_unsupported_type')
    })

    // Attacker-controlled bytes that pass the msword route but don't parse fail cleanly.
    it('returns failed (not a throw) on a corrupt .doc', async () => {
        const r = await extractAttachment({ content: buf('this is not really an OLE document'), contentType: 'application/msword', filename: 'bad.doc' })
        expect(r.status).toBe('failed')
        expect(r.detectedType).toBe('doc')
        expect(r.reason).toBeDefined()
    })
})

// XLSX handler (modern OOXML Excel) ------------------------------------------
// Real fixtures generated with exceljs. Sheets flatten to `=== name ===` + tab-joined rows.

describe('attachextract — xlsx handler', () => {
    it('flattens every sheet with headers, keeping cell values across sheets', async () => {
        const r = await extractAttachment({ content: fixture('sample.xlsx'), contentType: XLSX_TYPE, filename: 'book.xlsx' })
        expect(r.status).toBe('extracted')
        expect(r.detectedType).toBe('xlsx')
        expect(r.routedBy).toBe('content-type')
        // Both sheets, each under its own header.
        expect(r.extractedText).toContain('=== Q1 ===')
        expect(r.extractedText).toContain('=== Notes ===')
        expect(r.extractedText).toContain('Region\tRevenue')
        expect(r.extractedText).toContain('West\t4200')
        expect(r.extractedText).toContain('Ada Lovelace')
    })

    // A formula cell must extract its computed VALUE, not the "=SUM(...)" formula string.
    it('extracts the computed value of a formula, not the formula text', async () => {
        const r = await extractAttachment({ content: fixture('sample.xlsx'), contentType: XLSX_TYPE })
        expect(r.extractedText).toContain('Total\t7300')
        expect(r.extractedText).not.toContain('SUM(')
    })

    // A workbook with only empty sheets parses fine but yields no rows.
    it('reports an empty workbook as extracted_empty (present-empty)', async () => {
        const r = await extractAttachment({ content: fixture('empty.xlsx'), contentType: XLSX_TYPE })
        expect(r.status).toBe('extracted_empty')
        expect(r.extractedText).toBe('')
    })

    // xlsx is a zip; mislabeled octet-stream is rescued by the .xlsx extension.
    it('routes a mislabeled .xlsx by its extension', async () => {
        const r = await extractAttachment({ content: fixture('sample.xlsx'), contentType: 'application/octet-stream', filename: 'sample.xlsx' })
        expect(r.routedBy).toBe('extension')
        expect(r.detectedType).toBe('xlsx')
        expect(r.status).toBe('extracted')
    })

    // Zip bytes without a confirming extension must NOT be claimed as xlsx (could be docx/pptx/jar).
    it('does not claim zip bytes as xlsx without a .xlsx extension', async () => {
        const r = await extractAttachment({ content: fixture('sample.xlsx'), contentType: 'application/octet-stream', filename: 'archive.zip' })
        expect(r.detectedType).toBeUndefined()
        expect(r.status).toBe('skipped_unsupported_type')
    })

    // Attacker-controlled bytes that pass the xlsx route but don't parse fail cleanly.
    it('returns failed (not a throw) on a corrupt xlsx', async () => {
        const r = await extractAttachment({ content: buf('PK\x03\x04 not a real workbook'), contentType: XLSX_TYPE, filename: 'bad.xlsx' })
        expect(r.status).toBe('failed')
        expect(r.detectedType).toBe('xlsx')
    })
})

// Output cap -----------------------------------------------------------------
// Input is byte-capped, but output isn't proportional to input — cap it centrally so a
// pathological/large document can't dump megabytes of text into S3 + the search index.

describe('attachextract — output cap', () => {
    // Over-cap output is truncated to exactly MAX_OUTPUT_CHARS and flagged.
    it('truncates over-cap extracted text and sets truncated', async () => {
        const content = Buffer.alloc(2 * MAX_OUTPUT_CHARS, 0x41) // 2x the cap of 'A', well under the input cap
        const r = await extractAttachment({ content, contentType: 'text/plain', filename: 'big.txt' })
        expect(r.status).toBe('extracted')
        expect(r.extractedText?.length).toBe(MAX_OUTPUT_CHARS)
        expect(r.truncated).toBe(true)
    })

    // Under-cap output is untouched and carries no flag.
    it('leaves under-cap text whole with no truncated flag', async () => {
        const r = await extractAttachment({ content: buf('short body'), contentType: 'text/plain' })
        expect(r.extractedText).toBe('short body')
        expect(r.truncated).toBeUndefined()
    })

    // The empty contract is orthogonal to truncation: still '' , never flagged.
    it('does not flag an empty extraction as truncated', async () => {
        const r = await extractAttachment({ content: buf('   '), contentType: 'text/plain' })
        expect(r.status).toBe('extracted_empty')
        expect(r.extractedText).toBe('')
        expect(r.truncated).toBeUndefined()
    })
})

// Image-awareness signals (for a future OCR pass) ----------------------------
// We never read the images (no OCR), but we label docs that likely hold unread text.

describe('attachextract — image-awareness signals', () => {
    // Normal text PDF: real text, no image pages, no flag.
    it('does not flag a normal text PDF', async () => {
        const r = await extractAttachment({ content: fixture('sample.pdf'), contentType: 'application/pdf' })
        expect(r.status).toBe('extracted')
        expect(r.lowTextDensity).toBeFalsy()
        expect(r.pageCount).toBe(1)
        expect(r.emptyPageCount).toBe(0)
    })

    // Sparse PDF: real text is KEPT (not discarded), but flagged as likely image-heavy.
    it('keeps sparse PDF text but flags lowTextDensity', async () => {
        const r = await extractAttachment({ content: fixture('sparse.pdf'), contentType: 'application/pdf' })
        expect(r.status).toBe('extracted')
        expect(r.extractedText).toContain('Hi')
        expect(r.lowTextDensity).toBe(true)
    })

    // Text-less (scanned-style) PDF: extracted_empty, with page counts reported so an
    // OCR pass knows how many pages to render.
    it('reports a text-less PDF as extracted_empty with page counts', async () => {
        const r = await extractAttachment({ content: fixture('blank.pdf'), contentType: 'application/pdf' })
        expect(r.status).toBe('extracted_empty')
        expect(r.pageCount).toBe(1)
        expect(r.emptyPageCount).toBe(1)
    })

    // DOCX with an embedded image but little text: content is probably in the image.
    it('flags a docx with images and little text', async () => {
        const r = await extractAttachment({ content: fixture('image.docx'), contentType: DOCX_TYPE })
        expect(r.status).toBe('extracted')
        expect(r.extractedText).toContain('Report')
        expect(r.lowTextDensity).toBe(true)
    })

    // A text-rich docx (no media) is not flagged.
    it('does not flag a text-only docx', async () => {
        const r = await extractAttachment({ content: fixture('sample.docx'), contentType: DOCX_TYPE })
        expect(r.lowTextDensity).toBeFalsy()
    })
})

// EML handler (nested email) -------------------------------------------------

describe('attachextract — eml handler', () => {
    // Subject + text body of a forwarded email.
    it('extracts subject and text body', async () => {
        const r = await extractAttachment({ content: fixture('plain.eml'), contentType: 'message/rfc822', filename: 'plain.eml' })
        expect(r.status).toBe('extracted')
        expect(r.detectedType).toBe('eml')
        expect(r.extractedText).toContain('Meeting notes') // subject
        expect(r.extractedText).toContain('notes from our meeting') // body
    })

    // No text part -> the html body is flattened through the shared HTML path.
    it('falls back to the html body when there is no text part', async () => {
        const r = await extractAttachment({ content: fixture('html-only.eml'), contentType: 'message/rfc822' })
        expect(r.status).toBe('extracted')
        expect(r.extractedText).toContain('HTML only email body')
    })

    // A forwarded email's own attachment is recursed back through the pipeline.
    it('recurses into inner attachments as children', async () => {
        const r = await extractAttachment({ content: fixture('nested-attachment.eml'), contentType: 'message/rfc822' })
        expect(r.extractedText).toContain('Please find the attached report')
        expect(r.children).toHaveLength(1)
        expect(r.children?.[0].detectedType).toBe('text')
        expect(r.children?.[0].status).toBe('extracted')
        expect(r.children?.[0].extractedText).toContain('Quarterly report')
    })

    // Depth cap: an email nested inside an email is body-read, but we do NOT descend
    // into its own attachments (deep.txt must never appear).
    it('stops at one level of nesting (MAX_NESTING_DEPTH)', async () => {
        const r = await extractAttachment({ content: fixture('eml-in-eml.eml'), contentType: 'message/rfc822' })
        expect(r.children).toHaveLength(1)
        const inner = r.children![0]
        expect(inner.detectedType).toBe('eml')
        expect(inner.extractedText).toContain('Inner email body text')
        expect(inner.children).toBeUndefined() // deep.txt NOT descended into
        expect(JSON.stringify(r)).not.toContain('descended') // deep.txt content never surfaces
    })

    // Routing: recognized by extension when the type lies.
    it('routes an eml by extension', async () => {
        const r = await extractAttachment({ content: fixture('plain.eml'), contentType: 'application/octet-stream', filename: 'forward.eml' })
        expect(r.routedBy).toBe('extension')
        expect(r.detectedType).toBe('eml')
    })

    // NB: there is no "inner attachment over 10 MB" test on purpose — an inner attachment
    // is always a subset of the (<=10 MB) outer email, so the outer size gate trips first
    // and the inner cap is unreachable in a single eml. Each child still runs the same
    // gated pipeline via extractAt(depth + 1).
})