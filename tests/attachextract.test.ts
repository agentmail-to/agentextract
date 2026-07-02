import { readFileSync } from 'node:fs'
import { join } from 'node:path'

import { describe, it, expect } from 'vitest'

import { extractAttachment, detectRoute, MAX_INPUT_BYTES, EXTRACTION_VERSION } from '../attachextract'

// Real fixtures generated once with macOS textutil (.docx) and cupsfilter (.pdf).
// vitest runs from the repo root, so resolve against cwd.
const fixture = (name: string) => readFileSync(join(process.cwd(), 'tests', 'fixtures', name))
const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

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
        expect(r.extractedText).toBeUndefined()
        expect(r.charset).toBeDefined()
    })
})

// Result contract ------------------------------------------------------------

describe('attachextract — result contract', () => {
    it('always stamps the extraction version and echoes filename + byteSize', async () => {
        const content = buf('hello')
        const r = await extractAttachment({ content, contentType: 'text/plain', filename: 'hi.txt' })
        expect(r.extractionVersion).toBe(EXTRACTION_VERSION)
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
        expect(r.extractedText).toBeUndefined()
    })

    // A docx mislabeled octet-stream is caught by its extension.
    it('routes a docx by extension when the type lies', async () => {
        const r = await extractAttachment({ content: fixture('sample.docx'), contentType: 'application/octet-stream', filename: 'sample.docx' })
        expect(r.routedBy).toBe('extension')
        expect(r.detectedType).toBe('docx')
        expect(r.status).toBe('extracted')
    })
})
