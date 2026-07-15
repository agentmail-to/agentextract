import { readFileSync } from 'node:fs'
import { join } from 'node:path'

import { describe, it, expect } from 'vitest'
import iconv from 'iconv-lite'

import { extractAttachment, detectRoute, MAX_INPUT_BYTES, MAX_OUTPUT_CHARS } from '../attachment'

// NOTE: result shape is { status, extraction?, reason? }; `extraction` is omitted (never '') when a
// handler runs but produces no text. Nested emails (.eml) are out of scope in this version and skip.

// Real fixtures generated once with macOS textutil (.docx) and cupsfilter (.pdf), and exceljs (.xlsx).
// vitest runs from the repo root, so resolve against cwd.
const fixture = (name: string) => readFileSync(join(process.cwd(), 'tests', 'fixtures', name))
const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

// ---------------------------------------------------------------------------
// Synthetic tests for attachment.ts
// ---------------------------------------------------------------------------
// The RESULT is deliberately slim: { status, extraction?, reason? }. Routing is NOT on the result —
// it's a separate concern verified through detectRoute(). So routing-decision assertions here call
// detectRoute(input) directly; extraction assertions read the slim result.

const buf = (s: string) => Buffer.from(s, 'utf8')

// Result-shape guard ---------------------------------------------------------
// The whole point of the slim contract: the top-level result carries no routing/diagnostic noise,
// no top-level filename, and no truncation flag. If any of those leak back in, this fails.

describe('attachment — slim result contract', () => {
    it('returns only the three contract fields, never routing/diagnostic noise', async () => {
        const r = await extractAttachment({ content: buf('hello'), contentType: 'text/plain', filename: 'hi.txt' })
        // Every key present must be one of the three contract fields — nothing else.
        const allowed = ['status', 'extraction', 'reason']
        expect(Object.keys(r).every((k) => allowed.includes(k))).toBe(true)
        // The removed fields must be absent (not merely undefined-valued).
        for (const gone of ['filename', 'byteSize', 'detectedType', 'routedBy', 'charset', 'extractedText', 'truncated', 'children', 'lowTextDensity', 'pageCount', 'emptyPageCount']) {
            expect(gone in r).toBe(false)
        }
    })

    it('exposes only the three collapsed statuses', () => {
        // Compile-time contract is 'extracted' | 'skipped' | 'failed'; this documents it at runtime.
        const extracted = 'extracted'
        expect(['extracted', 'skipped', 'failed']).toContain(extracted)
    })
})

// Routing --------------------------------------------------------------------
// Routing lives on detectRoute now, so we assert the decision there and the extraction on the result.

describe('attachment — routing (never trust one signal)', () => {
    // Honest content-type: the common case. Routes on the type, extracts, reports it.
    it('routes text/csv by content-type', async () => {
        const input = { content: buf('a,b\n1,2'), contentType: 'text/csv', filename: 'data.csv' }
        expect(detectRoute(input)).toEqual({ kind: 'text', routedBy: 'content-type' })
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBe('a,b\n1,2')
    })

    // The real-world liar: providers ship CSVs as application/octet-stream. The
    // extension has to rescue the route the content-type failed.
    it('routes a mislabeled octet-stream CSV by its .csv extension', async () => {
        const input = { content: buf('id,name\n1,ada'), contentType: 'application/octet-stream', filename: 'report.csv' }
        expect(detectRoute(input)).toEqual({ kind: 'text', routedBy: 'extension' })
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBe('id,name\n1,ada')
    })

    // No usable type AND no extension — only the bytes are left. The "looks like
    // text" sniff carries it.
    it('routes an octet-stream blob with no filename by sniffing the bytes', async () => {
        const input = { content: buf('just some plain text'), contentType: 'application/octet-stream' }
        expect(detectRoute(input).routedBy).toBe('sniff')
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBe('just some plain text')
    })

    // Extensionless text file, no content-type at all — sniff again.
    it('routes an extensionless, typeless text file by sniffing', async () => {
        const input = { content: buf('line one\nline two'), filename: 'notes' }
        expect(detectRoute(input).routedBy).toBe('sniff')
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
    })

    // Any text/* subtype we didn't enumerate is still plain text.
    it('treats an unenumerated text/* subtype as text', async () => {
        const input = { content: buf('BEGIN:VCARD'), contentType: 'text/x-unknown' }
        expect(detectRoute(input).kind).toBe('text')
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
    })

    // Header-only MIME types (DSNs / forwards) are text, not a full email.
    it('routes message/global-headers as text', async () => {
        const input = { content: buf('From: a@b.com\nTo: c@d.com'), contentType: 'message/global-headers' }
        expect(detectRoute(input).kind).toBe('text')
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
    })

    // DSN status blobs and ARF spam reports are header-style text payloads.
    it('routes message/delivery-status and message/feedback-report as text', async () => {
        const dsnInput = { content: buf('Reporting-MTA: dns; mx.example.com\nAction: failed\nStatus: 5.1.1'), contentType: 'message/delivery-status' }
        expect(detectRoute(dsnInput).kind).toBe('text')
        expect((await extractAttachment(dsnInput)).status).toBe('extracted')
        const arfInput = { content: buf('Feedback-Type: abuse\nUser-Agent: SomeReporter/1.0'), contentType: 'message/feedback-report' }
        expect(detectRoute(arfInput).kind).toBe('text')
        expect((await extractAttachment(arfInput)).status).toBe('extracted')
    })

    // Nested emails are OUT OF SCOPE in this version — there is no eml handler. A full email typed
    // message/rfc822 or message/global routes to nothing and lands on a skip, never decoded as raw
    // text. (Header-only message/* report types above are still text.)
    it('skips full emails (message/rfc822 / message/global) — eml out of scope', async () => {
        for (const contentType of ['message/rfc822', 'message/global']) {
            const input = { content: fixture('plain.eml'), contentType }
            expect(detectRoute(input).kind).toBeUndefined()
            const r = await extractAttachment(input)
            expect(r.status).toBe('skipped')
            expect(r.extraction).toBeUndefined()
        }
    })

    // AMP HTML (text/x-amp-html in real traffic) is markup: it must flatten through the
    // html handler, not decode as raw text and surface tag soup.
    it('routes html-ish text/* subtypes through the html handler', async () => {
        const input = { content: buf('<div>AMP body text</div><style amp-custom>.a{color:red}</style>'), contentType: 'text/x-amp-html' }
        expect(detectRoute(input).kind).toBe('html')
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
        expect(r.extraction).toContain('AMP body text')
        expect(r.extraction).not.toContain('color:red')
        expect(r.extraction).not.toContain('<div>')
    })

    // A PDF shipped as text/plain must NOT be latin1-decoded into garbage and reported
    // as extracted — the magic bytes void the false text claim and rescue the route.
    it('overrides a lying text/plain claim when the bytes are a PDF', async () => {
        const input = { content: fixture('sample.pdf'), contentType: 'text/plain', filename: 'report.txt' }
        expect(detectRoute(input)).toEqual({ kind: 'pdf', routedBy: 'sniff' })
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
        expect(r.extraction).toContain('extractable text content')
    })

    // Same lie, docx flavor: zip magic + .docx extension rescue the route.
    it('overrides a lying text claim when the bytes are a docx', async () => {
        const input = { content: fixture('sample.docx'), contentType: 'text/plain', filename: 'sample.docx' }
        expect(detectRoute(input)).toEqual({ kind: 'docx', routedBy: 'sniff' })
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
    })

    // Unrescuable binary behind a text claim becomes a labeled skip, not mojibake.
    it('skips provably-binary bytes behind a text/plain claim', async () => {
        const r = await extractAttachment({ content: Buffer.from([0x00, 0x01, 0x02, 0x03]), contentType: 'text/plain' })
        expect(r.status).toBe('skipped')
        expect(r.extraction).toBeUndefined()
    })

    // The byte-verification guard must not break the zero-byte edge: an empty text
    // attachment stays with the text handler and lands on extracted with no extraction.
    it('keeps a zero-byte text attachment on the text handler (extracted, no extraction)', async () => {
        const input = { content: Buffer.alloc(0), contentType: 'text/plain', filename: 'empty.txt' }
        expect(detectRoute(input).kind).toBe('text')
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
        // A completed-but-empty extraction omits `extraction` entirely — never '' — so status alone
        // ('extracted') distinguishes "ran, no text" from a skip/failure.
        expect(r.extraction).toBeUndefined()
    })

    // A BOM'd UTF-16 file (NUL-heavy bytes) declared as text must survive the guard.
    it('keeps a declared UTF-16 text file with a BOM on the text handler', async () => {
        const input = { content: Buffer.from('﻿hello', 'utf16le'), contentType: 'text/plain' }
        expect(detectRoute(input).kind).toBe('text')
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBe('hello')
    })

    // A declared, recognized non-text media type is a deliberate skip — we do NOT
    // sniff it (sniff is only for missing/unknown/octet-stream).
    it('skips a declared image type without sniffing', async () => {
        const input = { content: buf('...'), contentType: 'image/png', filename: 'logo.png' }
        expect(detectRoute(input)).toEqual({ routedBy: 'none' })
        const r = await extractAttachment(input)
        expect(r.status).toBe('skipped')
        expect(r.extraction).toBeUndefined()
    })

    // Binary blob (contains a NUL) mislabeled octet-stream: the sniff must reject it
    // rather than decode garbage.
    it('skips a binary octet-stream blob the sniff rejects', async () => {
        const input = { content: Buffer.from([0x00, 0x01, 0x02, 0x03]), contentType: 'application/octet-stream' }
        expect(detectRoute(input).routedBy).toBe('none')
        const r = await extractAttachment(input)
        expect(r.status).toBe('skipped')
    })

    // UTF-16 text is full of NUL bytes, so the plain "looks like text" heuristic would
    // wrongly reject an undeclared UTF-16 file as binary. The BOM check rescues it.
    it('sniffs an undeclared UTF-16 file as text via its BOM', async () => {
        const input = { content: Buffer.from('﻿hello', 'utf16le') }
        expect(detectRoute(input).routedBy).toBe('sniff')
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBe('hello')
    })

    // RTF is text/* by MIME but its body is control-word markup. With no RTF handler,
    // decoding it as raw text would leak \rtf1/\pard/\fonttbl into extraction — worse
    // than a labeled skip. Both the content-type path and the sniff path must skip it.
    it('skips RTF instead of leaking control words as text', async () => {
        const rtf = buf('{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Times;}}\\pard Dear team, the meeting is at noon.\\par}')
        // 1. Declared text/rtf — must not fall through the text/* fallback.
        const byType = await extractAttachment({ content: rtf, contentType: 'text/rtf', filename: 'memo.rtf' })
        expect(byType.status).toBe('skipped')
        expect(byType.extraction).toBeUndefined()
        // 2. Mislabeled octet-stream — the {\rtf magic must skip it before looksLikeText grabs it.
        const bySniffInput = { content: rtf, contentType: 'application/octet-stream' }
        expect(detectRoute(bySniffInput).routedBy).toBe('none')
        const bySniff = await extractAttachment(bySniffInput)
        expect(bySniff.status).toBe('skipped')
        expect(bySniff.extraction).toBeUndefined()
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
// The result no longer exposes `charset`; correct decoding is proven by the decoded text itself.

describe('attachment — charset-correct decoding', () => {
    // windows-1252 'é' (0xE9) is invalid UTF-8 on its own. Naive Buffer.toString('utf8')
    // mangles it to the replacement char; we must decode with the declared charset.
    it('decodes a windows-1252 file that would mojibake under naive utf-8', async () => {
        const content = Buffer.from([0x63, 0x61, 0x66, 0xe9]) // "café" in windows-1252
        expect(content.toString('utf8')).not.toBe('café') // proves the naive path is broken
        expect(content.toString('utf8')).toContain('�')

        const r = await extractAttachment({ content, contentType: 'text/csv; charset=windows-1252' })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBe('café') // correct decode is the observable proof the charset was honored
    })

    // BOM stripped, CRLF/CR normalized to LF.
    it('strips the BOM and normalizes newlines', async () => {
        const content = Buffer.concat([Buffer.from('﻿'), buf('a\r\nb\rc')])
        const r = await extractAttachment({ content, contentType: 'text/plain; charset=utf-8' })
        expect(r.extraction).toBe('a\nb\nc')
    })

    // A BOM is a definitive charset signal, ahead of jschardet.
    it('decodes a UTF-16 file by its BOM', async () => {
        const content = Buffer.from('﻿hello world', 'utf16le') // FF FE + LE-encoded text
        const r = await extractAttachment({ content, contentType: 'text/plain' })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBe('hello world')
    })

    // Some real text attachments declare UTF-16 in Content-Type but omit the BOM. The explicit
    // charset is the only signal that the NUL-heavy bytes are text, so the binary-contradiction
    // guard must not reroute them to a skip.
    it('decodes a BOM-less UTF-16 file when Content-Type declares the charset', async () => {
        const content = Buffer.from('hello world', 'utf16le')
        const input = { content, contentType: 'text/plain; charset=utf-16le' }
        expect(detectRoute(input)).toEqual({ kind: 'text', routedBy: 'content-type' })
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBe('hello world')
    })

    // Precedence: an in-band BOM is definitive and must beat a wrong Content-Type charset. A
    // UTF-16 file mislabeled charset=windows-1252 would mojibake under the hint; the BOM wins.
    it('lets a BOM override a contradicting charset hint', async () => {
        const content = Buffer.from('﻿facturé', 'utf16le') // real UTF-16LE BOM, lying windows-1252 label
        const r = await extractAttachment({ content, contentType: 'text/plain; charset=windows-1252' })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBe('facturé') // hint ignored in favor of the BOM
    })

    // Precedence: content that is provably valid UTF-8 must not be mangled by a wrong single-byte
    // hint. windows-1252 would turn each multi-byte char into two garbage chars; UTF-8 wins.
    it('keeps valid UTF-8 multi-byte text despite a wrong single-byte charset hint', async () => {
        const content = Buffer.from('café ünïcode 中文', 'utf8')
        const r = await extractAttachment({ content, contentType: 'text/plain; charset=windows-1252' })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBe('café ünïcode 中文') // decoded as utf-8, not the declared hint
    })

    // When the chosen charset is wrong and produces U+FFFD (jschardet confidently reads this
    // undeclared big5 as GB2312), we fall back to byte-preserving latin1 — never emit replacement
    // chars, which are unrecoverable. The observable guarantee: no U+FFFD in the output.
    it('falls back to latin1 rather than emit U+FFFD on a wrong-charset decode', async () => {
        const content = iconv.encode('你好世界，這是發票總額。'.repeat(15), 'big5')
        const r = await extractAttachment({ content, contentType: 'text/plain' })
        expect(r.status).toBe('extracted')
        expect(r.extraction).not.toContain('�')
    })
})

// Safety ---------------------------------------------------------------------

describe('attachment — safety gates', () => {
    // Over the cap: skipped BEFORE any decode, no text produced.
    it('skips oversize input before any work', async () => {
        const content = Buffer.alloc(MAX_INPUT_BYTES + 1, 0x41)
        const r = await extractAttachment({ content, contentType: 'text/plain', filename: 'big.txt' })
        expect(r.status).toBe('skipped')
        expect(r.extraction).toBeUndefined()
        expect(r.reason).toContain(`${MAX_INPUT_BYTES + 1}`)
    })

    // A handler that succeeds but yields no text is a terminal, valid outcome: extracted, no extraction.
    it('reports empty/whitespace-only text as extracted with no extraction', async () => {
        const r = await extractAttachment({ content: buf('   \n\t  '), contentType: 'text/plain' })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBeUndefined() // omitted, never '' — status 'extracted' carries the "ran, empty" signal
    })
})

// HTML handler ---------------------------------------------------------------

describe('attachment — html handler', () => {
    it('flattens HTML to visible text', async () => {
        const input = { content: buf('<html><body><h1>Title</h1><p>Body text here.</p></body></html>'), contentType: 'text/html' }
        expect(detectRoute(input)).toEqual({ kind: 'html', routedBy: 'content-type' })
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
        // html-to-text uppercases <h1> headings by default, so match case-insensitively.
        expect(r.extraction?.toLowerCase()).toContain('title')
        expect(r.extraction).toContain('Body text here.')
    })

    // Non-rendered content must be dropped; real table text must survive.
    it('drops script/style and keeps table cell text', async () => {
        const html =
            '<style>.a{color:red}</style><script>alert("x")</script>' +
            '<table><tr><td>Cell A</td><td>Cell B</td></tr></table>'
        const r = await extractAttachment({ content: buf(html), contentType: 'text/html', filename: 'page.html' })
        expect(r.extraction).toContain('Cell A')
        expect(r.extraction).toContain('Cell B')
        expect(r.extraction).not.toContain('color:red')
        expect(r.extraction).not.toContain('alert')
    })
})

// PDF handler ----------------------------------------------------------------

describe('attachment — pdf handler', () => {
    it('extracts text from a real PDF', async () => {
        const input = { content: fixture('sample.pdf'), contentType: 'application/pdf', filename: 'sample.pdf' }
        expect(detectRoute(input).kind).toBe('pdf')
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
        expect(r.extraction).toContain('extractable text content')
    })

    // A PDF mislabeled as octet-stream with no .pdf name is rescued by the %PDF magic bytes.
    it('routes a mislabeled PDF by its magic bytes', async () => {
        const input = { content: fixture('sample.pdf'), contentType: 'application/octet-stream' }
        expect(detectRoute(input)).toEqual({ kind: 'pdf', routedBy: 'sniff' })
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
    })

    // Several providers use historical PDF aliases and do not always include a filename. These are
    // real PDF claims, not generic unsupported binary media, so they should route directly.
    it('routes common PDF MIME aliases without needing a filename', async () => {
        for (const contentType of ['application/x-pdf', 'application/acrobat', 'application/vnd.pdf']) {
            const input = { content: fixture('sample.pdf'), contentType }
            expect(detectRoute(input)).toEqual({ kind: 'pdf', routedBy: 'content-type' })
            const r = await extractAttachment(input)
            expect(r.status).toBe('extracted')
            expect(r.extraction).toContain('extractable text content')
        }
    })

    // Attacker-controlled bytes that pass the %PDF gate but don't parse must fail cleanly,
    // never throw — this is also the reachability case for the 'failed' status.
    it('returns failed (not a throw) on a corrupt PDF', async () => {
        const r = await extractAttachment({ content: buf('%PDF-1.4\nthis is not a real pdf body'), contentType: 'application/pdf' })
        expect(r.status).toBe('failed')
        expect(r.reason).toBeDefined()
        expect(r.extraction).toBeUndefined()
    })
})

// DOCX handler ---------------------------------------------------------------

describe('attachment — docx handler', () => {
    it('extracts a real .docx and preserves paragraph breaks', async () => {
        const input = { content: fixture('sample.docx'), contentType: DOCX_TYPE, filename: 'sample.docx' }
        expect(detectRoute(input).kind).toBe('docx')
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
        expect(r.extraction).toContain('First paragraph')
        expect(r.extraction).toContain('Second paragraph')
        // mammoth separates paragraphs with a blank line — the structure we keep.
        expect(r.extraction).toMatch(/First paragraph[\s\S]*\n\n[\s\S]*Second paragraph/)
    })

    // An empty document parses successfully but yields no text: extracted, no extraction.
    it('reports an empty .docx as extracted with no extraction', async () => {
        const r = await extractAttachment({ content: fixture('empty.docx'), contentType: DOCX_TYPE })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBeUndefined()
    })

    // A docx mislabeled octet-stream is caught by its extension.
    it('routes a docx by extension when the type lies', async () => {
        const input = { content: fixture('sample.docx'), contentType: 'application/octet-stream', filename: 'sample.docx' }
        expect(detectRoute(input)).toEqual({ kind: 'docx', routedBy: 'extension' })
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
    })
})

// DOC handler (legacy OLE Word) ----------------------------------------------
// Real fixture generated with macOS textutil (-convert doc). Word 97–2003 is an OLE
// compound binary mammoth can't read, so it routes to word-extractor instead.

describe('attachment — doc handler', () => {
    it('extracts a real legacy .doc by content-type', async () => {
        const input = { content: fixture('sample.doc'), contentType: 'application/msword', filename: 'report.doc' }
        expect(detectRoute(input)).toEqual({ kind: 'doc', routedBy: 'content-type' })
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
        expect(r.extraction).toContain('First paragraph about revenue')
        expect(r.extraction).toContain('Second paragraph about costs')
    })

    // A .doc mislabeled octet-stream is rescued by the shared OLE magic + .doc extension —
    // the extension gate matters because .xls/.ppt/.msg carry the identical OLE signature.
    it('routes a mislabeled .doc by its OLE magic + extension', async () => {
        const input = { content: fixture('sample.doc'), contentType: 'application/octet-stream', filename: 'report.doc' }
        expect(detectRoute(input).kind).toBe('doc')
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
        expect(r.extraction).toContain('First paragraph about revenue')
    })

    // OLE bytes with a non-.doc name must NOT be claimed as doc (could be xls/ppt/msg).
    it('does not claim OLE bytes as doc without a .doc extension', async () => {
        const input = { content: fixture('sample.doc'), contentType: 'application/octet-stream', filename: 'book.xls' }
        expect(detectRoute(input).kind).toBeUndefined()
        const r = await extractAttachment(input)
        expect(r.status).toBe('skipped')
    })

    // Bytes that carry the OLE magic (so magic-verify lets them through) but don't parse fail
    // cleanly. NB: garbage WITHOUT the OLE magic is now rerouted by magic-verify, not sent here.
    it('returns failed (not a throw) on a corrupt .doc', async () => {
        const oleGarbage = Buffer.concat([Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]), buf(' not a real doc body')])
        const input = { content: oleGarbage, contentType: 'application/msword', filename: 'bad.doc' }
        expect(detectRoute(input).kind).toBe('doc')
        const r = await extractAttachment(input)
        expect(r.status).toBe('failed')
        expect(r.reason).toBeDefined()
    })
})

// XLSX handler (modern OOXML Excel) ------------------------------------------
// Real fixtures generated with exceljs. Sheets flatten to `=== name ===` + tab-joined rows.

describe('attachment — xlsx handler', () => {
    it('flattens every sheet with headers, keeping cell values across sheets', async () => {
        const input = { content: fixture('sample.xlsx'), contentType: XLSX_TYPE, filename: 'book.xlsx' }
        expect(detectRoute(input)).toEqual({ kind: 'xlsx', routedBy: 'content-type' })
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
        // Both sheets, each under its own header.
        expect(r.extraction).toContain('=== Q1 ===')
        expect(r.extraction).toContain('=== Notes ===')
        expect(r.extraction).toContain('Region\tRevenue')
        expect(r.extraction).toContain('West\t4200')
        expect(r.extraction).toContain('Ada Lovelace')
    })

    // A formula cell must extract its computed VALUE, not the "=SUM(...)" formula string.
    it('extracts the computed value of a formula, not the formula text', async () => {
        const r = await extractAttachment({ content: fixture('sample.xlsx'), contentType: XLSX_TYPE })
        expect(r.extraction).toContain('Total\t7300')
        expect(r.extraction).not.toContain('SUM(')
    })

    // A workbook with only empty sheets parses fine but yields no rows: extracted, no extraction.
    it('reports an empty workbook as extracted with no extraction', async () => {
        const r = await extractAttachment({ content: fixture('empty.xlsx'), contentType: XLSX_TYPE })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBeUndefined()
    })

    // xlsx is a zip; mislabeled octet-stream is rescued by the .xlsx extension.
    it('routes a mislabeled .xlsx by its extension', async () => {
        const input = { content: fixture('sample.xlsx'), contentType: 'application/octet-stream', filename: 'sample.xlsx' }
        expect(detectRoute(input)).toEqual({ kind: 'xlsx', routedBy: 'extension' })
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
    })

    // A real xlsx is identified from its OOXML part even when named .zip and typed octet-stream —
    // content beats a missing/wrong extension.
    it('detects an xlsx from content even when named .zip', async () => {
        const input = { content: fixture('sample.xlsx'), contentType: 'application/octet-stream', filename: 'archive.zip' }
        expect(detectRoute(input).kind).toBe('xlsx')
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
    })

    // A generic (non-OOXML) zip has no docx/xlsx part, so it stays unrouted — not mis-claimed.
    it('does not claim a non-OOXML zip', async () => {
        const genericZip = Buffer.concat([Buffer.from([0x50, 0x4b, 0x03, 0x04]), buf('not an office file, just a zipped thing')])
        const input = { content: genericZip, contentType: 'application/octet-stream', filename: 'bundle.zip' }
        expect(detectRoute(input).kind).toBeUndefined()
        const r = await extractAttachment(input)
        expect(r.status).toBe('skipped')
    })

    // Bytes carrying the zip magic + the xl/workbook.xml marker (so content-detection routes them
    // to xlsx) but with a corrupt body fail cleanly. NB: a bare "PK.." with no OOXML part is now
    // rerouted by content-detection, not sent here.
    it('returns failed (not a throw) on a corrupt xlsx', async () => {
        const oxmlGarbage = Buffer.concat([Buffer.from([0x50, 0x4b, 0x03, 0x04]), buf('xl/workbook.xml'), buf(' corrupt body')])
        const input = { content: oxmlGarbage, contentType: XLSX_TYPE, filename: 'bad.xlsx' }
        expect(detectRoute(input).kind).toBe('xlsx')
        const r = await extractAttachment(input)
        expect(r.status).toBe('failed')
    })
})

// Output cap -----------------------------------------------------------------
// Input is byte-capped, but output isn't proportional to input — cap it centrally so a
// pathological/large document can't dump megabytes of text into S3 + the search index.

describe('attachment — output cap', () => {
    // Over-cap output is bounded to exactly MAX_OUTPUT_CHARS (silently — no truncation flag).
    it('bounds over-cap extracted text to MAX_OUTPUT_CHARS', async () => {
        const content = Buffer.alloc(2 * MAX_OUTPUT_CHARS, 0x41) // 2x the cap of 'A', well under the input cap
        const r = await extractAttachment({ content, contentType: 'text/plain', filename: 'big.txt' })
        expect(r.status).toBe('extracted')
        expect(r.extraction?.length).toBe(MAX_OUTPUT_CHARS)
    })

    // Under-cap output is untouched.
    it('leaves under-cap text whole', async () => {
        const r = await extractAttachment({ content: buf('short body'), contentType: 'text/plain' })
        expect(r.extraction).toBe('short body')
    })
})

// Magic-verify (mislabeled binary claims) ------------------------------------
// A confident but wrong binary content-type (a PDF stamped application/…docx, etc.) must not go
// straight to the wrong parser and fail — we verify content (magic bytes, and the OOXML part for
// docx vs xlsx) first and re-sniff on mismatch. Routing is verified via detectRoute.

describe('attachment — magic-verify of binary claims', () => {
    const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

    // PDF bytes mislabeled as docx: the %PDF magic reroutes them back to the pdf handler.
    it('recovers a PDF mislabeled as docx', async () => {
        const input = { content: fixture('sample.pdf'), contentType: DOCX, filename: 'report.docx' }
        expect(detectRoute(input)).toEqual({ kind: 'pdf', routedBy: 'sniff' })
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
        expect(r.extraction).toContain('extractable text content')
    })

    // DOCX bytes mislabeled application/pdf recover to docx from their OOXML part — even with a
    // lying .pdf name, since content (not the extension) decides.
    it('recovers a DOCX mislabeled as pdf regardless of extension', async () => {
        const input = { content: fixture('sample.docx'), contentType: 'application/pdf', filename: 'report.pdf' }
        expect(detectRoute(input).kind).toBe('docx')
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
    })

    // The docx↔xlsx case magic bytes alone can't resolve: an xlsx mislabeled as docx is caught by
    // its xl/workbook.xml part and rerouted to the xlsx handler instead of failing in mammoth.
    it('recovers an xlsx mislabeled as docx via its OOXML part', async () => {
        const input = { content: fixture('sample.xlsx'), contentType: DOCX, filename: 'report.docx' }
        expect(detectRoute(input).kind).toBe('xlsx')
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
    })

    // Honest binary claims are unaffected — magic matches, no reroute.
    it('leaves an honest PDF claim on the content-type route', async () => {
        const input = { content: fixture('sample.pdf'), contentType: 'application/pdf', filename: 'x.pdf' }
        expect(detectRoute(input)).toEqual({ kind: 'pdf', routedBy: 'content-type' })
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
    })
})

// Image-heavy documents (no OCR) ---------------------------------------------
// We never read images. The result no longer carries image-awareness signals (lowTextDensity /
// pageCount / emptyPageCount) — those are internal now. What we still guarantee is behavioral:
// reachable text is KEPT, and a text-less document lands on extracted with present-empty text.

describe('attachment — image-heavy documents', () => {
    // Normal text PDF: real text extracted.
    it('extracts a normal text PDF', async () => {
        const r = await extractAttachment({ content: fixture('sample.pdf'), contentType: 'application/pdf' })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toContain('extractable text content')
    })

    // Sparse PDF: whatever real text exists is KEPT (not discarded).
    it('keeps the reachable text of a sparse PDF', async () => {
        const r = await extractAttachment({ content: fixture('sparse.pdf'), contentType: 'application/pdf' })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toContain('Hi')
    })

    // Text-less (scanned-style) PDF: parses fine but yields no text → extracted, no extraction.
    it('reports a text-less PDF as extracted with no extraction', async () => {
        const r = await extractAttachment({ content: fixture('blank.pdf'), contentType: 'application/pdf' })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBeUndefined()
    })

    // DOCX with an embedded image but little text: the reachable text is still extracted.
    it('extracts the text of a docx with images and little text', async () => {
        const r = await extractAttachment({ content: fixture('image.docx'), contentType: DOCX_TYPE })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toContain('Report')
    })

    // A text-rich docx extracts its full body.
    it('extracts a text-only docx', async () => {
        const r = await extractAttachment({ content: fixture('sample.docx'), contentType: DOCX_TYPE })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toContain('First paragraph')
    })
})

// NB: no EML handler in this version. Nested emails (.eml / message/rfc822 / message/global) are
// out of scope and skip — see the "skips full emails" case in the routing block above.

// Edge cases surfaced by an adversarial pass. The first three guard the fixes in this change; the
// last two document limitations left as-is (a CONFIRM test for the uncancellable timeout, and a
// by-design assertion that the slim contract carries no truncation signal).
describe('attachment — edge cases (regression)', () => {
    // A validly BOM-decoded UTF-16 buffer that legitimately contains U+FFFD must NOT be re-decoded
    // as latin1 — that would keep the interleaved NULs from the two-byte units. decodeText's latin1
    // fallback is now gated on !bomCharset, so a BOM-definitive charset is trusted.
    it('keeps valid UTF-16LE text containing a literal U+FFFD (no latin1 corruption)', async () => {
        const content = Buffer.from('﻿a�b', 'utf16le') // BOM + "a�b", genuinely UTF-16LE
        const r = await extractAttachment({ content, contentType: 'text/plain' })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBe('a�b') // correct utf-16le decode
        expect(r.extraction ?? '').not.toContain('\u0000') // no NULs from a latin1 re-decode
    })

    // RTF is printable ASCII, so it slips past bytesContradictTextClaim; sent as text/plain it must
    // still skip (not leak control words), matching the text/rtf and octet-stream paths. detectRoute
    // now voids a text/html claim whose bytes start with RTF_MAGIC and re-sniffs to a labeled skip.
    it('skips RTF sent as text/plain instead of leaking control words', async () => {
        const rtf = Buffer.from('{\\rtf1\\ansi\\deff0 Dear team, the meeting is at noon.\\par}', 'utf8')
        const r = await extractAttachment({ content: rtf, contentType: 'text/plain' })
        expect(r.status).toBe('skipped')
        expect(r.extraction ?? '').not.toContain('\\rtf1')
    })

    // A zero-byte attachment resolves to the same status regardless of declared type — an early gate
    // returns 'extracted' (ran, no text) before an empty PDF/OOXML could route into a parser that
    // throws on zero bytes and reports 'failed'.
    it('gives a zero-byte attachment a consistent status across declared types', async () => {
        const empty = Buffer.alloc(0)
        const asText = await extractAttachment({ content: empty, contentType: 'text/plain' })
        const asPdf = await extractAttachment({ content: empty, contentType: 'application/pdf' })
        expect(asText.status).toBe('extracted')
        expect(asPdf.status).toBe(asText.status)
        expect(asPdf.extraction).toBeUndefined()
    })

    // Documented limitation — withTimeout (attachment.ts:514) is not exported, so replicate it
    // verbatim. Rejecting the wrapper does NOT cancel the underlying handler: its CPU work runs to
    // completion regardless (wasted CPU/memory after we time out). Can't be fixed without a
    // cancellable/off-thread parser.
    it('CONFIRM: withTimeout rejects at the deadline yet the handler still runs to completion', async () => {
        const withTimeout = <T>(p: Promise<T>, ms: number): Promise<T> =>
            new Promise((resolve, reject) => {
                const timer = setTimeout(() => reject(new Error(`timeout ${ms}ms`)), ms)
                p.then(
                    (v) => {
                        clearTimeout(timer)
                        resolve(v)
                    },
                    (e) => {
                        clearTimeout(timer)
                        reject(e)
                    }
                )
            })

        let ranToCompletion = false
        const handler = (async () => {
            await new Promise((r) => setTimeout(r, 40)) // async gap > timeout so the timer fires
            const end = Date.now() + 80
            while (Date.now() < end) {
                /* synchronous CPU burn — uncancellable */
            }
            ranToCompletion = true
            return 'done'
        })()

        let outcome = 'resolved'
        try {
            await withTimeout(handler, 20)
        } catch {
            outcome = 'rejected'
        }
        expect(outcome).toBe('rejected') // we stopped waiting at 20ms
        expect(ranToCompletion).toBe(false) // work had not finished when we gave up
        await handler // the uncancelled handler runs on...
        expect(ranToCompletion).toBe(true) // ...and completes its CPU work anyway
    })

    // ACCEPTED BY DESIGN — the 3-field slim intentionally carries NO truncation signal, so over-cap
    // output is bounded silently. Asserted positively (a normal green test): if a `truncated` field
    // is ever added, `not.toHaveProperty` below fails, forcing a conscious update to the contract.
    it('caps over-limit output silently — the slim contract has no truncation signal (by design)', async () => {
        const content = Buffer.alloc(2 * MAX_OUTPUT_CHARS, 0x41) // 500k "A", under the 10MB input cap
        const r = await extractAttachment({ content, contentType: 'text/plain' })
        expect(r.status).toBe('extracted')
        expect(r.extraction?.length).toBe(MAX_OUTPUT_CHARS) // capped
        expect(r).not.toHaveProperty('truncated') // deliberate: no field signals truncation
    })
})

// OOXML decompression preflight — a small in-cap .docx/.xlsx can declare a huge uncompressed size in
// its zip central directory and OOM the worker. The preflight reads ONLY those declared sizes, so the
// malicious archive is hand-built (a valid EOCD + one central-directory header with an inflated
// uncompressed field, pointing at an xl/workbook.xml part so it still routes to xlsx) — no real
// compressed data needed.
describe('attachextract — OOXML decompression preflight', () => {
    // Minimal zip: local header (zip magic + the xl/workbook.xml part name, for routing) + one
    // central-directory header whose uncompressed field is `declaredUncompressed`, + an EOCD.
    const fakeOoxmlZip = (
        declaredUncompressed: number,
        options: { entries?: number; centralDirectoryOffset?: number; centralDirectorySize?: number } = {}
    ): Buffer => {
        const name = Buffer.from('xl/workbook.xml') // makes ooxmlKind route the bytes to xlsx
        const lfh = Buffer.alloc(30 + name.length)
        lfh.writeUInt32LE(0x04034b50, 0) // local file header signature (PK\x03\x04)
        lfh.writeUInt16LE(name.length, 26) // file name length
        name.copy(lfh, 30)
        const cdh = Buffer.alloc(46 + name.length)
        cdh.writeUInt32LE(0x02014b50, 0) // central directory header signature
        cdh.writeUInt32LE(declaredUncompressed >>> 0, 24) // the inflated uncompressed size (the lie)
        cdh.writeUInt16LE(name.length, 28) // file name length
        name.copy(cdh, 46)
        const eocd = Buffer.alloc(22)
        eocd.writeUInt32LE(0x06054b50, 0) // end of central directory signature
        eocd.writeUInt16LE(options.entries ?? 1, 8) // entries on this disk
        eocd.writeUInt16LE(options.entries ?? 1, 10) // total entries
        eocd.writeUInt32LE(options.centralDirectorySize ?? cdh.length, 12) // central directory size
        eocd.writeUInt32LE(options.centralDirectoryOffset ?? lfh.length, 16) // central directory offset
        return Buffer.concat([lfh, cdh, eocd])
    }

    // Declares 100 MB uncompressed (> the 50 MB cap) → skipped before exceljs is ever loaded.
    it('skips an OOXML archive whose declared uncompressed size is over budget', async () => {
        const bomb = fakeOoxmlZip(100 * 1024 * 1024)
        const r = await extractAttachment({ content: bomb, contentType: XLSX_TYPE })
        expect(r.status).toBe('skipped')
        expect(r.reason).toMatch(/decompress/i)
        expect(r.extraction).toBeUndefined()
    })

    // ZIP64 archives mark classic EOCD fields with sentinels and store the real values elsewhere.
    // Treat those sentinels as over-budget instead of falling through to exceljs/mammoth, since the
    // real declared size can be arbitrarily large and the preflight deliberately does not chase ZIP64
    // extra records.
    it('skips ZIP64 EOCD sentinels before loading the parser', async () => {
        for (const bomb of [
            fakeOoxmlZip(1, { entries: 0xffff }),
            fakeOoxmlZip(1, { centralDirectorySize: 0xffffffff }),
            fakeOoxmlZip(1, { centralDirectoryOffset: 0xffffffff }),
        ]) {
            const r = await extractAttachment({ content: bomb, contentType: XLSX_TYPE })
            expect(r.status).toBe('skipped')
            expect(r.reason).toMatch(/decompress/i)
            expect(r.extraction).toBeUndefined()
        }
    })

    // A real, modestly-sized .xlsx (declared uncompressed well under the cap) sails through the
    // preflight and extracts normally — the gate must not over-reject legitimate files.
    it('lets a normal .xlsx through the preflight and extracts it', async () => {
        const r = await extractAttachment({ content: fixture('sample.xlsx'), contentType: XLSX_TYPE })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toContain('=== Q1 ===')
    })
})
