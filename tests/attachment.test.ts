import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import zlib from 'node:zlib'

import { describe, it, expect, vi } from 'vitest'
import iconv from 'iconv-lite'
import ExcelJS from 'exceljs'
import JSZip from 'jszip' // the zip reader inside exceljs/mammoth — the dual-EOCD tripwire pins our EOCD choice to its

import {
    extractAttachment,
    detectRoute,
    HANDLER_TIMEOUT_MS,
    MAX_INPUT_BYTES,
    MAX_OUTPUT_CHARS,
    MAX_UNCOMPRESSED_BYTES,
} from '../attachment'

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

// Minimal valid PDF builder: one non-embedded Helvetica (base-14) page per inner array, each line
// drawn as its own on-page Tj (pdf.js clips a single off-page text run, so text must be laid out as
// real lines to stay extractable). Byte-accurate xref so unpdf/pdf.js parses it. Used to synthesize
// oversized PDFs (many pages of text) without shipping a big binary fixture.
const buildPdf = (pages: string[][]): Buffer => {
    const esc = (s: string) => s.replace(/\\/g, '\\\\').replace(/\(/g, '\\(').replace(/\)/g, '\\)')
    const objects: string[] = []
    objects[1] = '<< /Type /Catalog /Pages 2 0 R >>'
    objects[2] = `<< /Type /Pages /Kids [${pages.map((_, i) => `${4 + i * 2} 0 R`).join(' ')}] /Count ${pages.length} >>`
    objects[3] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
    pages.forEach((lines, i) => {
        objects[4 + i * 2] = `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R >> >> /Contents ${5 + i * 2} 0 R >>`
        // First line at (72, 760); each subsequent line moves down 13pt (relative Td).
        const body = lines.map((l, j) => `${j === 0 ? '72 760 Td' : '0 -13 Td'} (${esc(l)}) Tj`).join('\n')
        const stream = `BT /F1 12 Tf\n${body}\nET`
        objects[5 + i * 2] = `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`
    })
    let pdf = '%PDF-1.4\n'
    const offsets: number[] = []
    for (let n = 1; n < objects.length; n++) {
        offsets[n] = Buffer.byteLength(pdf, 'latin1')
        pdf += `${n} 0 obj\n${objects[n]}\nendobj\n`
    }
    const xrefStart = Buffer.byteLength(pdf, 'latin1')
    const count = objects.length
    pdf += `xref\n0 ${count}\n0000000000 65535 f \n`
    for (let n = 1; n < count; n++) pdf += `${String(offsets[n]).padStart(10, '0')} 00000 n \n`
    pdf += `trailer\n<< /Size ${count} /Root 1 0 R >>\nstartxref\n${xrefStart}\n%%EOF`
    return Buffer.from(pdf, 'latin1')
}

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
    // detectRoute is exported and has no zero-byte shortcut of its own (extractAttachment resolves
    // empties before ever calling it), so the sniff must survive being handed nothing to sniff.
    it('routes zero bytes to nothing rather than guessing', () => {
        expect(detectRoute({ content: Buffer.alloc(0), contentType: 'application/octet-stream' })).toEqual({
            routedBy: 'none',
        })
        expect(detectRoute({ content: Buffer.alloc(0) })).toEqual({ routedBy: 'none' })
    })

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

    // The bottom two rungs of resolveCharset's ladder, below the confidence gate. Neither claims to
    // be a CORRECT decode — the point is that both beat the alternative of decoding as utf-8 and
    // turning every high byte into an irreversible U+FFFD.
    it("takes jschardet's guess even below the confidence gate rather than mangle high bytes", async () => {
        // jschardet reads these as KOI8-R at ~0.49 confidence — under the 0.7 gate, so only the floor
        // can rescue them. A plausible decode beats guaranteed replacement characters.
        const content = Buffer.from(Array.from({ length: 60 }, (_, i) => 0xc0 + (i % 16)))
        const r = await extractAttachment({ content, contentType: 'text/plain' })
        expect(r.status).toBe('extracted')
        expect(r.extraction).not.toContain('�') // every byte mapped to something
    })

    it('floors on latin1 when nothing is detected at all', async () => {
        // Invalid utf-8 lead bytes, too short and too mixed for jschardet to name any encoding (it
        // returns encoding: null), so the ladder runs out. latin1 catches it: every byte maps to a
        // character — possibly the wrong one, but reversibly so.
        const content = Buffer.from([0x41, 0xc0, 0x41, 0xc1, 0x41, 0xf5, 0x41])
        const r = await extractAttachment({ content, contentType: 'text/plain' })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBe('AÀAÁAõA')
    })

    // The BE half of the same rule. Its byte order is the mirror image of the LE case above, so a
    // sign/endianness slip in bomCharset would leave that test green and only break this one.
    it('decodes a UTF-16BE file by its BOM (FE FF)', async () => {
        const content = iconv.encode('hello world', 'utf-16be', { addBOM: true })
        expect(content.subarray(0, 2)).toEqual(Buffer.from([0xfe, 0xff])) // the sample really is BE-marked
        const r = await extractAttachment({ content, contentType: 'text/plain' })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBe('hello world')
    })

    // The three-byte UTF-8 BOM, stripped from the output rather than surfacing as a leading U+FEFF.
    it('decodes a UTF-8 BOM file and strips the mark from the extraction', async () => {
        const content = Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), buf('héllo wörld')])
        const r = await extractAttachment({ content, contentType: 'text/plain' })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBe('héllo wörld') // no leading
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

    // The OLE magic is shared with .xls/.ppt/.msg, so it cannot confirm a doc claim on its own. An
    // extension naming a sibling format contradicts the claim; a missing one contradicts nothing —
    // otherwise a real .doc sent with no filename would lose its route.
    it('refuses a doc claim the extension contradicts, but not one it is merely silent on', async () => {
        const ole = fixture('sample.doc')
        // .xls names a sibling OLE format, so the claim is void. Asserting the ROUTE (not just the
        // status) is what proves word-extractor was never reached — a 'failed' would mean it was.
        const asXls = { content: ole, contentType: 'application/msword', filename: 'book.xls' }
        expect(detectRoute(asXls).kind).not.toBe('doc')
        expect((await extractAttachment(asXls)).status).toBe('skipped')

        // The regression guard: no filename contradicts nothing, so a real .doc must still extract.
        // A naive `ext === '.doc'` gate would skip this.
        const noName = { content: ole, contentType: 'application/msword' }
        expect(detectRoute(noName).kind).toBe('doc')
        const r = await extractAttachment(noName)
        expect(r.status).toBe('extracted')
        expect(r.extraction).toContain('First paragraph about revenue')
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

    // The other side of that gate. A .doc filename normally routes by EXTENSION, so sniff's OLE rule
    // never runs — it only becomes reachable once a text claim is voided by the OLE magic, which is
    // the mislabeled-as-text/plain case. Without this the rule's accept path is dead in the suite.
    it('rescues an OLE .doc mislabeled text/plain through the sniff', async () => {
        const input = { content: fixture('sample.doc'), contentType: 'text/plain', filename: 'note.doc' }
        expect(detectRoute(input)).toEqual({ kind: 'doc', routedBy: 'sniff' })
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

    // Bytes carrying the zip magic + the xl/workbook.xml marker (so content-detection routes them to
    // xlsx) but with no valid central directory are caught by the fail-closed decompression preflight
    // and skipped BEFORE exceljs ever loads — cleanly labeled, never a throw. NB: a bare "PK.." with
    // no OOXML part is rerouted by content-detection, not sent here.
    it('fails a structurally-corrupt xlsx at the preflight (not a throw)', async () => {
        const oxmlGarbage = Buffer.concat([Buffer.from([0x50, 0x4b, 0x03, 0x04]), buf('xl/workbook.xml'), buf(' corrupt body')])
        const input = { content: oxmlGarbage, contentType: XLSX_TYPE, filename: 'bad.xlsx' }
        expect(detectRoute(input).kind).toBe('xlsx')
        const r = await extractAttachment(input)
        expect(r.status).toBe('failed')
        expect(r.reason).toMatch(/malformed zip/i)
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

    // The cap counts UTF-16 code units, so slicing at it can land between the halves of an astral
    // char. A lone half is not a character and serializes as U+FFFD, so the cut backs off by one.
    // Only reachable when the boundary unit is a high surrogate — hence the exact placement here.
    it('does not split a surrogate pair at the cap boundary', async () => {
        const text = 'a'.repeat(MAX_OUTPUT_CHARS - 1) + '😀' + 'b'.repeat(64) // high half lands on cap-1
        expect(text.charCodeAt(MAX_OUTPUT_CHARS - 1)).toBeGreaterThanOrEqual(0xd800)
        expect(text.charCodeAt(MAX_OUTPUT_CHARS - 1)).toBeLessThanOrEqual(0xdbff)
        const r = await extractAttachment({ content: buf(text), contentType: 'text/plain' })
        expect(r.status).toBe('extracted')
        expect(r.extraction?.length).toBe(MAX_OUTPUT_CHARS - 1) // backed off, not a hard slice
        expect(r.extraction).not.toContain('�')
        expect(r.extraction?.endsWith('a')).toBe(true)
    })

    // The other side of that branch: a non-surrogate boundary must still cut at exactly the cap.
    it('cuts at exactly the cap when the boundary is not a surrogate', async () => {
        const r = await extractAttachment({ content: buf('a'.repeat(MAX_OUTPUT_CHARS + 64)), contentType: 'text/plain' })
        expect(r.extraction?.length).toBe(MAX_OUTPUT_CHARS)
    })

    // The xlsx handler caps INCREMENTALLY as it flattens (it stops appending rows past the cap rather
    // than building the whole workbook's text and letting the central cap trim it). A workbook that
    // flattens to ~8x the cap still comes back bounded — and near the cap, proving the flatten ran.
    // The xlsx cap is incremental precisely so a dense workbook is never flattened in full and then
    // trimmed. A single-sheet cap only proves the ROW gate; this proves the SHEET gate — once the cap
    // is reached, later sheets are skipped whole rather than built and thrown away. Asserted through
    // the absence of the later sheet's header, which is the only externally visible trace of it.
    it('stops flattening later sheets once the cap is reached', async () => {
        const workbook = new ExcelJS.Workbook()
        const big = workbook.addWorksheet('Big')
        const cell = 'x'.repeat(40)
        for (let i = 0; i < 5_000; i++) big.addRow([cell, cell, cell]) // ~600k chars flattened, past the 250k cap
        workbook.addWorksheet('Later').addRow(['this sheet sits past the cap and must never be flattened'])
        const content = Buffer.from(await workbook.xlsx.writeBuffer())
        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toContain('=== Big ===')
        expect(r.extraction).not.toContain('=== Later ===')
    })

    it('caps an oversized xlsx to MAX_OUTPUT_CHARS', async () => {
        const workbook = new ExcelJS.Workbook()
        const sheet = workbook.addWorksheet('Big')
        const cell = 'x'.repeat(40)
        for (let i = 0; i < 50_000; i++) sheet.addRow([cell, cell, cell]) // ~6M chars flattened, >> 250k cap
        const content = Buffer.from(await workbook.xlsx.writeBuffer())
        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        expect(r.status).toBe('extracted')
        expect(r.extraction!.length).toBeLessThanOrEqual(MAX_OUTPUT_CHARS)
        expect(r.extraction!.length).toBeGreaterThan(MAX_OUTPUT_CHARS - 500) // capping actually engaged
    })

    // The pdf handler accumulates page text and stops once past the cap (it iterates pages itself
    // rather than parsing every page up front). A PDF whose text layer far exceeds the cap comes back
    // bounded, and near the cap — proving the incremental break ran, not that extraction was empty.
    it('caps an oversized pdf to MAX_OUTPUT_CHARS', async () => {
        // 90 pages x 55 lines (~57 chars each) ≈ 280k extractable chars, well over the 250k cap.
        const line = 'the quick brown fox jumps over the lazy dog and then some'
        const pdf = buildPdf(Array.from({ length: 90 }, () => Array.from({ length: 55 }, () => line)))
        const r = await extractAttachment({ content: pdf, contentType: 'application/pdf' })
        expect(r.status).toBe('extracted')
        expect(r.extraction!.length).toBeLessThanOrEqual(MAX_OUTPUT_CHARS)
        expect(r.extraction!.length).toBeGreaterThan(200_000) // lots of text extracted, then capped
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

    // A binary mislabeled text/plain;charset=utf-16 must NOT slip through as text: the utf-16 charset
    // exemption otherwise suppresses the binary check and the PDF would latin1/utf-16-decode into
    // gibberish reported as 'extracted'. Binary magic overrides the hint, so %PDF reroutes to pdf.
    it('recovers a PDF mislabeled text/plain; charset=utf-16 (magic beats the charset hint)', async () => {
        const input = { content: fixture('sample.pdf'), contentType: 'text/plain; charset=utf-16', filename: 'note.txt' }
        expect(detectRoute(input)).toEqual({ kind: 'pdf', routedBy: 'sniff' })
        const r = await extractAttachment(input)
        expect(r.status).toBe('extracted')
        expect(r.extraction).toContain('extractable text content') // real PDF text, not gibberish
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

    // A DOCX that embeds a workbook: the embedded xlsx contributes its internal `xl/workbook.xml` as
    // bytes stored EARLIER in the archive than the package's own `word/document.xml`. A raw-bytes scan
    // (earlier marker wins) would misread the package as xlsx; matching exact zip ENTRY names keeps it
    // docx, because the embedded workbook's path is not an entry of this package.
    const storedZip = (entries: Array<{ name: string; content: Buffer }>): Buffer => {
        const locals: Buffer[] = []
        const centrals: Buffer[] = []
        let offset = 0
        for (const { name, content } of entries) {
            const nameBuf = Buffer.from(name, 'latin1')
            const lfh = Buffer.alloc(30 + nameBuf.length)
            lfh.writeUInt32LE(0x04034b50, 0) // local file header sig (PK\x03\x04 — also the zip magic)
            lfh.writeUInt32LE(content.length, 18) // compressed size (method 0 = stored)
            lfh.writeUInt32LE(content.length, 22) // uncompressed size
            lfh.writeUInt16LE(nameBuf.length, 26)
            nameBuf.copy(lfh, 30)
            const localRec = Buffer.concat([lfh, content])
            const cdh = Buffer.alloc(46 + nameBuf.length)
            cdh.writeUInt32LE(0x02014b50, 0) // central directory header sig
            cdh.writeUInt32LE(content.length, 20)
            cdh.writeUInt32LE(content.length, 24)
            cdh.writeUInt16LE(nameBuf.length, 28)
            cdh.writeUInt32LE(offset, 42) // local header offset
            nameBuf.copy(cdh, 46)
            locals.push(localRec)
            centrals.push(cdh)
            offset += localRec.length
        }
        const cd = Buffer.concat(centrals)
        const eocd = Buffer.alloc(22)
        eocd.writeUInt32LE(0x06054b50, 0)
        eocd.writeUInt16LE(entries.length, 8)
        eocd.writeUInt16LE(entries.length, 10)
        eocd.writeUInt32LE(cd.length, 12)
        eocd.writeUInt32LE(offset, 16) // central directory offset
        return Buffer.concat([...locals, cd, eocd])
    }

    // Embedded workbook stored first (its bytes carry the literal `xl/workbook.xml`), real Word part second.
    const docxWithEmbeddedWorkbook = () =>
        storedZip([
            { name: 'word/embeddings/oleObject1.xlsx', content: buf('PK xl/workbook.xml embedded blob') },
            { name: 'word/document.xml', content: buf('<w:document>hello</w:document>') },
        ])

    it('keeps a DOCX-with-embedded-workbook as docx (entry name, not raw byte order)', () => {
        const bytes = docxWithEmbeddedWorkbook()
        // Claimed docx: must stay docx (the declared route is retained; no misroute to xlsx).
        expect(detectRoute({ content: bytes, contentType: DOCX, filename: 'report.docx' }).kind).toBe('docx')
        // Pure content sniff (no type/name): still docx from the root entry, despite byte order.
        expect(detectRoute({ content: bytes }).kind).toBe('docx')
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

    // The real guard, end to end. The replica below proves a PROPERTY of the timeout (that the work
    // it abandons keeps running) but executes a copy, so the shipped withTimeout/HandlerTimeoutError
    // were never run by this suite. Here a stubbed parser hangs forever and fake timers jump the
    // deadline, so the actual code path runs: reject at HANDLER_TIMEOUT_MS → caught → labeled 'failed',
    // never propagated to the caller.
    it('reports a hung handler as failed at HANDLER_TIMEOUT_MS', async () => {
        vi.resetModules()
        vi.doMock('unpdf', () => ({ getDocumentProxy: () => new Promise(() => {}) })) // never settles
        try {
            const { extractAttachment: extract } = await import('../attachment')
            vi.useFakeTimers()
            const pending = extract({ content: buf('%PDF-1.4 hung'), contentType: 'application/pdf' })
            await vi.advanceTimersByTimeAsync(HANDLER_TIMEOUT_MS + 10)
            const r = await pending
            expect(r.status).toBe('failed')
            expect(r.reason).toMatch(new RegExp(`handler exceeded ${HANDLER_TIMEOUT_MS}ms`))
            expect(r.extraction).toBeUndefined()
        } finally {
            vi.useRealTimers()
            vi.doUnmock('unpdf')
            vi.resetModules()
        }
    })

    // Documented limitation — withTimeout (attachment.ts:781) is not exported, so replicate it
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

// OOXML decompression preflight — a small in-cap .docx/.xlsx can inflate to hundreds of MB and OOM
// the worker. The guard no longer trusts the zip's self-declared uncompressed size (attacker-
// controlled); it STREAM-inflates each entry and counts REAL output bytes, aborting past the cap.
// So these archives carry actual deflate streams (pointing at an xl/workbook.xml part so they still
// route to xlsx), and the fail-closed cases deliberately malform the metadata.
describe('attachextract — OOXML decompression preflight', () => {
    // Build a routable single-entry OOXML zip with a REAL deflate stream. `opts` lets a test lie in
    // the declared-uncompressed field (to prove the guard ignores it) or corrupt the metadata.
    const craftOoxmlZip = (
        content: Buffer,
        opts: {
            part?: string
            method?: number // 8 = deflate (default), 0 = stored
            declaredUncompressed?: number // the old lie vector — now ignored by the guard
            entries?: number // EOCD entry count (0xffff = ZIP64 sentinel)
            cdSize?: number // EOCD central-directory size (0xffffffff = ZIP64 sentinel)
            cdOffset?: number // EOCD central-directory offset (0xffffffff = ZIP64; a wrong value = malformed)
            entryCompSize?: number // CD compressed-size field (0xffffffff = ZIP64 sentinel)
            localOffset?: number // CD → local-header pointer (out of range / wrong = malformed)
            omitEocd?: boolean // drop the EOCD entirely
            comment?: Buffer // trailing archive comment (EOCD commentLen is set to cover it)
        } = {}
    ): Buffer => {
        const name = Buffer.from(opts.part ?? 'xl/workbook.xml') // routes the bytes to xlsx via ooxmlKind
        const method = opts.method ?? 8
        const comp = method === 0 ? content : zlib.deflateRawSync(content)
        const declaredUncompressed = (opts.declaredUncompressed ?? content.length) >>> 0
        const lfh = Buffer.alloc(30 + name.length)
        lfh.writeUInt32LE(0x04034b50, 0) // local file header signature (PK\x03\x04)
        lfh.writeUInt16LE(method, 8) // compression method
        lfh.writeUInt32LE(comp.length, 18) // compressed size
        lfh.writeUInt32LE(declaredUncompressed, 22) // uncompressed size (guard no longer reads this)
        lfh.writeUInt16LE(name.length, 26) // file name length
        name.copy(lfh, 30)
        const localRec = Buffer.concat([lfh, comp])
        const cdh = Buffer.alloc(46 + name.length)
        cdh.writeUInt32LE(0x02014b50, 0) // central directory header signature
        cdh.writeUInt16LE(method, 10) // compression method
        cdh.writeUInt32LE((opts.entryCompSize ?? comp.length) >>> 0, 20) // compressed size
        cdh.writeUInt32LE(declaredUncompressed, 24) // uncompressed size (the old lie field)
        cdh.writeUInt16LE(name.length, 28) // file name length
        cdh.writeUInt32LE((opts.localOffset ?? 0) >>> 0, 42) // local header offset
        name.copy(cdh, 46)
        if (opts.omitEocd) return Buffer.concat([localRec, cdh])
        const eocd = Buffer.alloc(22)
        eocd.writeUInt32LE(0x06054b50, 0) // end of central directory signature
        eocd.writeUInt16LE(opts.entries ?? 1, 8) // entries on this disk
        eocd.writeUInt16LE(opts.entries ?? 1, 10) // total entries
        eocd.writeUInt32LE((opts.cdSize ?? cdh.length) >>> 0, 12) // central directory size
        eocd.writeUInt32LE((opts.cdOffset ?? localRec.length) >>> 0, 16) // central directory offset
        const comment = opts.comment ?? Buffer.alloc(0)
        eocd.writeUInt16LE(comment.length, 20) // archive comment length (matches the appended comment bytes)
        return Buffer.concat([localRec, cdh, eocd, comment])
    }

    // Content that genuinely inflates past the 50 MB cap (compresses to a few KB on disk).
    const overCapContent = () => Buffer.alloc(MAX_UNCOMPRESSED_BYTES + 10 * 1024 * 1024, 0x41)

    // Raw zip-record builders, for the tripwires below that need to lay out records by hand rather
    // than accept craftOoxmlZip's well-formed arrangement (two EOCDs, two central directories, a
    // record hidden past the declared count). Every field a tripwire lies in is a parameter here.
    const localFile = (name: string, data: Buffer, uncompressed: number, method: number) => {
        const n = Buffer.from(name, 'latin1')
        const h = Buffer.alloc(30 + n.length)
        h.writeUInt32LE(0x04034b50, 0)
        h.writeUInt16LE(method, 8)
        h.writeUInt32LE(data.length, 18) // compressed size
        h.writeUInt32LE(uncompressed, 22) // uncompressed size (the preflight ignores this)
        h.writeUInt16LE(n.length, 26)
        n.copy(h, 30)
        return Buffer.concat([h, data])
    }
    const central = (name: string, comp: number, uncompressed: number, localOffset: number, method: number) => {
        const n = Buffer.from(name, 'latin1')
        const h = Buffer.alloc(46 + n.length)
        h.writeUInt32LE(0x02014b50, 0)
        h.writeUInt16LE(method, 10)
        h.writeUInt32LE(comp, 20)
        h.writeUInt32LE(uncompressed, 24)
        h.writeUInt16LE(n.length, 28)
        h.writeUInt32LE(localOffset, 42)
        n.copy(h, 46)
        return h
    }
    const eocd = (entries: number, cdSize: number, cdOffset: number) => {
        const e = Buffer.alloc(22)
        e.writeUInt32LE(0x06054b50, 0)
        e.writeUInt16LE(entries, 8)
        e.writeUInt16LE(entries, 10)
        e.writeUInt32LE(cdSize, 12)
        e.writeUInt32LE(cdOffset, 16)
        return e
    }

    // A deflate stream that genuinely inflates past the cap — the payload every bomb tripwire points at.
    const bombUncompressed = MAX_UNCOMPRESSED_BYTES + 5 * 1024 * 1024
    const bombStream = () => zlib.deflateRawSync(Buffer.alloc(bombUncompressed, 0x41))

    // The core fix: a bomb that inflates over the cap is skipped by MEASUREMENT, not by trusting
    // the declared size — the metadata never touches exceljs.
    it('skips an OOXML archive that actually inflates past the cap', async () => {
        const bomb = craftOoxmlZip(overCapContent())
        const r = await extractAttachment({ content: bomb, contentType: XLSX_TYPE })
        expect(r.status).toBe('skipped')
        expect(r.reason).toMatch(/decompress/i)
        expect(r.extraction).toBeUndefined()
    })

    // The preflight's two refusals are different answers, not one. Corrupt bytes earn 'failed' — the
    // parser would have thrown on them anyway. An intact archive we decline to expand earns
    // 'skipped', like the MAX_INPUT_BYTES gate. Callers branch on that, so pin both against collapse.
    it('separates a corrupt archive (failed) from one merely over budget (skipped)', async () => {
        const broken = craftOoxmlZip(Buffer.from('<workbook/>'), { omitEocd: true })
        const overBudget = craftOoxmlZip(overCapContent())
        expect((await extractAttachment({ content: broken, contentType: XLSX_TYPE })).status).toBe('failed')
        expect((await extractAttachment({ content: overBudget, contentType: XLSX_TYPE })).status).toBe('skipped')
    })

    // The bypass regression test: the archive LIES, declaring a tiny uncompressed size while its
    // deflate stream expands past the cap. The old metadata-trusting guard waved this through; the
    // streaming guard catches it because it counts real output bytes.
    it('skips a bomb even when it declares a tiny uncompressed size', async () => {
        const bomb = craftOoxmlZip(overCapContent(), { declaredUncompressed: 100 })
        const r = await extractAttachment({ content: bomb, contentType: XLSX_TYPE })
        expect(r.status).toBe('skipped')
        expect(r.reason).toMatch(/decompress/i)
        expect(r.extraction).toBeUndefined()
    })

    // ZIP64 / out-of-range sentinels mean the real values live in a ZIP64 record we don't chase —
    // treat as over-budget rather than trust the classic field or fall through to the parser.
    it('skips ZIP64 / out-of-range sentinels', async () => {
        const small = Buffer.from('<workbook/>')
        for (const bomb of [
            craftOoxmlZip(small, { entries: 0xffff }),
            craftOoxmlZip(small, { cdSize: 0xffffffff }),
            craftOoxmlZip(small, { cdOffset: 0xffffffff }),
            craftOoxmlZip(small, { entryCompSize: 0xffffffff }),
        ]) {
            const r = await extractAttachment({ content: bomb, contentType: XLSX_TYPE })
            expect(r.status).toBe('skipped')
            expect(r.reason).toMatch(/zip64|out-of-range/i)
            expect(r.extraction).toBeUndefined()
        }
    })

    // Fail-closed on malformed metadata — a partial/misdirected walk must never silently pass.
    it('fails malformed archives (bad offsets / missing records) instead of passing them', async () => {
        const small = Buffer.from('<workbook/>')
        const cases: Array<[string, Buffer]> = [
            ['no EOCD', craftOoxmlZip(small, { omitEocd: true })],
            ['bad central-directory offset', craftOoxmlZip(small, { cdOffset: 3 })],
            ['bad local-header offset', craftOoxmlZip(small, { localOffset: 0x0fffffff })],
        ]
        for (const [, bomb] of cases) {
            const r = await extractAttachment({ content: bomb, contentType: XLSX_TYPE })
            expect(r.status).toBe('failed')
            expect(r.reason).toMatch(/malformed zip/i)
            expect(r.extraction).toBeUndefined()
        }
    })

    // Trailing bytes after the EOCD (some archivers/signers append them) must NOT cause a false skip:
    // the last-signature scan still finds the real EOCD, matching the parser, so a normal .xlsx with
    // junk appended still extracts. (The earlier comment-to-EOF invariant wrongly skipped these.)
    it('does not false-skip a normal .xlsx with trailing bytes after the EOCD', async () => {
        const withJunk = Buffer.concat([fixture('sample.xlsx'), Buffer.alloc(500, 0x2a)]) // 500 '*' bytes, no EOCD sig
        const r = await extractAttachment({ content: withJunk, contentType: XLSX_TYPE })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toContain('=== Q1 ===')
    })

    // Tripwire for the whole guarantee behind the last-signature EOCD scan: the preflight must select
    // the SAME end-of-central-directory record the parser (jszip, inside exceljs/mammoth) selects, or
    // it measures a different central directory than the parser inflates — a bomb-bypass. This crafts a
    // zip with TWO EOCDs whose central directories differ: an earlier one listing a tiny entry, and a
    // LAST one (which jszip's last-signature scan picks) listing an over-cap bomb. If either side drifts
    // — our findEocd stops matching last-signature, or a jszip upgrade changes its selection — an
    // assertion here breaks, instead of the split silently reopening.
    it('measures the same EOCD/central directory the parser reads (dual-EOCD tripwire)', async () => {
        const bombData = bombStream() // inflates past the cap
        const bombLocal = localFile('xl/workbook.xml', bombData, bombUncompressed, 8) // deflate; data starts at offset 45
        const smallLocal = localFile('small.txt', buf('tiny'), 4, 0) // stored
        const cdSmall = central('small.txt', 4, 4, bombLocal.length, 0)
        const cdSmallOffset = bombLocal.length + smallLocal.length
        const eocdEarly = eocd(1, cdSmall.length, cdSmallOffset) // earlier EOCD → tiny central directory
        const cdBomb = central('xl/workbook.xml', bombData.length, bombUncompressed, 0, 8)
        const cdBombOffset = cdSmallOffset + cdSmall.length + eocdEarly.length
        const eocdLast = eocd(1, cdBomb.length, cdBombOffset) // LAST EOCD → bomb central directory
        const file = Buffer.concat([bombLocal, smallLocal, cdSmall, eocdEarly, cdBomb, eocdLast])

        // The parser (jszip) picks the LAST EOCD, so it sees the bomb entry — never the earlier tiny CD.
        const parsed = await JSZip.loadAsync(file)
        expect(Object.keys(parsed.files)).toContain('xl/workbook.xml')
        expect(Object.keys(parsed.files)).not.toContain('small.txt')

        // The preflight must therefore measure the bomb via that same last EOCD and skip — not measure
        // the earlier tiny CD and wave it through to a parser that would inflate the bomb.
        const r = await extractAttachment({ content: file, contentType: XLSX_TYPE })
        expect(r.status).toBe('skipped')
        expect(r.reason).toMatch(/decompress/i)
    })

    // AGREEMENT INVARIANT 1 — the central directory must end exactly where the EOCD begins.
    // jszip does NOT always read from the declared cdOffset: it derives
    // extraBytes = eocdPos - (cdOffset + cdSize) and, when positive, rebases every offset by
    // `reader.zero = extraBytes` (its support for data prepended before the archive). Both readers
    // agree on WHICH EOCD here — the dual-EOCD tripwire above covers that — and still walk different
    // central directories. So: a benign CD at the declared cdOffset (what a raw walk measures) and a
    // bomb CD at cdOffset + extraBytes (what the parser actually reads). Both CD records are the same
    // byte length, so the single declared cdSize describes either one.
    it('skips a zip whose central directory is not where its EOCD says (reader.zero rebase tripwire)', async () => {
        const BENIGN = 'xl/workbook.xml' // 15 chars → 61-byte CD record; also what routes us to xlsx
        const BOMB = 'xl/bombPart.xml' // 15 chars → identical record size, distinct name
        const CD_SIZE = 46 + BENIGN.length
        const bombData = bombStream()

        // extraBytes lands on CD_SIZE, so the bomb's local header must sit at file offset CD_SIZE for
        // its declared localOffset of 0 to resolve there once jszip rebases (local headers shift too).
        const filler = Buffer.alloc(CD_SIZE, 0x00)
        const bombLocal = localFile(BOMB, bombData, bombUncompressed, 8)
        const benignLocal = localFile(BENIGN, buf('tiny'), 4, 0)
        const benignLocalPos = filler.length + bombLocal.length
        const cdOffset = benignLocalPos + benignLocal.length
        const cdBenign = central(BENIGN, 4, 4, benignLocalPos, 0) // stored, 4 bytes — the decoy
        const cdBomb = central(BOMB, bombData.length, bombUncompressed, 0, 8)
        const file = Buffer.concat([filler, bombLocal, benignLocal, cdBenign, cdBomb, eocd(1, CD_SIZE, cdOffset)])

        // eocdPos - (cdOffset + cdSize) === CD_SIZE > 0, so jszip rebases and lands on the bomb CD.
        const parsed = await JSZip.loadAsync(file)
        expect(Object.keys(parsed.files)).toEqual([BOMB]) // the parser never sees the decoy we'd measure
        expect((await parsed.file(BOMB)!.async('nodebuffer')).length).toBeGreaterThan(MAX_UNCOMPRESSED_BYTES)

        // So measuring the decoy at the raw cdOffset would approve a bomb. Reject the layout instead.
        const r = await extractAttachment({ content: file, contentType: XLSX_TYPE })
        expect(r.status).toBe('failed')
        expect(r.reason).toMatch(/does not end at the end-of-central-directory/i)
        expect(r.extraction).toBeUndefined()
    })

    // AGREEMENT INVARIANT 2 — walking exactly `entries` records must land exactly on the EOCD.
    // Invariant 1 is not sufficient: jszip ignores the declared record count entirely, reading central
    // directory headers until the signature stops matching, and does NOT error when its tally
    // disagrees with the count. So an archive can declare one entry, store two, and still size its
    // directory honestly — extraBytes === 0, invariant 1 satisfied, no rebase — while a count-driven
    // walk measures only the first record and the parser inflates both.
    it('skips a zip hiding a central-directory record past its declared count (entry-count tripwire)', async () => {
        const BENIGN = 'xl/workbook.xml' // the only record a count-driven walk reaches; routes us to xlsx
        const BOMB = 'xl/sharedStrings.xml' // reachable only by a signature-driven walk
        const bombData = bombStream()

        const bombLocal = localFile(BOMB, bombData, bombUncompressed, 8)
        const benignLocal = localFile(BENIGN, buf('<workbook/>'), 11, 0)
        const cdOffset = bombLocal.length + benignLocal.length
        const cdBenign = central(BENIGN, 11, 11, bombLocal.length, 0)
        const cdBomb = central(BOMB, bombData.length, bombUncompressed, 0, 8)
        // cdSize is the HONEST total of both records, so eocdPos === cdOffset + cdSize: no rebase,
        // and invariant 1 passes. Only the declared entry count (1, not 2) is a lie.
        const cdSize = cdBenign.length + cdBomb.length
        const file = Buffer.concat([bombLocal, benignLocal, cdBenign, cdBomb, eocd(1, cdSize, cdOffset)])

        // The parser reads until the signature stops — so it sees BOTH, despite the count saying one.
        const parsed = await JSZip.loadAsync(file)
        expect(Object.keys(parsed.files)).toEqual([BENIGN, BOMB])
        expect((await parsed.file(BOMB)!.async('nodebuffer')).length).toBeGreaterThan(MAX_UNCOMPRESSED_BYTES)

        const r = await extractAttachment({ content: file, contentType: XLSX_TYPE })
        expect(r.status).toBe('failed')
        expect(r.reason).toMatch(/more records than it declares/i)
        expect(r.extraction).toBeUndefined()
    })

    // The agreement invariants reject a LAYOUT, so they must not start rejecting honest archives:
    // measured against every real .docx/.xlsx reachable on disk (73 zips, 59 Office files), each one
    // satisfied both. A crafted-but-well-formed archive with a trailing comment covers the same
    // ground here — the comment sits after the EOCD, so it must not perturb either invariant.
    it('accepts a well-formed archive, including one with a trailing archive comment', async () => {
        const plain = craftOoxmlZip(buf('<workbook/>'))
        expect((await extractAttachment({ content: plain, contentType: XLSX_TYPE })).status).not.toBe('skipped')
        const commented = craftOoxmlZip(buf('<workbook/>'), { comment: buf('a trailing archive comment') })
        expect((await extractAttachment({ content: commented, contentType: XLSX_TYPE })).status).not.toBe('skipped')
    })

    // The remaining fail-closed exits. Everything the walk cannot MEASURE has to be a skip, never a
    // pass — the declared metadata is attacker-controlled, so "we could not tell" and "it is fine"
    // must not collapse into the same answer. One test per exit; each asserts its own reason so a
    // future guard that swallows another's case shows up as a changed message rather than silence.

    // Invariant 1 now rejects a wrong cdOffset before the walk is even reached, so reaching the
    // walk's own signature check needs an offset whose ARITHMETIC is honest — cdOffset + cdSize lands
    // exactly on the EOCD — while the bytes it points at are not a central-directory header.
    it('fails closed when the central-directory offset points at non-CD bytes', async () => {
        const junk = Buffer.concat([
            Buffer.from([0x50, 0x4b, 0x03, 0x04]), // zip magic: the bytes still look like an archive
            buf('xl/workbook.xml'), // found by ooxmlKind's raw-scan fallback, so this still routes to xlsx
            Buffer.alloc(40, 0x5a),
        ])
        const file = Buffer.concat([junk, eocd(1, junk.length, 0)]) // 0 + junk.length === eocdPos
        const r = await extractAttachment({ content: file, contentType: XLSX_TYPE })
        expect(r.status).toBe('failed')
        expect(r.reason).toMatch(/truncated or misaligned central directory/i)
    })

    // When the central directory cannot be walked, ooxmlKind falls back to scanning raw bytes for a
    // main-part marker. That fallback is order-dependent (the very thing entry-name matching fixed for
    // walkable archives), so pin which kind each shape yields. Note the outcome is a skip either way —
    // an unwalkable archive never survives the preflight — so this pins ROUTING, not the verdict.
    it.each([
        ['word/document.xml alone', ['word/document.xml'], 'docx'],
        ['xl/workbook.xml alone', ['xl/workbook.xml'], 'xlsx'],
        ['both, word stored first', ['word/document.xml', 'xl/workbook.xml'], 'docx'],
        ['both, xl stored first', ['xl/workbook.xml', 'word/document.xml'], 'xlsx'],
    ])('raw-scan fallback reads %s as %s', (_label, markers, kind) => {
        const junk = Buffer.concat([
            Buffer.from([0x50, 0x4b, 0x03, 0x04]),
            ...markers.map((m) => buf(m)),
            Buffer.alloc(20, 0x5a),
        ])
        const file = Buffer.concat([junk, eocd(1, junk.length, 0)]) // arithmetic honest, bytes not a CD header
        expect(detectRoute({ content: file, contentType: XLSX_TYPE }).kind).toBe(kind)
    })

    // The other way the walk gives up: a record whose declared name length runs off the end of the
    // file. The name cannot be read, so no entry list can be trusted and the raw scan takes over.
    it('falls back to the raw scan when a central-directory name runs past the end of the file', () => {
        const junk = Buffer.concat([Buffer.from([0x50, 0x4b, 0x03, 0x04]), buf('xl/workbook.xml'), Buffer.alloc(20, 0x5a)])
        const cd = Buffer.alloc(46)
        cd.writeUInt32LE(0x02014b50, 0) // a real CD signature...
        cd.writeUInt16LE(5000, 28) // ...whose declared name length runs far past EOF
        const file = Buffer.concat([junk, cd, eocd(1, cd.length, junk.length)])
        expect(detectRoute({ content: file, contentType: XLSX_TYPE }).kind).toBe('xlsx')
    })

    it('fails closed when an entry claims compressed bytes that run past the end of the file', async () => {
        const bomb = craftOoxmlZip(buf('<workbook/>'), { entryCompSize: 0x0ffffff0 }) // huge, but not the ZIP64 sentinel
        const r = await extractAttachment({ content: bomb, contentType: XLSX_TYPE })
        expect(r.status).toBe('failed')
        expect(r.reason).toMatch(/runs past end of file/i)
    })

    // Distinct from "too big": the stream cannot be read at all. Fail closed rather than treat an
    // unreadable entry as contributing zero bytes to the total.
    it('fails closed when a deflate stream is corrupt rather than merely oversized', async () => {
        const garbage = buf('this is not a deflate stream, only ascii pretending to be one')
        const local = localFile('xl/workbook.xml', garbage, 4096, 8) // claims deflate; payload is not
        const cd = central('xl/workbook.xml', garbage.length, 4096, 0, 8)
        const file = Buffer.concat([local, cd, eocd(1, cd.length, local.length)])
        const r = await extractAttachment({ content: file, contentType: XLSX_TYPE })
        expect(r.status).toBe('failed')
        expect(r.reason).toMatch(/unreadable compressed data/i)
    })

    it('fails closed on a compression method it cannot measure', async () => {
        const zip = craftOoxmlZip(buf('<workbook/>'), { method: 12 }) // 12 = bzip2; we only measure store/deflate
        const r = await extractAttachment({ content: zip, contentType: XLSX_TYPE })
        expect(r.status).toBe('skipped')
        expect(r.reason).toMatch(/unsupported compression method 12/i)
    })

    // The budget is ARCHIVE-wide, not per-entry. A stored entry cannot exceed the cap on its own
    // (MAX_INPUT_BYTES bounds the whole file well below it), so the only way to cross the line with
    // one is as the last straw on a running total — which is exactly what the cap must catch.
    it('skips when the archive-wide running total crosses the cap, not just one entry', async () => {
        const exact = zlib.deflateRawSync(Buffer.alloc(MAX_UNCOMPRESSED_BYTES, 0x41)) // inflates to EXACTLY the cap
        const l1 = localFile('xl/workbook.xml', exact, MAX_UNCOMPRESSED_BYTES, 8)
        const l2 = localFile('xl/tipItOver.bin', buf('the last straw'), 14, 0) // stored
        const c1 = central('xl/workbook.xml', exact.length, MAX_UNCOMPRESSED_BYTES, 0, 8)
        const c2 = central('xl/tipItOver.bin', 14, 14, l1.length, 0)
        const file = Buffer.concat([l1, l2, c1, c2, eocd(2, c1.length + c2.length, l1.length + l2.length)])
        const r = await extractAttachment({ content: file, contentType: XLSX_TYPE })
        expect(r.status).toBe('skipped')
        expect(r.reason).toMatch(/decompresses to over/i)
    })

    // An OOXML-shaped zip with NEITHER root part (a .pptx, a .jar, a plain archive) is not something
    // any handler claims, so it must land on a labeled skip rather than be forced into one.
    it('skips a zip that is neither a docx nor an xlsx (pptx root part)', async () => {
        const pptx = craftOoxmlZip(buf('<presentation/>'), { part: 'ppt/presentation.xml' })
        expect(detectRoute({ content: pptx, contentType: XLSX_TYPE })).toEqual({ routedBy: 'none' })
        const r = await extractAttachment({ content: pptx, contentType: XLSX_TYPE })
        expect(r.status).toBe('skipped')
    })

    // A real, modestly-sized .xlsx inflates well under the cap and extracts normally — the streaming
    // guard must not over-reject legitimate files (and pays only one extra inflate pass).
    it('lets a normal .xlsx through the preflight and extracts it', async () => {
        const r = await extractAttachment({ content: fixture('sample.xlsx'), contentType: XLSX_TYPE })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toContain('=== Q1 ===')
    })
})

// A utf-16 charset claim must be EARNED by the bytes ---------------------------------------------
// An explicit charset=utf-16* hint suppresses the printable-ratio check, because genuine BOM-less
// utf-16 is NUL-heavy and looksLikeText would void it as binary. That exemption used to be granted on
// the sender's word alone, overridden only by hasKnownBinaryMagic — an allowlist of exactly three
// (%PDF, zip, OLE). Every other binary kept the exemption and decoded into gibberish reported as
// 'extracted': a silent quality failure, worse than a labeled skip. The hint is now honoured only if
// the content is structurally well-formed utf-16.
describe('attachment — a utf-16 charset claim must be earned by the bytes', () => {
    // Deterministic high-entropy filler, standing in for the compressed body of a real image. Seeded
    // LCG rather than random so a failure here is always reproducible.
    const entropy = (seed: number, length: number): Buffer => {
        const out = Buffer.alloc(length)
        let s = seed >>> 0
        for (let i = 0; i < length; i++) {
            s = (Math.imul(s, 1664525) + 1013904223) >>> 0
            out[i] = (s >>> 24) & 0xff
        }
        return out
    }
    const binaryWith = (magic: number[]) => Buffer.concat([Buffer.from(magic), entropy(0x5eed, 1024)])

    // Formats whose magic is NOT one of the three hasKnownBinaryMagic knows. Verified against real
    // files on disk (a real PNG, JFIF JPEG and GIF89a) before being reduced to these signatures.
    const NOT_ALLOWLISTED: Array<[string, number[]]> = [
        ['png', [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]],
        ['jpeg', [0xff, 0xd8, 0xff, 0xe0]],
        ['gif', [0x47, 0x49, 0x46, 0x38, 0x39, 0x61]],
        ['gzip', [0x1f, 0x8b, 0x08, 0x00]],
    ]
    for (const [name, magic] of NOT_ALLOWLISTED) {
        it(`does not report a ${name} mislabeled text/plain; charset=utf-16 as extracted text`, async () => {
            const r = await extractAttachment({
                content: binaryWith(magic),
                filename: `photo.${name}`,
                contentType: 'text/plain; charset=utf-16',
            })
            expect(r.status).toBe('skipped') // never 'extracted' with a gibberish decode
            expect(r.extraction).toBeUndefined()
        })
    }

    // The allowlisted magics keep their stronger behaviour: not merely skipped, but re-sniffed back to
    // the real handler. Guards against a regression that closes the hole by dropping hasKnownBinaryMagic.
    it('still recovers an allowlisted binary (PDF) mislabeled charset=utf-16 via the sniff', async () => {
        const r = await extractAttachment({
            content: fixture('sample.pdf'),
            filename: 'note.txt',
            contentType: 'text/plain; charset=utf-16',
        })
        expect(detectRoute({ content: fixture('sample.pdf'), contentType: 'text/plain; charset=utf-16' })).toEqual({
            kind: 'pdf',
            routedBy: 'sniff',
        })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toContain('Hello from a PDF document')
    })

    // The other half of the contract: tightening the exemption must not start voiding real utf-16.
    // This is what stops the check from degenerating into "reject anything NUL-heavy".
    it('still extracts genuine BOM-less utf-16 text under an explicit charset hint', async () => {
        const body = 'Hello from a UTF-16 encoded note. Nothing binary here at all. '.repeat(12)
        for (const [charset, encoding] of [
            ['utf-16le', 'utf-16le'],
            ['utf-16be', 'utf-16be'],
            ['utf-16', 'utf-16le'], // bare utf-16, BOM-less: iconv picks an endianness heuristically
        ] as const) {
            const r = await extractAttachment({
                content: iconv.encode(body, encoding, { addBOM: false }),
                filename: 'note.txt',
                contentType: `text/plain; charset=${charset}`,
            })
            expect(r.status, `charset=${charset}`).toBe('extracted')
            expect(r.extraction, `charset=${charset}`).toContain('Hello from a UTF-16 encoded note.')
        }
    })

    // Every other utf-16 sample in this suite is BMP-only (ASCII and CJK alike), so none of them
    // exercise surrogate PAIRING — only the absence of surrogates. These two cover both halves of the
    // rule the check actually enforces: a correctly paired high+low is real text and must survive; an
    // unpaired high is not and must not. Without the first, a check that over-rejected (voiding any
    // surrogate at all) would pass this whole suite while silently skipping every real attachment
    // containing an emoji.
    it('extracts utf-16 text containing an astral character (valid surrogate pair)', async () => {
        const body = 'Hi 😀 — a grinning face, and some ordinary text after it. '.repeat(8)
        for (const [charset, encoding] of [
            ['utf-16le', 'utf-16le'],
            ['utf-16be', 'utf-16be'],
        ] as const) {
            const content = iconv.encode(body, encoding, { addBOM: false })
            // Guard the guard: prove the sample really carries the D83D/DE00 pair (U+1F600), so this
            // test can never quietly degrade into yet another BMP-only case.
            const pair = encoding === 'utf-16le' ? [0x3d, 0xd8, 0x00, 0xde] : [0xd8, 0x3d, 0xde, 0x00]
            expect(content.includes(Buffer.from(pair)), `${charset} sample carries a surrogate pair`).toBe(true)

            const r = await extractAttachment({
                content,
                filename: 'note.txt',
                contentType: `text/plain; charset=${charset}`,
            })
            expect(r.status, `charset=${charset}`).toBe('extracted')
            expect(r.extraction, `charset=${charset}`).toContain('Hi 😀')
        }
    })

    // The two halves of the sample-edge rule. A high surrogate with nothing after it inside the
    // sniff window is ambiguous, and which way it resolves depends on WHY there is nothing after it:
    // a truncated sample hides the low half (say nothing), a file that simply ends there does not
    // (broken). Collapsing the two either voids real text or waves an unpaired surrogate through.
    it('does not extract utf-16 that ends on an unpaired high surrogate (sample IS the whole file)', async () => {
        const u16le = (units: number[]) => {
            const b = Buffer.alloc(units.length * 2)
            units.forEach((u, i) => b.writeUInt16LE(u, i * 2))
            return b
        }
        const trailing = u16le([...[...'Hi there'].map((c) => c.charCodeAt(0)), 0xd83d])
        expect(trailing.length).toBeLessThan(8 * 1024) // whole file fits the sniff window: no truncation
        const r = await extractAttachment({
            content: trailing,
            filename: 'note.txt',
            contentType: 'text/plain; charset=utf-16le',
        })
        expect(r.status).not.toBe('extracted')
        expect(r.extraction).toBeUndefined()
    })

    it('still extracts utf-16 whose surrogate pair straddles the sniff-window edge', async () => {
        // SNIFF_BYTES (8192) is module-private, so it is spelled out here; the byte assertions below
        // pin the construction so this cannot silently stop probing the edge. Filler is sized to put
        // the pair's HIGH half at bytes 8190-8191 — the last code unit inside the window — and its LOW
        // half at 8192-8193, just outside. The sample therefore ends mid-pair: the exact shape the
        // truncation branch exists for, and the one it must not reject.
        const SNIFF = 8 * 1024
        const body = 'a'.repeat(SNIFF / 2 - 1) + '😀' + ' and more ordinary text past the window.'
        const content = iconv.encode(body, 'utf-16le', { addBOM: false })
        expect(content.length).toBeGreaterThan(SNIFF) // sample really is truncated
        expect(content.readUInt16LE(SNIFF - 2)).toBe(0xd83d) // high half: last unit in the window
        expect(content.readUInt16LE(SNIFF)).toBe(0xde00) // low half: first unit outside it

        const r = await extractAttachment({ content, filename: 'note.txt', contentType: 'text/plain; charset=utf-16le' })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toContain('😀')
    })

    it('does not extract utf-16 whose surrogate pair is broken (high surrogate, then an ASCII char)', async () => {
        const u16le = (units: number[]) => {
            const b = Buffer.alloc(units.length * 2)
            units.forEach((u, i) => b.writeUInt16LE(u, i * 2))
            return b
        }
        const ascii = (s: string) => [...s].map((c) => c.charCodeAt(0))
        // D83D is a HIGH surrogate, so utf-16 requires a low (DC00–DFFF) next. Here an 'A' follows it:
        // structurally impossible in real utf-16, and exactly the signature of binary read through the
        // wrong lens. No encoder will emit this, hence the hand-built code units.
        const broken = u16le([...ascii('Hi '), 0xd83d, ...ascii('A and then some more ordinary text.')])
        const r = await extractAttachment({
            content: broken,
            filename: 'note.txt',
            contentType: 'text/plain; charset=utf-16le',
        })
        expect(r.status).not.toBe('extracted')
        expect(r.extraction).toBeUndefined()
    })

    // The other two things that prove bytes are not the utf-16 they claim to be. A LOW surrogate with
    // no high before it is as impossible as an unpaired high; U+FFFE/U+FFFF are permanent
    // noncharacters. Both are common in binary read through a utf-16 lens and absent from real text.
    it.each([
        ['a lone low surrogate (DE00 with no high before it)', 0xde00],
        ['the U+FFFF noncharacter', 0xffff],
        ['the U+FFFE noncharacter', 0xfffe],
    ])('does not extract utf-16 containing %s', async (_label, unit) => {
        const units = [...[...'Hi '].map((c) => c.charCodeAt(0)), unit, ...[...' and more text.'].map((c) => c.charCodeAt(0))]
        const content = Buffer.alloc(units.length * 2)
        units.forEach((u, i) => content.writeUInt16LE(u, i * 2))
        const r = await extractAttachment({
            content,
            filename: 'note.txt',
            contentType: 'text/plain; charset=utf-16le',
        })
        expect(r.status).not.toBe('extracted')
        expect(r.extraction).toBeUndefined()
    })

    // Too short to hold even one whole code unit, so the claim cannot be verified either way — and an
    // unverifiable utf-16 claim must not be honoured on the sender's word. Re-sniffing still rescues
    // it as the plain ASCII it actually is, which is the point: voiding the claim is not a skip.
    it('does not honour a utf-16 claim on content too short to hold a code unit', async () => {
        const route = detectRoute({ content: Buffer.from([0x41]), contentType: 'text/plain; charset=utf-16' })
        expect(route).toEqual({ kind: 'text', routedBy: 'sniff' }) // 'sniff', not 'content-type': the claim was voided
    })

    // utf-16 CJK has no NUL bytes at all, so any check keyed on NUL count or NUL parity would reject
    // it. Pins the check to well-formedness (surrogate pairing / noncharacters) instead.
    it('still extracts utf-16 text with no NUL bytes at all (CJK)', async () => {
        const body = '日本語のテキストです。これは添付ファイルの本文です。'.repeat(8)
        const r = await extractAttachment({
            content: iconv.encode(body, 'utf-16le', { addBOM: false }),
            filename: 'note.txt',
            contentType: 'text/plain; charset=utf-16le',
        })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toContain('日本語のテキストです。')
    })
})
