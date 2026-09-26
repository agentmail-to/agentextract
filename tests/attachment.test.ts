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
    MAX_PDF_PAGES,
    MAX_UNCOMPRESSED_BYTES,
} from '../attachment'

// NOTE: result shape is { status, extraction?, reason?, truncated?, emptyReason? }; `extraction` is
// omitted (never '') when a handler runs but produces no text, `truncated` says whether the document
// continues past it, and `emptyReason` says WHY there is no text on exactly the results that have
// none. Nested emails (.eml) are out of scope in this version and skip.

// Real fixtures generated once with macOS textutil (.docx) and cupsfilter (.pdf), and exceljs (.xlsx).
// vitest runs from the repo root, so resolve against cwd.
const fixture = (name: string) => readFileSync(join(process.cwd(), 'tests', 'fixtures', name))
const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

const stubExcelReaderHooks = (reader: object, omit?: string) => {
    const hooks: Record<string, unknown> = {
        _parseRels: async () => undefined,
        _parseWorkbook: async () => undefined,
        _parseSharedStrings: async function* () {},
        _parseStyles: async () => undefined,
        _parseWorksheet: () => [],
        _parseHyperlinks: () => [],
    }
    if (omit) delete hooks[omit]
    Object.assign(reader, hooks)
}

// ---------------------------------------------------------------------------
// Synthetic tests for attachment.ts
// ---------------------------------------------------------------------------
// The RESULT is deliberately slim: { status, extraction?, reason?, truncated? }. Routing is NOT on
// the result — it's a separate concern verified through detectRoute(). So routing-decision
// assertions here call detectRoute(input) directly; extraction assertions read the slim result.

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
// The whole point of the slim contract: the top-level result carries no routing/diagnostic noise
// and no top-level filename. If any of that leaks back in, this fails. `truncated` is the one field
// deliberately allowed back — a partial extraction that reads as complete is worse than a missing one.

describe('attachment — slim result contract', () => {
    it('returns only the four contract fields, never routing/diagnostic noise', async () => {
        const r = await extractAttachment({ content: buf('hello'), contentType: 'text/plain', filename: 'hi.txt' })
        // Every key present must be one of the four contract fields — nothing else.
        const allowed = ['status', 'extraction', 'reason', 'truncated']
        expect(Object.keys(r).every((k) => allowed.includes(k))).toBe(true)
        // The removed fields must be absent (not merely undefined-valued).
        for (const gone of ['filename', 'byteSize', 'detectedType', 'routedBy', 'charset', 'extractedText', 'children', 'lowTextDensity', 'pageCount', 'emptyPageCount']) {
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

    it.each(['ucs-2', 'ucs2', 'utf16le', 'utf16', 'utf16be'])(
        'accepts the iconv UTF-16 alias %s in Content-Type',
        async (charset) => {
            const encoding = charset.endsWith('be') ? 'utf-16be' : 'utf-16le'
            const content = iconv.encode('hello world', encoding, { addBOM: false })
            const r = await extractAttachment({ content, contentType: `text/plain; charset=${charset}` })
            expect(r).toMatchObject({ status: 'extracted', extraction: 'hello world', truncated: false })
        }
    )

    // U+FFFD is also a legitimate code point in explicitly declared, BOM-less UTF-16. Its presence
    // must not activate the latin1 recovery path and expose an interleaved NUL after every character.
    it('keeps a literal U+FFFD in BOM-less UTF-16 declared by Content-Type', async () => {
        const content = iconv.encode('before � after', 'utf-16le', { addBOM: false })
        const r = await extractAttachment({ content, contentType: 'text/plain; charset=utf-16le' })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBe('before � after')
        expect(r.extraction).not.toContain('\u0000')
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
        expect(r.reason).toBe('too-large')
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

    it('destroys the PDF loading task after both successful and failed page reads', async () => {
        const destroy = vi.fn(async () => undefined)
        let fail = false
        vi.resetModules()
        vi.doMock('unpdf', () => ({
            getDocumentProxy: async () => ({
                numPages: 1,
                loadingTask: { destroy },
                getPage: async () => ({
                    getTextContent: async () => {
                        if (fail) throw new Error('page failed')
                        return { items: [{ str: 'page text' }] }
                    },
                }),
            }),
        }))
        try {
            const { extractAttachment: extract } = await import('../attachment')
            const input = { content: buf('%PDF-1.4 mock'), contentType: 'application/pdf' }
            expect((await extract(input)).status).toBe('extracted')
            fail = true
            expect((await extract(input)).status).toBe('failed')
            expect(destroy).toHaveBeenCalledTimes(2)
        } finally {
            vi.doUnmock('unpdf')
            vi.resetModules()
        }
    })

    it('preserves successful output and the primary parse error when PDF teardown rejects', async () => {
        let fail = false
        vi.resetModules()
        vi.doMock('unpdf', () => ({
            getDocumentProxy: async () => ({
                numPages: 1,
                loadingTask: { destroy: async () => Promise.reject(new Error('teardown failed')) },
                getPage: async () => ({
                    getTextContent: async () => {
                        if (fail) throw new Error('page failed')
                        return { items: [{ str: 'page text' }] }
                    },
                }),
            }),
        }))
        try {
            const { extractAttachment: extract } = await import('../attachment')
            const input = { content: buf('%PDF-1.4 mock'), contentType: 'application/pdf' }
            expect(await extract(input)).toMatchObject({
                status: 'extracted',
                extraction: 'page text',
                truncated: false,
            })
            fail = true
            // An error escaping a third-party parser is the file's fault by default, which is what
            // 'malformed' says. The point of the case is that teardown rejecting afterwards does not
            // replace it, so the reason belongs to the parse and not to the cleanup.
            expect(await extract(input)).toMatchObject({ status: 'failed', reason: 'malformed' })
        } finally {
            vi.doUnmock('unpdf')
            vi.resetModules()
        }
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

    it.each([
        ['empty', ''],
        ['declaration-only', '<?xml version="1.0"?>'],
        ['comment-only', '<!-- optional styles omitted -->'],
    ])('accepts an %s styles part without losing worksheet rows', async (_label, styles) => {
        const workbook = new ExcelJS.Workbook()
        workbook.addWorksheet('Styles').addRow(['KEEP'])
        const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer())
        zip.file('xl/styles.xml', styles)

        const r = await extractAttachment({
            content: await zip.generateAsync({ type: 'nodebuffer' }),
            contentType: XLSX_TYPE,
        })
        expect(r).toMatchObject({ status: 'extracted', extraction: '=== Styles ===\nKEEP', truncated: false })
    })

    it('accepts an empty optional shared-string part when no cell references it', async () => {
        const workbook = new ExcelJS.Workbook()
        workbook.addWorksheet('Numbers').addRow([42])
        const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer())
        zip.file('xl/sharedStrings.xml', '')

        const r = await extractAttachment({
            content: await zip.generateAsync({ type: 'nodebuffer' }),
            contentType: XLSX_TYPE,
        })
        expect(r).toMatchObject({ status: 'extracted', extraction: '=== Numbers ===\n42', truncated: false })
    })

    it('keeps an explicitly empty shared-string cell in its row position', async () => {
        const workbook = new ExcelJS.Workbook()
        workbook.addWorksheet('Shared').addRow(['A', 'B', 'C'])
        const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer())
        const shared = await zip.file('xl/sharedStrings.xml')!.async('string')
        zip.file('xl/sharedStrings.xml', shared.replace('<t>B</t>', '<t></t>'))

        const r = await extractAttachment({
            content: await zip.generateAsync({ type: 'nodebuffer' }),
            contentType: XLSX_TYPE,
        })
        expect(r).toMatchObject({ status: 'extracted', extraction: '=== Shared ===\nA\t\tC', truncated: false })
    })

    it('does not revive blank rows or trailing columns from empty shared strings', async () => {
        const blank = new ExcelJS.Workbook()
        const blankSheet = blank.addWorksheet('Blank')
        blankSheet.addRow(['A', 'B'])
        blankSheet.addRow(['A', 'B'])
        const blankZip = await JSZip.loadAsync(await blank.xlsx.writeBuffer())
        const shared = await blankZip.file('xl/sharedStrings.xml')!.async('string')
        blankZip.file('xl/sharedStrings.xml', shared.replace(/<t>[AB]<\/t>/g, '<t></t>'))
        const blankResult = await extractAttachment({
            content: await blankZip.generateAsync({ type: 'nodebuffer' }),
            contentType: XLSX_TYPE,
        })
        expect(blankResult).toMatchObject({ status: 'extracted', truncated: false })
        expect(blankResult.extraction).toBeUndefined()

        const trailing = new ExcelJS.Workbook()
        trailing.addWorksheet('Trailing').addRow(['KEEP', 'DROP'])
        const trailingZip = await JSZip.loadAsync(await trailing.xlsx.writeBuffer())
        const trailingShared = await trailingZip.file('xl/sharedStrings.xml')!.async('string')
        trailingZip.file('xl/sharedStrings.xml', trailingShared.replace('<t>DROP</t>', '<t></t>'))
        const trailingResult = await extractAttachment({
            content: await trailingZip.generateAsync({ type: 'nodebuffer' }),
            contentType: XLSX_TYPE,
        })
        expect(trailingResult).toMatchObject({
            status: 'extracted',
            extraction: '=== Trailing ===\nKEEP',
            truncated: false,
        })
    })

    it('matches ExcelJS decimal parseInt semantics for shared-string indexes', async () => {
        const workbook = new ExcelJS.Workbook()
        workbook.addWorksheet('Indexes').addRow(['ZERO', 'ONE', 'TWO', 'THREE'])
        const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer())
        const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
        zip.file(
            'xl/worksheets/sheet1.xml',
            sheet
                .replace('<c r="A1" t="s"><v>0</v></c>', '<c r="A1" t="s"><v>1e3</v></c>')
                .replace('<c r="B1" t="s"><v>1</v></c>', '<c r="B1" t="s"><v>1.5</v></c>')
                .replace('<c r="C1" t="s"><v>2</v></c>', '<c r="C1" t="s"><v>1abc</v></c>')
                .replace('<c r="D1" t="s"><v>3</v></c>', '<c r="D1" t="s"><v>0x10</v></c>')
        )

        const r = await extractAttachment({
            content: await zip.generateAsync({ type: 'nodebuffer' }),
            contentType: XLSX_TYPE,
        })
        expect(r).toMatchObject({
            status: 'extracted',
            extraction: '=== Indexes ===\nONE\tONE\tONE\tZERO',
            truncated: false,
        })
    })

    it('ignores a shared-string cell with no index instead of treating it as index zero', async () => {
        const workbook = new ExcelJS.Workbook()
        workbook.addWorksheet('Missing index').addRow(['ZERO', 42])
        const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer())
        const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
        zip.file('xl/worksheets/sheet1.xml', sheet.replace('<c r="A1" t="s"><v>0</v></c>', '<c r="A1" t="s"/>'))

        const r = await extractAttachment({
            content: await zip.generateAsync({ type: 'nodebuffer' }),
            contentType: XLSX_TYPE,
        })
        expect(r).toMatchObject({ status: 'extracted', extraction: '=== Missing index ===\n42', truncated: false })
    })

    it('bounds a hostile shared-string index instead of accumulating its digits', async () => {
        const workbook = new ExcelJS.Workbook()
        workbook.addWorksheet('Index').addRow(['VALUE'])
        const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer())
        const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
        zip.file('xl/worksheets/sheet1.xml', sheet.replace('<v>0</v>', `<v>${'9'.repeat(100_000)}</v>`))

        const r = await extractAttachment({
            content: await zip.generateAsync({ type: 'nodebuffer' }),
            contentType: XLSX_TYPE,
        })
        expect(r.status).toBe('failed')
        expect(r.reason).toBe('malformed')
    })

    it('fails explicitly when cells reference a missing shared-string table', async () => {
        const workbook = new ExcelJS.Workbook()
        workbook.addWorksheet('S').addRow(['Alpha', 42])
        workbook.addWorksheet('T').addRow(['Beta'])
        const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer())
        zip.remove('xl/sharedStrings.xml')

        const r = await extractAttachment({
            content: await zip.generateAsync({ type: 'nodebuffer' }),
            contentType: XLSX_TYPE,
        })
        expect(r.status).toBe('failed')
        expect(r.extraction).toBeUndefined()
        expect(r.reason).toBe('malformed')
    })

    // A formula cell must extract its computed VALUE, not the "=SUM(...)" formula string.
    it('extracts the computed value of a formula, not the formula text', async () => {
        const r = await extractAttachment({ content: fixture('sample.xlsx'), contentType: XLSX_TYPE })
        expect(r.extraction).toContain('Total\t7300')
        expect(r.extraction).not.toContain('SUM(')
    })

    it('extracts shared-formula continuation results instead of their raw value objects', async () => {
        const workbook = new ExcelJS.Workbook()
        const sheet = workbook.addWorksheet('Shared')
        sheet.getCell('A1').value = 2
        sheet.fillFormula('B1:B2', 'A1+1', [3, 6])

        const r = await extractAttachment({
            content: Buffer.from(await workbook.xlsx.writeBuffer()),
            contentType: XLSX_TYPE,
        })
        expect(r).toMatchObject({ status: 'extracted', truncated: false })
        expect(r.extraction).toBe('=== Shared ===\n2\t3\n6')
        expect(r.extraction).not.toContain('{"formula"')
    })

    it('keeps a zero-valued shared-formula continuation result', async () => {
        const workbook = new ExcelJS.Workbook()
        const sheet = workbook.addWorksheet('Shared zero')
        sheet.getCell('A1').value = 1
        sheet.fillFormula('B1:B2', 'A1-1', [1, 0])

        const r = await extractAttachment({
            content: Buffer.from(await workbook.xlsx.writeBuffer()),
            contentType: XLSX_TYPE,
        })
        expect(r).toMatchObject({ status: 'extracted', truncated: false })
        expect(r.extraction).toBe('=== Shared zero ===\n1\t1\n0')
    })

    it('preserves boolean and date result types for streamed formula cells', async () => {
        const workbook = new ExcelJS.Workbook()
        workbook.properties.date1904 = true
        const sheet = workbook.addWorksheet('Typed formulas')
        const date = new Date(Date.UTC(2024, 0, 2))
        sheet.getCell('A1').value = { formula: 'DATE(2024,1,2)', result: date }
        sheet.getCell('A2').value = { formula: '1=1', result: true }
        sheet.getCell('A3').value = { formula: '1=0', result: false }

        const r = await extractAttachment({
            content: Buffer.from(await workbook.xlsx.writeBuffer()),
            contentType: XLSX_TYPE,
        })
        expect(r).toMatchObject({ status: 'extracted', truncated: false })
        expect(r.extraction).toBe(`=== Typed formulas ===\n${date.toString()}\ntrue\nfalse`)
    })

    it.each([
        ['type', (xml: string) => xml.replace('t="b"', 't="&#98;"')],
        ['address', (xml: string) => xml.replace('r="A1"', 'r="&#x41;1"')],
    ])('decodes an XML-encoded formula %s attribute before matching the streamed cell', async (_label, encode) => {
        const workbook = new ExcelJS.Workbook()
        workbook.addWorksheet('Encoded').getCell('A1').value = { formula: '1=1', result: true }
        const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer())
        const name = 'xl/worksheets/sheet1.xml'
        zip.file(name, encode(await zip.file(name)!.async('string')))

        const r = await extractAttachment({
            content: await zip.generateAsync({ type: 'nodebuffer' }),
            contentType: XLSX_TYPE,
        })
        expect(r).toMatchObject({ status: 'extracted', extraction: '=== Encoded ===\ntrue', truncated: false })
    })

    it('does not coerce non-finite formula results into dates or booleans', async () => {
        const workbook = new ExcelJS.Workbook()
        const sheet = workbook.addWorksheet('Invalid')
        sheet.getCell('A1').value = { formula: 'DATE(2024,1,2)', result: 45_293 }
        sheet.getCell('A1').numFmt = 'yyyy-mm-dd'
        sheet.getCell('B1').value = { formula: '1=1', result: true }
        const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer())
        const name = 'xl/worksheets/sheet1.xml'
        const xml = await zip.file(name)!.async('string')
        zip.file(
            name,
            xml
                .replace(/(<c r="A1"[^>]*)(>.*?<v>)[^<]*(<\/v>)/, '$1 t="e"$2#N/A$3')
                .replace(/(<c r="B1"[^>]*t="b"[^>]*>.*?<v>)[^<]*(<\/v>)/, '$1not-a-number$2')
        )

        const r = await extractAttachment({
            content: await zip.generateAsync({ type: 'nodebuffer' }),
            contentType: XLSX_TYPE,
        })
        expect(r).toMatchObject({ status: 'extracted', truncated: false })
        expect(r.extraction).not.toContain('Invalid Date')
        expect(r.extraction).not.toContain('true')
    })

    it('ignores boolean-looking cells outside worksheet rows', async () => {
        const workbook = new ExcelJS.Workbook()
        workbook.addWorksheet('Scoped').getCell('A1').value = { formula: '1+0', result: 1 }
        const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer())
        const name = 'xl/worksheets/sheet1.xml'
        const xml = await zip.file(name)!.async('string')
        zip.file(name, xml.replace('<sheetData>', '<extLst><c r="A1" t="b"/></extLst><sheetData>'))

        const r = await extractAttachment({
            content: await zip.generateAsync({ type: 'nodebuffer' }),
            contentType: XLSX_TYPE,
        })
        expect(r).toMatchObject({ status: 'extracted', extraction: '=== Scoped ===\n1', truncated: false })
    })

    it('concatenates visible inline rich-text runs without phonetic guides', async () => {
        const workbook = new ExcelJS.Workbook()
        workbook.addWorksheet('Rich').getCell('A1').value = 'placeholder'
        const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer())
        const name = 'xl/worksheets/sheet1.xml'
        const xml = await zip.file(name)!.async('string')
        zip.file(
            name,
            xml.replace(
                /<c r="A1"[^>]*>.*?<\/c>/,
                '<c r="A1" t="inlineStr"><is><r><t>Alpha &amp; </t></r><r><t>Beta</t></r>' +
                    '<rPh sb="0" eb="1"><t>PHONETIC</t></rPh></is></c>'
            )
        )

        const r = await extractAttachment({
            content: await zip.generateAsync({ type: 'nodebuffer' }),
            contentType: XLSX_TYPE,
        })
        expect(r).toMatchObject({ status: 'extracted', extraction: '=== Rich ===\nAlpha & Beta', truncated: false })
    })

    it('ignores worksheet-shaped cells in foreign extension namespaces', async () => {
        const mutate = async (value: ExcelJS.CellValue, payload: string) => {
            const workbook = new ExcelJS.Workbook()
            workbook.addWorksheet('Scoped').getCell('A1').value = value
            const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer())
            const name = 'xl/worksheets/sheet1.xml'
            const xml = await zip.file(name)!.async('string')
            zip.file(
                name,
                xml.replace(
                    '</sheetData>',
                    `</sheetData><extLst><ext uri="probe"><foreign xmlns="urn:foreign">${payload}</foreign></ext></extLst>`
                )
            )
            return extractAttachment({ content: await zip.generateAsync({ type: 'nodebuffer' }), contentType: XLSX_TYPE })
        }

        const inline = await mutate(
            'REAL',
            '<row><c r="A1" t="inlineStr"><is><t>FORGED</t></is></c></row>'
        )
        const boolean = await mutate({ formula: '1+0', result: 1 }, '<row><c r="A1" t="b"/></row>')

        expect(inline.extraction).toBe('=== Scoped ===\nREAL')
        expect(boolean.extraction).toBe('=== Scoped ===\n1')
    })

    it('tolerates an undeclared extension prefix without losing the worksheet', async () => {
        const workbook = new ExcelJS.Workbook()
        workbook.addWorksheet('Extensions').getCell('A1').value = 'REAL'
        const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer())
        const name = 'xl/worksheets/sheet1.xml'
        const xml = await zip.file(name)!.async('string')
        zip.file(name, xml.replace('</worksheet>', '<extLst><ext uri="probe"><xr:extra/></ext></extLst></worksheet>'))

        const r = await extractAttachment({
            content: await zip.generateAsync({ type: 'nodebuffer' }),
            contentType: XLSX_TYPE,
        })
        expect(r).toMatchObject({ status: 'extracted', extraction: '=== Extensions ===\nREAL', truncated: false })
    })

    it('ignores ExcelJS null row events outside sheetData', async () => {
        const workbook = new ExcelJS.Workbook()
        const sheet = workbook.addWorksheet('Rows')
        sheet.addRow(['first'])
        sheet.addRow(['second'])
        const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer())
        const name = 'xl/worksheets/sheet1.xml'
        const xml = await zip.file(name)!.async('string')
        zip.file(name, xml.replace('</sheetData>', '</sheetData><extLst><ext uri="probe"><row r="99"/></ext></extLst>'))

        const r = await extractAttachment({
            content: await zip.generateAsync({ type: 'nodebuffer' }),
            contentType: XLSX_TYPE,
        })
        expect(r).toMatchObject({
            status: 'extracted',
            extraction: '=== Rows ===\nfirst\nsecond',
            truncated: false,
        })
    })

    it.each([
        ['worksheet', 'xl/worksheets/sheet1.xml'],
        ['shared strings', 'xl/sharedStrings.xml'],
        ['styles', 'xl/styles.xml'],
    ])('enforces the XML depth ceiling in streamed %s', async (_label, name) => {
        const workbook = new ExcelJS.Workbook()
        const sheet = workbook.addWorksheet('Deep')
        sheet.getCell('A1').value = 'text'
        sheet.getCell('B1').value = 1
        sheet.getCell('B1').numFmt = '0.00'
        const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer())
        const xml = await zip.file(name)!.async('string')
        const close = xml.lastIndexOf('</')
        const nested = '<extLst>'.repeat(300) + '</extLst>'.repeat(300)
        zip.file(name, xml.slice(0, close) + nested + xml.slice(close))

        const r = await extractAttachment({
            content: await zip.generateAsync({ type: 'nodebuffer' }),
            contentType: XLSX_TYPE,
        })
        expect(r.status).toBe('failed')
        expect(r.reason).toBe('malformed')
    })

    it('rejects an incomplete shared-string table instead of silently dropping later cells', async () => {
        const workbook = new ExcelJS.Workbook()
        const sheet = workbook.addWorksheet('Data')
        sheet.addRow(['first'])
        sheet.addRow(['second'])
        const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer())
        const name = 'xl/sharedStrings.xml'
        const xml = await zip.file(name)!.async('string')
        const firstEnd = xml.indexOf('</si>') + '</si>'.length
        zip.file(name, `${xml.slice(0, firstEnd)}<si><t>second`)

        const r = await extractAttachment({
            content: await zip.generateAsync({ type: 'nodebuffer' }),
            contentType: XLSX_TYPE,
        })
        expect(r.status).toBe('failed')
        expect(r.reason).toBe('malformed')
    })

    it('rejects an incomplete styles table at natural EOF', async () => {
        const workbook = new ExcelJS.Workbook()
        const cell = workbook.addWorksheet('Styled').getCell('A1')
        cell.value = 1
        cell.numFmt = '0.00'
        const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer())
        const name = 'xl/styles.xml'
        const xml = await zip.file(name)!.async('string')
        zip.file(name, xml.slice(0, xml.lastIndexOf('</styleSheet>')))

        const r = await extractAttachment({
            content: await zip.generateAsync({ type: 'nodebuffer' }),
            contentType: XLSX_TYPE,
        })
        expect(r.status).toBe('failed')
        expect(r.reason).toBe('malformed')
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

    it('routes and extracts an xlsx whose workbook root differs only in ASCII case', async () => {
        const zip = await JSZip.loadAsync(fixture('sample.xlsx'))
        const workbook = await zip.file('xl/workbook.xml')!.async('nodebuffer')
        zip.remove('xl/workbook.xml')
        zip.file('xl/Workbook.xml', workbook)
        const content = await zip.generateAsync({ type: 'nodebuffer' })

        expect(detectRoute({ content, contentType: XLSX_TYPE }).kind).toBe('xlsx')
        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toContain('=== Q1 ===')
        expect(r.extraction).toContain('West\t4200')
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
        expect(r.reason).toBe('malformed')
    })
})

// Output cap -----------------------------------------------------------------
// Input is byte-capped, but output isn't proportional to input — cap it centrally so a
// pathological/large document can't dump megabytes of text into S3 + the search index.

// An empty extraction used to be one result object for three unrelated situations, so a caller had
// no way to tell "this page is a scan, send it to OCR" from "this file is empty". These pin the
// distinction, and pin that it is drawn from PROOF rather than guessed: only a format that can show
// it holds non-text content may say 'no-text-layer'.
// `reason` is a CODE, not prose. The strings it replaced interpolated values the caller already
// had — its own byte count, its own content type, constants this package exports — so they read as
// diagnostics while carrying nothing to branch on without a regex. These pin the contract itself:
// which values exist, and that the field appears on exactly the results that were refused.
describe('attachment — reason is a code', () => {
    const REASONS = new Set([
        'too-large',
        'expands-too-large',
        'unsupported-format',
        'unrecognized',
        'password-protected',
        'unsupported-zip-feature',
        'malformed',
        'wrong-document-shape',
        'timed-out',
        'internal',
    ])

    it('never sets a reason on a successful extraction', async () => {
        for (const [content, filename] of [
            [fixture('sample.pdf'), 'a.pdf'],
            [fixture('sample.xlsx'), 'a.xlsx'],
            [Buffer.from('hello'), 'a.txt'],
            [Buffer.alloc(0), 'empty.txt'], // extracted-but-empty still carries no REFUSAL reason
        ] as [Buffer, string][]) {
            const r = await extractAttachment({ content, filename })
            expect(r.status).toBe('extracted')
            expect(r.reason).toBeUndefined()
        }
    })

    // Every refusal must be branchable. A free-text reason could not promise this, which is why the
    // suite used to match it with two dozen regexes.
    it('sets a known code on every refusal', async () => {
        const refusals = [
            { content: Buffer.alloc(MAX_INPUT_BYTES + 1, 0x41), filename: 'big.txt' },
            { content: Buffer.from([0x89, 0x50, 0x4e, 0x47]), contentType: 'image/png', filename: 'a.png' },
            { content: Buffer.from([0x00, 0x01, 0x02, 0x03]), filename: 'mystery.bin' },
            { content: Buffer.from('PK\x03\x04 not really a zip'), filename: 'broken.docx', contentType: DOCX_TYPE },
        ]
        for (const input of refusals) {
            const r = await extractAttachment(input)
            expect(r.status).not.toBe('extracted')
            expect(REASONS.has(r.reason as string)).toBe(true)
        }
    })

    // NOTE there is no case here for 'internal' itself. It had one — asserting
    // `REASONS.has('internal')`, a property of the literal set declared just above it, which can
    // never fail. Reaching that code needs an invariant to actually trip, which 'fails fast when the
    // pinned ExcelJS streaming hooks are unavailable' already does; a second copy here would be a
    // worse version of it. Deleted rather than replaced.

    // A reason is either a decline or an inability, never both. Pairing them by hand at each return
    // site is what let the 0xffff sentinel come back `skipped` from the budget and `failed` from the
    // rewrite for the same fact.
    it('derives a single status from each reason', async () => {
        const seen = new Map<string, Set<string>>()
        const inputs = [
            { content: Buffer.alloc(MAX_INPUT_BYTES + 1, 0x41), filename: 'big.txt' },
            { content: Buffer.from([0x89, 0x50, 0x4e, 0x47]), contentType: 'image/png', filename: 'a.png' },
            { content: Buffer.from([0x00, 0x01, 0x02]), filename: 'mystery.bin' },
            { content: Buffer.from('PK\x03\x04 not a zip'), filename: 'broken.docx', contentType: DOCX_TYPE },
        ]
        for (const input of inputs) {
            const r = await extractAttachment(input)
            if (r.reason === undefined) continue
            const statuses = seen.get(r.reason) ?? new Set<string>()
            statuses.add(r.status)
            seen.set(r.reason, statuses)
        }
        for (const [reason, statuses] of seen) expect([reason, statuses.size]).toEqual([reason, 1])
    })
})

describe('attachment — why an extraction is empty', () => {
    it('reports no-text-layer for a PDF whose pages carry no text', async () => {
        const r = await extractAttachment({ content: fixture('blank.pdf'), filename: 'scan.pdf' })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBeUndefined()
        expect(r.emptyReason).toBe('no-text-layer')
    })

    it('reports no-text-content for a zero-byte attachment', async () => {
        const r = await extractAttachment({ content: Buffer.alloc(0), filename: 'empty.txt' })
        expect(r.status).toBe('extracted')
        expect(r.emptyReason).toBe('no-text-content')
    })

    // The case that makes this more than a PDF feature: whitespace-only text reached the same result
    // object as a scan, and the two want opposite follow-up actions.
    it('reports no-text-content for a whitespace-only text attachment', async () => {
        const r = await extractAttachment({ content: Buffer.from('   \n\t  '), filename: 'blank.txt' })
        expect(r.status).toBe('extracted')
        expect(r.emptyReason).toBe('no-text-content')
    })

    // REGRESSION, and the sharpest of them: both values are claims about the WHOLE document, so a
    // read that stopped early cannot support either. A PDF full of text with a zero cap reported
    // 'no-text-layer' — telling a caller to OCR a document that has a text layer — and a .docx the
    // same way reported 'no-text-content', telling them to discard it. `truncated` alone is the
    // honest answer there: we stopped before finding text and cannot say whether there is any.
    it('reports no empty reason for a read that stopped early', async () => {
        for (const [content, filename] of [
            [fixture('sample.pdf'), 'a.pdf'],
            [fixture('sample.docx'), 'a.docx'],
        ] as [Buffer, string][]) {
            const r = await extractAttachment({ content, filename }, { maxOutputChars: 0 })
            expect(r.status).toBe('extracted')
            expect(r.truncated).toBe(true)
            expect(r.extraction).toBeUndefined()
            expect(r.emptyReason).toBeUndefined()
        }
    })

    // Exactly one of the two is present on every successful result, so a consumer never has to test
    // for both — and never sees a reason contradicting text it was also handed.
    it('omits emptyReason whenever there is text', async () => {
        const r = await extractAttachment({ content: fixture('sample.pdf'), filename: 'a.pdf' })
        expect(r.extraction).toBeTruthy()
        expect(r.emptyReason).toBeUndefined()
    })

    // A password-protected OOXML file is an OLE container, so routing saw OLE where it wanted PK and
    // called the file unrecognized — true of the bytes, and wrong about the file.
    it('names a password-protected Office file rather than calling it unrecognized', async () => {
        const ole = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
        const marker = Buffer.from('EncryptedPackage', 'utf16le')
        const content = Buffer.concat([ole, Buffer.alloc(400), marker, Buffer.alloc(2_048)])
        const r = await extractAttachment({ content, filename: 'locked.docx' })
        expect(r.status).toBe('skipped')
        expect(r.reason).toBe('password-protected')
    })

    // The marker is what earns the label; an OLE file without it is still a legacy binary we may or
    // may not read, and must not be relabeled as locked.
    it('does not call an ordinary OLE file password-protected', async () => {
        const r = await extractAttachment({ content: fixture('sample.doc'), filename: 'sample.doc' })
        expect(r.status).toBe('extracted')
    })

    // REGRESSION: the check sat inside the no-handler branch, so an encrypted package named .doc —
    // an OLE container, which carries exactly the magic the legacy handler routes on — reached
    // word-extractor, threw, and came back 'malformed', which the codes define as never retryable.
    // Whether we happen to have a handler says nothing about whether the content is ours to read.
    it('names a locked file even when its extension routes to a handler', async () => {
        const ole = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
        const marker = Buffer.from('EncryptedPackage', 'utf16le')
        const content = Buffer.concat([ole, Buffer.alloc(400), marker, Buffer.alloc(2_048)])
        const r = await extractAttachment({ content, filename: 'locked.doc', contentType: 'application/msword' })
        expect(r.status).toBe('skipped')
        expect(r.reason).toBe('password-protected')
    })

    // pdf.js refuses an encrypted document by throwing, so there is nothing to sniff — the error is
    // the signal. Same situation, same answer: skipped and locked, not failed and broken.
    it('names a password-protected PDF rather than calling it malformed', async () => {
        vi.resetModules()
        vi.doMock('unpdf', () => ({
            getDocumentProxy: async () => {
                const error = new Error('No password given')
                error.name = 'PasswordException'
                throw error
            },
        }))
        try {
            const { extractAttachment: extract } = await import('../attachment')
            const r = await extract({ content: buf('%PDF-1.4 locked'), contentType: 'application/pdf' })
            expect(r.status).toBe('skipped')
            expect(r.reason).toBe('password-protected')
        } finally {
            vi.doUnmock('unpdf')
            vi.resetModules()
        }
    })
})

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

    // That early gate is the one 'extracted' return that never reaches the cap logic, so it is the
    // one that can forget `truncated`. Asserted as presence, not just value: `undefined === false`
    // is false either way, and only `'truncated' in r` catches an omitted key.
    it('reports truncated false on a zero-byte attachment rather than omitting it', async () => {
        for (const contentType of ['text/plain', 'application/pdf', XLSX_TYPE, undefined]) {
            const r = await extractAttachment({ content: Buffer.alloc(0), contentType })
            expect(r.status).toBe('extracted')
            expect('truncated' in r).toBe(true)
            expect(r.truncated).toBe(false)
        }
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
            expect(r.reason).toBe('timed-out')
            expect(r.extraction).toBeUndefined()
        } finally {
            vi.useRealTimers()
            vi.doUnmock('unpdf')
            vi.resetModules()
        }
    })

    // The other side of that race, and the reason HANDLER_DEADLINE_MARGIN_MS exists: a parser slow
    // enough to pass the handler's deadline but not the timeout must come back as a partial
    // extraction, NOT as 'failed'. With the two instants equal it never can — handlers stop at
    // `Date.now() > deadline`, strictly after the moment the timer already fired at — so ten seconds
    // of successfully read pages are discarded.
    //
    // Fake timers are load-bearing here, not a speed trick. The per-handler deadline tests elsewhere
    // in this file stub Date.now() and leave setTimeout real, which decouples the two clocks and is
    // exactly why none of them can see this; one synthetic clock driving both is the real
    // relationship. It also keeps this off the wall clock — a ~9s real-time test is a CI flake
    // waiting to happen.
    it('reports a slow yielding handler as truncated, not failed, inside the timeout', async () => {
        const PAGE_MS = 250
        const PAGES = 60 // 15s of work — far past the deadline, so the stop is unambiguously its
        vi.resetModules()
        vi.doMock('unpdf', () => ({
            getDocumentProxy: async () => ({
                numPages: PAGES,
                // Yields between pages, like the real one: the deadline check sits at the top of the
                // loop, so the handler can act on it.
                getPage: (n: number) =>
                    new Promise((resolve) =>
                        setTimeout(
                            () =>
                                resolve({
                                    getTextContent: async () => ({ items: [{ str: `page ${n}`, hasEOL: true }] }),
                                }),
                            PAGE_MS
                        )
                    ),
            }),
        }))
        try {
            const { extractAttachment: extract } = await import('../attachment')
            vi.useFakeTimers()
            const pending = extract({ content: buf('%PDF-1.4 slow'), contentType: 'application/pdf' })
            await vi.advanceTimersByTimeAsync(HANDLER_TIMEOUT_MS + PAGE_MS)
            const r = await pending

            expect(r.status).toBe('extracted') // 'failed' here is the regression
            expect(r.truncated).toBe(true)
            expect(r.extraction).toContain('page 1') // pages read before the stop survive
            expect(r.extraction).not.toContain(`page ${PAGES}`) // and it really did stop short
        } finally {
            vi.useRealTimers()
            vi.doUnmock('unpdf')
            vi.resetModules()
        }
    })

    // Documented limitation — withTimeout is not exported, so replicate it
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

    // Formerly pinned the opposite, with a note that adding `truncated` must force a conscious
    // contract update. This is that update.
    it('reports over-limit output as truncated rather than capping it silently', async () => {
        const content = Buffer.alloc(2 * MAX_OUTPUT_CHARS, 0x41) // 500k "A", under the 10MB input cap
        const r = await extractAttachment({ content, contentType: 'text/plain' })
        expect(r.status).toBe('extracted')
        expect(r.extraction?.length).toBe(MAX_OUTPUT_CHARS) // still capped
        expect(r.truncated).toBe(true) // ...but no longer silently
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

    it('labels an unexpected decompression-preflight rejection instead of throwing', async () => {
        const createInflate = vi.spyOn(zlib, 'createInflateRaw').mockImplementationOnce(() => {
            throw new Error('inflate setup failed')
        })
        try {
            const r = await extractAttachment({ content: fixture('sample.xlsx'), contentType: XLSX_TYPE })
            // Labeled rather than thrown, which is the whole assertion. It lands on 'malformed'
            // because an unrecognized throw is attributed to the bytes — see failureReason, and the
            // note there that a genuine bug of ours reaching here is reported the same way.
            expect(r).toEqual({ status: 'failed', reason: 'malformed' })
        } finally {
            createInflate.mockRestore()
        }
    })

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
        expect(r.reason).toBe('expands-too-large')
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
        expect(r.reason).toBe('expands-too-large')
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
            expect(r.reason).toBe('unsupported-zip-feature')
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
            expect(r.reason).toBe('malformed')
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
        expect(r.reason).toBe('expands-too-large')
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
        expect(r.reason).toBe('malformed')
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
        expect(r.reason).toBe('malformed')
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
        expect(r.reason).toBe('malformed')
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
        expect(r.reason).toBe('malformed')
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
        expect(r.reason).toBe('malformed')
    })

    it('fails closed on a compression method it cannot measure', async () => {
        const zip = craftOoxmlZip(buf('<workbook/>'), { method: 12 }) // 12 = bzip2; we only measure store/deflate
        const r = await extractAttachment({ content: zip, contentType: XLSX_TYPE })
        expect(r.status).toBe('skipped')
        expect(r.reason).toBe('unsupported-zip-feature')
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
        expect(r.reason).toBe('expands-too-large')
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

    // Every pair of zero bytes is a decoded U+0000. It satisfies surrogate/noncharacter checks and
    // used to turn an ordinary NUL-filled binary into thousands of invisible "text" characters.
    it('rejects NUL-filled binary under a utf-16 charset claim', async () => {
        const content = Buffer.alloc(9_000)
        const r = await extractAttachment({
            content,
            filename: 'payload.bin',
            contentType: 'text/plain; charset=utf-16',
        })
        expect(r.status).toBe('skipped')
        expect(r.extraction).toBeUndefined()
    })

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

// Extract options: maxOutputChars + trailer ----------------------------------
// The cap may only tighten, the trailer only appears on a real cut, and omitting both must
// reproduce pre-options behaviour exactly.

describe('attachment — extract options', () => {
    // 500k "A" against the 250k ceiling: over-cap under every setting below, so each case isolates
    // the option under test rather than the input size.
    const oversized = () => Buffer.alloc(2 * MAX_OUTPUT_CHARS, 0x41)
    const asText = { contentType: 'text/plain' }

    it('tightens the cap when asked', async () => {
        const r = await extractAttachment({ content: oversized(), ...asText }, { maxOutputChars: 100 })
        expect(r.extraction).toHaveLength(100)
        expect(r.truncated).toBe(true)
    })

    it('clamps a cap above the ceiling instead of honouring it', async () => {
        // The footgun this stops: a streaming handler reads until the cap, so a caller must not be
        // able to widen how far we read into a document.
        const r = await extractAttachment({ content: oversized(), ...asText }, { maxOutputChars: 10 * MAX_OUTPUT_CHARS })
        expect(r.extraction).toHaveLength(MAX_OUTPUT_CHARS)
    })

    it('ignores an unusable or negative cap rather than producing an empty extraction', async () => {
        for (const bad of [Number.NaN, Number.POSITIVE_INFINITY, undefined]) {
            const r = await extractAttachment({ content: oversized(), ...asText }, { maxOutputChars: bad })
            expect(r.extraction).toHaveLength(MAX_OUTPUT_CHARS)
        }
        // Negative clamps to 0 — a cap of "no text at all" is coherent, so it is honoured, not
        // ignored. -Infinity is the same request, so it must land the same place: it used to take
        // the non-finite fallback and come back as the full 250k, the one non-monotonic seam here.
        for (const none of [-5, Number.NEGATIVE_INFINITY]) {
            const zero = await extractAttachment({ content: oversized(), ...asText }, { maxOutputChars: none })
            expect(zero.extraction).toBeUndefined() // '' collapses to omitted, per the contract
            expect(zero.status).toBe('extracted')
        }
    })

    it.each(['abc', {}, null, true])('falls back to the ceiling for a non-number JS cap: %j', async (bad) => {
        const r = await extractAttachment(
            { content: oversized(), ...asText },
            { maxOutputChars: bad as never }
        )
        expect(r.extraction).toHaveLength(MAX_OUTPUT_CHARS)
        expect(r.truncated).toBe(true)
    })

    it('floors a fractional cap', async () => {
        const r = await extractAttachment({ content: oversized(), ...asText }, { maxOutputChars: 10.9 })
        expect(r.extraction).toHaveLength(10)
    })

    it('appends the trailer only when the text was actually cut', async () => {
        const trailer = '\n[truncated]'
        const cut = await extractAttachment({ content: oversized(), ...asText }, { maxOutputChars: 100, trailer })
        expect(cut.extraction?.endsWith(trailer)).toBe(true)

        const whole = await extractAttachment({ content: buf('short'), ...asText }, { trailer })
        expect(whole.extraction).toBe('short')
        expect(whole.truncated).toBe(false)
    })

    it('keeps the trailer outside cap accounting — the cap bounds text, not the returned string', async () => {
        // Documented deliberately: a consumer sizing a buffer on maxOutputChars must add the trailer.
        const trailer = '\n[truncated]'
        const r = await extractAttachment({ content: oversized(), ...asText }, { maxOutputChars: 100, trailer })
        expect(r.extraction).toHaveLength(100 + trailer.length)
        expect(r.extraction?.slice(0, 100)).toBe('A'.repeat(100))
    })

    it('reports truncated independently of the trailer, so a consumer need not parse the text', async () => {
        const r = await extractAttachment({ content: oversized(), ...asText }, { maxOutputChars: 100 })
        expect(r.truncated).toBe(true)
        expect(r.extraction?.endsWith('A')).toBe(true) // no trailer supplied, none appended
    })

    it('omits truncated on skipped and failed — there is no text to have cut', async () => {
        const skipped = await extractAttachment({ content: Buffer.alloc(MAX_INPUT_BYTES + 1), ...asText })
        expect(skipped.status).toBe('skipped')
        expect('truncated' in skipped).toBe(false)
    })

    it('reproduces pre-options behaviour when both options are omitted', async () => {
        // AC#6, stated as a test: the default path must be byte-identical to passing nothing.
        const content = oversized()
        const bare = await extractAttachment({ content, ...asText })
        const empty = await extractAttachment({ content, ...asText }, {})
        expect(bare).toEqual(empty)
        expect(bare.extraction).toHaveLength(MAX_OUTPUT_CHARS)
    })

    it('does not split a surrogate pair at a tightened cap boundary', async () => {
        // The existing surrogate guard has to follow the resolved cap, not the module constant.
        const content = buf('a' + '😀'.repeat(50)) // 😀 is a surrogate PAIR: cap 2 lands mid-pair
        const r = await extractAttachment({ content, ...asText }, { maxOutputChars: 2 })
        expect(r.extraction).toBe('a') // the lone high half was dropped, not emitted as U+FFFD
        expect(r.extraction).not.toContain('�')
    })
})

// xlsx streaming: entry-order workaround --------------------------------------
// Regressions for reorderForStreaming; the entry-loss bug it works around is documented at its
// definition. Each asserts over repeated reads because that bug is a race — a single green read
// proves nothing at a partial failure rate. Before the reorder the multi-sheet case dropped a sheet
// or threw on most reads; the no-shared-strings case dropped sheets on 35 of 50.

// Shared by the determinism block and the sheet-identity block below, which assert on the same
// naming convention (S1.., s1r1..) and must not drift apart. exceljs's own writer emits
// xl/workbook.xml LAST, which is the layout that used to throw.
const workbookWith = async (sheets: number, rows: number) => {
    const workbook = new ExcelJS.Workbook()
    for (let s = 1; s <= sheets; s++) {
        const sheet = workbook.addWorksheet(`S${s}`)
        for (let r = 1; r <= rows; r++) sheet.addRow([`s${s}r${r}`, r])
    }
    return Buffer.from(await workbook.xlsx.writeBuffer())
}

describe('attachment — xlsx streaming determinism', () => {
    const READS = 12 // enough to catch a partial-rate race; the workbooks are tiny

    it('returns every worksheet, on every read', async () => {
        const content = await workbookWith(5, 10)
        for (let i = 0; i < READS; i++) {
            const r = await extractAttachment({ content, contentType: XLSX_TYPE })
            expect(r.status).toBe('extracted')
            for (let s = 1; s <= 5; s++) expect(r.extraction).toContain(`=== S${s} ===`)
        }
    })

    it('returns byte-identical extraction across repeated reads of the same workbook', async () => {
        const content = await workbookWith(4, 25)
        const first = await extractAttachment({ content, contentType: XLSX_TYPE })
        for (let i = 0; i < READS; i++) {
            expect(await extractAttachment({ content, contentType: XLSX_TYPE })).toEqual(first)
        }
    })

    it('decodes shared strings across inflate chunk boundaries without replacement characters', async () => {
        const workbook = new ExcelJS.Workbook()
        const sheet = workbook.addWorksheet('Unicode')
        const values = Array.from({ length: 2_000 }, (_, i) => `row ${i}: café — 東京 😀`)
        for (const value of values) sheet.addRow([value])

        const r = await extractAttachment({
            content: Buffer.from(await workbook.xlsx.writeBuffer()),
            contentType: XLSX_TYPE,
        })
        expect(r).toMatchObject({ status: 'extracted', truncated: false })
        expect(r.extraction).not.toContain('\uFFFD')
        expect(r.extraction).toContain(values[0])
        expect(r.extraction).toContain(values[999])
        expect(r.extraction).toContain(values[1_999])
    })

    it.each([0.9, 0.7, 0.5])('fails a worksheet whose XML stops at %s of its original length', async (ratio) => {
        const zip = await JSZip.loadAsync(await workbookWith(1, 40))
        const name = 'xl/worksheets/sheet1.xml'
        const xml = await zip.file(name)!.async('nodebuffer')
        zip.file(name, xml.subarray(0, Math.floor(xml.length * ratio)))

        const r = await extractAttachment({
            content: await zip.generateAsync({ type: 'nodebuffer' }),
            contentType: XLSX_TYPE,
        })
        expect(r.status).toBe('failed')
        expect(r.extraction).toBeUndefined()
    })

    it('fails when a truncated later worksheet yields no rows but still counts as emitted', async () => {
        const zip = await JSZip.loadAsync(await workbookWith(2, 10))
        const name = 'xl/worksheets/sheet2.xml'
        const xml = await zip.file(name)!.async('nodebuffer')
        zip.file(name, xml.subarray(0, Math.floor(xml.length * 0.2)))

        const r = await extractAttachment({
            content: await zip.generateAsync({ type: 'nodebuffer' }),
            contentType: XLSX_TYPE,
        })
        expect(r.status).toBe('failed')
        expect(r.extraction).toBeUndefined()
    })

    // The ORDINARY case, where tab order and file order agree — so this pins that the reorder does
    // not scramble a normal workbook, and nothing more. It cannot see where the two authorities
    // disagree, because exceljs's writer never emits a workbook in which they do: that is what the
    // sheet-identity block below builds by hand.
    it('keeps worksheets in workbook order', async () => {
        const r = await extractAttachment({ content: await workbookWith(4, 3), contentType: XLSX_TYPE })
        const headers = r.extraction!.match(/^=== .* ===$/gm)
        expect(headers).toEqual(['=== S1 ===', '=== S2 ===', '=== S3 ===', '=== S4 ==='])
    })

    // No strings means no xl/sharedStrings.xml, so ordering alone cannot set the flag that keeps the
    // reader off the lossy path — the handler injects an empty table. Numeric-only cells are the
    // realistic shape of such a workbook.
    it('returns every worksheet when the workbook has no shared-string table', async () => {
        const workbook = new ExcelJS.Workbook()
        for (let s = 1; s <= 3; s++) {
            const sheet = workbook.addWorksheet(`N${s}`)
            for (let r = 1; r <= 10; r++) sheet.addRow([s * 1000 + r, r * 2, r * 3])
        }
        const content = Buffer.from(await workbook.xlsx.writeBuffer())
        for (let i = 0; i < READS; i++) {
            const r = await extractAttachment({ content, contentType: XLSX_TYPE })
            expect(r.status).toBe('extracted')
            for (let s = 1; s <= 3; s++) expect(r.extraction).toContain(`=== N${s} ===`)
        }
    })

    // OPC part names compare ASCII-case-insensitively, while exceljs dispatches this control part by
    // exact spelling. The private rewrite canonicalizes it so shared-string indices resolve to text.
    it('preserves text from a case-variant shared-string part', async () => {
        const workbook = new ExcelJS.Workbook()
        const sheet = workbook.addWorksheet('Data')
        sheet.addRow(['Alpha Beta'])
        sheet.addRow(['Gamma', 42])
        const zip = await JSZip.loadAsync(Buffer.from(await workbook.xlsx.writeBuffer()))
        const strings = await zip.file('xl/sharedStrings.xml')!.async('nodebuffer')
        zip.remove('xl/sharedStrings.xml')
        zip.file('xl/SharedStrings.xml', strings)

        const r = await extractAttachment({
            content: await zip.generateAsync({ type: 'nodebuffer' }),
            contentType: XLSX_TYPE,
        })
        expect(r.status).toBe('extracted')
        expect(r.truncated).toBe(false)
        expect(r.extraction).toBe('=== Data ===\nAlpha Beta\nGamma\t42')
    })

    // The two shapes where streaming deliberately does NOT reproduce workbook.xlsx.load(). Pinned so
    // they stay deliberate: both are improvements for a search index, but both change what a merged
    // workbook's rows and columns look like against main, and an unpinned improvement is
    // indistinguishable from an accident.
    it('emits a merged cell once instead of proxying it into every slave cell', async () => {
        const workbook = new ExcelJS.Workbook()
        const sheet = workbook.addWorksheet('Merged')
        sheet.getCell('A1').value = 'TITLE'
        sheet.mergeCells('A1:C1') // horizontal
        sheet.getCell('A2').value = 'a'
        sheet.getCell('B2').value = 'b'
        sheet.getCell('C2').value = 'c'
        sheet.getCell('A3').value = 'SIDE'
        sheet.mergeCells('A3:A5') // vertical
        sheet.getCell('B3').value = 'x'

        const r = await extractAttachment({
            content: Buffer.from(await workbook.xlsx.writeBuffer()),
            contentType: XLSX_TYPE,
        })
        // load() gave "TITLE\tTITLE\tTITLE" and two trailing "SIDE" rows carrying nothing else.
        expect(r.extraction).toBe('=== Merged ===\nTITLE\na\tb\tc\nSIDE\tx')
    })

    it('emits an error-valued formula as empty rather than stringifying the value object', async () => {
        const workbook = new ExcelJS.Workbook()
        const sheet = workbook.addWorksheet('Err')
        sheet.getCell('A1').value = { formula: '1/0', result: { error: '#DIV/0!' } } as never
        sheet.getCell('B1').value = 'ok'

        const r = await extractAttachment({
            content: Buffer.from(await workbook.xlsx.writeBuffer()),
            contentType: XLSX_TYPE,
        })
        expect(r.extraction).toBe('=== Err ===\n\tok') // load() put "[object Object]" in that cell
        expect(r.extraction).not.toContain('[object Object]')
    })

    // The reorder rewrites the archive before the parser sees it, so it must not disturb the values.
    it('preserves cell values, formula results and sheet names through the reorder', async () => {
        const r = await extractAttachment({ content: fixture('sample.xlsx'), contentType: XLSX_TYPE })
        expect(r.extraction).toContain('=== Q1 ===')
        expect(r.extraction).toContain('=== Notes ===')
        expect(r.extraction).toContain('West\t4200')
        expect(r.extraction).toContain('Total\t7300') // formula RESULT survived
        expect(r.extraction).not.toContain('SUM(')
    })
})

// Sheet identity --------------------------------------------------------------
// Order, names and membership come from xl/workbook.xml + its rels, not from the archive. Every
// fixture here needs the two to DISAGREE, and exceljs's writer never produces that — it emits
// <sheets> in file order with relative rel targets — so each is built by editing the parts of a
// real workbook afterwards, which is how all three shapes reach us in the wild anyway.

describe('attachment — xlsx sheet identity comes from the workbook', () => {
    const rebuild = async (content: Buffer, edit: (zip: JSZip) => Promise<void>): Promise<Buffer> => {
        const zip = await JSZip.loadAsync(content)
        await edit(zip)
        return zip.generateAsync({ type: 'nodebuffer' })
    }

    const part = (zip: JSZip, name: string) => zip.file(name)!.async('string')
    const headers = (extraction?: string) => extraction?.match(/^=== .* ===$/gm)

    // Dragging a tab in Excel reorders <sheets> and leaves the sheetN.xml parts exactly where they
    // were. load() followed <sheets>; reading entry order does not.
    it('emits sheets in tab order when it disagrees with archive order', async () => {
        const dragged = await rebuild(await workbookWith(3, 2), async (zip) => {
            const xml = await part(zip, 'xl/workbook.xml')
            zip.file(
                'xl/workbook.xml',
                xml.replace(/<sheets>(.*?)<\/sheets>/, (_, inner: string) => {
                    const elements = inner.match(/<sheet\b[^>]*\/>/g) ?? []
                    return `<sheets>${elements.reverse().join('')}</sheets>`
                })
            )
        })

        const r = await extractAttachment({ content: dragged, contentType: XLSX_TYPE })
        expect(headers(r.extraction)).toEqual(['=== S3 ===', '=== S2 ===', '=== S1 ==='])
        // The ROWS have to move with the header — reordering names alone would mislabel every sheet.
        expect(r.extraction).toMatch(/=== S3 ===\ns3r1/)
        expect(r.extraction).toMatch(/=== S1 ===\ns1r1/)
    })

    // A legal absolute Target. exceljs compares rel.Target to the single string
    // `worksheets/sheetN.xml`, so this spelling matches nothing and every sheet lost its name.
    it('resolves an absolute rel Target rather than falling back to positional names', async () => {
        const absolute = await rebuild(await workbookWith(3, 2), async (zip) => {
            const xml = await part(zip, 'xl/_rels/workbook.xml.rels')
            const absolutized = xml.replace(/Target="(worksheets\/sheet\d+\.xml)"/g, 'Target="/xl/$1"')
            zip.file('xl/_rels/workbook.xml.rels', absolutized)
        })

        const r = await extractAttachment({ content: absolute, contentType: XLSX_TYPE })
        expect(headers(r.extraction)).toEqual(['=== S1 ===', '=== S2 ===', '=== S3 ==='])
        expect(r.extraction).not.toMatch(/=== Sheet\d+ ===/) // the old fallback naming
    })

    // A worksheet part no <sheet> references — a stale part left by an editing tool. The archive
    // holds it, the workbook does not, and load() ignored it; entry-order dispatch emitted it as a
    // sheet of its own with content duplicated from wherever it was copied.
    it('drops a worksheet part the workbook does not reference', async () => {
        const withOrphan = await rebuild(await workbookWith(2, 2), async (zip) => {
            zip.file('xl/worksheets/sheet7.xml', await part(zip, 'xl/worksheets/sheet1.xml'))
        })

        const r = await extractAttachment({ content: withOrphan, contentType: XLSX_TYPE })
        expect(r.status).toBe('extracted') // the orphan must not trip the lost-worksheet backstop
        expect(headers(r.extraction)).toEqual(['=== S1 ===', '=== S2 ==='])
        expect(r.extraction!.match(/s1r1/g)).toHaveLength(1) // emitted once, not duplicated
    })

    // The fallback path, for an archive whose workbook relationships cannot be read. The reader-only
    // rewrite injects an empty relationships part, setting exceljs's workbookRels flag and keeping
    // every worksheet inline instead of spooling it to a temp file. Archive order supplies identity.
    it('returns every worksheet deterministically when the rels part is missing', async () => {
        const noRels = await rebuild(await workbookWith(3, 2), async (zip) => {
            zip.remove('xl/_rels/workbook.xml.rels')
        })

        for (let i = 0; i < 12; i++) {
            const r = await extractAttachment({ content: noRels, contentType: XLSX_TYPE })
            expect(r.status).toBe('extracted')
            expect(headers(r.extraction)).toEqual(['=== Sheet1 ===', '=== Sheet2 ===', '=== Sheet3 ==='])
            for (let s = 1; s <= 3; s++) expect(r.extraction).toContain(`s${s}r1`)
        }
    })

    it.each([
        ['empty', ''],
        ['rootless', '<root/>'],
    ])('never enters ExcelJS temp-file spooling for a present-but-%s rels part', async (_label, relsBody) => {
        const content = await rebuild(await workbookWith(6, 1), async (zip) => {
            zip.file('xl/_rels/workbook.xml.rels', relsBody)
        })
        const tmp = require('tmp') as { file: (...args: unknown[]) => unknown }
        const originalFile = tmp.file
        let spoolCalls = 0
        tmp.file = (..._args) => {
            spoolCalls += 1
            throw new Error('ExcelJS temp-file spool reached')
        }
        try {
            const r = await extractAttachment({ content, contentType: XLSX_TYPE }, { maxOutputChars: 12 })
            expect(r.status).toBe('extracted')
            expect(r.truncated).toBe(true)
            expect(spoolCalls).toBe(0)
        } finally {
            tmp.file = originalFile
        }
    })

    // MEMBERSHIP is the archive's, not the workbook's. The workbook still declares three sheets here
    // — only the relationship placing one of them is gone — so dropping that part would DELETE its
    // text, and invisibly: the backstop compares `seen` against the resolved list, so removing a
    // sheet from both sides keeps them equal and it never fires. This is the same hazard the
    // lost-worksheet backstop exists for, reached from the other end.
    it('keeps a declared sheet whose relationship cannot be resolved', async () => {
        const brokenRel = await rebuild(await workbookWith(3, 2), async (zip) => {
            const xml = await part(zip, 'xl/_rels/workbook.xml.rels')
            zip.file('xl/_rels/workbook.xml.rels', xml.replace(/<Relationship[^>]*sheet2\.xml"[^>]*\/>/, ''))
        })

        const r = await extractAttachment({ content: brokenRel, contentType: XLSX_TYPE })
        expect(r.status).toBe('extracted')
        // Named S1/S3 where the workbook could say so, and the unplaceable one kept under the
        // fallback name — it loses its tab position, which is unknowable, but never its rows.
        expect(headers(r.extraction)).toHaveLength(3)
        for (let s = 1; s <= 3; s++) expect(r.extraction).toContain(`s${s}r1`)
    })

    // Targets are URI references, so this spelling is legal and names sheet2.xml. Undecoded it
    // resolves to nothing, which before the membership fix silently deleted the sheet.
    it('resolves a percent-encoded rel Target', async () => {
        const encoded = await rebuild(await workbookWith(3, 2), async (zip) => {
            const xml = await part(zip, 'xl/_rels/workbook.xml.rels')
            zip.file('xl/_rels/workbook.xml.rels', xml.replace('worksheets/sheet2.xml', 'worksheets/sheet%32.xml'))
        })

        const r = await extractAttachment({ content: encoded, contentType: XLSX_TYPE })
        expect(headers(r.extraction)).toEqual(['=== S1 ===', '=== S2 ===', '=== S3 ==='])
        expect(r.extraction).toMatch(/=== S2 ===\ns2r1/)
    })

    // The three shapes below are all one mistake: reading "we could not account for this" as "there
    // was nothing there". Each returned a clean `extracted` while deleting rows.

    // A <sheets> yielding no declaration at all is a workbook we did not understand, not a workbook
    // with no sheets. Honoring the empty answer drops every worksheet part — the whole document,
    // reported as success with no text whatsoever.
    it('falls back to the archive when the workbook declares no sheets at all', async () => {
        const noDeclarations = await rebuild(await workbookWith(3, 2), async (zip) => {
            const xml = await part(zip, 'xl/workbook.xml')
            zip.file('xl/workbook.xml', xml.replace(/<sheets>.*?<\/sheets>/, '<sheets/>'))
        })

        const r = await extractAttachment({ content: noDeclarations, contentType: XLSX_TYPE })
        expect(headers(r.extraction)).toHaveLength(3)
        for (let s = 1; s <= 3; s++) expect(r.extraction).toContain(`s${s}r1`)
    })

    // OPC compares part URIs ASCII-case-insensitively, so this Target legally names the entry stored
    // as xl/worksheets/sheet2.xml. Matching case-sensitively left it unresolvable: the rows survived
    // through the rescue, but under a fallback name and at the end, so the sheet lost its identity.
    it('resolves a Target that differs from the stored entry only in case', async () => {
        const content = await rebuild(await workbookWith(3, 2), async (zip) => {
            const xml = await part(zip, 'xl/_rels/workbook.xml.rels')
            zip.file('xl/_rels/workbook.xml.rels', xml.replace('worksheets/sheet2.xml', 'Worksheets/sheet2.xml'))
        })

        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        // Name and tab position both, not merely the rows — that is what case-sensitivity cost.
        expect(headers(r.extraction)).toEqual(['=== S1 ===', '=== S2 ===', '=== S3 ==='])
        expect(r.extraction).toMatch(/=== S2 ===\ns2r1/)
    })

    // OPC applies the same case-insensitive rule to the stored part URI, while exceljs's streaming
    // reader dispatches only a lowercase xl/worksheets/sheetN.xml. The in-memory rebuild bridges the
    // two by canonicalizing its private copy; the source archive is never mutated.
    it('extracts a worksheet whose stored part name differs only in case', async () => {
        const content = await rebuild(await workbookWith(3, 2), async (zip) => {
            const sheet = await zip.file('xl/worksheets/sheet2.xml')!.async('nodebuffer')
            zip.remove('xl/worksheets/sheet2.xml')
            zip.file('xl/Worksheets/Sheet2.xml', sheet)
        })

        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        expect(r.status).toBe('extracted')
        expect(headers(r.extraction)).toEqual(['=== S1 ===', '=== S2 ===', '=== S3 ==='])
        expect(r.extraction).toMatch(/=== S2 ===\ns2r1/)
    })

    // OPC relationships and content types, not filename conventions, identify a worksheet part.
    // Streaming exceljs does not dispatch a custom path, so the rebuild gives only the streamed copy
    // a canonical name and emits the measured entry once.
    it('extracts a relationship-declared worksheet stored at a custom part path', async () => {
        const content = await rebuild(await workbookWith(3, 2), async (zip) => {
            const from = 'xl/worksheets/sheet2.xml'
            const to = 'xl/custom/quarterly-data.xml'
            const sheet = await zip.file(from)!.async('nodebuffer')
            zip.remove(from)
            zip.file(to, sheet)

            const rels = await part(zip, 'xl/_rels/workbook.xml.rels')
            zip.file('xl/_rels/workbook.xml.rels', rels.replace('worksheets/sheet2.xml', 'custom/quarterly-data.xml'))

            const types = await part(zip, '[Content_Types].xml')
            zip.file('[Content_Types].xml', types.replace('/xl/worksheets/sheet2.xml', '/xl/custom/quarterly-data.xml'))
        })

        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        expect(r.status).toBe('extracted')
        expect(headers(r.extraction)).toEqual(['=== S1 ===', '=== S2 ===', '=== S3 ==='])
        expect(r.extraction).toMatch(/=== S2 ===\ns2r1/)
    })

    it('retains a custom worksheet when content-type metadata exceeds the resolver cap', async () => {
        const content = await rebuild(await workbookWith(1, 2), async (zip) => {
            const from = 'xl/worksheets/sheet1.xml'
            const to = 'xl/custom/data.xml'
            const sheet = await zip.file(from)!.async('nodebuffer')
            zip.remove(from)
            zip.file(to, sheet)

            const rels = await part(zip, 'xl/_rels/workbook.xml.rels')
            zip.file('xl/_rels/workbook.xml.rels', rels.replace('worksheets/sheet1.xml', 'custom/data.xml'))

            const types = await part(zip, '[Content_Types].xml')
            zip.file(
                '[Content_Types].xml',
                types
                    .replace('/xl/worksheets/sheet1.xml', '/xl/custom/data.xml')
                    .replace('</Types>', `<!--${'x'.repeat(4 * 1024 * 1024)}--></Types>`)
            )
        })

        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        expect(r).toMatchObject({ status: 'extracted', extraction: '=== S1 ===\ns1r1\t1\ns1r2\t2', truncated: false })
    })

    it('matches a UTF-8 flagged non-ASCII worksheet part name to its relationship target', async () => {
        const content = await rebuild(await workbookWith(1, 2), async (zip) => {
            const from = 'xl/worksheets/sheet1.xml'
            const to = 'xl/custom/café.xml'
            const sheet = await zip.file(from)!.async('nodebuffer')
            zip.remove(from)
            zip.file(to, sheet)

            const rels = await part(zip, 'xl/_rels/workbook.xml.rels')
            zip.file('xl/_rels/workbook.xml.rels', rels.replace('worksheets/sheet1.xml', 'custom/café.xml'))
            const types = await part(zip, '[Content_Types].xml')
            zip.file('[Content_Types].xml', types.replace('/xl/worksheets/sheet1.xml', '/xl/custom/café.xml'))
        })

        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        expect(r).toMatchObject({ status: 'extracted', truncated: false })
        expect(r.extraction).toBe('=== S1 ===\ns1r1\t1\ns1r2\t2')
    })

    it('keeps non-ASCII part names distinct when only Unicode case folding would merge them', async () => {
        const content = await rebuild(await workbookWith(2, 1), async (zip) => {
            for (const [from, to] of [
                ['xl/worksheets/sheet1.xml', 'xl/custom/Ä.xml'],
                ['xl/worksheets/sheet2.xml', 'xl/custom/ä.xml'],
            ]) {
                const sheet = await zip.file(from)!.async('nodebuffer')
                zip.remove(from)
                zip.file(to, sheet)
            }

            const rels = await part(zip, 'xl/_rels/workbook.xml.rels')
            zip.file(
                'xl/_rels/workbook.xml.rels',
                rels.replace('worksheets/sheet1.xml', 'custom/Ä.xml').replace('worksheets/sheet2.xml', 'custom/ä.xml')
            )
            const types = await part(zip, '[Content_Types].xml')
            zip.file(
                '[Content_Types].xml',
                types
                    .replace('/xl/worksheets/sheet1.xml', '/xl/custom/Ä.xml')
                    .replace('/xl/worksheets/sheet2.xml', '/xl/custom/ä.xml')
            )
        })

        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        expect(r).toMatchObject({ status: 'extracted', truncated: false })
        expect(r.extraction).toContain('=== S1 ===\ns1r1')
        expect(r.extraction).toContain('=== S2 ===\ns2r1')
    })

    it('resolves case-variant workbook metadata before placing a custom worksheet', async () => {
        const content = await rebuild(await workbookWith(1, 2), async (zip) => {
            const from = 'xl/worksheets/sheet1.xml'
            const to = 'xl/custom/data.xml'
            const sheet = await zip.file(from)!.async('nodebuffer')
            zip.remove(from)
            zip.file(to, sheet)

            const relsName = 'xl/_rels/workbook.xml.rels'
            const rels = (await part(zip, relsName)).replace('worksheets/sheet1.xml', 'custom/data.xml')
            zip.remove(relsName)
            zip.file('XL/_rels/Workbook.xml.rels', rels)
            const types = await part(zip, '[Content_Types].xml')
            zip.file('[Content_Types].xml', types.replace('/xl/worksheets/sheet1.xml', '/xl/custom/data.xml'))
        })

        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        expect(r).toMatchObject({ status: 'extracted', truncated: false })
        expect(r.extraction).toContain('=== S1 ===\ns1r1')
    })

    it.each(['uppercase Override media type', 'Default extension mapping'])(
        'authorizes a custom worksheet through %s',
        async (variant) => {
            const content = await rebuild(await workbookWith(1, 2), async (zip) => {
                const from = 'xl/worksheets/sheet1.xml'
                const to = variant.startsWith('Default') ? 'xl/custom/data.foo' : 'xl/custom/data.xml'
                const sheet = await zip.file(from)!.async('nodebuffer')
                zip.remove(from)
                zip.file(to, sheet)

                const rels = await part(zip, 'xl/_rels/workbook.xml.rels')
                zip.file('xl/_rels/workbook.xml.rels', rels.replace('worksheets/sheet1.xml', to.slice(3)))
                const types = await part(zip, '[Content_Types].xml')
                if (variant.startsWith('Default')) {
                    zip.file(
                        '[Content_Types].xml',
                        types
                            .replace(/<Override PartName="\/xl\/worksheets\/sheet1\.xml"[^>]*\/>/, '')
                            .replace(
                                '</Types>',
                                '<Default Extension="foo" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
                            )
                    )
                } else {
                    zip.file(
                        '[Content_Types].xml',
                        types
                            .replace('/xl/worksheets/sheet1.xml', '/xl/custom/data.xml')
                            .replace(
                                'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
                                'APPLICATION/VND.OPENXMLFORMATS-OFFICEDOCUMENT.SPREADSHEETML.WORKSHEET+XML'
                            )
                    )
                }
            })

            const r = await extractAttachment({ content, contentType: XLSX_TYPE })
            expect(r).toMatchObject({ status: 'extracted', truncated: false })
            expect(r.extraction).toContain('=== S1 ===\ns1r1')
        }
    )

    // Relationship type and content type are declarations from the same untrusted package, not
    // independent corroboration. Even when both call styles.xml a worksheet, a structural part must
    // never replace the real sheet or be emitted twice by the rebuilt archive.
    it.each([
        ['a reader-control part', 'styles.xml'],
        ['another non-worksheet XML part', 'theme/theme1.xml'],
    ])('rescues real rows when forged declarations point a worksheet at %s', async (_label, target) => {
        const content = await rebuild(await workbookWith(1, 2), async (zip) => {
            const rels = await part(zip, 'xl/_rels/workbook.xml.rels')
            zip.file('xl/_rels/workbook.xml.rels', rels.replace('Target="worksheets/sheet1.xml"', `Target="${target}"`))

            const types = await part(zip, '[Content_Types].xml')
            zip.file(
                '[Content_Types].xml',
                types.replace(
                    '</Types>',
                    `<Override PartName="/xl/${target}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`
                )
            )
        })

        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toContain('s1r1')
        expect(r.extraction!.match(/s1r1/g)).toHaveLength(1)
    })

    // Resolving OUTSIDE xl/worksheets is normal — chartsheets live there — but the declaration has to
    // be accountable. The first resolves to nothing; the last two resolve to a part that EXISTS,
    // which is why "does the archive hold it?" was the wrong question on its own — neither holds this
    // sheet's rows, and the part that does goes unclaimed.
    it.each([
        ['absolute without the xl/ prefix', '/worksheets/sheet2.xml', undefined],
        ['a real part that is not a sheet', 'styles.xml', undefined],
        ['a stray backup beside the real part', 'worksheets/sheet2.xml.bak', 'xl/worksheets/sheet2.xml.bak'],
    ])('keeps a declared sheet whose Target is %s', async (_label, target, plant) => {
        const content = await rebuild(await workbookWith(3, 2), async (zip) => {
            if (plant) zip.file(plant, await part(zip, 'xl/worksheets/sheet2.xml'))
            const xml = await part(zip, 'xl/_rels/workbook.xml.rels')
            zip.file('xl/_rels/workbook.xml.rels', xml.replace('worksheets/sheet2.xml', target))
        })

        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        for (let s = 1; s <= 3; s++) expect(r.extraction).toContain(`s${s}r1`)
    })

    // The control for the rule above: a chartsheet is a real declaration with no worksheet part, so
    // it must NOT count as unplaced — otherwise every workbook holding one flips into rescue mode and
    // resurrects the orphans this change exists to drop. Both halves asserted at once.
    it('drops an orphan part while a declared chartsheet reads as accounted for', async () => {
        const content = await rebuild(await workbookWith(1, 2), async (zip) => {
            zip.file('xl/chartsheets/sheet1.xml', '<chartsheet/>')
            zip.file('xl/worksheets/sheet9.xml', await part(zip, 'xl/worksheets/sheet1.xml'))
            const rels = await part(zip, 'xl/_rels/workbook.xml.rels')
            zip.file(
                'xl/_rels/workbook.xml.rels',
                rels.replace(
                    '</Relationships>',
                    '<Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.xml"/></Relationships>'
                )
            )
            const xml = await part(zip, 'xl/workbook.xml')
            zip.file('xl/workbook.xml', xml.replace('</sheets>', '<sheet name="Chart" sheetId="2" r:id="rIdChart"/></sheets>'))
        })

        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        // sheet9 is a copy of sheet1's rows, so a resurrected orphan shows up as a second header —
        // the row text alone could not tell the two apart.
        expect(headers(r.extraction)).toEqual(['=== S1 ==='])
    })

    // The control above passes for the wrong reason if "is this a chart sheet?" is answered from the
    // TARGET's path, which the producer writes. Here the relationship still says Type=worksheet and
    // only the path is redirected onto a planted chartsheet part: a path whitelist reads it as an
    // accounted-for chart sheet, and the worksheet part actually holding the rows is dropped.
    it('classifies a sheet by relationship type, not by the path its target happens to take', async () => {
        const content = await rebuild(await workbookWith(3, 2), async (zip) => {
            zip.file('xl/chartsheets/sheet1.xml', '<chartsheet/>')
            const rels = await part(zip, 'xl/_rels/workbook.xml.rels')
            zip.file(
                'xl/_rels/workbook.xml.rels',
                rels.replace('Target="worksheets/sheet2.xml"', 'Target="chartsheets/sheet1.xml"')
            )
        })

        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        for (let s = 1; s <= 3; s++) expect(r.extraction).toContain(`s${s}r1`)
    })

    // The mirror of the test above, and the reason neither signal is trusted alone: here the TYPE
    // says chartsheet while the target still points at the worksheet part holding the rows. Trusting
    // the type by itself accepted the declaration as accounted-for and dropped that part as an orphan.
    it('treats a non-worksheet type contradicted by its target as unplaced', async () => {
        const content = await rebuild(await workbookWith(3, 2), async (zip) => {
            const rels = await part(zip, 'xl/_rels/workbook.xml.rels')
            zip.file(
                'xl/_rels/workbook.xml.rels',
                rels.replace(
                    /(<Relationship[^>]*)Type="[^"]*\/worksheet"([^>]*Target="worksheets\/sheet2\.xml")/,
                    '$1Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"$2'
                )
            )
        })

        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        for (let s = 1; s <= 3; s++) expect(r.extraction).toContain(`s${s}r1`)
    })

    // And the third way the pair can disagree: the type names a chart sheet, the target resolves to a
    // real part, and that part is neither a worksheet nor a chart sheet. "Not a worksheet" was too
    // weak a test — the kinds have to match, or the sheet that does hold the rows goes unclaimed.
    it('treats a non-worksheet type pointing at an unrelated part as unplaced', async () => {
        const content = await rebuild(await workbookWith(3, 2), async (zip) => {
            const rels = await part(zip, 'xl/_rels/workbook.xml.rels')
            zip.file(
                'xl/_rels/workbook.xml.rels',
                rels.replace(
                    /(<Relationship[^>]*)Type="[^"]*\/worksheet"([^>]*)Target="worksheets\/sheet2\.xml"/,
                    '$1Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"$2Target="styles.xml"'
                )
            )
        })

        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        for (let s = 1; s <= 3; s++) expect(r.extraction).toContain(`s${s}r1`)
    })

    // And the fourth way the pair can disagree: a type that is not a relationship type at all, only
    // spelled to end like one. A type is an exact identifier the format defines, so matching its tail
    // let any stranger claim the authority — planted with a chart-sheet part to point at, so the
    // target agrees and nothing but the type's origin is wrong.
    it('rejects a foreign relationship type that merely ends like a chartsheet one', async () => {
        const content = await rebuild(await workbookWith(3, 2), async (zip) => {
            zip.file('xl/chartsheets/fake.xml', '<chartsheet/>')
            const rels = await part(zip, 'xl/_rels/workbook.xml.rels')
            zip.file(
                'xl/_rels/workbook.xml.rels',
                rels.replace(
                    /(<Relationship[^>]*)Type="[^"]*\/worksheet"([^>]*)Target="worksheets\/sheet2\.xml"/,
                    '$1Type="https://invalid.example/relationships/chartsheet"$2Target="chartsheets/fake.xml"'
                )
            )
        })

        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        for (let s = 1; s <= 3; s++) expect(r.extraction).toContain(`s${s}r1`)
    })

    // ISO/IEC 29500 Strict is a legal .xlsx — Excel offers it as "Strict Open XML Workbook" — and
    // re-homes the same vocabulary under purl.oclc.org. Recognizing only Transitional matched no
    // <sheet> at all, so a Strict workbook silently lost its tab order and names to the archive's.
    it('reads sheet identity from an ISO Strict workbook', async () => {
        const strict = await rebuild(await workbookWith(3, 2), async (zip) => {
            const xml = await part(zip, 'xl/workbook.xml')
            zip.file(
                'xl/workbook.xml',
                xml
                    .replace(/<sheets>(.*?)<\/sheets>/, (_, inner: string) => {
                        const elements = inner.match(/<sheet\b[^>]*\/>/g) ?? []
                        return `<sheets>${elements.reverse().join('')}</sheets>`
                    })
                    .replace(
                        'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                        'http://purl.oclc.org/ooxml/spreadsheetml/main'
                    )
                    .replace(
                        'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                        'http://purl.oclc.org/ooxml/officeDocument/relationships'
                    )
            )
        })

        const r = await extractAttachment({ content: strict, contentType: XLSX_TYPE })
        // Reversed tab order, under the workbook's own names — the archive fallback would give
        // Sheet1, Sheet2, Sheet3 in file order, which is what this used to return.
        expect(headers(r.extraction)).toEqual(['=== S3 ===', '=== S2 ===', '=== S1 ==='])
    })

    // Element identity is (namespace, local name) plus position. Matching bare local names let a
    // foreign <foo:sheet> in an extension block claim a real sheet's part before the genuine
    // declaration reached it, so the sheet came back under the injected name and tab position.
    it('ignores a foreign element whose local name is sheet', async () => {
        const content = await rebuild(await workbookWith(3, 2), async (zip) => {
            const xml = await part(zip, 'xl/workbook.xml')
            const rId = /<sheet [^>]*name="S2"[^>]*r:id="([^"]+)"/.exec(xml)![1]
            zip.file(
                'xl/workbook.xml',
                xml.replace(
                    '<sheets>',
                    `<extLst xmlns:foo="urn:example:foreign"><foo:sheet name="Injected" r:id="${rId}"/></extLst><sheets>`
                )
            )
        })

        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        expect(headers(r.extraction)).toEqual(['=== S1 ===', '=== S2 ===', '=== S3 ==='])
    })

    // The two metadata parts are read whole, so their size has to be bounded independently of the
    // 50 MB archive budget — one part spending all of it costs that twice over, as a Buffer and then
    // as the string the parse needs. Over the cap degrades to the archive's own parts; it never fails.
    it('falls back to the archive rather than materializing an oversized workbook part', async () => {
        const padded = await rebuild(await workbookWith(2, 2), async (zip) => {
            const xml = await part(zip, 'xl/workbook.xml')
            zip.file('xl/workbook.xml', xml.replace('<sheets>', `<!--${' '.repeat(5 * 1024 * 1024)}--><sheets>`))
        })

        const r = await extractAttachment({ content: padded, contentType: XLSX_TYPE })
        // Archive order and fallback names — the workbook's own naming is what we declined to read.
        expect(headers(r.extraction)).toEqual(['=== Sheet1 ===', '=== Sheet2 ==='])
        for (let s = 1; s <= 2; s++) expect(r.extraction).toContain(`s${s}r1`)
    })

    it('retains a custom worksheet when workbook metadata exceeds the resolver cap', async () => {
        const padded = await rebuild(await workbookWith(1, 2), async (zip) => {
            const from = 'xl/worksheets/sheet1.xml'
            const to = 'xl/custom/data.xml'
            const sheet = await zip.file(from)!.async('nodebuffer')
            zip.remove(from)
            zip.file(to, sheet)

            const rels = await part(zip, 'xl/_rels/workbook.xml.rels')
            zip.file('xl/_rels/workbook.xml.rels', rels.replace('worksheets/sheet1.xml', 'custom/data.xml'))
            const types = await part(zip, '[Content_Types].xml')
            zip.file('[Content_Types].xml', types.replace('/xl/worksheets/sheet1.xml', '/xl/custom/data.xml'))
            const workbook = await part(zip, 'xl/workbook.xml')
            zip.file('xl/workbook.xml', workbook.replace('<sheets>', `<!--${'x'.repeat(5 * 1024 * 1024)}--><sheets>`))
        })

        const r = await extractAttachment({ content: padded, contentType: XLSX_TYPE })
        expect(r).toMatchObject({ status: 'extracted', extraction: '=== Sheet ===\ns1r1\t1\ns1r2\t2', truncated: false })
    })

    it('keeps an Override from inheriting a worksheet Default in metadata fallback', async () => {
        const content = await rebuild(await workbookWith(1, 1), async (zip) => {
            const from = 'xl/worksheets/sheet1.xml'
            const to = 'xl/custom/data.xml'
            const sheet = await zip.file(from)!.async('nodebuffer')
            zip.remove(from)
            zip.file(to, sheet)

            const rels = await part(zip, 'xl/_rels/workbook.xml.rels')
            zip.file('xl/_rels/workbook.xml.rels', rels.replace('worksheets/sheet1.xml', 'custom/data.xml'))
            const types = await part(zip, '[Content_Types].xml')
            zip.file(
                '[Content_Types].xml',
                types
                    .replace(
                        'Extension="xml" ContentType="application/xml"',
                        'Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"'
                    )
                    .replace('/xl/worksheets/sheet1.xml', '/xl/custom/data.xml')
                    .replace(
                        'PartName="/xl/custom/data.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"',
                        'PartName="/xl/custom/data.xml" ContentType="application/xml"'
                    )
            )
            const workbook = await part(zip, 'xl/workbook.xml')
            zip.file('xl/workbook.xml', workbook.replace('<sheets>', `<!--${'x'.repeat(5 * 1024 * 1024)}--><sheets>`))
        })

        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        expect(r.status).toBe('failed')
        expect(r.extraction).toBeUndefined()
        expect(r.reason).toBe('wrong-document-shape')
    })

    it('resolves escaped OPC worksheet paths without decoding them away from the ZIP item', async () => {
        const content = await rebuild(await workbookWith(1, 1), async (zip) => {
            const from = 'xl/worksheets/sheet1.xml'
            const to = 'xl/worksheets/custom%20sheet.xml'
            const sheet = await zip.file(from)!.async('nodebuffer')
            zip.remove(from)
            zip.file(to, sheet)

            const rels = await part(zip, 'xl/_rels/workbook.xml.rels')
            zip.file('xl/_rels/workbook.xml.rels', rels.replace('worksheets/sheet1.xml', 'worksheets/custom%20sheet.xml'))
            const types = await part(zip, '[Content_Types].xml')
            zip.file('[Content_Types].xml', types.replace('/xl/worksheets/sheet1.xml', '/xl/worksheets/custom%20sheet.xml'))
        })

        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        expect(r).toMatchObject({ status: 'extracted', extraction: '=== S1 ===\ns1r1\t1', truncated: false })
    })

    // parseWorkbookParts uses saxes with namespaces enabled, whose prefix resolution scans every
    // open tag. The depth guard must make this fallback quickly; without it 20k nested elements take
    // several seconds synchronously and the handler timeout cannot fire while the event loop is held.
    it('bounds deeply nested workbook metadata before saxes becomes quadratic', async () => {
        const content = await rebuild(await workbookWith(2, 2), async (zip) => {
            const xml = await part(zip, 'xl/workbook.xml')
            const nested = '<extLst>'.repeat(20_000) + '</extLst>'.repeat(20_000)
            zip.file('xl/workbook.xml', xml.replace('<sheets>', `${nested}<sheets>`))
        })

        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        expect(r.status).toBe('extracted')
        for (let s = 1; s <= 2; s++) expect(r.extraction).toContain(`s${s}r1`)
    }, 2_000)

})

// The rebuild vs the decompression budget --------------------------------------
// The budget measures each central-directory record once, then reorderForStreaming decides what
// unzipper actually gets. That only binds if the rebuild can never write a measured entry more than
// once — and zip permits duplicate entry names, so resolving layout through a name lookup collapsed
// every reference onto the first entry carrying that name and re-emitted its bytes per reference.
// Measured at 21x on a 111 KB input: 1.5 MB budgeted, 31.5 MB rebuilt. Scaled inside the existing
// gates that is a 5 MB attachment forcing a 0.8 GB allocation, which is an OOM the API side cannot
// contain — not a JS throw, so its per-attachment catch never sees it.

describe('attachment — the rebuild cannot amplify what the budget measured', () => {
    // Hand-built, because JSZip dedupes by name and duplicate names are the whole point.
    const buildZip = (files: { name: string; data: Buffer }[]) => {
        const locals: Buffer[] = []
        const centrals: Buffer[] = []
        let offset = 0
        for (const file of files) {
            const name = Buffer.from(file.name, 'latin1')
            const deflated = zlib.deflateRawSync(file.data)
            const local = Buffer.alloc(30)
            local.writeUInt32LE(0x04034b50, 0)
            local.writeUInt16LE(20, 4)
            local.writeUInt16LE(8, 8)
            local.writeUInt32LE(zlib.crc32(file.data), 14)
            local.writeUInt32LE(deflated.length, 18)
            local.writeUInt32LE(file.data.length, 22)
            local.writeUInt16LE(name.length, 26)
            locals.push(local, name, deflated)
            const central = Buffer.alloc(46)
            central.writeUInt32LE(0x02014b50, 0)
            central.writeUInt16LE(20, 4)
            central.writeUInt16LE(20, 6)
            central.writeUInt16LE(8, 10)
            central.writeUInt32LE(zlib.crc32(file.data), 16)
            central.writeUInt32LE(deflated.length, 20)
            central.writeUInt32LE(file.data.length, 24)
            central.writeUInt16LE(name.length, 28)
            central.writeUInt32LE(offset, 42)
            centrals.push(central, name)
            offset += 30 + name.length + deflated.length
        }
        const directory = Buffer.concat(centrals)
        const end = Buffer.alloc(22)
        end.writeUInt32LE(0x06054b50, 0)
        end.writeUInt16LE(files.length, 8)
        end.writeUInt16LE(files.length, 10)
        end.writeUInt32LE(directory.length, 12)
        end.writeUInt32LE(offset, 16)
        return Buffer.concat([...locals, directory, end])
    }

    it('matches Mammoth by selecting the last duplicate DOCX main part', async () => {
        const document = (value: string) =>
            buf(
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
                    `<w:p><w:r><w:t>${value}</w:t></w:r></w:p></w:body></w:document>`
            )
        const content = buildZip([
            { name: '[Content_Types].xml', data: buf('<Types/>') },
            { name: 'word/document.xml', data: document('FIRST') },
            { name: 'word/document.xml', data: document('LAST') },
        ])

        const r = await extractAttachment({ content, contentType: DOCX_TYPE })
        expect(r).toMatchObject({ status: 'extracted', extraction: 'LAST\n\n', truncated: false })
    })

    // What checkDecompressionBudget counts: every central-directory record's uncompressed size, once.
    const budgetMeasures = (buf: Buffer) => {
        const eocd = buf.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]))
        let p = buf.readUInt32LE(eocd + 16)
        let total = 0
        for (let i = buf.readUInt16LE(eocd + 10); i > 0; i--) {
            total += buf.readUInt32LE(p + 24)
            p += 46 + buf.readUInt16LE(p + 28) + buf.readUInt16LE(p + 30) + buf.readUInt16LE(p + 32)
        }
        return total
    }

    // What unzipper would actually expand: every LOCAL record of the rebuilt archive.
    const rebuildInflatesTo = (buf: Buffer) => {
        let p = 0
        let total = 0
        while (p + 30 <= buf.length && buf.readUInt32LE(p) === 0x04034b50) {
            const method = buf.readUInt16LE(p + 8)
            const compSize = buf.readUInt32LE(p + 18)
            const start = p + 30 + buf.readUInt16LE(p + 26) + buf.readUInt16LE(p + 28)
            total += method === 0 ? compSize : zlib.inflateRawSync(buf.subarray(start, start + compSize)).length
            p = start + compSize
        }
        return total
    }

    it.each(['xl/sharedStrings.xml', 'xl/SharedStrings.xml'])(
        'fails closed when a duplicate control part is stored as %s',
        async (duplicateName) => {
            const xml = (value: string) => Buffer.from(value, 'latin1')
            const content = buildZip([
                {
                    name: '[Content_Types].xml',
                    data: xml(
                        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
                            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
                            '</Types>'
                    ),
                },
                {
                    name: 'xl/workbook.xml',
                    data: xml(
                        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
                            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
                            '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>'
                    ),
                },
                {
                    name: 'xl/_rels/workbook.xml.rels',
                    data: xml(
                        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
                            '</Relationships>'
                    ),
                },
                {
                    name: 'xl/worksheets/sheet1.xml',
                    data: xml(
                        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
                            '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>'
                    ),
                },
                {
                    name: 'xl/sharedStrings.xml',
                    data: xml(
                        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>KEEP</t></si></sst>'
                    ),
                },
                {
                    name: duplicateName,
                    data: xml(
                        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>'
                    ),
                },
            ])

            const r = await extractAttachment({ content, contentType: XLSX_TYPE })
            expect(r.status).toBe('failed')
            expect(r.reason).toBe('malformed')
        }
    )

    it('writes each measured entry at most once, even when many entries share a name', async () => {
        const DUPLICATES = 20
        const rows = Array.from(
            { length: 20_000 },
            (_, r) => `<row r="${r + 1}"><c t="inlineStr"><is><t>padding cell ${r}</t></is></c></row>`
        ).join('')
        const sheetXml = (body: string) =>
            Buffer.from(
                `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${body}</worksheet>`,
                'latin1'
            )
        // One heavy entry and 20 trivial ones, ALL named xl/worksheets/sheet1.xml. No rels part, so
        // resolution takes the archive fallback — the path that used to map all 21 onto the heavy one.
        const content = buildZip([
            { name: '[Content_Types].xml', data: Buffer.from('<Types/>', 'latin1') },
            {
                name: 'xl/workbook.xml',
                data: Buffer.from('<workbook><sheets><sheet name="A" sheetId="1" r:id="rId1"/></sheets></workbook>', 'latin1'),
            },
            { name: 'xl/worksheets/sheet1.xml', data: sheetXml(`<sheetData>${rows}</sheetData>`) },
            ...Array.from({ length: DUPLICATES }, () => ({
                name: 'xl/worksheets/sheet1.xml',
                data: sheetXml('<sheetData/>'),
            })),
        ])

        // Capture the bytes the handler hands the reader. Asserting on the extraction instead would
        // measure the wrong thing — the amplification is in the archive, before a cap can apply.
        let rebuilt = Buffer.alloc(0)
        vi.resetModules()
        vi.doMock('exceljs', () => ({
            default: {
                stream: {
                    xlsx: {
                        WorkbookReader: class {
                            stream: NodeJS.ReadableStream
                            constructor(stream: NodeJS.ReadableStream) {
                                this.stream = stream
                                stubExcelReaderHooks(this)
                            }
                            async *[Symbol.asyncIterator]() {
                                const chunks: Buffer[] = []
                                for await (const chunk of this.stream) chunks.push(chunk as Buffer)
                                rebuilt = Buffer.concat(chunks)
                            }
                        },
                    },
                },
            },
        }))
        try {
            const { extractAttachment: extract } = await import('../attachment')
            await extract({ content, contentType: XLSX_TYPE })
        } finally {
            vi.doUnmock('exceljs')
            vi.resetModules()
        }

        // The only entries the rebuild adds without central records are fixed empty control parts
        // of ours, not anything the input controls.
        expect(rebuildInflatesTo(rebuilt)).toBeLessThanOrEqual(budgetMeasures(content) + 300)
    })

    // The same "a name is not a key" mistake, on the other side of the ledger: layout stopped
    // duplicating entries, but the rescue that saves unplaced ones still excluded them BY NAME, so
    // the entry nobody laid out was thrown away with the one that was. Here rather than beside the
    // other sheet-identity tests because only this block can build an archive JSZip refuses to.
    it('rescues an entry sharing a claimed name rather than excluding it by name', async () => {
        const sheet = (v: string) =>
            Buffer.from(
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
                    `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${v}</t></is></c></row></sheetData></worksheet>`,
                'latin1'
            )
        // Alpha resolves to sheet1.xml; Beta's r:id matches no relationship, so it goes unplaced and
        // the rescue engages. Two DISTINCT entries share that name, and only one is ever laid out.
        const content = buildZip([
            { name: '[Content_Types].xml', data: Buffer.from('<Types/>', 'latin1') },
            {
                name: 'xl/workbook.xml',
                data: Buffer.from(
                    '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' +
                        '<sheet name="Alpha" sheetId="1" r:id="rId1"/><sheet name="Beta" sheetId="2" r:id="rIdMISSING"/>' +
                        '</sheets></workbook>',
                    'latin1'
                ),
            },
            {
                name: 'xl/_rels/workbook.xml.rels',
                data: Buffer.from(
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
                        '</Relationships>',
                    'latin1'
                ),
            },
            { name: 'xl/worksheets/sheet1.xml', data: sheet('AAA-claimed') },
            { name: 'xl/worksheets/sheet1.xml', data: sheet('BBB-the-victim') },
        ])

        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        expect(r.extraction).toContain('AAA-claimed')
        expect(r.extraction).toContain('BBB-the-victim')
    })
})

// Post-cap emptiness, CRC pin, and the lost-worksheet backstop --------------------------

describe('attachment — emptiness is decided after the cap', () => {
    // The pdf handler reports `empty` from its PRE-cap page join, so on a tight cap it answers
    // `false` about text the central slice then reduces to nothing; `??` would emit that as '', or
    // as a bare trailer with no document text. Routed through a PDF because pdf is the only handler
    // that sets `empty` — a text/plain input cannot exercise this.
    const pdfWithText = () => buildPdf([['some real extractable text on page one']])

    it('omits extraction when the cap slices a pdf to nothing, rather than returning an empty string', async () => {
        const r = await extractAttachment({ content: pdfWithText(), contentType: 'application/pdf' }, { maxOutputChars: 0 })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBeUndefined() // never '', per the contract
        expect('extraction' in r).toBe(false)
    })

    it('never returns a bare trailer with no document text in front of it', async () => {
        const trailer = '\n[truncated]'
        const r = await extractAttachment(
            { content: pdfWithText(), contentType: 'application/pdf' },
            { maxOutputChars: 0, trailer }
        )
        expect(r.extraction).toBeUndefined()
        expect(r.extraction ?? '').not.toBe(trailer)
    })

    // A handler's own `empty: true` must still be honoured: the guard only stops a stale `false`
    // from overriding a genuinely-empty slice.
    it('still honours a handler that reports itself empty', async () => {
        const r = await extractAttachment({ content: fixture('blank.pdf'), contentType: 'application/pdf' })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBeUndefined()
    })
})

// xlsx deadline enforcement ---------------------------------------------------
// The deadline check has to run once per ROW, not once per row that produced text: an all-empty row
// is skipped and trips no cap, so a check below that skip leaves the budget unenforced exactly where
// the loop is cheapest to spin.

describe('attachment — xlsx honours the deadline on contentless rows', () => {
    // addRow(['']) emits a <row> yielding no non-empty cells — measured: all 30 reach the reader and
    // all are skipped by the handler. addRow([]) emits nothing, so it cannot exercise this. The
    // trailing marker row is what makes the stop observable from outside.
    const sheetOfBlankRows = async (blanks: number) => {
        const workbook = new ExcelJS.Workbook()
        const sheet = workbook.addWorksheet('Blanks')
        for (let i = 0; i < blanks; i++) sheet.addRow([''])
        sheet.addRow(['TRAILING-MARKER'])
        return Buffer.from(await workbook.xlsx.writeBuffer())
    }

    it('stops on a sheet of blank-but-present rows instead of iterating to the end', async () => {
        const content = await sheetOfBlankRows(400)

        // Blunt, like the pdf deadline test: the FIRST Date.now() sets the deadline, every later
        // call lands past it. exceljs/unzipper may read the clock themselves, so a mock letting N
        // rows through would depend on how often they do.
        const base = Date.now()
        let calls = 0
        const clock = vi.spyOn(Date, 'now').mockImplementation(() => {
            calls += 1
            return calls === 1 ? base : base + HANDLER_TIMEOUT_MS + 1
        })
        const stopped = await extractAttachment({ content, contentType: XLSX_TYPE })
        clock.mockRestore()

        expect(stopped.status).toBe('extracted') // a deadline stop is not a failure
        expect(stopped.truncated).toBe(true)
        // The point of the test: the blank rows must not have been walked through to reach this.
        expect(stopped.extraction ?? '').not.toContain('TRAILING-MARKER')

        // Control: a real clock reads the same bytes to the end and flags nothing. Without it the
        // assertions above would also pass if the handler simply failed on this workbook.
        const whole = await extractAttachment({ content, contentType: XLSX_TYPE })
        expect(whole.truncated).toBe(false)
        expect(whole.extraction).toContain('TRAILING-MARKER')
    })

    it('returns earlier sheet text when the deadline expires across row-less sheets', async () => {
        const content = await workbookWith(2, 1)
        let readKept = false
        const row = {
            eachCell: (_options: unknown, callback: (cell: { text: string }) => void) => {
                callback({ text: 'kept' })
                readKept = true
            },
        }
        const worksheet = (rows: unknown[]) => ({
            async *[Symbol.asyncIterator]() {
                for (const value of rows) yield value
            },
        })

        vi.resetModules()
        vi.doMock('exceljs', () => ({
            default: {
                stream: {
                    xlsx: {
                        WorkbookReader: class {
                            constructor() {
                                stubExcelReaderHooks(this)
                            }
                            async *[Symbol.asyncIterator]() {
                                yield worksheet([row])
                                yield worksheet([]) // never enters the per-row deadline check
                            }
                        },
                    },
                },
            },
        }))
        const base = Date.now()
        const clock = vi
            .spyOn(Date, 'now')
            .mockImplementation(() => (readKept ? base + HANDLER_TIMEOUT_MS : base))
        try {
            const { extractAttachment: extract } = await import('../attachment')
            const r = await extract({ content, contentType: XLSX_TYPE })
            expect(r.status).toBe('extracted')
            expect(r.truncated).toBe(true)
            expect(r.extraction).toContain('kept')
        } finally {
            clock.mockRestore()
            vi.doUnmock('exceljs')
            vi.resetModules()
        }
    })

    it('checks the deadline while parsing non-row worksheet markup', async () => {
        const workbook = new ExcelJS.Workbook()
        workbook.addWorksheet('Markup').addRow(['kept'])
        const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer())
        const name = 'xl/worksheets/sheet1.xml'
        const xml = await zip.file(name)!.async('string')
        zip.file(name, xml.replace('</worksheet>', `<extLst><!--${'x'.repeat(512 * 1024)}--></extLst></worksheet>`))
        const content = await zip.generateAsync({ type: 'nodebuffer' })

        const base = Date.now()
        let calls = 0
        const clock = vi
            .spyOn(Date, 'now')
            .mockImplementation(() => (++calls < 8 ? base : base + HANDLER_TIMEOUT_MS + 1))
        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        clock.mockRestore()

        expect(r.status).toBe('extracted')
        expect(r.truncated).toBe(true)
        expect(r.extraction).toContain('kept')
    })
})

describe('attachment — xlsx incremental output cap', () => {
    it('charges sheet headers and separators before reading later worksheets', async () => {
        const content = await workbookWith(20, 1)
        const row = {
            eachCell: (_options: unknown, callback: (cell: { text: string }) => void) => callback({ text: 'x' }),
        }
        const worksheet = {
            async *[Symbol.asyncIterator]() {
                yield row
            },
        }

        vi.resetModules()
        vi.doMock('exceljs', () => ({
            default: {
                stream: {
                    xlsx: {
                        WorkbookReader: class {
                            constructor() {
                                stubExcelReaderHooks(this)
                            }
                            async *[Symbol.asyncIterator]() {
                                for (let i = 0; i < 4; i++) yield worksheet
                                throw new Error('reader advanced beyond the header-bound cap')
                            }
                        },
                    },
                },
            },
        }))
        try {
            const { extractAttachment: extract } = await import('../attachment')
            const r = await extract({ content, contentType: XLSX_TYPE }, { maxOutputChars: 20 })
            expect(r.status).toBe('extracted')
            expect(r.truncated).toBe(true)
            expect(r.extraction).toBe('=== S1 ===\nx\n\n=== S2')
        } finally {
            vi.doUnmock('exceljs')
            vi.resetModules()
        }
    })
})

describe('attachment — xlsx lost-worksheet backstop', () => {
    // Pins the precomputed CRC against a fresh computation, so the constant and the literal it
    // describes cannot drift if the injected XML is edited.
    it('the injected control parts match their precomputed CRCs', () => {
        const parts: Array<[string, number]> = [
            [
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>',
                0x2949bd0b,
            ],
            [
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
                0x9f1f1b86,
            ],
        ]
        for (const [body, crc] of parts) expect(zlib.crc32(Buffer.from(body, 'latin1'))).toBe(crc)
    })

    // The reorder closes the two KNOWN triggers; this proves the CLASS is closed. Deleting a
    // worksheet part cannot simulate it — that lowers the archive's count too, so there is no
    // mismatch. The failure is a reader under-yielding against an intact archive, so that is stubbed.
    it('fails rather than silently returning a workbook with a missing worksheet', async () => {
        const workbook = new ExcelJS.Workbook()
        for (const name of ['Alpha', 'Beta', 'Gamma']) workbook.addWorksheet(name).addRow([`${name} data`])
        const content = Buffer.from(await workbook.xlsx.writeBuffer())

        vi.resetModules()
        vi.doMock('exceljs', async () => {
            const actual = await vi.importActual<{ default: typeof ExcelJS }>('exceljs')
            const Real = actual.default.stream.xlsx.WorkbookReader
            // Yields the first two of the archive's three worksheets, then stops — exactly the shape
            // of the upstream entry loss, minus the nondeterminism.
            class UnderYieldingReader extends Real {
                async *[Symbol.asyncIterator]() {
                    let yielded = 0
                    for await (const worksheet of super[Symbol.asyncIterator]()) {
                        if (yielded++ >= 2) return
                        yield worksheet
                    }
                }
            }
            return {
                ...actual,
                default: {
                    ...actual.default,
                    stream: { xlsx: { ...actual.default.stream.xlsx, WorkbookReader: UnderYieldingReader } },
                },
            }
        })

        const { extractAttachment: withLossyReader } = await import('../attachment')
        const r = await withLossyReader({ content, contentType: XLSX_TYPE })
        vi.doUnmock('exceljs')
        vi.resetModules()

        expect(r.status).toBe('failed')
        expect(r.reason).toBe('internal')
    })

    it('fails fast when the pinned ExcelJS streaming hooks are unavailable', async () => {
        const content = await workbookWith(1, 1)
        vi.resetModules()
        vi.doMock('exceljs', () => ({
            default: {
                stream: {
                    xlsx: {
                        WorkbookReader: class {
                            constructor() {
                                stubExcelReaderHooks(this, '_parseStyles')
                            }
                            async *[Symbol.asyncIterator]() {}
                        },
                    },
                },
            },
        }))
        try {
            const { extractAttachment: extract } = await import('../attachment')
            const r = await extract({ content, contentType: XLSX_TYPE })
            expect(r.status).toBe('failed')
            expect(r.reason).toBe('internal')
        } finally {
            vi.doUnmock('exceljs')
            vi.resetModules()
        }
    })

    // The backstop must not fire on a deliberate stop: truncation leaves later sheets unread BY
    // DESIGN, and turning that into a failure would break every capped large workbook.
    it('does not fire when extraction stopped early at the cap', async () => {
        const workbook = new ExcelJS.Workbook()
        const big = workbook.addWorksheet('Big')
        const cell = 'x'.repeat(40)
        for (let i = 0; i < 5_000; i++) big.addRow([cell, cell, cell]) // flattens past the 250k cap
        workbook.addWorksheet('Later').addRow(['never reached'])
        const content = Buffer.from(await workbook.xlsx.writeBuffer())

        const r = await extractAttachment({ content, contentType: XLSX_TYPE })
        expect(r.status).toBe('extracted') // not 'failed', despite 1 of 2 worksheets read
        expect(r.truncated).toBe(true)
        expect(r.extraction).not.toContain('=== Later ===')
    })
})

// xlsx archive-rewrite limits -------------------------------------------------
// Built by hand: no writer emits a 65534-entry workbook, and this needs the count exactly. Every
// entry is empty and stored, keeping the fixture ~5.6 MB (inside MAX_INPUT_BYTES) with a
// decompressed total of zero, so it reaches the rewrite instead of a gate in front of it.

describe('attachment — xlsx archive rewrite limits', () => {
    // `dropped` names that many entries so the rewrite's third group discards them: the unanchored
    // regex exceljs dispatches on matches these, the anchored WORKSHEET_PART does not, so they are
    // neither laid out nor kept. That is what makes the pre-drop and post-drop counts differ.
    const zipWithEntryCount = (count: number, dropped = 0): Buffer => {
        // One entry must be xl/workbook.xml so the archive routes as .xlsx; the rest are filler.
        const names = Array.from({ length: count }, (_, i) =>
            i === 0
                ? 'xl/workbook.xml'
                : i <= dropped
                  ? `xl/worksheets/sheet${i}.xml.bak`
                  : `p/${i.toString(36).padStart(5, '0')}`
        )
        const locals: Buffer[] = []
        const centrals: Buffer[] = []
        const offsets: number[] = []
        let offset = 0
        for (const entry of names) {
            const name = Buffer.from(entry, 'latin1')
            const local = Buffer.alloc(30) // method 0, crc 0, both sizes 0 — an empty stored entry
            local.writeUInt32LE(0x04034b50, 0)
            local.writeUInt16LE(20, 4)
            local.writeUInt16LE(name.length, 26)
            locals.push(local, name)
            offsets.push(offset)
            offset += 30 + name.length
        }
        names.forEach((entry, i) => {
            const name = Buffer.from(entry, 'latin1')
            const central = Buffer.alloc(46)
            central.writeUInt32LE(0x02014b50, 0)
            central.writeUInt16LE(20, 4)
            central.writeUInt16LE(20, 6)
            central.writeUInt16LE(name.length, 28)
            central.writeUInt32LE(offsets[i], 42)
            centrals.push(central, name)
        })
        const directory = Buffer.concat(centrals)
        const end = Buffer.alloc(22)
        end.writeUInt32LE(0x06054b50, 0)
        end.writeUInt16LE(count, 8)
        end.writeUInt16LE(count, 10)
        end.writeUInt32LE(directory.length, 12)
        end.writeUInt32LE(offset, 16)
        return Buffer.concat([...locals, directory, end])
    }

    // These fixtures are ~5.6 MB with 65k entries by construction — the size IS the test. For the
    // accepting cases, replace ExcelJS with a drain-only reader and freeze timers: the behavior under
    // test is the rewritten archive's entry count, not whether a loaded CI worker parses 65k empty
    // entries inside the production 10-second handler timeout.
    const BOUNDARY_TIMEOUT_MS = 30_000

    const rewriteAtBoundary = async (content: Buffer) => {
        let rebuilt = Buffer.alloc(0)
        vi.resetModules()
        vi.doMock('exceljs', () => ({
            default: {
                stream: {
                    xlsx: {
                        WorkbookReader: class {
                            stream: NodeJS.ReadableStream
                            constructor(stream: NodeJS.ReadableStream) {
                                this.stream = stream
                                stubExcelReaderHooks(this)
                            }
                            async *[Symbol.asyncIterator]() {
                                const chunks: Buffer[] = []
                                for await (const chunk of this.stream) chunks.push(chunk as Buffer)
                                rebuilt = Buffer.concat(chunks)
                            }
                        },
                    },
                },
            },
        }))
        vi.useFakeTimers()
        try {
            const { extractAttachment: extract } = await import('../attachment')
            const result = await extract({ content, contentType: XLSX_TYPE })
            const eocd = rebuilt.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]))
            return { result, entryCount: eocd < 0 ? undefined : rebuilt.readUInt16LE(eocd + 10) }
        } finally {
            vi.useRealTimers()
            vi.doUnmock('exceljs')
            vi.resetModules()
        }
    }

    // The reason is asserted, not just the status: this archive's directory read perfectly well — we
    // declined to REWRITE it — and reporting that as "could not be read" sent an investigation at the
    // wrong half of the preflight. Two distinct causes must not collapse into one message.
    it(
        'refuses an archive whose rewrite would land on the 0xffff entry-count sentinel',
        async () => {
            // 65533 real entries + two injected control parts = 65535 = 0xffff.
            const r = await extractAttachment({ content: zipWithEntryCount(0xffff - 2), contentType: XLSX_TYPE })
            // `skipped`, not `failed`: the archive is intact and we decline the variant, which is
            // the same answer the decompression budget gives ZIP64 for the same reason. The status
            // is derived from the reason now, so one fact cannot come back under two statuses.
            expect(r.status).toBe('skipped')
            expect(r.reason).toBe('unsupported-zip-feature')
        },
        BOUNDARY_TIMEOUT_MS
    )

    it(
        'accepts one entry below that boundary, so the refusal is the sentinel and not the size',
        async () => {
            // 65532 + 2 = 65534, a legal count. Same shape, same ~5.6 MB, one fewer entry: without
            // this the test above would also pass if the rewrite simply gave up on large archives.
            const { result, entryCount } = await rewriteAtBoundary(zipWithEntryCount(0xffff - 3))
            expect(result.status).toBe('extracted')
            expect(entryCount).toBe(0xfffe)
        },
        BOUNDARY_TIMEOUT_MS
    )

    // The count that matters is the one the EOCD is written from, which is post-drop. Checking the
    // pre-drop count refused archives whose actual rewrite lands comfortably under the sentinel —
    // the same 65535 as the first test, minus two entries the third group discards.
    it(
        'counts the entries it will write, not the ones it was given',
        async () => {
            const { result, entryCount } = await rewriteAtBoundary(zipWithEntryCount(0xffff - 2, 2))
            expect(result.status).toBe('extracted')
            expect(entryCount).toBe(0xfffd)
        },
        BOUNDARY_TIMEOUT_MS
    )
})

// PDF truncation signals ------------------------------------------------------
// The pdf handler stops for three reasons, each of which has to reach the caller as `truncated`.
// All three are asserted because they fail differently — content-driven cap, structural page
// ceiling, time-driven deadline.

describe('attachment — pdf reports every way it can stop early', () => {
    const PDF = 'application/pdf'

    // Two mechanisms produce the flag — the handler's own break, and the page of overshoot tripping
    // the entry point's over-cap check — so breaking the handler's flag alone does NOT fail this.
    // Not quite redundant: the handler's running total counts a trailing page join the final text
    // lacks, so a document landing within ~2 chars of the cap is flagged only by the handler.
    it('reports truncated when the output cap stops it', async () => {
        // 90 pages x 55 lines (~57 chars each) ≈ 280k extractable chars, past the 250k cap.
        const line = 'the quick brown fox jumps over the lazy dog and then some'
        const pdf = buildPdf(Array.from({ length: 90 }, () => Array.from({ length: 55 }, () => line)))
        const r = await extractAttachment({ content: pdf, contentType: PDF })
        expect(r.status).toBe('extracted')
        expect(r.truncated).toBe(true)
        expect(r.extraction!.length).toBeLessThanOrEqual(MAX_OUTPUT_CHARS)
    })

    it('does not count a trailing page separator or flag a complete near-cap PDF', async () => {
        const pdf = buildPdf([['a complete single-page document near its cap']])
        const whole = await extractAttachment({ content: pdf, contentType: PDF })
        const length = whole.extraction!.length
        const trailer = '\n[truncated]'

        for (const cap of [length, length + 1]) {
            const r = await extractAttachment({ content: pdf, contentType: PDF }, { maxOutputChars: cap, trailer })
            expect(r.truncated).toBe(false)
            expect(r.extraction).toBe(whole.extraction)
            expect(r.extraction).not.toContain(trailer)
        }
    })

    // Pages past MAX_PDF_PAGES are never read, so their text is missing whether or not the cap was
    // reached — 2001 pages of one word each stays far under the cap and would otherwise look
    // complete. This flag is the only signal that says otherwise.
    it('reports truncated when the page ceiling stops it, even far under the output cap', async () => {
        const pdf = buildPdf(Array.from({ length: MAX_PDF_PAGES + 1 }, (_, i) => [`page${i + 1}`]))
        const r = await extractAttachment({ content: pdf, contentType: PDF })
        expect(r.status).toBe('extracted')
        expect(r.truncated).toBe(true)
        expect(r.extraction!.length).toBeLessThan(MAX_OUTPUT_CHARS) // nowhere near the cap
        expect(r.extraction).toContain('page1')
        expect(r.extraction).not.toContain(`page${MAX_PDF_PAGES + 1}`) // the page past the ceiling
    })

    it('reports truncated when the deadline stops it', async () => {
        const pdf = buildPdf(Array.from({ length: 20 }, (_, i) => [`page ${i + 1} of the document`]))

        // extractAttachment's FIRST Date.now() sets the deadline; every later call lands past it, so
        // the page loop stops on its first check. Deliberately blunt — pdf.js may read the clock
        // itself, so a mock letting N pages through would depend on how often it does.
        const base = Date.now()
        let calls = 0
        const clock = vi.spyOn(Date, 'now').mockImplementation(() => {
            calls += 1
            return calls === 1 ? base : base + HANDLER_TIMEOUT_MS + 1
        })
        const stopped = await extractAttachment({ content: pdf, contentType: PDF })
        clock.mockRestore()

        expect(stopped.status).toBe('extracted') // a deadline stop is not a failure
        expect(stopped.truncated).toBe(true)

        // Control: the same bytes with a real clock are complete and NOT flagged — without this the
        // assertion above would also pass if `truncated` were hardcoded true.
        const whole = await extractAttachment({ content: pdf, contentType: PDF })
        expect(whole.truncated).toBe(false)
        expect(whole.extraction).toContain('page 20 of the document')
    })

    it('leaves truncated false for a pdf that is read completely', async () => {
        const r = await extractAttachment({ content: buildPdf([['short document']]), contentType: PDF })
        expect(r.status).toBe('extracted')
        expect(r.truncated).toBe(false)
        expect(r.extraction).toContain('short document')
    })
})
