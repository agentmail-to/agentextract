import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'

import { extractAttachment, HANDLER_TIMEOUT_MS } from '../attachment'

// ---------------------------------------------------------------------------
// .docx on malformed and hostile archives
// ---------------------------------------------------------------------------
// THE ERROR POLICY, stated once. saxes is a conformant parser; the DOM parser it replaced recovered
// from some malformedness and threw on other. So the swap moves the line in BOTH directions, and
// this file is where that is pinned rather than discovered. The policy:
//
//   text was read, then the parse broke  -> `extracted` + `truncated: true`. A prefix is honest, and
//                                           the contract already has a word for "the document
//                                           continues past this point".
//   nothing was read                     -> `failed`, which a caller can see and retry.
//
// Deliberately NOT a saxes error handler that reports and parses on. Measured, that recovery emits
// close-tag text as content and descends into elements mammoth drops — silent wrong output, which is
// the one outcome this library fails over everywhere else.
//
// Measured against mammoth on the eleven cases below: seven agree, two we are STRICTER on (a raw
// 0x0B, and the same character as a numeric reference — U+000B is not a legal XML character, and
// xmldom passed it through where saxes refuses), and two we are MORE AVAILABLE on (trailing junk
// after the root, and an unclosed root, both of which made mammoth throw away the whole document).
// Each case below names which.
//
// An undeclared namespace prefix is NOT in this file, deliberately. It used to be, asserting
// `truncated` — but the fixture put the offending element after `</w:document>`, so it was measuring
// trailing junk under a misleading name and duplicating the case above it. Undeclared prefixes no
// longer break the parse at all (see the resolvePrefix block in docx-streaming.test.ts), which is
// what a test measuring the right thing would have said.

const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const zipWith = async (name: string, content: string | Buffer, compression: 'DEFLATE' | 'STORE' = 'DEFLATE') => {
    const zip = new JSZip()
    zip.file(
        '[Content_Types].xml',
        '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'
    )
    zip.file(
        '_rels/.rels',
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    )
    zip.file(name, content)
    return Buffer.from(await zip.generateAsync({ type: 'nodebuffer', compression }))
}

const docxOf = (documentXml: string) => zipWith('word/document.xml', documentXml)
const body = (inner: string) => `<?xml version="1.0"?><w:document ${W}><w:body>${inner}</w:body></w:document>`
const para = (s: string) => `<w:p><w:r><w:t>${s}</w:t></w:r></w:p>`

const extract = async (documentXml: string) =>
    extractAttachment({ content: await docxOf(documentXml), contentType: DOCX_TYPE })

const VT = String.fromCharCode(11) // U+000B — not a legal XML 1.0 character

describe('docx — malformed XML', () => {
    // The two we are STRICTER on. Both are the same illegal character by two spellings, and both
    // sit in the FIRST text node — so nothing was read before the parse broke, and the policy says
    // `failed`. The same character further into a document would keep everything before it. Pinned
    // as a deliberate coverage loss, not an oversight: a document that is not well-formed XML now
    // gets a labeled refusal instead of text the previous parser guessed at.
    it.each([
        ['a raw control character in a w:t', body(para(`a${VT}b`))],
        ['the same character as a numeric reference', body(para('a&#11;b'))],
    ])('fails on %s, where the DOM parser passed it through', async (_name, documentXml) => {
        const r = await extract(documentXml)
        expect(r.status).toBe('failed')
        expect(r.extraction).toBeUndefined()
    })

    // The two we are MORE AVAILABLE on: mammoth threw away the entire document for both of these,
    // and the text before the fault is perfectly good.
    it.each([
        ['trailing junk after the root element', `${body(para('one'))}<junk/>`],
        ['an unclosed root element', `<?xml version="1.0"?><w:document ${W}><w:body>${para('one')}`],
    ])('keeps the text read before %s, and flags it truncated', async (_name, documentXml) => {
        const r = await extract(documentXml)
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBe('one\n\n')
        expect(r.truncated).toBe(true)
    })

    // Same policy, a case where mammoth also recovered. The text matches it exactly; what differs is
    // that we say the document was cut and it did not.
    it('keeps the text and flags truncated on a mismatched close tag', async () => {
        const r = await extract(
            `<?xml version="1.0"?><w:document ${W}><w:body>${para('one')}</w:zzz></w:body></w:document>`
        )
        expect(r.status).toBe('extracted')
        expect(r.truncated).toBe(true)
    })

    // The degradation the two `failed` cases above only show the worst end of. The same illegal
    // character, moved off the first text node, costs everything after it and nothing before —
    // which is what makes "stricter than mammoth" a bounded loss rather than a cliff.
    it('keeps everything before an illegal character that appears partway through', async () => {
        const r = await extract(body(`${para('first')}${para('second')}${para(`third${VT}broken`)}${para('fourth')}`))
        expect(r.status).toBe('extracted')
        expect(r.truncated).toBe(true)
        expect(r.extraction).toBe('first\n\nsecond\n\n')
    })

    // mammoth threw on both of these too, so this is agreement rather than a change.
    it.each([
        ['an undefined entity', body(para('a&nbsp;b'))],
        ['a duplicate attribute', `<?xml version="1.0"?><w:document ${W}><w:body><w:p><w:r><w:t xml:space="preserve" xml:space="default">x</w:t></w:r></w:p></w:body></w:document>`],
    ])('fails on %s', async (_name, documentXml) => {
        expect((await extract(documentXml)).status).toBe('failed')
    })

    // mammoth threw "Could not find the body element: are you sure this is a docx file?" for both.
    // Without this check a zip whose main part is not WordprocessingML would come back `extracted`
    // with no text — a silent nothing rather than a labeled refusal.
    it.each([
        ['a w:document with no w:body', `<?xml version="1.0"?><w:document ${W}><w:notBody>${para('x')}</w:notBody></w:document>`],
        ['a foreign root element', '<?xml version="1.0"?><nope xmlns="http://example.com/x"><p>hi</p></nope>'],
    ])('fails on %s rather than extracting nothing', async (_name, documentXml) => {
        const r = await extract(documentXml)
        expect(r.status).toBe('failed')
        expect(r.reason).toMatch(/w:body/)
    })
})

describe('docx — hostile payloads', () => {
    // saxes never parses entity declarations: its entity table is the five predefined ones on a null
    // prototype, and nothing ever writes to it. So the expansion cannot happen — the reference is
    // simply undefined and the parse fails. Asserted on TIME as well as status, because "bounded"
    // is the actual property; a parser that expanded this would not return at all.
    it('refuses a billion-laughs payload without expanding it', async () => {
        const entities =
            '<!ENTITY l0 "ha">' +
            [1, 2, 3, 4].map((n) => `<!ENTITY l${n} "${`&l${n - 1};`.repeat(10)}">`).join('')
        const started = Date.now()
        const r = await extract(
            `<?xml version="1.0"?><!DOCTYPE d [${entities}]><w:document ${W}><w:body>${para('&l4;')}</w:body></w:document>`
        )
        expect(Date.now() - started).toBeLessThan(HANDLER_TIMEOUT_MS)
        expect(r.status).toBe('failed')
        expect(r.extraction).toBeUndefined()
    })

    // 1000 seeded iterations. The property is not "produces good text" — it is that the degraded
    // state is always REACHABLE and always LABELED, which is what makes shipping without a fallback
    // reader safe. Seeded so a failure is reproducible; the long unseeded runs live in eval/.
    it('never throws, always labels, and always returns in time on mutated bytes', async () => {
        let seed = 0x5eed
        const rnd = () => (seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff

        // Two archives, because they fuzz different layers. The deflated one mostly exercises the
        // zip walk and the decompression preflight; the stored one puts the mutated bytes straight
        // in front of saxes, which is where the interesting failures are.
        const deflated = await docxOf(body([...Array(20)].map((_, i) => para(`paragraph ${i}`)).join('')))
        const stored = await zipWith(
            'word/document.xml',
            body([...Array(20)].map((_, i) => para(`paragraph ${i}`)).join('')),
            'STORE'
        )

        for (let i = 0; i < 1000; i++) {
            const source = i % 2 === 0 ? deflated : stored
            const mutated = Buffer.from(source)
            for (let f = 0; f < 1 + ((rnd() * 4) | 0); f++) {
                mutated[(rnd() * mutated.length) | 0] ^= 1 << ((rnd() * 8) | 0)
            }

            const started = Date.now()
            const r = await extractAttachment({ content: mutated, contentType: DOCX_TYPE })
            expect(Date.now() - started).toBeLessThan(HANDLER_TIMEOUT_MS)
            expect(['extracted', 'skipped', 'failed']).toContain(r.status)
            // The contract everywhere else in this library: text is present or the field is absent,
            // never ''. A mutation must not be able to produce the third state.
            expect(r.extraction).not.toBe('')
        }
    })
})

describe('docx — archives the handler never sees', () => {
    // Routing refuses this before the handler: ooxmlKind answers 'docx' only for an archive whose
    // central directory NAMES word/document.xml, so a renamed main part is unroutable. mammoth would
    // have followed _rels/.rels here — this is the one place we read the conventional path instead —
    // but routing already declined such a package before this change, so it is a pinned agreement
    // rather than a regression. If it ever needs fixing, the fix belongs in ooxmlKind.
    it('declines an archive whose main part is not at the conventional path', async () => {
        const content = await zipWith('word/document2.xml', body(para('unreachable')))
        const r = await extractAttachment({ content, contentType: DOCX_TYPE })
        expect(r.status).toBe('skipped')
        expect(r.reason).toMatch(/unsupported type/)
    })

    // The decompression preflight inflates every entry to measure it, so a corrupt deflate stream is
    // caught there — before the handler's own inflate can see it. The value of pinning it is that
    // the handler's inflate-error path is proven unreachable-but-safe: if a future preflight change
    // stops catching this, the reason string moves and this test says so.
    it('fails a corrupt deflate stream at the preflight, not in the handler', async () => {
        const good = await docxOf(body(para('x')))
        const corrupt = Buffer.from(good)
        const at = corrupt.indexOf('word/document.xml') + 'word/document.xml'.length + 10
        corrupt[at] ^= 0xff
        corrupt[at + 1] ^= 0xff
        corrupt[at + 2] ^= 0xff

        const r = await extractAttachment({ content: corrupt, contentType: DOCX_TYPE })
        expect(r.status).toBe('failed')
        expect(r.reason).toBe('malformed zip: unreadable compressed data')
    })
})
