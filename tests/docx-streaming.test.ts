import { describe, it, expect, vi } from 'vitest'
import JSZip from 'jszip'

import { extractAttachment, HANDLER_TIMEOUT_MS, MAX_OUTPUT_CHARS } from '../attachment'

// ---------------------------------------------------------------------------
// .docx streaming reader — element semantics, caps, and the deadline
// ---------------------------------------------------------------------------
// tests/docx-fidelity.test.ts proves the reader agrees with mammoth on a real Word corpus. This file
// covers what that corpus cannot: the corpus contains no w:tab, no w:br, no w:cr, no tracked
// changes, no fields, no w:sdt and no w:hyperlink at all — real Word documents simply don't carry
// every construct — and no corpus document is anywhere near the output cap or the deadline.
//
// THE CONTRACT BEING PINNED, in one place so the tests below can just reference it (mammoth's
// raw-text.js is 13 lines, and each of these was measured against it, not read off the spec):
//
//   w:t          -> its text, verbatim, never trimmed
//   w:tab        -> '\t';  w:noBreakHyphen -> U+2011;  w:softHyphen -> U+00AD
//   w:p          -> its children, then TWO newlines. Including the last paragraph, so a document
//                   ends with '\n\n'.
//   w:br, w:cr   -> NOTHING. Two different mechanisms in mammoth (br is a document node with no raw
//                   mapping, cr isn't whitelisted at all) reaching the same result.
//   anything else with no handler -> the WHOLE SUBTREE disappears, children included. This is the
//                   one that makes a naive "emit every w:t" reader wrong rather than merely different.
//   bare text outside a w:t -> dropped.
//   tables       -> one paragraph per cell paragraph. No tabs, no row markers: a 2x2 table is
//                   byte-identical to four consecutive paragraphs.
//
// Everything here builds its own archive, because the point is to reach constructs no real document
// in the corpus contains.

const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

const NS = [
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
    'xmlns:v="urn:schemas-microsoft-com:vml"',
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
].join(' ')

// JSZip output satisfies both decompression-budget invariants, so a fixture built here reaches the
// handler rather than being turned away in front of it. STORE covers the stored-entry path.
const docxFrom = async (documentXml: string, compression: 'DEFLATE' | 'STORE' = 'DEFLATE') => {
    const zip = new JSZip()
    zip.file(
        '[Content_Types].xml',
        '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'
    )
    zip.file(
        '_rels/.rels',
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    )
    zip.file('word/document.xml', documentXml)
    return Buffer.from(await zip.generateAsync({ type: 'nodebuffer', compression }))
}

const buildDocx = (body: string, compression: 'DEFLATE' | 'STORE' = 'DEFLATE') =>
    docxFrom(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}</w:body></w:document>`, compression)

const extract = async (body: string, options?: { maxOutputChars?: number }) =>
    extractAttachment({ content: await buildDocx(body), contentType: DOCX_TYPE }, options)

const para = (inner: string) => `<w:p><w:r>${inner}</w:r></w:p>`
const text = (s: string) => para(`<w:t>${s}</w:t>`)

describe('docx — the whitelist drops unrecognised subtrees', () => {
    // mc:Choice is the DrawingML branch and mc:Fallback the VML one; Word emits both for the same
    // text box. Taking both would DOUBLE the text, and taking Choice instead of Fallback would take
    // the branch mammoth discards. Corpus-covered by text-box.docx, kept here because a 20-document
    // snapshot failure names the document, not the branch.
    it('takes only the mc:Fallback branch of mc:AlternateContent', async () => {
        const r = await extract(
            `<w:p><w:r><mc:AlternateContent>` +
                `<mc:Choice Requires="wps"><w:drawing><wps:txbx><w:txbxContent>${text('CHOICE')}</w:txbxContent></wps:txbx></w:drawing></mc:Choice>` +
                `<mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent>${text('FALLBACK')}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>` +
                `</mc:AlternateContent></w:r></w:p>`
        )
        expect(r.extraction).toContain('FALLBACK')
        expect(r.extraction).not.toContain('CHOICE')
    })

    it('takes only the first selected child from sdt and AlternateContent', async () => {
        const sdt = await extract(
            `<w:p><w:sdt>` +
                `<w:sdtContent><w:r><w:t>FIRST</w:t></w:r></w:sdtContent>` +
                `<w:sdtContent><w:r><w:t>SECOND</w:t></w:r></w:sdtContent>` +
                `</w:sdt></w:p>`
        )
        const alternate = await extract(
            `<w:p><w:r><mc:AlternateContent>` +
                `<mc:Fallback><w:r><w:t>FIRST</w:t></w:r></mc:Fallback>` +
                `<mc:Fallback><w:r><w:t>SECOND</w:t></w:r></mc:Fallback>` +
                `</mc:AlternateContent></w:r></w:p>`
        )
        expect(sdt.extraction).toBe('FIRST\n\n')
        expect(alternate.extraction).toBe('FIRST\n\n')
    })

    // A w:pict's text is hoisted out and reinserted as a SIBLING of the enclosing paragraph, so it
    // lands AFTER that paragraph's break rather than inside it. Emitting it inline would glue two
    // words into one nonword token — 'HelloBoxed' — which is exactly the shape an LLM reads wrong.
    it('emits w:pict content after the enclosing paragraph, not inside it', async () => {
        const r = await extract(
            `<w:p><w:r><w:t>Hello</w:t></w:r><w:r><w:pict><v:shape><v:textbox><w:txbxContent>${text('Boxed')}</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>`
        )
        expect(r.extraction).toBe('Hello\n\nBoxed\n\n')
    })

    // The other half, and the counter-intuitive one: an extra that never reaches a w:p is silently
    // discarded. It looks like a bug worth "fixing", so it is pinned as the contract instead.
    it('drops a w:pict that sits outside any paragraph', async () => {
        const r = await extract(
            `<w:pict><v:shape><v:textbox><w:txbxContent>${text('Orphan')}</w:txbxContent></v:textbox></v:shape></w:pict>${text('after')}`
        )
        expect(r.extraction).toBe('after\n\n')
    })

    it('does not charge a large orphan picture to the output cap', async () => {
        const hidden = Array.from({ length: 4_000 }, () => text('x'.repeat(50))).join('')
        const r = await extract(
            `<w:pict><v:shape><v:textbox><w:txbxContent>${hidden}</w:txbxContent></v:textbox></v:shape></w:pict>${text('visible')}`,
            { maxOutputChars: 1_000 }
        )
        expect(r).toMatchObject({ status: 'extracted', extraction: 'visible\n\n', truncated: false })
    })

    // Over-tightening the whitelist loses real body text silently. None of these containers appears
    // anywhere in the fidelity corpus, so nothing else would catch it.
    it.each([
        ['w:ins', `<w:p><w:ins><w:r><w:t>KEEP</w:t></w:r></w:ins></w:p>`],
        ['w:smartTag', `<w:p><w:smartTag><w:r><w:t>KEEP</w:t></w:r></w:smartTag></w:p>`],
        ['w:hyperlink', `<w:p><w:hyperlink r:id="rId1"><w:r><w:t>KEEP</w:t></w:r></w:hyperlink></w:p>`],
        ['w:sdt', `<w:p><w:sdt><w:sdtPr><w:alias w:val="DROP"/></w:sdtPr><w:sdtContent><w:r><w:t>KEEP</w:t></w:r></w:sdtContent></w:sdt></w:p>`],
        ['w:object', `<w:p><w:r><w:object><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>KEEP</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:object></w:r></w:p>`],
    ])('recurses into %s', async (_name, body) => {
        const r = await extract(body)
        expect(r.extraction).toContain('KEEP')
        expect(r.extraction).not.toContain('DROP')
    })

    // The load-bearing one. A reader that recursed into unknown elements — the obvious way to write
    // this — would pass every other test in this file and still over-extract on real documents.
    it('drops an unrecognised element WITH its children rather than recursing', async () => {
        const r = await extract(`${text('kept')}<w:zzz>${text('hidden')}</w:zzz>`)
        // Not just the text: no stray '\n\n' from the paragraph inside the dropped subtree either.
        expect(r.extraction).toBe('kept\n\n')
    })

    // Plain-object prototype names are valid XML element names. They are not entries in the literal
    // map: looking them up without an own-property guard used to append native function source.
    it('treats Object.prototype names as unknown elements, never literals', async () => {
        const r = await extract(`${text('before')}<constructor/><toString/><x:wrap xmlns:x="constructor"/>${text('after')}`)
        expect(r.extraction).toBe('before\n\nafter\n\n')
        expect(r.extraction).not.toMatch(/function|native code/)
    })

    // Character data lives only inside w:t. Whitespace between elements is common in
    // pretty-printed markup, so a reader that accumulated all text would leak indentation.
    it('drops bare character data outside a w:t', async () => {
        const r = await extract('<w:p><w:r>stray<w:t>kept</w:t>more</w:r></w:p>')
        expect(r.extraction).toBe('kept\n\n')
    })

    // Field instructions are markup, not content. Leaking them puts ` HYPERLINK "http://..." ` into
    // the body of an extract that an agent then reads as prose.
    it('suppresses w:instrText across a complex field, keeping only the display run', async () => {
        const r = await extract(
            `<w:p>` +
                `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
                `<w:r><w:instrText xml:space="preserve"> HYPERLINK "http://example.com" </w:instrText></w:r>` +
                `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
                `<w:r><w:t>link</w:t></w:r>` +
                `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
                `</w:p>`
        )
        expect(r.extraction).toBe('link\n\n')
    })

    // KNOWN mammoth DEFECT, MATCHED ON PURPOSE. w:fldSimple has no handler in mammoth and is not on
    // its ignore list, so the whole subtree — including the cached result a reader sees on screen —
    // is dropped. Cross-references, page numbers, DOCPROPERTY and merge fields all land here. This
    // change is a pure port, so it keeps the defect; recovering it is one entry in DOCX_CONTAINERS
    // and belongs in its own change, once the corpus says how many documents it affects.
    it('drops a w:fldSimple subtree, matching mammoth rather than fixing it here', async () => {
        const r = await extract(
            `<w:p><w:r><w:t>page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>7</w:t></w:r></w:fldSimple></w:p>`
        )
        expect(r.extraction).toBe('page \n\n')
    })

    // Deleted text must not resurface in an extract; inserted text must.
    it('drops tracked deletions and keeps tracked insertions', async () => {
        const r = await extract(
            `<w:p>` +
                `<w:del><w:r><w:delText>GONE</w:delText></w:r></w:del>` +
                `<w:ins><w:r><w:t>ADDED</w:t></w:r></w:ins>` +
                `</w:p>`
        )
        expect(r.extraction).toBe('ADDED\n\n')
    })

    // The subtlest paragraph rule. A deleted paragraph MARK means the break is gone, not the text:
    // the content runs into the next paragraph. Emitting the usual '\n\n' would invent a break the
    // author deleted, splitting one sentence into two.
    it('merges a paragraph whose mark was deleted into the next one', async () => {
        const r = await extract(
            `<w:p><w:pPr><w:rPr><w:del/></w:rPr></w:pPr><w:r><w:t>first half </w:t></w:r></w:p>${text('second half')}`
        )
        expect(r.extraction).toBe('first half second half\n\n')
    })

    it('keeps a deleted-mark paragraph text box after the following merged paragraph', async () => {
        const deleted =
            `<w:p><w:pPr><w:rPr><w:del/></w:rPr></w:pPr>` +
            `<w:r><w:t>one</w:t></w:r>` +
            `<w:r><w:pict><v:shape><v:textbox><w:txbxContent>${text('boxed')}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>` +
            `</w:p>`
        const r = await extract(`${deleted}${text('two')}`)
        expect(r.extraction).toBe('onetwo\n\nboxed\n\n')
        expect(r.truncated).toBe(false)
    })

    it('does not let a nested text-box paragraph steal a deleted paragraph extra', async () => {
        const box = (value: string) =>
            `<w:r><w:pict><v:shape><v:textbox><w:txbxContent>${text(value)}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`
        const deleted =
            `<w:p><w:pPr><w:rPr><w:del/></w:rPr></w:pPr>` +
            `<w:r><w:t>DELTEXT</w:t></w:r>${box('DEL-BOX')}</w:p>`
        const following = `<w:p><w:r><w:t>NEXT</w:t></w:r>${box('NEXT-BOX')}</w:p>`

        const r = await extract(deleted + following)
        expect(r.extraction).toBe('DELTEXTNEXT\n\nDEL-BOX\n\nNEXT-BOX\n\n')
    })

    // ACCEPTED DIVERGENCE. mammoth stashes a deleted-mark paragraph for the next paragraph and
    // drops that text when there is no next one. The streaming reader emits as it goes and keeps the
    // trailing text. This is the safer direction for attachment extraction, but it is not exact
    // mammoth fidelity and must stay visible as such.
    it('keeps a final deleted-mark paragraph that mammoth drops', async () => {
        const deleted =
            '<w:p><w:pPr><w:rPr><w:del/></w:rPr></w:pPr><w:r><w:t>TRAILING</w:t></w:r></w:p>'
        const r = await extract(`${text('before')}${deleted}`)
        expect(r.extraction).toBe('before\n\nTRAILING')
    })

    // Same marker, one level out: a deleted table row takes the whole row with it.
    it('drops a table row marked deleted, keeping the rows around it', async () => {
        const r = await extract(
            `<w:tbl>` +
                `<w:tr><w:trPr><w:del/></w:trPr><w:tc>${text('dropped cell')}</w:tc></w:tr>` +
                `<w:tr><w:tc>${text('live cell')}</w:tc></w:tr>` +
                `</w:tbl>`
        )
        expect(r.extraction).toBe('live cell\n\n')
    })
})

describe('docx — leaf and character contract', () => {
    // TWO newlines, and a trailing one on the last paragraph. A reader that joined paragraphs with
    // '\n\n' instead of suffixing each would differ only in the final break — invisible in a
    // `toContain` assertion, and different from mammoth on every document ever extracted.
    it('suffixes every paragraph with two newlines, including the last', async () => {
        expect((await extract(`${text('a')}${text('b')}`)).extraction).toBe('a\n\nb\n\n')
    })

    // mammoth selects the document's w:body child. Paragraphs beside it are malformed package
    // debris, not body content, and must not be prepended or appended to the extraction.
    it('extracts only content inside the first direct w:body', async () => {
        const content = await docxFrom(
            `<?xml version="1.0"?><w:document ${NS}>${text('BEFORE')}<w:body>${text('INSIDE')}</w:body>${text('AFTER')}</w:document>`
        )
        const r = await extractAttachment({ content, contentType: DOCX_TYPE })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBe('INSIDE\n\n')
    })

    it.each([
        ['w:body', `<w:body>${text('HIDDEN')}</w:body>`],
        ['w:document', `<w:document>${text('HIDDEN')}</w:document>`],
    ])('drops a nested %s subtree inside the real body', async (_name, nested) => {
        const r = await extract(`${text('before')}${nested}${text('after')}`)
        expect(r.extraction).toBe('before\n\nafter\n\n')
    })

    // An empty paragraph is exactly '\n\n', which the entry point then trims to nothing rather than
    // storing ''. Both halves matter: the break is produced, and the result still omits `extraction`.
    it('produces only breaks for an empty paragraph, and omits the extraction', async () => {
        const r = await extract('<w:p/>')
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBeUndefined()
    })

    // The "obvious improvement" that would silently change every document ever extracted. mammoth
    // emits nothing for either — w:br is a document node with no raw-text mapping, w:cr is not
    // whitelisted at all — so a line break inside a paragraph simply is not represented.
    it('emits nothing for w:br or w:cr', async () => {
        const r = await extract(para('<w:t>a</w:t><w:br/><w:cr/><w:br w:type="page"/><w:t>b</w:t>'))
        expect(r.extraction).toBe('ab\n\n')
    })

    it('maps w:tab, w:noBreakHyphen and w:softHyphen to their literals', async () => {
        const r = await extract(para('<w:t>a</w:t><w:tab/><w:t>b</w:t><w:noBreakHyphen/><w:softHyphen/><w:t>c</w:t>'))
        expect(r.extraction).toBe(`a\tb‑­c\n\n`)
    })

    it('maps w:sym through the same dingbat table and F0 fallback as mammoth', async () => {
        const r = await extract(para('<w:t>a</w:t><w:sym w:font="Wingdings" w:char="F0FC"/><w:t>b</w:t>'))
        expect(r.extraction).toBe('a✓b\n\n')
    })

    // Leading and trailing spaces inside a w:t are content. A "helpful" trim would silently join
    // words across run boundaries, which is how formatted text is stored.
    it('never trims whitespace inside or around a run', async () => {
        const r = await extract(para('<w:t xml:space="preserve">  spaced  </w:t><w:t xml:space="preserve">out  </w:t>'))
        expect(r.extraction).toBe('  spaced  out  \n\n')
    })

    // Entities decode; CDATA does NOT. The second is counter-intuitive: mammoth's DOM walk only
    // handles element and text nodes, so a CDATA section vanishes. We reproduce it by NOT wiring
    // saxes' cdata event — i.e. by absence, which is exactly the kind of thing a later reader
    // "fixes" without realising it is a contract.
    it('decodes XML entities but drops CDATA, as mammoth does', async () => {
        expect((await extract(text('a &amp; b &lt;c&gt;'))).extraction).toBe('a & b <c>\n\n')
        expect((await extract(para('<w:t><![CDATA[cdata]]></w:t>'))).extraction).toBeUndefined()
    })

    // xmldom normalised all four to '\n' before mammoth ever saw the markup. saxes does the
    // XML-standard \r\n and \r itself but leaves NEL (U+0085) and LINE SEPARATOR (U+2028) alone, so
    // without an explicit normalisation this is the one test that fails.
    it('normalises CRLF, CR, NEL and LINE SEPARATOR to a plain newline', async () => {
        const CR = String.fromCharCode(13)
        const r = await extract(text(`a${CR}\nb${CR}cd e`))
        expect(r.extraction).toBe('a\nb\nc\nd\ne\n\n')
    })

    // Catches decoding each inflate chunk independently. A chunk boundary lands mid-sequence in any
    // document with non-ASCII text, and chunk.toString('utf8') turns that one character into two
    // U+FFFD. Deliberately not tuned to the chunk size: a long run of 4-byte characters puts a
    // boundary inside one of them whatever the chunking turns out to be.
    it('decodes multi-byte UTF-8 across inflate chunk boundaries', async () => {
        const emoji = '\u{1F600}'.repeat(20_000) // 80 KB of 4-byte sequences, several chunks deep
        const r = await extract(text(emoji))
        expect(r.extraction).not.toContain('�')
        expect(r.extraction).toBe(`${emoji}\n\n`)
    })

    // The entry point's surrogate guard only fires when the handler OVERSHOOTS the cap; a handler
    // that stopped exactly on it would slip a lone half through, which serialises as U+FFFD. The
    // existing cap tests go through text/plain and cannot reach this path.
    it('never splits a surrogate pair when the cap lands mid-character', async () => {
        const r = await extract(text('\u{1F600}'.repeat(4_000)), { maxOutputChars: 1_001 })
        expect(r.truncated).toBe(true)
        expect(r.extraction).not.toContain('�')
        expect([...(r.extraction ?? '')].every((c) => c === '\u{1F600}')).toBe(true)
    })
})

// Undeclared namespace prefixes. saxes fails the whole parse on one — for ATTRIBUTES as well as
// elements — where the DOM parser this replaced shrugged. Under the error policy that turns into
// `extracted` + `truncated`, i.e. everything after the offending element is silently gone. These are
// not exotic: Word 2013 and later stamp w15:paraId on EVERY w:p, with w16cid:durableId beside it, so
// a document carrying those without their xmlns loses everything after its first paragraph.
//
// An allowlist of known prefixes cannot close that — there is always another prefix, and the first
// version of this reader proved the point by naming wne: in a comment while omitting it from the
// map. resolvePrefix closes the open class instead: anything unclaimed resolves to a sentinel URI
// that OOXML_PREFIXES does not map, which is exactly the shape mammoth gives an unmapped namespace.
describe('docx — an undeclared prefix must not end the parse', () => {
    const withAttr = (attr: string) =>
        `${text('before')}<w:p ${attr}><w:r><w:t>middle</w:t></w:r></w:p>${text('after')}`

    // Attributes, which is the common case and the one that reads as impossible until you hit it:
    // the element is perfectly well-formed WordprocessingML and only its attribute is unbound.
    it.each([
        ['w15:paraId', 'w15:paraId="12AB34CD"'],
        ['w16cid:durableId', 'w16cid:durableId="99"'],
        ['a prefix no allowlist would think to carry', 'zz:whatever="1"'],
    ])('reads the whole document despite an undeclared %s attribute', async (_name, attr) => {
        const r = await extract(withAttr(attr))
        expect(r.extraction).toBe('before\n\nmiddle\n\nafter\n\n')
        expect(r.truncated).toBe(false)
    })

    // Elements, where the prefix being unmapped must still mean "drop this subtree" — the fix must
    // not turn an unknown namespace into a recursed-into container.
    it('drops an undeclared-prefix element with its children, and keeps reading after it', async () => {
        const r = await extract(`${text('before')}<zz:wrap>${text('hidden')}</zz:wrap>${text('after')}`)
        expect(r.extraction).toBe('before\n\nafter\n\n')
        expect(r.truncated).toBe(false)
    })

    // Why the assumed-prefix map still has to exist alongside the sentinel: `w` is one of the three
    // prefixes OOXML_PREFIXES actually maps, so resolving an undeclared one to the sentinel would
    // drop the entire document body. mammoth throws outright on this document; we recover it.
    it('still recovers a document that does not declare w: at all', async () => {
        const content = await docxFrom(
            `<?xml version="1.0"?><w:document><w:body>${text('recovered')}</w:body></w:document>`
        )
        const r = await extractAttachment({ content, contentType: DOCX_TYPE })
        expect(r.extraction).toBe('recovered\n\n')
    })

    // ...and the assumed map must never win over a real declaration, or a strict-format document
    // would be read against the transitional URI.
    it('lets an in-document xmlns shadow the assumed prefix', async () => {
        const content = await docxFrom(
            `<?xml version="1.0"?><w:document xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"><w:body>${text('strict')}</w:body></w:document>`
        )
        const r = await extractAttachment({ content, contentType: DOCX_TYPE })
        expect(r.extraction).toBe('strict\n\n')
    })
})

describe('docx — the cap and the deadline stop the read', () => {
    // Long enough to cross the default cap, with a marker last so "did it stop reading" is
    // observable from outside rather than inferred from a length.
    const oversized = (marker = 'TRAILING-MARKER') =>
        `${Array.from({ length: 6_000 }, (_, i) => text(`paragraph ${i} ${'x'.repeat(50)}`)).join('')}${text(marker)}`

    it('reports truncated when the output cap stops it, and not when it does not', async () => {
        const cut = await extract(oversized())
        expect(cut.status).toBe('extracted')
        expect(cut.truncated).toBe(true)
        expect(cut.extraction?.length).toBeLessThanOrEqual(MAX_OUTPUT_CHARS)

        const whole = await extract(text('short document'))
        expect(whole.truncated).toBe(false)
        expect(whole.extraction).toBe('short document\n\n')
    })

    // The claim is that the cap stops the READ, not that it trims the result. If it only trimmed,
    // the marker would still have been parsed — and on a ceiling-sized document that is the whole
    // difference between 68 ms and 812 ms, and between 175 MB and 607 MB.
    it('stops parsing at the cap rather than reading on and trimming', async () => {
        const started = Date.now()
        const r = await extract(oversized())
        expect(r.extraction).not.toContain('TRAILING-MARKER')
        expect(Date.now() - started).toBeLessThan(HANDLER_TIMEOUT_MS)
    })

    it('omits the extraction entirely, never returning an empty string, at a zero cap', async () => {
        const r = await extract(text('some real text'), { maxOutputChars: 0 })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBeUndefined()
        expect(r.truncated).toBe(true)
    })

    // A single paragraph of many runs, deliberately: with the deadline checked per PARAGRAPH rather
    // than per chunk, a one-paragraph document would spin unbudgeted, which is the docx analogue of
    // the xlsx sheet-of-blank-rows case.
    it('reports truncated when the deadline stops it', async () => {
        const body = `<w:p>${Array.from({ length: 4_000 }, (_, i) => `<w:r><w:t>run ${i} </w:t></w:r>`).join('')}</w:p>`
        const content = await buildDocx(body)

        // extractAttachment's FIRST Date.now() sets the deadline; every later call lands past it, so
        // the read stops on its first check. Deliberately blunt — zlib and saxes may read the clock
        // themselves, so a mock letting N chunks through would depend on how often they do.
        const base = Date.now()
        let calls = 0
        const clock = vi.spyOn(Date, 'now').mockImplementation(() => {
            calls += 1
            return calls === 1 ? base : base + HANDLER_TIMEOUT_MS + 1
        })
        const stopped = await extractAttachment({ content, contentType: DOCX_TYPE })
        clock.mockRestore()

        expect(stopped.status).toBe('extracted') // a deadline stop is not a failure
        expect(stopped.truncated).toBe(true)

        // Control: the same bytes on a real clock are complete and NOT flagged — without this the
        // assertion above would also pass if `truncated` were hardcoded true.
        const whole = await extractAttachment({ content, contentType: DOCX_TYPE })
        expect(whole.truncated).toBe(false)
        expect(whole.extraction).toContain('run 3999')
    })

    // The memory claim in one assertion: a single text node far larger than the cap must not be
    // accumulated whole before the cap applies. saxes buffers one text node, so this is also the
    // bound on that claim — it is the shape that costs the most, and it still stops at the cap.
    it('caps a document that is one enormous run', async () => {
        const r = await extract(text('lorem ipsum dolor sit amet '.repeat(200_000))) // ~5.4M chars
        expect(r.status).toBe('extracted')
        expect(r.truncated).toBe(true)
        expect(r.extraction?.length).toBeLessThanOrEqual(MAX_OUTPUT_CHARS)
        expect(r.extraction?.startsWith('lorem ipsum')).toBe(true)
    })

    // A text box is hoisted after its enclosing paragraph only when its w:pict closes. If extraction
    // stops while that frame is open, returning its value inline would produce `HelloBox...`, which
    // is not a prefix of the complete `Hello\n\nBox...` output.
    it('keeps a truncated text-box extraction as a prefix of complete output', async () => {
        const box = 'boxed '.repeat(30_000)
        const body =
            `<w:p><w:r><w:t>Hello</w:t></w:r><w:r><w:pict><v:shape><v:textbox><w:txbxContent>` +
            `${text(box)}</w:txbxContent></v:textbox></v:shape></w:pict></w:r><w:r><w:t>After</w:t></w:r></w:p>`
        const whole = await extract(body, { maxOutputChars: MAX_OUTPUT_CHARS })
        const partial = await extract(body, { maxOutputChars: 100 })

        expect(whole.truncated).toBe(false)
        expect(partial.truncated).toBe(true)
        expect(whole.extraction!.startsWith(partial.extraction!)).toBe(true)
    })

    it('drains the enclosing paragraph after the cap so text-box-only output stays useful', async () => {
        const boxParagraphs = Array.from({ length: 500 }, (_, i) => text(`boxed-${String(i).padStart(3, '0')}`)).join('')
        const body =
            `<w:p><w:r><w:t>Intro</w:t></w:r><w:r><w:pict><v:shape><v:textbox><w:txbxContent>` +
            `${boxParagraphs}</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>`
        const r = await extract(body, { maxOutputChars: 100 })

        expect(r).toMatchObject({ status: 'extracted', truncated: true })
        expect(r.extraction).toContain('Intro\n\nboxed-000')
        expect(r.extraction!.length).toBe(100)
    })

    // saxes resolves namespaces by scanning its open-tag stack. Without an explicit depth ceiling,
    // a tiny deeply nested document turns that into quadratic synchronous CPU and blocks the timer
    // that is supposed to contain it.
    it('refuses pathological XML nesting before namespace resolution becomes quadratic', async () => {
        const nested = '<w:p>'.repeat(300) + '</w:p>'.repeat(300)
        const r = await extract(nested)
        expect(r.status).toBe('failed')
        expect(r.reason).toMatch(/XML nesting exceeds 256 elements/)
    })
})

describe('docx — archive-level reads', () => {
    // Some producers store small parts uncompressed. Full fidelity is required on that path, and it
    // is a different branch of the handler (no inflate stream at all).
    it('reads a stored, uncompressed word/document.xml', async () => {
        const content = await buildDocx(`${text('stored entry')}`, 'STORE')
        const r = await extractAttachment({ content, contentType: DOCX_TYPE })
        expect(r.status).toBe('extracted')
        expect(r.extraction).toBe('stored entry\n\n')
    })

    // Each table level contributes w:tbl/w:tr/w:tc. Twenty levels put the deepest text at XML depth
    // 65: valid structure that the former 64-element guard rejected before the parser reached it.
    it('reads text through twenty nested tables', async () => {
        let nested = text('deeply nested')
        for (let i = 0; i < 20; i++) nested = `<w:tbl><w:tr><w:tc>${nested}</w:tc></w:tr></w:tbl>`
        const r = await extract(nested)
        expect(r.status).toBe('extracted')
        expect(r.truncated).toBe(false)
        expect(r.extraction).toBe('deeply nested\n\n')
    })

    // Stored entries used to arrive as one attachment-sized chunk, so neither cap nor deadline was
    // checked again until the entire XML part had been parsed. A clock that expires between 16 KiB
    // slices makes the incremental stop observable without relying on wall-clock timing.
    it('checks the deadline between chunks of a stored document part', async () => {
        const body = `${text('prefix')}${Array.from({ length: 5_000 }, () => text('stored paragraph')).join('')}${text('TRAILING-MARKER')}`
        const content = await buildDocx(body, 'STORE')
        const base = Date.now()
        let calls = 0
        const clock = vi.spyOn(Date, 'now').mockImplementation(() => (++calls <= 2 ? base : base + HANDLER_TIMEOUT_MS + 1))
        const r = await extractAttachment({ content, contentType: DOCX_TYPE })
        clock.mockRestore()

        expect(r.status).toBe('extracted')
        expect(r.truncated).toBe(true)
        expect(r.extraction).toContain('prefix')
        expect((r.extraction ?? '').includes('TRAILING-MARKER')).toBe(false)
    })

    // A table flattens to one paragraph per cell paragraph: no tabs, no row markers, no column
    // structure. Documented as a limit rather than a bug — it is what mammoth did, and what the
    // README now states — but pinned here because it is the thing most likely to be "improved"
    // without realising it changes every extracted invoice.
    it('flattens a table to bare paragraphs, losing row and column structure', async () => {
        const cell = (s: string) => `<w:tc>${text(s)}</w:tc>`
        const r = await extract(
            `<w:tbl><w:tr>${cell('r1c1')}${cell('r1c2')}</w:tr><w:tr>${cell('r2c1')}${cell('r2c2')}</w:tr></w:tbl>`
        )
        expect(r.extraction).toBe('r1c1\n\nr1c2\n\nr2c1\n\nr2c2\n\n')
        // ...i.e. indistinguishable from four consecutive paragraphs:
        expect(r.extraction).toBe((await extract(['r1c1', 'r1c2', 'r2c1', 'r2c2'].map(text).join(''))).extraction)
    })
})
