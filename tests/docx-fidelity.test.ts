import { readFileSync, readdirSync } from 'node:fs'
import { join } from 'node:path'

import { describe, it, expect } from 'vitest'
import mammoth from 'mammoth'

import { extractAttachment } from '../attachment'

// ---------------------------------------------------------------------------
// .docx fidelity — our reader against mammoth, over a real Word corpus
// ---------------------------------------------------------------------------
// This library read .docx through mammoth.extractRawText until the streaming rewrite. That reader
// materialised an xmldom tree AND a document model before any cap could apply; the replacement walks
// word/document.xml with a SAX parser and stops at the cap. The output contract is unchanged, and
// this file is what says so.
//
// THE INTERLOCK. Each case pins ONE golden from TWO sides: the snapshot records what we produce, and
// mammoth is then asserted to produce the same string. Neither side can move alone —
//   - a regression in our reader fails the snapshot;
//   - `vitest -u`, which would launder that regression by rewriting the snapshot, then fails the
//     mammoth assertion, because mammoth still produces the old value.
// The only way to change behaviour is to add an ACCEPTED_DIVERGENCES entry with a written reason,
// which is a reviewed, committed edit. That is the mechanical form of "any structural regressions
// understood and accepted".
//
// NO skipIf ON THE ORACLE. `import mammoth` at the top is deliberate: if mammoth is missing the
// suite fails loudly rather than degrading to a snapshot pinned only by the code that generated it,
// which is self-referential and reports green. mammoth is a devDependency, so it is always present;
// retiring the oracle should be a reviewed deletion of this file, not a condition that skips itself.

// Vendored from mammoth 1.12.0 test/test-data — see docx-corpus/README.md for why they are copied
// rather than read out of node_modules, and what each one covers.
const CORPUS_DIR = join(process.cwd(), 'tests', 'fixtures', 'docx-corpus')
// The three fixtures the rest of the suite already uses. They are macOS textutil output, a dialect
// the vendored corpus (all Word) does not otherwise represent; elsewhere they are only asserted with
// `toContain`, which cannot catch a paragraph-break regression.
const FIXTURE_DIR = join(process.cwd(), 'tests', 'fixtures')
const LOCAL_FIXTURES = ['sample.docx', 'empty.docx', 'image.docx']

const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

// Documents where our reader deliberately diverges from mammoth. Each entry must say WHY, and holds
// what mammoth produces — the snapshot holds ours. An empty table means the port is exact.
const ACCEPTED_DIVERGENCES: Record<string, { reason: string; mammoth: string | null }> = {}

// The entry point omits `extraction` entirely rather than ever returning '' (a ran-but-empty
// extraction is `{ status: 'extracted' }` with no text). That collapse is a library-wide contract
// applied identically to whatever reader is installed, so the oracle gets it too — otherwise
// empty.docx, whose text is exactly '\n\n', reads as a divergence in every run.
const asExtraction = (text: string): string | null => (text.trim().length === 0 ? null : text)

const corpusFiles = () => readdirSync(CORPUS_DIR).filter((name) => name.endsWith('.docx')).sort()

const cases = [
    ...corpusFiles().map((name) => [name, join(CORPUS_DIR, name)] as const),
    ...LOCAL_FIXTURES.map((name) => [`fixtures/${name}`, join(FIXTURE_DIR, name)] as const),
]

describe('docx fidelity — our reader against mammoth', () => {
    it.each(cases)('%s', async (name, path) => {
        const content = readFileSync(path)

        const result = await extractAttachment({ content, contentType: DOCX_TYPE })
        // No corpus document is anywhere near MAX_OUTPUT_CHARS, so anything but a complete
        // extraction means a gate fired that should not have — worth failing on directly rather
        // than letting it show up as a confusing snapshot diff.
        expect(result.status).toBe('extracted')
        expect(result.truncated).toBe(false)

        const ours = result.extraction ?? null
        expect(ours).toMatchSnapshot()

        const { value } = await mammoth.extractRawText({ buffer: content })
        const theirs = asExtraction(value)

        const divergence = ACCEPTED_DIVERGENCES[name]
        if (divergence) {
            expect(theirs).toBe(divergence.mammoth)
            return
        }
        // Paragraph array first: vitest diffs it element-wise, so a lost or hoisted paragraph reads
        // as one changed line instead of a wall of escaped text. Then the exact string, which is
        // what catches trailing breaks and whitespace the split would swallow.
        expect((theirs ?? '').split('\n\n')).toEqual((ours ?? '').split('\n\n'))
        expect(theirs).toBe(ours)
    })

    // A fixture added without a golden, or deleted while its golden lingers, are both silent holes:
    // it.each simply enumerates fewer cases and the suite still passes. Counting them here is what
    // makes the corpus itself part of the contract.
    it('runs every corpus document, and only documents that exist', () => {
        expect(corpusFiles()).toHaveLength(17)
        expect(cases).toHaveLength(20)
        for (const [, path] of cases) expect(readFileSync(path).length).toBeGreaterThan(0)
    })
})
