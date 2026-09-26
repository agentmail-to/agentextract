# AgentExtract

Email quote & signature extractor. Given an email's `text` and/or `html`, it strips the quoted
history — the "On … wrote:" reply chain, forwarded blocks, `>`-quotes — along with trailing
boilerplate noise, leaving just the sender's new content. It's pure string/regex with no DOM parser
and no ML, so it runs cheaply on a Lambda Node.js runtime.

AgentExtract is maintained and used by [AgentMail](https://www.agentmail.to).

## Install

```sh
npm install agentextract
```

## Usage

```ts
import { extractEmailBody } from 'agentextract'

const { extractedText, extractedHtml } = extractEmailBody({
  text: 'Sounds good!\n\nOn Mon, Jun 1 Bob <bob@x.com> wrote:\n> old quoted message',
  html: '<div>Sounds good!</div><blockquote>old quoted message</blockquote>',
})
// extractedText === 'Sounds good!'
```

## API

`extractEmailBody({ text?, html? })` → `{ extractedText?, extractedHtml? }` is the main entry
point. It runs the quote cut and noise-strip on whichever fields you pass; a field that throws is
dropped from the result rather than failing the whole call.

The underlying stages are exported too, if you need them on their own:

- `extractNewContent(text)` / `extractFromHtml(html)` — the raw cut, before noise-stripping
- `stripNoise(text)` / `stripNoiseHtml(html)` — trailing-boilerplate removal
- `isTrueDsn(text)` — DSN / bounce detection

## Attachments

`extractAttachment(input)` pulls the text out of an email attachment — plain text, HTML, PDF, Word
(`.docx` + legacy `.doc`), and Excel (`.xlsx`). It never throws on bad or attacker-controlled input;
failures come back as a labeled `failed` or `skipped` status. Nested emails (`.eml`), images/OCR,
legacy `.xls`/`.ppt`, and archives are out of scope in this version and are skipped.

```ts
import { extractAttachment } from 'agentextract'

const result = await extractAttachment({
  content: buffer, // the raw attachment bytes
  filename: 'report.pdf',
  contentType: 'application/pdf',
})
// result.status === 'extracted'
// result.extraction === 'Q3 revenue …'
// result.truncated === false
```

The heavy parsers (`unpdf`, `exceljs`, `saxes`, …) are lazy-loaded per handler, so importing
`extractAttachment` costs nothing until you actually call it on a matching attachment. It's also
available on its own subpath — `import { extractAttachment } from 'agentextract/attachment'` — if
you want to reach it without touching the body-extraction entry point.

An optional second argument tunes the extraction. Both fields are optional and omitting them
reproduces the default behaviour exactly:

```ts
const result = await extractAttachment(input, {
  maxOutputChars: 50_000, // tighten the output cap; clamped to MAX_OUTPUT_CHARS, never loosened
  trailer: '\n[truncated — the source document continues past this point.]',
})
```

`reason` is a **code, not a sentence** — one of `too-large`, `expands-too-large`,
`unsupported-format`, `unrecognized`, `password-protected`, `unsupported-zip-feature`, `malformed`,
`wrong-document-shape`, `timed-out` or `internal`. It is set on every `skipped` and `failed` result
and on no `extracted` one, so it is safe to branch on. Two are worth knowing: `timed-out` is the only
failure worth retrying, and `internal` means one of our invariants tripped or a pinned dependency
moved — not a bad file, so retrying or re-requesting the attachment cannot help.

When a **complete** read yields no text, `extraction` is omitted and `emptyReason` says why:
`'no-text-layer'` means the document has content but none of it is text — a scan, a photographed
page — so OCR is the next step, while `'no-text-content'` means it was read and genuinely holds
nothing. Exactly one of `extraction` and `emptyReason` is present whenever `truncated` is false.
Both are claims about the whole document, so neither is reported for a truncated read: a result with
no text, no `emptyReason` and `truncated: true` means we stopped before finding text and cannot say
whether there is any. Only a handler that can *prove* the distinction reports `'no-text-layer'`;
today PDF (its first page paints an image) and `.docx` (the document references a picture). Where a
handler cannot prove EITHER — a multi-page PDF whose first page is blank, `.xlsx` and HTML, whose
images live where these readers do not look, and `.doc`, whose parser reports text or nothing and
nothing about the rest — **neither value is reported**, and that absence means exactly what it says:
we cannot tell. Treat a missing `emptyReason` as "may be worth OCR", never as "empty".

Password-protected **OOXML** files (`.docx`/`.xlsx`/`.pptx`, which Office wraps in an OLE container)
and password-protected PDFs are `skipped` with reason `password-protected`. Legacy `.doc` files
encrypted the old way — the `fEncrypted` FIB flag, with no `EncryptedPackage` stream — are **not**
detected and come back `failed` / `malformed`.

`trailer` is appended to `extraction` only when the text was actually cut, and sits **outside** cap
accounting — the cap bounds extracted text, so the returned string may exceed it by the trailer's
length. `result.truncated` reports the same fact programmatically, whether or not a trailer was
supplied, so a consumer never has to parse the text to find out.

### Resource limits & the safety boundary

`extractAttachment` never throws on hostile input, but its in-process guards are **soft** — they
bound what this library accumulates, not what the OS lets a parser allocate. Treat real CPU/OOM
containment as the host's job (a Lambda memory limit, or a terminable worker/subprocess). These
guards reduce blast radius; they are **not** a sandbox.

- **Input size** — attachments over `MAX_INPUT_BYTES` (10 MB) are skipped before any decode or parse.
- **Decompression** — OOXML (`.docx`/`.xlsx`) archives are stream-inflated and **measured**; one that
  actually expands past `MAX_UNCOMPRESSED_BYTES` (50 MB) is skipped before the parser loads. ZIP64 or
  out-of-range metadata is likewise `skipped`; malformed metadata is `failed`. This archive-wide
  preflight runs before the handler timeout starts. Its work is still bounded by the 10 MB input gate, the ZIP
  entry-count ceiling and the 50 MB inflate ceiling, but a maximal 65k-entry directory can spend time
  there that is not charged to `HANDLER_TIMEOUT_MS`.
- **XML nesting** — DOCX, streamed XLSX worksheets/control tables, and namespace-aware XLSX identity
  parsing refuse trees deeper than 256 elements. `saxes` namespace resolution scans the open-tag
  stack, so this converts otherwise-quadratic attacker-controlled nesting into a fixed bound while
  leaving room for legitimately nested Word tables. After identity fallback, ExcelJS may reparse
  `workbook.xml` and its relationships without namespace resolution; that linear path is not depth-
  capped. The parser work inside one chunk is synchronous, so the deadline cannot replace the
  structural ceiling where namespace mode is used.
- **Output** — extracted text is capped at `MAX_OUTPUT_CHARS` (250k), or lower via `maxOutputChars`.
  Cutting sets `truncated` on the result, so a partial extraction is never mistaken for a complete
  one. The PDF, `.docx` and `.xlsx` handlers apply the cap **incrementally** as they build and stop at
  the next format-safe boundary. DOCX may drain an enclosing text box or table whose output order is
  not known until it closes, but retains only cap-clipped variants while doing so. None materializes
  the full document. Only the HTML handler (html-to-text) returns a complete string that is then
  trimmed, so for that one the cap is **post-materialization** and peak memory follows the whole
  document.
- **Timeout** — `HANDLER_TIMEOUT_MS` (10 s) stops *awaiting* a slow async parse. It cannot cancel
  synchronous CPU already running inside a parser, so handlers that yield between units of work (PDF
  per page, `.docx` per inflate chunk, `.xlsx` per row) also check the deadline themselves and stop;
  the HTML, `.doc` and text handlers cannot. Their
  deadline sits `HANDLER_DEADLINE_MARGIN_MS` (1 s) *inside* the timeout, which is the window a handler
  has to return what it read — so a self-stopped parse comes back `extracted` with `truncated` set
  rather than being raced to `failed`. A single page, inflate chunk or row that overruns the margin
  still times out.
- **PDF** — page count and accumulated output are bounded (`MAX_PDF_PAGES`, `MAX_OUTPUT_CHARS`), but
  pdf.js's internal per-page decompression is **not** bounded in-library (no hook exists).
- **`.docx`** — `word/document.xml` is located in the archive's central directory, inflated on its
  own, and read with a streaming SAX parser (`saxes`) rather than loaded into a DOM. Parser memory
  tracks one inflate chunk plus bounded retained text; deferred text-box frames stop retaining new
  text once the output cap is crossed, tables retain at most two cap-clipped merge variants, and
  paragraphs reachable only through orphan text boxes are discarded instead of copied through each
  enclosing frame. The XML depth ceiling separately bounds the remaining frame overhead.
  Measured at a 1024 MB heap on a 45 MB `document.xml` inside a 3.65 MB archive, against the previous
  DOM-based reader:

  | concurrency | before | after |
  |---|---|---|
  | 1 | 812 ms / 607 MB | **68 ms / 175 MB** |
  | 2 | 1552 ms / 970 MB | **59 ms / 164 MB** |
  | 4 | 3512 ms / 1217 MB | **58 ms / 168 MB** |

  The old reader built 43.4M characters and kept 250k. The sharpest case is not the large archive: a
  0.41 MB attachment holding one 43 MB `<w:t>` peaked the old reader at 1270 MB, against 199 MB here.
  Two residual bounds, stated rather than glossed: peak is not independent of *input* size (the API
  takes a `Buffer`), and `saxes` buffers one text node whole, so a single enormous run still costs
  about twice its own size. Malformed XML is also stricter than before — a document the old reader
  silently half-read now comes back either `truncated` or `failed`.
- **`.docx` scope** — text comes from the document body *and* from the parts around it: footnotes,
  endnotes, comments, headers and footers are all extracted. They are appended after the body, in
  that order, as ordinary paragraphs — the order is the output cap's priority order rather than the
  document's reading order, which cannot be reconstructed (a footnote's reference sits inline while
  its body lives in another part, and a header is repeated per section rather than positioned once).
  A document that hits the cap therefore loses its page furniture before a footnote, and a footnote
  before a body paragraph. A header repeated across sections is emitted once. A malformed auxiliary
  part costs only its own text and sets `truncated`; only the body is load-bearing.
  **Table structure is not preserved**: each cell's paragraphs are emitted in reading order with the
  same blank-line separator as body paragraphs, so a 2×2 table is indistinguishable from four
  consecutive paragraphs. List bullets and numbers are dropped; the item text remains. Vertical-
  merge continuation cells are omitted as they were by the previous reader. These are compatibility
  targets rather than a claim of byte-for-byte identity for every malformed OOXML tree; intentional
  recovery differences are pinned in the DOCX streaming tests. A consumer reading an invoice or a
  contract should still know the column a figure sat in is gone.
- **`.xlsx`** — read row-by-row through `exceljs`'s streaming reader rather than loaded whole, so
  peak memory tracks the shared-string table plus one row instead of a live object per cell
  (measured: 294 MB → 171 MB, and 3.7x faster, on a 5 MB / 38 MB-uncompressed workbook). It is not
  independent of document size — a workbook with a very large string table still costs — so `.xlsx`
  remains the format most likely to reach the host's memory limit.
  That reader loses zip entries unless `xl/sharedStrings.xml` and `xl/_rels/workbook.xml.rels` are
  parsed before the first worksheet ([exceljs #2790](https://github.com/exceljs/exceljs/issues/2790),
  [#3064](https://github.com/exceljs/exceljs/issues/3064)), so the archive's entry order is rewritten
  in memory first; when either control part is legitimately absent, the reader-only copy supplies an
  empty equivalent so `exceljs` never falls into its temporary-file spool path. See
  `reorderForStreaming`. That rewrite is also what keeps the decompression
  budget binding on this format: the streaming reader walks local file headers, not the central
  directory the budget measured, and only the rebuilt copy is guaranteed to carry exactly the
  measured entries — so an archive that cannot be rebuilt is `failed`, never streamed as it arrived.
  Behind that, a workbook read to completion that yields fewer worksheets than the workbook declares
  returns `failed` rather than a partial workbook reported as `extracted` — silent partial output is
  the one outcome worth failing over, since a caller can retry a failure but cannot tell a truncated
  document from a complete one.
- **`.xlsx` sheet identity** — which sheets exist, in what order, and under what names comes from
  `xl/workbook.xml` and its relationships, not from the archive's layout. Sheets are emitted in tab
  order (which a dragged tab changes without moving any `sheetN.xml`), named as the workbook names
  them (`exceljs` matches relationship targets against a single spelling and silently fails to name a
  sheet whose target is written as an absolute package path), and worksheet parts the workbook does
  not reference are dropped rather than emitted as sheets of their own. Dropping is the narrow case,
  though — a part is an orphan only when the workbook placed every sheet it declared and none of them
  claimed this part's name. If any declaration could not be placed, or the part shares its name with
  one that was, it ships under a fallback name having lost only its tab position. What a declaration
  points at must be agreed on by the relationship's *type* and its target: neither is trusted alone,
  since the target's path is a string the producing application chose and a type contradicted by its
  target is not a resolution. Both OOXML flavours are recognized, Transitional and ISO Strict, and
  targets are matched case-insensitively as OPC requires. A workbook whose sheet list cannot be read
  at all — one from which not a single declaration parses, or whose workbook part is too large to
  read whole — falls back to the archive's own parts, in entry order, named `Sheet1`, `Sheet2`, ….
  The rule throughout is that a sheet may lose its position or its name, never its rows.

### Accepted dependency advisory

ExcelJS 4.4.0 requires `uuid@^8.3.0`, so consumers currently install `uuid@8.3.2`, which npm flags
under [GHSA-w5hq-g745-h8pq](https://github.com/advisories/GHSA-w5hq-g745-h8pq). The affected APIs are
UUID v3/v5/v6 writes into caller-provided buffers. ExcelJS loads UUID v4 for its conditional-formatting
write transform; AgentExtract only exercises the streaming read path, and never calls the affected
APIs. The advisory is therefore accepted until ExcelJS publishes a compatible dependency update.
An `overrides` entry would only alter this repository's root install and would not protect consumers,
so the published dependency graph is left honest rather than making local audit output misleading.

## What it does that off-the-shelf engines don't

Benchmarked against TalonJS and Mailgun Talon on a real corpus of ~41k messages (exact-match vs a
gpt-4o answer key, graded on the cut):

| Capability | AgentExtract | TalonJS | Mailgun Talon |
|---|---|---|---|
| Quote cutting (On-wrote / `>` / From: / Original-Msg) | yes | yes | yes |
| Glued "On…wrote:" orphan fix | **yes** | no | no |
| Foreign-language attributions (12+ langs) + Chinese/Arabic | **yes** | limited | limited |
| DSN / bounce keep-whole | **yes** | no | no |
| Forward keep-whole | **yes** | no | no |
| Inline-reply keep-whole | **yes** | no | partial |
| HTML trailing-signature reattach | **yes** | no | no |
| Noise-strip (mobile sigs / disclaimers / footers) | **yes** | no | no |

Per-feature exact-match % (AgentExtract / TalonJS / Talon): on_wrote 96/51/49, no_quote 100/99/99,
from_header 89/84/81, gt_quote 85/77/76, foreign_verb 87/61/26, inline 91/38/22, dsn 93/40/40.

## Tests

```sh
npm install
npm test
```
