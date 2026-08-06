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

The heavy parsers (`unpdf`, `mammoth`, `exceljs`, …) are lazy-loaded per handler, so importing
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
  actually expands past `MAX_UNCOMPRESSED_BYTES` (50 MB) is skipped before the parser loads. Malformed
  or ZIP64 metadata is treated as over-budget (fail-closed), not trusted.
- **Output** — extracted text is capped at `MAX_OUTPUT_CHARS` (250k), or lower via `maxOutputChars`.
  Cutting sets `truncated` on the result, so a partial extraction is never mistaken for a complete
  one. The PDF and `.xlsx` handlers apply the cap **incrementally** as they build — and stop reading
  the document once they reach it — so neither ever materializes in full. The `.docx` (mammoth) and
  HTML (html-to-text) handlers return a complete string that is then trimmed, so for those the cap is
  **post-materialization** and peak memory follows the whole document.
- **Timeout** — `HANDLER_TIMEOUT_MS` (10 s) stops *awaiting* a slow async parse. It cannot cancel
  synchronous CPU already running inside a parser, so handlers that yield between units of work (PDF
  per page, `.xlsx` per row) also check the deadline themselves and stop; the others cannot. Their
  deadline sits `HANDLER_DEADLINE_MARGIN_MS` (1 s) *inside* the timeout, which is the window a handler
  has to return what it read — so a self-stopped parse comes back `extracted` with `truncated` set
  rather than being raced to `failed`. A single page or row that overruns the margin still times out.
- **PDF** — page count and accumulated output are bounded (`MAX_PDF_PAGES`, `MAX_OUTPUT_CHARS`), but
  pdf.js's internal per-page decompression is **not** bounded in-library (no hook exists).
- **`.xlsx`** — read row-by-row through `exceljs`'s streaming reader rather than loaded whole, so
  peak memory tracks the shared-string table plus one row instead of a live object per cell
  (measured: 294 MB → 171 MB, and 3.7x faster, on a 5 MB / 38 MB-uncompressed workbook). It is not
  independent of document size — a workbook with a very large string table still costs — so `.xlsx`
  remains the format most likely to reach the host's memory limit.
  That reader loses zip entries unless `xl/sharedStrings.xml` and `xl/_rels/workbook.xml.rels` are
  parsed before the first worksheet ([exceljs #2790](https://github.com/exceljs/exceljs/issues/2790),
  [#3064](https://github.com/exceljs/exceljs/issues/3064)), so the archive's entry order is rewritten
  in memory first. See `reorderForStreaming`. That rewrite is also what keeps the decompression
  budget binding on this format: the streaming reader walks local file headers, not the central
  directory the budget measured, and only the rebuilt copy is guaranteed to carry exactly the
  measured entries — so an archive that cannot be rebuilt is `failed`, never streamed as it arrived.
  Behind that, a workbook read to completion that yields fewer worksheets than the archive holds
  returns `failed` rather than a partial workbook reported as `extracted` — silent partial output is
  the one outcome worth failing over, since a caller can retry a failure but cannot tell a truncated
  document from a complete one.

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
