# talon-extractor

In-house email quote & signature extractor. Given an email's `text` and/or `html`, it strips the
quoted history (the "On … wrote:" reply chain, forwarded blocks, `>`-quotes) and boilerplate noise,
leaving the sender's actual new content. Pure string/regex — no DOM parser, no ML — so it runs cheaply
on a Lambda Node.js runtime. Drop-in replacement for [talonjs](https://www.npmjs.com/package/talonjs).

## Usage

```ts
import { extractContent } from './draft_talon'

const { extracted_text, extracted_html } = extractContent({
  text: 'Sounds good!\n\nOn Mon, Jun 1 Bob <bob@x.com> wrote:\n> old quoted message',
  html: '<div>Sounds good!</div><blockquote>old quoted message</blockquote>',
})
// extracted_text === 'Sounds good!'
```

`extractContent({ text?, html? })` returns `{ extracted_text?, extracted_html? }`.

Also exported: `extractNewContent` / `extractFromHtml` (the raw cut, before noise-stripping),
`stripNoise` / `stripNoiseHtml` (trailing-boilerplate removal), and `isTrueDsn` (bounce detection).

## What it does that off-the-shelf engines don't

Benchmarked against TalonJS and Mailgun Talon on a real corpus of ~41k messages (exact-match vs a
gpt-4o answer key, graded on the cut):

| Capability | this | TalonJS | Mailgun Talon |
|---|---|---|---|
| Quote cutting (On-wrote / `>` / From: / Original-Msg) | yes | yes | yes |
| Glued "On…wrote:" orphan fix | **yes** | no | no |
| Foreign-language attributions (12+ langs) + Chinese/Arabic | **yes** | limited | limited |
| DSN / bounce keep-whole | **yes** | no | no |
| Forward keep-whole | **yes** | no | no |
| Inline-reply keep-whole | **yes** | no | partial |
| HTML trailing-signature reattach | **yes** | no | no |
| Noise-strip (mobile sigs / disclaimers / footers) | **yes** | no | no |

Per-feature exact-match % (this / TalonJS / Talon): on_wrote 96/51/49, no_quote 100/99/99,
from_header 89/84/81, gt_quote 85/77/76, foreign_verb 87/61/26, inline 91/38/22, dsn 93/40/40.

## Tests

```sh
npm install
npm test
```

The evaluation/audit harness (the `eval_*.ts` scripts) is intentionally **not** included here — it
depends on a ~1.2GB message corpus and large answer-key files that don't belong in source control.
