# docx fidelity corpus

Seventeen real Word-authored `.docx`, vendored verbatim from **mammoth 1.12.0**, `test/test-data/`.

## Why these files are here rather than read from `node_modules`

`tests/docx-fidelity.test.ts` pins our `.docx` output against `mammoth.extractRawText` — the reader
this library used until the streaming rewrite. That suite needs a corpus that is:

1. **Reviewable.** The acceptance criterion for the rewrite is "output diffed against mammoth across
   a document corpus, with any structural regressions understood and accepted". That is only
   checkable if the corpus and its expected output are in the diff.
2. **Stable.** mammoth declares no `files` field in its `package.json`, so `test/` ships by accident
   rather than by contract. A future release that adds one would delete the regression suite
   silently. mammoth is also a `devDependency` now, kept solely as the fidelity oracle.
3. **Extensible.** The corpus grows with redacted real-world documents; splitting it between
   `node_modules` and here would be worse than one directory.

## What they cover

They are chosen for coverage of the WordprocessingML constructs that change the extracted text, not
for size. Notably: `strict-format.docx` binds the ISO-strict namespace URI (matching a literal `w:`
prefix would silently return nothing for it); `text-box.docx` carries a real `mc:AlternateContent`
with both a `mc:Choice` (DrawingML) and a `mc:Fallback` (VML `w:pict`) branch, which is the
double-count hazard and the paragraph-hoisting case at once; `tables.docx`, `footnotes.docx`,
`endnotes.docx` and `comments.docx` pin what is deliberately NOT extracted.

Five of them — `embedded-style-map`, `strikethrough`, `underline`, `external-picture`,
`tiny-picture-target-base-relative` — add no construct the others don't already have. They are kept
anyway: the value of a real corpus is the cases nobody thought to enumerate, and they cost ~90 KB
and no measurable runtime.

## Licence

mammoth is BSD-2-Clause, which permits redistribution in source form provided the notice and
disclaimer are retained. They are, in full:

> Copyright (c) 2013, Michael Williamson
> All rights reserved.
>
> Redistribution and use in source and binary forms, with or without
> modification, are permitted provided that the following conditions are met:
>
> 1. Redistributions of source code must retain the above copyright notice, this
>    list of conditions and the following disclaimer.
> 2. Redistributions in binary form must reproduce the above copyright notice,
>    this list of conditions and the following disclaimer in the documentation
>    and/or other materials provided with the distribution.
>
> THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
> ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
> WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
> DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR
> ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
> (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
> LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
> ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
> (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
> SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

Their contents are trivial sample prose ("Walking on imported air", "Ouch.") with no third-party
material.
