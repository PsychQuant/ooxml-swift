# Fixture 09: `style-define-on-first-use`

## Covers

`mdocx-grammar` Requirement **"Style references via typed enum with define-on-first-use"** (Requirement 9).

## What this fixture pins

Two paragraphs reference the same `WordStyle` value (`.titleBrown`). The op log emits exactly **one** `DefineStyle` op (carrying the style's properties: font, size, color, bold) **before** the two `InsertParagraph` ops. Subsequent references emit only the style-reference id `"titleBrown"`, never re-emitting `DefineStyle`. The docx shows both `<w:p>` elements sharing `<w:pStyle w:val="titleBrown"/>`.

## Files

- `.mdocx.swift` — two paragraphs, both `style: .titleBrown`.
- `.docx` — two `<w:p>` elements with identical `<w:pPr><w:pStyle w:val="titleBrown"/></w:pPr>`.
- `.normalized.docx` — byte-equal.
- `.oplog.jsonl` — one DefineStyle followed by two InsertParagraph + SetRuns pairs (5 ops total).

## What to inspect when this fixture fails

| Failure | Likely cause |
|---------|--------------|
| Phase B `oplog byte-diff` | Implementation emitted DefineStyle twice (once per Paragraph reference) — incorrect; OR DefineStyle emitted AFTER the first InsertParagraph — incorrect; OR DefineStyle missing properties (font/size/color/bold) |
| Phase B `docx byte-diff` | One paragraph missing `<w:pStyle>` (style propagation broken) |
