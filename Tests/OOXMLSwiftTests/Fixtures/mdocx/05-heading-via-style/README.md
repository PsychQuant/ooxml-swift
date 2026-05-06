# Fixture 05: `heading-via-style`

## Covers

`mdocx-grammar` Requirement **"No semantic shortcuts for OOXML-style attributes"** (Requirement 5).

## What this fixture pins

Headings, quotes, captions, list items — every paragraph kind that OOXML expresses through `<w:pStyle>` — is written as `Paragraph(style: .styleName)`. The DSL does NOT provide `Heading1(...)`, `Quote(...)`, `Caption(...)` wrapper components. Same rule applies to inline formatting: `Run(text:, bold:)` not `Bold(text)`.

## Edge case captured

Two paragraphs in one section: one with `style: .heading1` producing `<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>`, one without producing default-style body paragraph. Demonstrates style is per-paragraph (not section-default) and absent style means "default body paragraph" not "uninitialized".

## What to inspect when this fixture fails

| Failure | Likely cause |
|---------|--------------|
| Phase B `docx byte-diff` | Implementation invented a `Heading1(...)` wrapper (rejected); OR emitted style id with wrong casing; OR failed to omit `<w:pPr>` block on the unstyled body paragraph |
