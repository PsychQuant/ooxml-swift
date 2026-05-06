# Fixture 04: `mirrored-element-set`

## Covers

`mdocx-grammar` Requirement **"OOXML-mirror element naming"** (Requirement 4).

## What this fixture pins

DSL element names mirror OOXML term-of-art 1:1. `Paragraph` ↔ `<w:p>`, `Run` ↔ `<w:r>`, with no translation table needed in the reverse transcoder.

## Edge case captured

This is a meta-fixture: its purpose is to document the naming-policy contract by example. The full mirrored set (Section, Table, TableRow, TableCell, Hyperlink, Bookmark) is exercised by sister fixtures 06, 08, 10a/10b, 12, 13a/13b. Together they confirm the naming policy applies to every named element introduced by the spec.

## What to inspect when this fixture fails

| Failure | Likely cause |
|---------|--------------|
| Phase B `docx byte-diff` | Element renamed (e.g., `Run` → `Span` or `Format`) — explicitly rejected per Decision 4 |
| Other failures | Standard file-set / compile-pass issues |
