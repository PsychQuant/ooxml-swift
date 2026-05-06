# Fixture 11: `bullet-and-numbered-lists`

## Covers

`mdocx-grammar` Requirement **"Lists use Paragraph with numPr reference, not nested containers"** (Requirement 11).

## What this fixture pins

Numbered + bullet lists, with nested level — all expressed as `Paragraph(style: .listItem, numbering:, level:)`. NO `List { ListItem }` nested-container syntax. NumberingDefinition references travel by typed identifier (`.numbered1`, `.bulletA`).

## Edge case captured

- Two numbered items (numId=1) + two bullet items (numId=2) — proves multiple list contexts coexist in one section.
- Nested level (`level: 1` on `li4`) — proves indentation is per-paragraph via `<w:ilvl>`, not by structural nesting in the DSL.
- Each paragraph independently identifies its list context — matches OOXML's `<w:numPr>` shape.

## What to inspect when this fixture fails

| Failure | Likely cause |
|---------|--------------|
| Phase B `docx byte-diff` | Implementation introduced `List { ListItem }` nested syntax (rejected); OR `<w:numPr>` placement wrong (must be inside `<w:pPr>`); OR list context lost across paragraphs |
