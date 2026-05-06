# Fixture 10b: `table-3x3-with-formatting`

## Covers

`mdocx-grammar` Requirement **"Table grammar mirrors OOXML three-layer structure"** (Requirement 10), the formatted variant.

## What this fixture pins

3×3 table with header-row styling. Header row paragraphs use `style: .tableHeader` + `Run("...", bold: true)`. Body rows use plain paragraphs. Demonstrates the three-layer hierarchy scales naturally and cell content can carry the full paragraph-style + inline-format surface.

## Edge case captured

- Header row + body rows in the same table — proves the DSL doesn't enforce row uniformity.
- Per-cell paragraph IDs (`tbl1-r0-c0-p0` ... `tbl1-r2-c2-p0`) — proves IDs can be hierarchically structured for readability.
- Cell content uses both styled paragraph (header row) and plain paragraph (body rows) freely.

## What to inspect when this fixture fails

| Failure | Likely cause |
|---------|--------------|
| Phase B `docx byte-diff` | Implementation merged adjacent same-style cells; OR lost `<w:b/>` formatting on header runs; OR re-ordered table rows |
