# Fixture 10a: `table-1x1`

## Covers

`mdocx-grammar` Requirement **"Table grammar mirrors OOXML three-layer structure"** (Requirement 10).

## What this fixture pins

The minimal table: 1 row × 1 cell × 1 paragraph. Three-layer DSL hierarchy `Table { TableRow { TableCell { Paragraph { ... } } } }` maps 1:1 to OOXML `<w:tbl><w:tr><w:tc><w:p>...</w:p></w:tc></w:tr></w:tbl>`. Each of the four layers carries explicit `id:` per Requirement 8.

## Edge case captured

The smallest possible table — proves the three-layer hierarchy doesn't have implicit defaults. Even a single-cell table requires explicit `Table` + `TableRow` + `TableCell` containment. The DSL does NOT provide a `Table.singleCell(...)` shortcut.

## What to inspect when this fixture fails

| Failure | Likely cause |
|---------|--------------|
| Phase B `docx byte-diff` | Implementation collapsed three-layer hierarchy into a flat element; OR omitted required `<w:tr>` / `<w:tc>` wrappers |
