# Fixture 03: `tab-break-nbhyphen`

## Covers

`mdocx-grammar` Requirement **"Special-character inline atoms as standalone children"** (Requirement 3).

## What this fixture pins

`Tab()`, `Break()`, and `NoBreakHyphen()` are first-class siblings of `Run` and `String` in paragraph result-builder bodies. They emit no-text `<w:r>` elements wrapping `<w:tab/>`, `<w:br/>`, `<w:noBreakHyphen/>` respectively.

## Edge case captured

All three atoms appear in a single paragraph mixed with String literals, proving they compose without ceremony. Each is in its own `<w:r>` (not merged with adjacent text runs). The `<w:r>` wrappers carry no `<w:rPr>` (no formatting on bare atoms).

## What to inspect when this fixture fails

| Failure | Likely cause |
|---------|--------------|
| Phase B `docx byte-diff` | Atoms merged with surrounding text into one `<w:r>` (incorrect — each gets its own); OR atoms emitted as static factory methods on `Run` (`Run.tab()` style — explicitly rejected per spec); OR wrong OOXML element name (`<w:tab/>` vs `<w:tabStop/>` etc.) |
| `Swift compile-pass` failure | Removed `import WordDSLSwift` or syntax error |
