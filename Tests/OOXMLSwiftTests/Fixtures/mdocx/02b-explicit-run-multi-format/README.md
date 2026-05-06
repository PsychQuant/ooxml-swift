# Fixture 02b: `explicit-run-multi-format`

## Covers

`mdocx-grammar` Requirement **"Flat Run with implicit String literal inline grammar"** (Requirement 2, canonical at `openspec/specs/mdocx-grammar/spec.md`), the **explicit-Run-with-formatting** variant.

Sister fixture: **02a-plain-string** covers the implicit-String pure-plain case. Together 02a + 02b pin the entire inline grammar surface.

## What this fixture pins

The simplest expression of multi-format `Run(...)` plus mixed-with-Strings composition. The `.mdocx.swift` paragraph body contains:

1. A plain String `"prefix "` (implicit unstyled `<w:r><w:t>`).
2. An explicit `Run("styled phrase", bold: true, italics: true, color: "#663300")` (one `<w:r>` with `<w:rPr>` containing three formatting elements).
3. A plain String `" suffix"` (another implicit unstyled `<w:r><w:t>`).

The corresponding `<slug>.docx` shows three `<w:r>` elements in document order:

```xml
<w:r><w:t xml:space="preserve">prefix </w:t></w:r>
<w:r><w:rPr><w:b/><w:i/><w:color w:val="663300"/></w:rPr><w:t xml:space="preserve">styled phrase</w:t></w:r>
<w:r><w:t xml:space="preserve"> suffix</w:t></w:r>
```

## Edge cases captured

- **Multi-format on one Run**: `bold: true, italics: true, color:` proves that flag combinations don't require separate `Run(...)` calls — one Run call expresses the full format combination.
- **Run + String composition**: Plain Strings flank the formatted Run in the same paragraph body, confirming the two surfaces compose without ceremony (no need to wrap Strings in `Run(...)` just because a sibling is formatted).
- **`<w:rPr>` element ordering**: `<w:b/>` then `<w:i/>` then `<w:color>` — the order matches OOXML's canonical `<w:rPr>` child ordering. Phase B byte-diff catches implementations that emit a different order.

## What to inspect when this fixture fails

| Failure | Likely cause |
|---------|--------------|
| `directory layout: name does not match <NN>[<letter>]-<slug>` | Directory was renamed |
| `file set: missing <slug>.{docx,normalized.docx,mdocx.swift,README.md}` | A required file was deleted |
| `Swift compile-pass: no import declaration` | `import WordDSLSwift` removed |
| `Swift compile-pass: unbalanced braces` | Syntax error introduced |
| Phase B: `docx byte-diff` (only when `activatePhaseB == true`) | Likely causes (in priority order): (1) WordDSLSwift implementation re-ordered `<w:rPr>` children (must match `<w:b/>`, `<w:i/>`, `<w:color>` sequence); (2) implementation emitted color hex without leading `#` strip (the DSL accepts `"#663300"` but the OOXML attribute value must be `663300` without `#`); (3) implementation merged the explicit Run with adjacent String literals into one run (incorrect — Run and String produce separate `<w:r>` elements always); (4) implementation collapsed `xml:space="preserve"` (must keep on `<w:t>` elements containing leading/trailing whitespace) |

## Why `<slug>.normalized.docx` equals `<slug>.docx` byte-for-byte here

Same as fixtures 01 and 02a: the hand-crafted docx omits every field the normalizer strips. Normalize is a no-op.

## Source content

The exact values — `"prefix "`, `Run("styled phrase", bold: true, italics: true, color: "#663300")`, `" suffix"` — are chosen to match the spec's Example block "explicit Run carries multiple format flags" pattern, with English text instead of CJK so reviewers can compare against the spec's `Run("意識本質", bold: true, italics: true, color: "#663300")` example without depending on font availability.
