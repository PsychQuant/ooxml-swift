# Fixture 02a: `plain-string`

## Covers

`mdocx-grammar` Requirement **"Flat Run with implicit String literal inline grammar"** (Requirement 2, canonical at `openspec/specs/mdocx-grammar/spec.md`).

Sister fixture: **02b-explicit-run-multi-format** covers the explicit `Run(text, bold:, italics:, ...)` variant. Together 02a + 02b pin the entire "implicit String literal, explicit Run for any formatting" inline grammar surface.

## What this fixture pins

The minimal expression of the implicit-String rule. The `.mdocx.swift` paragraph body contains only three String literals — no explicit `Run(...)` calls. The compiler implicitly converts each String into an unstyled `<w:r><w:t>` run.

The corresponding `<slug>.docx` shows three `<w:r>` elements in the same document order. Each `<w:r>` contains a `<w:t xml:space="preserve">` element with the literal text. None of the `<w:r>` elements has a `<w:rPr>` — confirming the "unstyled" property of implicit string conversion.

## Edge case captured

Three Strings in a row (not just one) — proves that consecutive String literals each produce their own `<w:r><w:t>` element in document order, not merged into a single run. This matters because:

- The op log emits one `Run` operation per String literal in source order, not one merged operation.
- The reverse transcoder must split runs back into individual String literals (one per `<w:r>` with no rPr), not merge consecutive same-style runs into a single literal.

## What to inspect when this fixture fails

| Failure | Likely cause |
|---------|--------------|
| `directory layout: name does not match <NN>[<letter>]-<slug>` | Directory was renamed |
| `file set: missing <slug>.{docx,normalized.docx,mdocx.swift,README.md}` | A required file was deleted; rebuild via the Python in 01-dual-extension-recognition's README |
| `Swift compile-pass: no import declaration` | `import WordDSLSwift` removed from .mdocx.swift |
| `Swift compile-pass: unbalanced braces` | Syntax error introduced |
| Phase B: `docx byte-diff` (only when `activatePhaseB == true`) | The new `WordDSLSwift` implementation merged consecutive Strings into a single `<w:r>` (incorrect — must produce one `<w:r>` per String literal); OR added `<w:rPr>` to runs (incorrect — implicit String must produce unstyled runs); OR changed paraId numbering away from monotonic `00000001` |

## Why `<slug>.normalized.docx` equals `<slug>.docx` byte-for-byte here

Same as fixture 01: the hand-crafted docx omits every field the normalizer strips (no RSIDs, no theme, no settings keys, single monotonic paraId). Normalize is a no-op.

## Source content

The three Strings — `"本章探討"`, `"意識本質"`, `"的議題。"` — come from the spec's Example block "mixed inline content" (under Requirement "Flat Run with implicit String literal inline grammar"), modified to drop the explicit `Run(...)` middle so this fixture exercises the pure-String case. Sister fixture 02b uses the exact spec Example values for the explicit-Run case.
