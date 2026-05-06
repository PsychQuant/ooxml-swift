// Fixture 02b: explicit-run-multi-format
//
// Pins `mdocx-grammar` Requirement: "Flat Run with implicit String literal
// inline grammar" (canonical at openspec/specs/mdocx-grammar/spec.md), the
// explicit-Run-with-formatting variant.
//
// Edge case: a single `Run("text", bold:, italics:, color:)` call carrying
// multiple format flags simultaneously. Demonstrates the Decision-1 rule:
// any non-default formatting goes through `Run(...)` with named parameters
// — there are NO single-format wrapper components like `Bold(...)` or
// `Italic(...)`.
//
// What this fixture pins:
//   1. `Run("text", bold: true, italics: true, color: "#663300")` is the canonical multi-format form.
//   2. The compiler emits one `<w:r>` element with one `<w:rPr>` containing `<w:b/>`, `<w:i/>`, `<w:color w:val="663300"/>` in that order.
//   3. Surrounding plain Strings (the "prefix " and " suffix" in the same paragraph body) are emitted as separate unstyled `<w:r>` runs — confirming Run + String compose freely in the same paragraph body.
//   4. The reverse transcoder reads `<w:r>` with `<w:rPr>` and emits `Run(text:, ...)` with the corresponding format flags — never a String literal even if visually the text could be one.
//
// Sister fixture 02a covers the pure plain-String case; together 02a + 02b
// pin the entire "implicit String, explicit Run for any formatting" surface.

import WordDSLSwift

let document = WordDocument {
    Section(id: "main") {
        Paragraph(id: "p1") {
            "prefix "
            Run("styled phrase", bold: true, italics: true, color: "#663300")
            " suffix"
        }
    }
}
