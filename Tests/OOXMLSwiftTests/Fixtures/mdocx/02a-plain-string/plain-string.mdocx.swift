// Fixture 02a: plain-string
//
// Pins `mdocx-grammar` Requirement: "Flat Run with implicit String literal
// inline grammar" (canonical at openspec/specs/mdocx-grammar/spec.md).
//
// Edge case: paragraph body with ONLY plain `String` literals — no explicit
// `Run(...)` calls anywhere. The compiler implicitly converts each String
// to an unstyled `Run` containing that text. The corresponding docx shows
// three `<w:r><w:t>` elements with NO `<w:rPr>` formatting.
//
// What this fixture pins:
//   1. Plain String literals are valid paragraph body content (no Run wrapper required).
//   2. Multiple consecutive String literals produce multiple unstyled <w:r><w:t> runs in document order.
//   3. The reverse transcoder reads `<w:r>` with no `<w:rPr>` and emits a String literal (not `Run("...")`).
//
// Sister fixture 02b covers the explicit `Run(text, bold:, italics:, color:)` form;
// together they pin the entire "implicit String, explicit Run" inline grammar surface.

import WordDSLSwift

let document = WordDocument {
    Section(id: "main") {
        Paragraph(id: "p1") {
            "本章探討"
            "意識本質"
            "的議題。"
        }
    }
}
