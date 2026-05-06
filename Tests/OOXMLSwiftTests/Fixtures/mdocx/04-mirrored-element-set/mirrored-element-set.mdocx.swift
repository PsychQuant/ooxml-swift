// Fixture 04: mirrored-element-set
// Pins `mdocx-grammar` Requirement: "OOXML-mirror element naming".
// Edge case: minimal showcase that DSL element names mirror OOXML term-of-art (Paragraph↔w:p,
// Run↔w:r, with italic Run formatting demonstrating standard <w:rPr> shape).
// Other named elements (Table/TableRow/TableCell, Hyperlink, Bookmark, Section) are covered by
// fixtures 06, 08, 10a/10b, 12, 13a/13b — each independently confirms the naming policy.

import WordDSLSwift

let document = WordDocument {
    Section(id: "main") {
        Paragraph(id: "p1") {
            "Showcases mirrored OOXML element names: Paragraph and Run map 1:1 to "
            Run("w:p", italics: true)
            " and "
            Run("w:r", italics: true)
            "."
        }
    }
}
