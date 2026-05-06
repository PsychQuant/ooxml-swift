// Fixture 06: two-sections-marker-inversion
// Pins `mdocx-grammar` Requirement: "Section as DSL container with compile-time marker inversion".
// Edge case: two sequential Section{...} blocks at DSL level invert into OOXML's marker pattern:
// the first section's <w:sectPr> appears inside its last paragraph's <w:pPr>; the second
// section's <w:sectPr> appears as a direct child of <w:body> after its last paragraph.

import WordDSLSwift

let document = WordDocument {
    Section(id: "front", type: .continuous) {
        Paragraph(id: "f1") { "Front matter section content." }
    }
    Section(id: "main", type: .nextPage) {
        Paragraph(id: "m1") { "Main body section content." }
    }
}
