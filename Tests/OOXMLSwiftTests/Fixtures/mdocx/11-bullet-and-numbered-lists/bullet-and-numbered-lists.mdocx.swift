// Fixture 11: bullet-and-numbered-lists
// Pins `mdocx-grammar` Requirement: "Lists use Paragraph with numPr reference, not nested containers".
// Edge case: numbered list (numId 1) + bullet list (numId 2) + nested level (ilvl 1) — all expressed
// as Paragraph(style: .listItem, numbering: ..., level: ...) with NumberingDefinition references.
// The DSL does NOT provide List { ListItem } nested-container syntax.

import WordDSLSwift

let document = WordDocument {
    Section(id: "main") {
        Paragraph(id: "li1", style: .listItem, numbering: .numbered1, level: 0) {
            "Numbered item 1"
        }
        Paragraph(id: "li2", style: .listItem, numbering: .numbered1, level: 0) {
            "Numbered item 2"
        }
        Paragraph(id: "li3", style: .listItem, numbering: .bulletA, level: 0) {
            "Bullet item 1"
        }
        Paragraph(id: "li4", style: .listItem, numbering: .bulletA, level: 1) {
            "Sub-bullet item"
        }
    }
}
