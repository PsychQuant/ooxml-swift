// Fixture 05: heading-via-style
// Pins `mdocx-grammar` Requirement: "No semantic shortcuts for OOXML-style attributes".
// Edge case: heading is `Paragraph(style: .heading1)`, NOT `Heading1(...)`. Confirms the DSL
// has no semantic-shortcut wrapper for paragraph-style or run-formatting attributes.
// Body paragraph follows without style — proves heading-style is per-paragraph, not document-level.

import WordDSLSwift

let document = WordDocument {
    Section(id: "main") {
        Paragraph(id: "h1", style: .heading1) {
            "Chapter Title"
        }
        Paragraph(id: "p1") {
            "Body content under the heading."
        }
    }
}
