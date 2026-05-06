// Fixture 09: style-define-on-first-use
// Pins `mdocx-grammar` Requirement: "Style references via typed enum with define-on-first-use".
// Edge case: two paragraphs reference the same WordStyle value (.titleBrown). The op log emits
// exactly one DefineStyle op (carrying the style's properties) BEFORE the two InsertParagraph
// ops. Subsequent references emit only style-reference id, never re-emitting DefineStyle.
// Both .docx <w:p> elements share <w:pStyle w:val="titleBrown"/>.

import WordDSLSwift

let document = WordDocument {
    Section(id: "main") {
        Paragraph(id: "h1", style: .titleBrown) {
            "Title"
        }
        Paragraph(id: "h2", style: .titleBrown) {
            "Subtitle"
        }
    }
}
