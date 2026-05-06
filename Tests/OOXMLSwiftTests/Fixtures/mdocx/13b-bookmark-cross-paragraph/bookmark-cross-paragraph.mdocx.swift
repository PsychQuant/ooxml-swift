// Fixture 13b: bookmark-cross-paragraph
// Pins `mdocx-grammar` Requirement: "Bookmarks default to container with paired-marker escape hatch".
// Edge case: paired-marker form for cross-paragraph spans where start and end markers
// cannot share a parent. BookmarkStart(id:) and BookmarkEnd(id:) sit as siblings of paragraphs
// in Section's body. The id values must match. Sister fixture 13a covers the container form.

import WordDSLSwift

let document = WordDocument {
    Section(id: "main") {
        BookmarkStart(id: "ch1_span")
        Paragraph(id: "p1") {
            "First paragraph in span."
        }
        Paragraph(id: "p2") {
            "Second paragraph in span."
        }
        BookmarkEnd(id: "ch1_span")
    }
}
