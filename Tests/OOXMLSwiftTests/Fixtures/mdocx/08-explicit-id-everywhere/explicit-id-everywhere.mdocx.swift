// Fixture 08: explicit-id-everywhere
// Pins `mdocx-grammar` Requirement: "Mandatory explicit identifiers on structural elements".
// Edge case: every Section, Paragraph, Bookmark, Hyperlink-as-anchor-target carries explicit id:.
// Compiler refuses to compile any source missing id: on a structural element. The docx maps
// each id to its OOXML stable identifier (w14:paraId for paragraphs, w:bookmarkId/w:name for bookmarks).

import WordDSLSwift

let document = WordDocument {
    Section(id: "main") {
        Paragraph(id: "intro") {
            "see "
            Hyperlink(to: .anchor("body_anchor")) { "the body" }
            "."
        }
        Paragraph(id: "body_target") {
            Bookmark(id: "body_anchor") {
                "Body content."
            }
        }
    }
}
