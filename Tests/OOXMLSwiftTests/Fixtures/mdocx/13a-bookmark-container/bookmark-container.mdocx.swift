// Fixture 13a: bookmark-container
// Pins `mdocx-grammar` Requirement: "Bookmarks default to container with paired-marker escape hatch".
// Edge case: container form — Bookmark(id:) { ... } wraps a single Run inside a paragraph.
// The implementation injects <w:bookmarkStart> + <w:bookmarkEnd> markers around the wrapped runs.
// Sister fixture 13b covers the cross-paragraph paired-marker escape hatch.

import WordDSLSwift

let document = WordDocument {
    Section(id: "main") {
        Paragraph(id: "p1") {
            Bookmark(id: "intro_text") {
                "本章探討"
            }
        }
    }
}
