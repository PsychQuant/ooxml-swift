// Fixture 12: hyperlink-url-anchor-mailto
// Pins `mdocx-grammar` Requirement: "Hyperlinks are containers with target enum".
// Edge case: one paragraph with three hyperlinks — .url, .anchor, .mailto — demonstrating the
// HyperlinkTarget enum's three primary cases. URL + mailto produce <w:hyperlink r:id="...">
// references with rels entries; anchor produces <w:hyperlink w:anchor="..."> with no rels entry.

import WordDSLSwift

let document = WordDocument {
    Section(id: "main") {
        Paragraph(id: "p1") {
            "External: "
            Hyperlink(to: .url("https://example.com")) { "example.com" }
            ". Internal: "
            Hyperlink(to: .anchor("ch1_intro")) { "Chapter 1" }
            ". Email: "
            Hyperlink(to: .mailto("hello@example.com")) { "contact us" }
            "."
        }
    }
}
