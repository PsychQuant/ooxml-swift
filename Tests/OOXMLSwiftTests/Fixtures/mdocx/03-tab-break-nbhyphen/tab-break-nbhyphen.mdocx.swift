// Fixture 03: tab-break-nbhyphen
// Pins `mdocx-grammar` Requirement: "Special-character inline atoms as standalone children".
// Edge case: Tab(), Break(), NoBreakHyphen() compose as siblings of String/Run in paragraph body,
// not as static factory methods on Run. Each emits a no-text <w:r> wrapping the special-char element.

import WordDSLSwift

let document = WordDocument {
    Section(id: "main") {
        Paragraph(id: "p1") {
            "Header"
            Tab()
            "Right-aligned"
            Break()
            "Continued"
            NoBreakHyphen()
            "tail"
        }
    }
}
