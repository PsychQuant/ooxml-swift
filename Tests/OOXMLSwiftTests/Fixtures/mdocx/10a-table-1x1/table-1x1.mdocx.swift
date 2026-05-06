// Fixture 10a: table-1x1
// Pins `mdocx-grammar` Requirement: "Table grammar mirrors OOXML three-layer structure".
// Edge case: minimal 1×1 table — Table > TableRow > TableCell > Paragraph.
// Each layer carries explicit id: per Requirement 8.

import WordDSLSwift

let document = WordDocument {
    Section(id: "main") {
        Table(id: "tbl1") {
            TableRow(id: "tbl1-r0") {
                TableCell(id: "tbl1-r0-c0") {
                    Paragraph(id: "tbl1-r0-c0-p0") {
                        "cell content"
                    }
                }
            }
        }
    }
}
