// Fixture 10b: table-3x3-with-formatting
// Pins `mdocx-grammar` Requirement: "Table grammar mirrors OOXML three-layer structure".
// Edge case: 3×3 table with header-row formatting (style + bold). Demonstrates the three-layer
// hierarchy scales without ceremony, and that cell content can carry style + run-formatting.

import WordDSLSwift

let document = WordDocument {
    Section(id: "main") {
        Table(id: "tbl1") {
            TableRow(id: "tbl1-r0") {
                TableCell(id: "tbl1-r0-c0") {
                    Paragraph(id: "tbl1-r0-c0-p0", style: .tableHeader) {
                        Run("H1", bold: true)
                    }
                }
                TableCell(id: "tbl1-r0-c1") {
                    Paragraph(id: "tbl1-r0-c1-p0", style: .tableHeader) {
                        Run("H2", bold: true)
                    }
                }
                TableCell(id: "tbl1-r0-c2") {
                    Paragraph(id: "tbl1-r0-c2-p0", style: .tableHeader) {
                        Run("H3", bold: true)
                    }
                }
            }
            TableRow(id: "tbl1-r1") {
                TableCell(id: "tbl1-r1-c0") { Paragraph(id: "tbl1-r1-c0-p0") { "r1c1" } }
                TableCell(id: "tbl1-r1-c1") { Paragraph(id: "tbl1-r1-c1-p0") { "r1c2" } }
                TableCell(id: "tbl1-r1-c2") { Paragraph(id: "tbl1-r1-c2-p0") { "r1c3" } }
            }
            TableRow(id: "tbl1-r2") {
                TableCell(id: "tbl1-r2-c0") { Paragraph(id: "tbl1-r2-c0-p0") { "r2c1" } }
                TableCell(id: "tbl1-r2-c1") { Paragraph(id: "tbl1-r2-c1-p0") { "r2c2" } }
                TableCell(id: "tbl1-r2-c2") { Paragraph(id: "tbl1-r2-c2-p0") { "r2c3" } }
            }
        }
    }
}
