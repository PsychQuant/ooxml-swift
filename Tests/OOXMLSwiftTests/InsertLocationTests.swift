import XCTest
@testable import OOXMLSwift

final class InsertLocationTests: XCTestCase {

    func testParagraphIndexInsertAtStart() throws {
        var doc = WordDocument()
        doc.body.children = [.paragraph(Paragraph(text: "existing"))]
        try doc.insertParagraph(Paragraph(text: "new"), at: .paragraphIndex(0))
        XCTAssertEqual(doc.body.children.count, 2)
        if case .paragraph(let p) = doc.body.children[0] {
            XCTAssertEqual(p.runs.first?.text, "new")
        } else {
            XCTFail("Expected first child to be paragraph")
        }
    }

    func testParagraphIndexInvalidThrows() {
        var doc = WordDocument()
        XCTAssertThrowsError(try doc.insertParagraph(Paragraph(text: "x"), at: .paragraphIndex(999))) { error in
            guard case InsertLocationError.invalidParagraphIndex(let i) = error else {
                XCTFail("Expected invalidParagraphIndex, got \(error)"); return
            }
            XCTAssertEqual(i, 999)
        }
    }

    func testAfterImageIdNotFoundThrows() {
        var doc = WordDocument()
        doc.body.children = [.paragraph(Paragraph(text: "no image"))]
        XCTAssertThrowsError(try doc.insertParagraph(
            Paragraph(text: "caption"),
            at: .afterImageId("rId-nonexistent")
        )) { error in
            guard case InsertLocationError.imageIdNotFound(let rId) = error else {
                XCTFail("Expected imageIdNotFound, got \(error)"); return
            }
            XCTAssertEqual(rId, "rId-nonexistent")
        }
    }

    func testAfterImageIdFound() throws {
        var doc = WordDocument()
        var paraWithImage = Paragraph()
        var imageRun = Run(text: "")
        imageRun.drawing = Drawing(width: 100, height: 100, imageId: "rId22", name: "pic")
        paraWithImage.runs = [imageRun]
        doc.body.children = [
            .paragraph(Paragraph(text: "before")),
            .paragraph(paraWithImage)
        ]

        try doc.insertParagraph(Paragraph(text: "caption"), at: .afterImageId("rId22"))

        XCTAssertEqual(doc.body.children.count, 3)
        if case .paragraph(let p) = doc.body.children[2] {
            XCTAssertEqual(p.runs.first?.text, "caption")
        } else {
            XCTFail("Expected caption as third child")
        }
    }

    func testAfterTableIndexInsert() throws {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "before")),
            .table(Table(rowCount: 2, columnCount: 2)),
            .paragraph(Paragraph(text: "after"))
        ]
        try doc.insertParagraph(Paragraph(text: "inserted"), at: .afterTableIndex(0))
        XCTAssertEqual(doc.body.children.count, 4)
        if case .paragraph(let p) = doc.body.children[2] {
            XCTAssertEqual(p.runs.first?.text, "inserted")
        } else {
            XCTFail("Expected inserted paragraph at index 2")
        }
    }

    func testTableIndexOutOfRangeThrows() {
        var doc = WordDocument()
        doc.body.children = [.paragraph(Paragraph(text: "x"))]
        XCTAssertThrowsError(try doc.insertParagraph(Paragraph(text: "y"), at: .afterTableIndex(0))) { error in
            guard case InsertLocationError.tableIndexOutOfRange = error else {
                XCTFail("Expected tableIndexOutOfRange, got \(error)"); return
            }
        }
    }

    func testIntoTableCellInsert() throws {
        var doc = WordDocument()
        let cell = TableCell(paragraphs: [Paragraph(text: "original")])
        let row = TableRow(cells: [cell])
        let table = Table(rows: [row])
        doc.body.children = [.table(table)]

        try doc.insertParagraph(
            Paragraph(text: "added"),
            at: .intoTableCell(tableIndex: 0, row: 0, col: 0)
        )

        if case .table(let t) = doc.body.children[0] {
            XCTAssertEqual(t.rows[0].cells[0].paragraphs.count, 2)
            XCTAssertEqual(t.rows[0].cells[0].paragraphs[1].runs.first?.text, "added")
        } else {
            XCTFail("Expected table still at index 0")
        }
    }

    func testIntoTableCellOutOfRangeThrows() {
        var doc = WordDocument()
        let cell = TableCell(paragraphs: [])
        let row = TableRow(cells: [cell])
        let table = Table(rows: [row])
        doc.body.children = [.table(table)]

        XCTAssertThrowsError(try doc.insertParagraph(
            Paragraph(text: "x"),
            at: .intoTableCell(tableIndex: 0, row: 5, col: 0)
        )) { error in
            guard case InsertLocationError.tableCellOutOfRange = error else {
                XCTFail("Expected tableCellOutOfRange, got \(error)"); return
            }
        }
    }
}
