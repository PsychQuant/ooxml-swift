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

    // MARK: - afterText / beforeText (0.9.0)

    func testAfterTextInsertsAfterMatchingParagraph() throws {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "intro")),
            .paragraph(Paragraph(text: "圖 4-1：前後期報酬率分布")),
            .paragraph(Paragraph(text: "body text")),
        ]
        try doc.insertParagraph(
            Paragraph(text: "caption here"),
            at: .afterText("圖 4-1", instance: 1)
        )
        XCTAssertEqual(doc.body.children.count, 4)
        if case .paragraph(let p) = doc.body.children[2] {
            XCTAssertEqual(p.runs.first?.text, "caption here")
        } else {
            XCTFail("Expected caption at index 2")
        }
    }

    func testBeforeTextInsertsBeforeMatchingParagraph() throws {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "intro")),
            .paragraph(Paragraph(text: "section start")),
            .paragraph(Paragraph(text: "body text")),
        ]
        try doc.insertParagraph(
            Paragraph(text: "heading"),
            at: .beforeText("section start", instance: 1)
        )
        XCTAssertEqual(doc.body.children.count, 4)
        if case .paragraph(let p) = doc.body.children[1] {
            XCTAssertEqual(p.runs.first?.text, "heading")
        } else {
            XCTFail("Expected heading at index 1")
        }
    }

    func testTextNotFoundThrows() {
        var doc = WordDocument()
        doc.body.children = [.paragraph(Paragraph(text: "only this"))]
        XCTAssertThrowsError(try doc.insertParagraph(
            Paragraph(text: "x"),
            at: .afterText("missing", instance: 1)
        )) { error in
            guard case InsertLocationError.textNotFound(let s, let i) = error else {
                XCTFail("Expected textNotFound, got \(error)"); return
            }
            XCTAssertEqual(s, "missing")
            XCTAssertEqual(i, 1)
        }
    }

    func testAfterTextInstance2FindsSecondOccurrence() throws {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "圖 4-1 first")),
            .paragraph(Paragraph(text: "middle")),
            .paragraph(Paragraph(text: "圖 4-1 second")),
            .paragraph(Paragraph(text: "tail")),
        ]
        try doc.insertParagraph(
            Paragraph(text: "after 2nd"),
            at: .afterText("圖 4-1", instance: 2)
        )
        // Expected order: [圖 4-1 first, middle, 圖 4-1 second, after 2nd, tail]
        XCTAssertEqual(doc.body.children.count, 5)
        if case .paragraph(let p) = doc.body.children[3] {
            XCTAssertEqual(p.runs.first?.text, "after 2nd")
        } else {
            XCTFail("Expected 'after 2nd' at index 3")
        }
    }

    func testAfterTextCrossRunMatch() throws {
        // Simulates the thesis scenario where text is split across runs
        var doc = WordDocument()
        var para = Paragraph()
        para.runs = [
            Run(text: "圖 4-1："),
            Run(text: ""),  // phantom empty run
            Run(text: "前後期報酬率分布")
        ]
        doc.body.children = [.paragraph(para)]

        try doc.insertParagraph(
            Paragraph(text: "caption"),
            at: .afterText("圖 4-1：前後期", instance: 1)
        )
        XCTAssertEqual(doc.body.children.count, 2)
    }
}
