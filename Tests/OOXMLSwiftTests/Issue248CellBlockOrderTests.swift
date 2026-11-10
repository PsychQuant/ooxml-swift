import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// PsychQuant/ooxml-swift#248: a table cell's paragraphs and nested tables
/// keep their source order when the table is re-serialized from the typed
/// model.
///
/// `TableCell` stores paragraphs and nested tables in two arrays, so the
/// order between them used to be lost at read time and the writer emitted
/// every paragraph, then every table. Editing one cell rewrites the whole
/// table, which moved content in every OTHER cell that held a nested table.
/// All fixtures are synthetic.
final class Issue248CellBlockOrderTests: XCTestCase {

    // MARK: - Fixture plumbing

    private static let namespaces = [
        #"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#,
        #"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#,
        #"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#,
    ].joined(separator: " ")

    private var cleanup: [URL] = []

    override func tearDown() {
        for url in cleanup { try? FileManager.default.removeItem(at: url) }
        cleanup = []
        super.tearDown()
    }

    private func p(_ text: String) -> String { "<w:p><w:r><w:t>\(text)</w:t></w:r></w:p>" }
    private let emptyP = "<w:p/>"

    /// A one-cell table whose cell holds `content`.
    private func tbl(_ content: String) -> String {
        "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid><w:tr><w:tc>"
            + "<w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr>\(content)</w:tc></w:tr></w:tbl>"
    }

    /// Body: one paragraph, then a two-row table. Row 0 is the cell the tests
    /// edit; row 1 holds the cell under test.
    private func documentXML(cell: String) -> String {
        let outer = "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"5000\"/></w:tblGrid>"
            + "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"5000\" w:type=\"dxa\"/></w:tcPr>\(p("X0"))</w:tc></w:tr>"
            + "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"5000\" w:type=\"dxa\"/></w:tcPr>\(cell)</w:tc></w:tr>"
            + "</w:tbl>"
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
            + "<w:document \(Self.namespaces)><w:body>\(p("LEAD"))\(outer)<w:sectPr/></w:body></w:document>"
    }

    private func makeDocx(_ documentXML: String) throws -> URL {
        let parts: [String: Data] = [
            "[Content_Types].xml": Data(#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.utf8),
            "_rels/.rels": Data(#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.utf8),
            "word/document.xml": Data(documentXML.utf8),
            "word/_rels/document.xml.rels": Data(#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#.utf8),
        ]
        let archive = try Archive(accessMode: .create)
        for name in parts.keys.sorted() {
            let data = parts[name]!
            try archive.addEntry(with: name, type: .file, uncompressedSize: Int64(data.count), compressionMethod: .deflate) { position, size in
                let start = data.startIndex.advanced(by: Int(position))
                return data.subdata(in: start..<start.advanced(by: size))
            }
        }
        let url = FileManager.default.temporaryDirectory.appendingPathComponent("i248-\(UUID().uuidString).docx")
        try XCTUnwrap(archive.data).write(to: url)
        cleanup.append(url)
        return url
    }

    /// Save in overlay mode (the path che-word-mcp uses); returns the saved
    /// file and its `word/document.xml`.
    private func save(_ doc: WordDocument) throws -> (url: URL, xml: String) {
        let url = FileManager.default.temporaryDirectory.appendingPathComponent("i248-out-\(UUID().uuidString).docx")
        cleanup.append(url)
        try DocxWriter.write(doc, to: url)
        let archive = try Archive(url: url, accessMode: .read)
        var out = Data()
        _ = try archive.extract(try XCTUnwrap(archive["word/document.xml"])) { out.append($0) }
        return (url, String(decoding: out, as: UTF8.self))
    }

    /// The block children of a `<w:tc>`, in order: a paragraph is its text
    /// (`<empty>` when it has none), a table is `T[...]` with its cells'
    /// own signatures inside.
    private func signature(ofCell cell: XMLElement) -> String {
        var parts: [String] = []
        for child in cell.children ?? [] {
            guard let el = child as? XMLElement else { continue }
            switch el.name {
            case "w:p":
                let text = el.stringValue ?? ""
                parts.append(text.isEmpty ? "<empty>" : text)
            case "w:tbl":
                let cells = el.elements(forName: "w:tr").flatMap { $0.elements(forName: "w:tc") }
                parts.append("T[" + cells.map { signature(ofCell: $0) }.joined(separator: "|") + "]")
            default:
                continue
            }
        }
        return parts.joined(separator: ",")
    }

    /// Signature of the cell under test (row 1 of the first body table).
    private func signature(_ documentXML: String) throws -> String {
        let doc = try XMLDocument(xmlString: documentXML)
        let body = try XCTUnwrap(doc.rootElement()?.elements(forName: "w:body").first)
        let table = try XCTUnwrap(body.elements(forName: "w:tbl").first)
        let rows = table.elements(forName: "w:tr")
        XCTAssertEqual(rows.count, 2)
        return signature(ofCell: try XCTUnwrap(rows[1].elements(forName: "w:tc").first))
    }

    /// The typed cell under test of a read document.
    private func cellUnderTest(_ doc: WordDocument) throws -> TableCell {
        guard case .table(let table) = doc.body.children[1] else {
            XCTFail("fixture: body child 1 must be the table"); throw CocoaError(.featureUnsupported)
        }
        return table.rows[1].cells[0]
    }

    private func cellSignature(_ cell: TableCell) throws -> String {
        let doc = try XMLDocument(xmlString: "<w:r xmlns:w=\"x\">\(cell.toXML())</w:r>")
        return signature(ofCell: try XCTUnwrap(doc.rootElement()?.elements(forName: "w:tc").first))
    }

    // MARK: - The reported defect: editing another cell of the same table

    private var shapes: [(label: String, cell: String, expected: String)] {
        [
            ("table then empty paragraph", p("A") + tbl(p("N1")) + emptyP, "A,T[N1],<empty>"),
            ("table then a paragraph with text", p("A") + tbl(p("N1")) + p("TAIL"), "A,T[N1],TAIL"),
            ("table then two paragraphs", p("A") + tbl(p("N1")) + p("B") + p("C"), "A,T[N1],B,C"),
            ("paragraph between two tables", p("A") + tbl(p("N1")) + p("B") + tbl(p("N2")) + emptyP,
             "A,T[N1],B,T[N2],<empty>"),
            ("table first", tbl(p("N1")) + emptyP, "T[N1],<empty>"),
            ("no nested table", p("A") + p("B"), "A,B"),
            ("a nested table holding a nested table",
             p("A") + tbl(p("X") + tbl(p("DEEP")) + p("Y")) + p("B"), "A,T[X,T[DEEP],Y],B"),
        ]
    }

    /// Editing row 0 rewrites the whole table from the typed model. The cell
    /// in row 1 must come back with its blocks in source order, and stay
    /// that way over repeated read / edit / save.
    func testEditingAnotherCellKeepsEveryCellInSourceOrder() throws {
        for shape in shapes {
            let source = documentXML(cell: shape.cell)
            XCTAssertEqual(try signature(source), shape.expected, "fixture: \(shape.label)")
            var url = try makeDocx(source)
            for generation in 1...3 {
                var doc = try DocxReader.read(from: url)
                try doc.updateCell(tableIndex: 0, row: 0, col: 0, text: "X\(generation)")
                let saved = try save(doc)
                doc.close()
                XCTAssertTrue(saved.xml.contains(">X\(generation)<"), "the edit must land (\(shape.label))")
                XCTAssertEqual(try signature(saved.xml), shape.expected,
                               "\(shape.label), generation \(generation)")
                url = saved.url
            }
        }
    }

    /// The sixth shape of the report: a paragraph with formatting but no
    /// run sits BEFORE the table, and the trailing empty paragraph after it.
    /// The run-less paragraph is the one the pre-#248 writer could mistake
    /// for the trailing placeholder. Its formatting must stay on the
    /// paragraph in front of the table, which the text-only signature above
    /// cannot see.
    func testAFormattedEmptyParagraphBeforeTheTableKeepsItsPlaceAndItsFormatting() throws {
        let centered = #"<w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p>"#
        var url = try makeDocx(documentXML(cell: p("A") + centered + tbl(p("N1")) + emptyP))
        for generation in 1...3 {
            var doc = try DocxReader.read(from: url)
            try doc.updateCell(tableIndex: 0, row: 0, col: 0, text: "X\(generation)")
            let saved = try save(doc)
            doc.close()
            XCTAssertEqual(try signature(saved.xml), "A,<empty>,T[N1],<empty>", "generation \(generation)")

            let parsed = try XMLDocument(xmlString: saved.xml)
            let body = try XCTUnwrap(parsed.rootElement()?.elements(forName: "w:body").first)
            let table = try XCTUnwrap(body.elements(forName: "w:tbl").first)
            let cell = try XCTUnwrap(table.elements(forName: "w:tr")[1].elements(forName: "w:tc").first)
            let blocks = (cell.children ?? []).compactMap { $0 as? XMLElement }
                .filter { $0.name == "w:p" || $0.name == "w:tbl" }
            XCTAssertEqual(blocks.count, 4, "generation \(generation)")
            let isCentered = blocks.map {
                $0.name == "w:p" && !($0.elements(forName: "w:pPr").first?.elements(forName: "w:jc").isEmpty ?? true)
            }
            XCTAssertEqual(isCentered, [false, true, false, false],
                           "the centering belongs to the second block only, generation \(generation)")
            url = saved.url
        }
    }

    /// The reader records the kinds in document order.
    func testTheReaderRecordsTheBlockKindsInDocumentOrder() throws {
        let doc = try DocxReader.read(from: makeDocx(documentXML(cell: p("A") + tbl(p("N1")) + p("B") + p("C"))))
        let cell = try cellUnderTest(doc)
        XCTAssertEqual(cell._blockOrder, [.paragraph, .table, .paragraph, .paragraph])
        XCTAssertEqual(cell.paragraphs.count, 3)
        XCTAssertEqual(cell.nestedTables.count, 1)
    }

    // MARK: - Whitespace inside a nested table

    /// Reading the cell's children in document order also keeps the
    /// whitespace overlay's ordinals in step. The two-pass read parsed every
    /// paragraph of the cell before the nested table, so a preserved
    /// whitespace-only `<w:t>` inside the nested table came back empty when
    /// a paragraph with text followed the table (shape 2b of #133).
    func testPreservedWhitespaceInsideANestedTableSurvivesAnEditToAnotherCell() throws {
        let spaced = #"<w:p><w:r><w:t xml:space="preserve">  </w:t></w:r><w:r><w:t>N1</w:t></w:r></w:p>"#
        var doc = try DocxReader.read(from: makeDocx(documentXML(cell: p("A") + tbl(spaced) + p("TAIL"))))
        try doc.updateCell(tableIndex: 0, row: 0, col: 0, text: "X1")
        let xml = try save(doc).xml
        XCTAssertTrue(xml.contains(#"<w:t xml:space="preserve">  </w:t>"#),
                      "the preserved whitespace inside the nested table must survive:\n\(xml)")
        // `XMLDocument` drops whitespace-only text, so the signature shows
        // only the order; the raw-byte check above is the whitespace check.
        XCTAssertEqual(try signature(xml), "A,T[N1],TAIL")
    }

    /// A child of the cell the reader does not model (here a block-level
    /// content control) is skipped, and its `<w:t>` elements still have to be
    /// counted or every later lookup in the whitespace overlay is off by
    /// that many. With the old two-pass read this shape came out right by
    /// coincidence — the paragraph after the table was parsed first and made
    /// up the difference — so a single-pass read that did not count the
    /// skipped child lost whitespace the previous release kept.
    func testASkippedChildBeforeANestedTableDoesNotShiftTheWhitespaceOverlay() throws {
        let sdt = "<w:sdt><w:sdtContent><w:p><w:r><w:t>S</w:t></w:r></w:p></w:sdtContent></w:sdt>"
        let spaced = #"<w:p><w:r><w:t xml:space="preserve">  </w:t></w:r><w:r><w:t>N1</w:t></w:r></w:p>"#
        var doc = try DocxReader.read(from: makeDocx(documentXML(cell: sdt + tbl(spaced) + p("Q"))))
        try doc.updateCell(tableIndex: 0, row: 0, col: 0, text: "X1")
        let xml = try save(doc).xml
        XCTAssertTrue(xml.contains(#"<w:t xml:space="preserve">  </w:t>"#),
                      "the preserved whitespace inside the nested table must survive:\n\(xml)")
    }

    /// The same, with the whitespace in a paragraph of the cell itself after
    /// the nested table. This shape lost its whitespace before this issue
    /// too.
    func testASkippedChildDoesNotShiftTheWhitespaceOfALaterParagraph() throws {
        let sdt = "<w:sdt><w:sdtContent><w:p><w:r><w:t>S</w:t></w:r></w:p></w:sdtContent></w:sdt>"
        let spaced = #"<w:p><w:r><w:t xml:space="preserve">  </w:t></w:r><w:r><w:t>Q</w:t></w:r></w:p>"#
        var doc = try DocxReader.read(from: makeDocx(documentXML(cell: sdt + p("A") + tbl(p("N1")) + spaced)))
        try doc.updateCell(tableIndex: 0, row: 0, col: 0, text: "X1")
        let xml = try save(doc).xml
        XCTAssertTrue(xml.contains(#"<w:t xml:space="preserve">  </w:t>"#),
                      "the preserved whitespace after the nested table must survive:\n\(xml)")
    }

    // MARK: - A cell must still end with a paragraph

    /// A cell whose only child is a table has no paragraph in the source.
    /// One is added at the END, and the next generations keep exactly that.
    func testACellHoldingOnlyANestedTableEndsWithOneParagraph() throws {
        var url = try makeDocx(documentXML(cell: tbl(p("N1"))))
        for generation in 1...3 {
            var doc = try DocxReader.read(from: url)
            try doc.updateCell(tableIndex: 0, row: 0, col: 0, text: "X\(generation)")
            let saved = try save(doc)
            doc.close()
            XCTAssertEqual(try signature(saved.xml), "T[N1],<empty>", "generation \(generation)")
            url = saved.url
        }
    }

    /// A cell whose last block is a table gets one paragraph after it; the
    /// paragraph before the table stays before it.
    func testACellEndingWithATableKeepsItsParagraphBeforeTheTable() throws {
        var doc = try DocxReader.read(from: makeDocx(documentXML(cell: p("A") + tbl(p("N1")))))
        try doc.updateCell(tableIndex: 0, row: 0, col: 0, text: "X1")
        XCTAssertEqual(try signature(save(doc).xml), "A,T[N1],<empty>")
    }

    // MARK: - Edits to the cell itself

    /// Replacing a paragraph in place keeps the recorded order valid.
    func testEditingAParagraphInPlaceKeepsTheOrder() throws {
        var doc = try DocxReader.read(from: makeDocx(documentXML(cell: p("A") + tbl(p("N1")) + p("B") + p("C"))))
        try doc.updateCellParagraph(tableIndex: 0, row: 1, col: 0, paragraphIndex: 1, text: "B2")
        XCTAssertEqual(try signature(save(doc).xml), "A,T[N1],B2,C")
    }

    /// `updateCell` rewrites the cell's first paragraph. In a cell that
    /// starts with a table, that paragraph is the one AFTER the table, and
    /// the text now stays there — which is where Word puts it when you type
    /// into that cell. Before this issue the flattened output moved the text
    /// in front of the table and added an empty paragraph behind it; the
    /// second half of the test shows that shape, which a cell with no
    /// record still has.
    func testUpdatingACellThatStartsWithATableWritesIntoTheParagraphAfterIt() throws {
        var doc = try DocxReader.read(from: makeDocx(documentXML(cell: tbl(p("N1")) + emptyP)))
        try doc.updateCell(tableIndex: 0, row: 1, col: 0, text: "X")
        XCTAssertEqual(try signature(save(doc).xml), "T[N1],X")

        var withoutRecord = try cellUnderTest(doc)
        withoutRecord._blockOrder = nil
        XCTAssertEqual(try cellSignature(withoutRecord), "X,T[N1],<empty>")
    }

    /// Adding a paragraph changes how many blocks the cell holds, so the
    /// recorded order no longer describes it. The cell falls back to the
    /// shape it had before this issue (paragraphs, tables, a trailing
    /// paragraph) instead of emitting against a stale order: nothing is
    /// lost and the cell still ends with a paragraph.
    func testAddingAParagraphFallsBackToTheFlattenedShape() throws {
        let doc = try DocxReader.read(from: makeDocx(documentXML(cell: p("A") + tbl(p("N1")) + p("B"))))
        var cell = try cellUnderTest(doc)
        cell.paragraphs.append(Paragraph(text: "NEW"))
        XCTAssertEqual(try cellSignature(cell), "A,B,NEW,T[N1],<empty>")
    }

    /// Adding a nested table falls back the same way.
    func testAddingANestedTableFallsBackToTheFlattenedShape() throws {
        let doc = try DocxReader.read(from: makeDocx(documentXML(cell: p("A") + tbl(p("N1")) + p("B"))))
        var cell = try cellUnderTest(doc)
        cell.nestedTables.append(Table(rows: [TableRow(cells: [TableCell(text: "N2")])]))
        XCTAssertEqual(try cellSignature(cell), "A,B,T[N1],T[N2],<empty>")
    }

    /// Removing one paragraph and adding another leaves the COUNT where it
    /// was, but the record no longer describes the cell. The record is
    /// dropped by the first of the two writes, so the cell falls back
    /// instead of putting a paragraph on the wrong side of the table.
    func testRemovingThenAddingAParagraphFallsBack() throws {
        let doc = try DocxReader.read(from: makeDocx(documentXML(cell: p("A") + tbl(p("N1")) + p("B"))))

        var front = try cellUnderTest(doc)
        front.paragraphs.removeLast()
        front.paragraphs.insert(Paragraph(text: "NEWFIRST"), at: 0)
        XCTAssertNil(front._blockOrder)
        XCTAssertEqual(try cellSignature(front), "NEWFIRST,A,T[N1],<empty>")

        var back = try cellUnderTest(doc)
        back.paragraphs.remove(at: 0)
        back.paragraphs.append(Paragraph(text: "C"))
        XCTAssertNil(back._blockOrder)
        XCTAssertEqual(try cellSignature(back), "B,C,T[N1],<empty>")
    }

    /// The same for nested tables.
    func testRemovingThenAddingANestedTableFallsBack() throws {
        let doc = try DocxReader.read(from: makeDocx(documentXML(cell: p("A") + tbl(p("N1")) + p("B"))))
        var cell = try cellUnderTest(doc)
        cell.nestedTables.removeAll()
        cell.nestedTables.append(Table(rows: [TableRow(cells: [TableCell(text: "N2")])]))
        XCTAssertNil(cell._blockOrder)
        XCTAssertEqual(try cellSignature(cell), "A,B,T[N2],<empty>")
    }

    /// Replacing the whole array with one of the SAME length keeps the
    /// record: the new paragraphs take the positions the old ones had. That
    /// is the same write an in-place edit makes, and the two cannot be told
    /// apart.
    func testReplacingTheParagraphsWithTheSameNumberKeepsThePositions() throws {
        let doc = try DocxReader.read(from: makeDocx(documentXML(cell: p("A") + tbl(p("N1")) + p("B"))))
        var cell = try cellUnderTest(doc)
        cell.paragraphs = [Paragraph(text: "X"), Paragraph(text: "Y")]
        XCTAssertEqual(try cellSignature(cell), "X,T[N1],Y")
    }

    // MARK: - What the record must not change

    /// The record is not part of a cell's value: a cell built by a caller
    /// equals the same content read from a file.
    func testTheRecordedOrderIsNotPartOfEquality() throws {
        let doc = try DocxReader.read(from: makeDocx(documentXML(cell: p("A") + p("B"))))
        var read = try cellUnderTest(doc)
        var built = TableCell(paragraphs: read.paragraphs, properties: read.properties)
        XCTAssertNotNil(read._blockOrder)
        XCTAssertNil(built._blockOrder, "a caller-built cell records no order")
        XCTAssertEqual(read, built)
        read._blockOrder = nil
        built._blockOrder = [.paragraph, .paragraph]
        XCTAssertEqual(read, built)
    }

    /// A caller-built cell has no record and keeps the shape it had before.
    func testACallerBuiltCellKeepsItsPreviousShape() throws {
        var cell = TableCell(paragraphs: [Paragraph(text: "A"), Paragraph()])
        cell.nestedTables = [Table(rows: [TableRow(cells: [TableCell(text: "N1")])])]
        XCTAssertEqual(try cellSignature(cell), "A,T[N1],<empty>")
    }

    // MARK: - Tree-backed cells

    /// A tree-backed cell reads its order from its node.
    func testATreeBackedCellEmitsInNodeOrder() throws {
        let doc = try DocxReader.read(
            from: makeDocx(documentXML(cell: p("A") + tbl(p("N1")) + p("B") + tbl(p("N2")) + p("C"))),
            wireTreeBackedViews: true)
        let cell = try cellUnderTest(doc)
        XCTAssertNotNil(cell.xmlNode, "fixture: the cell must be tree-backed")
        XCTAssertEqual(try cellSignature(cell), "A,T[N1],B,T[N2],C")
    }
}
