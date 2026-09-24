import XCTest
@testable import OOXMLSwift

/// PsychQuant/macdoc#156: address one paragraph inside a table cell. A
/// multi-paragraph cell must not collapse, sibling paragraphs and the pPr
/// properties the typed model knows stay as they were, and a marker
/// repeated in other rows is never touched.
///
/// Not covered here: pPr children the typed model does not model (for
/// example `w:kinsoku`, `w:snapToGrid`). Typed edits re-emit document.xml
/// from the typed model and drop those children across the whole document
/// — a pre-existing gap of the typed write path, tracked as #168.
final class CellParagraphAddressingTests: XCTestCase {
    private func directory() throws -> URL {
        let url = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString)
        try FileManager.default.createDirectory(at: url, withIntermediateDirectories: true)
        addTeardownBlock { try? FileManager.default.removeItem(at: url) }
        return url
    }

    /// `texts` become consecutive runs: the first bold 14 pt, the rest italic.
    private func paragraph(_ texts: [String], _ configure: (inout ParagraphProperties) -> Void) -> Paragraph {
        let runs = texts.enumerated().map { index, text -> Run in
            var run = Run(text: text)
            if index == 0 { run.properties.bold = true; run.properties.fontSize = 28 }
            else { run.properties.italic = true }
            return run
        }
        var paragraph = Paragraph(runs: runs)
        configure(&paragraph.properties)
        return paragraph
    }

    /// Row 0 and row 2 repeat the same markers; row 1 column 1 is the
    /// four-paragraph checkbox cell from the issue, each paragraph with its
    /// own pPr, the third split across two differently formatted runs.
    private func formDocument() -> WordDocument {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "Before the table"))
        let checkboxes = [
            paragraph(["□(1)"]) { $0.indentation = Indentation(left: 240) },
            paragraph(["□(2)"]) { $0.alignment = .center },
            paragraph(["□", "(3)"]) { $0.spacing = Spacing(before: 120, after: 60) },
            paragraph(["□無報酬"]) { $0.keepNext = true }
        ]
        doc.appendTable(Table(rows: [
            TableRow(cells: [TableCell(text: "Label A"), TableCell(paragraphs: [Paragraph(text: "□(1)"), Paragraph(text: "□(2)")])]),
            TableRow(cells: [TableCell(text: "Label B"), TableCell(paragraphs: checkboxes)]),
            TableRow(cells: [TableCell(text: "Label C"), TableCell(paragraphs: [Paragraph(text: "□(1)")])])
        ]))
        return doc
    }

    private func cell(_ doc: WordDocument, row: Int, col: Int) -> TableCell {
        doc.getTables()[0].rows[row].cells[col]
    }

    private func assertOnlyParagraphChanged(_ doc: WordDocument, original: WordDocument, index: Int, text: String,
                                            file: StaticString = #filePath, line: UInt = #line) throws {
        let before = cell(original, row: 1, col: 1).paragraphs
        let after = cell(doc, row: 1, col: 1).paragraphs
        XCTAssertEqual(after.count, before.count, "cell must keep every paragraph", file: file, line: line)
        for (k, paragraph) in after.enumerated() {
            XCTAssertEqual(paragraph.properties, before[k].properties, "pPr of paragraph \(k)", file: file, line: line)
            if k == index {
                XCTAssertEqual(paragraph.runs.count, 1, file: file, line: line)
                XCTAssertEqual(paragraph.runs.first?.text, text, file: file, line: line)
                XCTAssertEqual(paragraph.runs.first?.properties, before[k].runs.first?.properties, "first run's formatting", file: file, line: line)
            } else {
                XCTAssertEqual(paragraph.runs.map(\.text), before[k].runs.map(\.text), "paragraph \(k) text", file: file, line: line)
                XCTAssertEqual(paragraph.runs.map(\.properties), before[k].runs.map(\.properties), "paragraph \(k) runs", file: file, line: line)
            }
        }
        for row in [0, 2] {
            XCTAssertEqual(try doc.cellParagraphTexts(tableIndex: 0, row: row, col: 1),
                           try original.cellParagraphTexts(tableIndex: 0, row: row, col: 1), "row \(row)", file: file, line: line)
        }
    }

    func testReadsEveryParagraphOfACell() throws {
        let doc = formDocument()
        XCTAssertEqual(try doc.cellParagraphTexts(tableIndex: 0, row: 1, col: 1), ["□(1)", "□(2)", "□(3)", "□無報酬"])
        XCTAssertEqual(try doc.cellParagraphTexts(tableIndex: 0, row: 0, col: 1), ["□(1)", "□(2)"])
        XCTAssertEqual(try doc.cellParagraphTexts(tableIndex: 0, row: 1, col: 0), ["Label B"])
    }

    func testUpdatesOneParagraphOfAnInMemoryCellAndSurvivesSaveAndReload() throws {
        let original = formDocument()
        var doc = original
        try doc.updateCellParagraph(tableIndex: 0, row: 1, col: 1, paragraphIndex: 0, text: "V(1)")
        XCTAssertEqual(try doc.cellParagraphTexts(tableIndex: 0, row: 1, col: 1), ["V(1)", "□(2)", "□(3)", "□無報酬"])
        try assertOnlyParagraphChanged(doc, original: original, index: 0, text: "V(1)")

        let url = try directory().appendingPathComponent("memory.docx")
        try DocxWriter.write(doc, to: url)
        var reloaded = try DocxReader.read(from: url)
        defer { reloaded.close() }
        try assertOnlyParagraphChanged(reloaded, original: original, index: 0, text: "V(1)")
    }

    /// DocxReader produces detached cells (no xmlNode): the edit goes
    /// through the typed model, and the known pPr properties (indentation,
    /// alignment, spacing, keepNext) and the sibling paragraphs survive the
    /// writer's re-emission and a reload. Unmodeled pPr children are outside
    /// this test's scope; see #168.
    func testReaderProducedCellKeepsKnownPPrAndSiblingParagraphsAcrossSaveAndReload() throws {
        let root = try directory()
        let seed = root.appendingPathComponent("seed.docx")
        try DocxWriter.write(formDocument(), to: seed)
        var original = try DocxReader.read(from: seed)
        defer { original.close() }
        XCTAssertNil(original.getTables()[0].rows[1].cells[1].xmlNode, "reader cells are expected to be detached")
        XCTAssertNil(original.getTables()[0].rows[1].cells[1].paragraphs[2].xmlNode)

        var doc = try DocxReader.read(from: seed)
        defer { doc.close() }
        try doc.updateCellParagraph(tableIndex: 0, row: 1, col: 1, paragraphIndex: 2, text: "V(3)")
        XCTAssertEqual(try doc.cellParagraphTexts(tableIndex: 0, row: 1, col: 1), ["□(1)", "□(2)", "V(3)", "□無報酬"])
        try assertOnlyParagraphChanged(doc, original: original, index: 2, text: "V(3)")

        let output = root.appendingPathComponent("edited.docx")
        try DocxWriter.write(doc, to: output)
        var reloaded = try DocxReader.read(from: output)
        defer { reloaded.close() }
        try assertOnlyParagraphChanged(reloaded, original: original, index: 2, text: "V(3)")
        let xml = String(decoding: try XCTUnwrap(RawPartChannel.readAllParts(from: output)["word/document.xml"]), as: UTF8.self)
        for token in ["<w:ind w:left=\"240\"/>", "<w:jc w:val=\"center\"/>", "w:before=\"120\"", "<w:keepNext/>", ">V(3)</w:t>"] {
            XCTAssertTrue(xml.contains(token), token)
        }
        XCTAssertEqual(xml.components(separatedBy: ">□(1)</w:t>").count - 1, 3, "the repeated marker in other rows is untouched")
    }

    func testEmptyParagraphGetsAPlainRunAndKeepsItsProperties() throws {
        var doc = WordDocument()
        var empty = Paragraph()
        empty.properties.alignment = .right
        doc.appendTable(Table(rows: [TableRow(cells: [TableCell(paragraphs: [Paragraph(text: "keep"), empty])])]))
        try doc.updateCellParagraph(tableIndex: 0, row: 0, col: 0, paragraphIndex: 1, text: "filled")
        let paragraphs = doc.getTables()[0].rows[0].cells[0].paragraphs
        XCTAssertEqual(paragraphs.map { $0.runs.map(\.text).joined() }, ["keep", "filled"])
        XCTAssertEqual(paragraphs[1].properties.alignment, .right)
        XCTAssertEqual(paragraphs[1].runs.first?.properties, RunProperties())
    }

    /// Every coordinate is checked before anything is written: the value is
    /// unchanged and document.xml is not marked dirty.
    func testInvalidAddressesThrowBeforeAnyWrite() throws {
        let root = try directory()
        let seed = root.appendingPathComponent("seed.docx")
        try DocxWriter.write(formDocument(), to: seed)
        var doc = try DocxReader.read(from: seed)
        defer { doc.close() }
        let before = doc.getTables()
        let addresses: [(table: Int, row: Int, col: Int, paragraph: Int)] = [
            (-1, 1, 1, 0), (1, 1, 1, 0), (0, -1, 1, 0), (0, 3, 1, 0),
            (0, 1, -1, 0), (0, 1, 2, 0), (0, 1, 1, -1), (0, 1, 1, 4)
        ]
        for address in addresses {
            XCTAssertThrowsError(try doc.updateCellParagraph(tableIndex: address.table, row: address.row, col: address.col,
                                                             paragraphIndex: address.paragraph, text: "X"), "\(address)") { error in
                switch error {
                case WordError.invalidIndex, WordError.invalidFormat: break
                default: XCTFail("unexpected error \(error) for \(address)")
                }
            }
            if address.paragraph == 0 {
                XCTAssertThrowsError(try doc.cellParagraphTexts(tableIndex: address.table, row: address.row, col: address.col), "\(address)")
            }
        }
        XCTAssertEqual(doc.getTables(), before)
        XCTAssertFalse(doc.modifiedParts.contains("word/document.xml"))
        XCTAssertThrowsError(try doc.updateCellParagraph(tableIndex: 0, row: 1, col: 1, paragraphIndex: 4, text: "X")) { error in
            guard case WordError.invalidFormat(let message) = error else { return XCTFail("\(error)") }
            XCTAssertTrue(message.contains("4 paragraph"), message)
        }
    }

    /// A tree-backed table ignores typed paragraph writes; refusing is the
    /// only honest answer until tree-level cell editing exists.
    func testTreeBackedTableIsRefusedInsteadOfSilentlyIgnored() throws {
        let w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        let tree = try XmlTreeReader.parse(Data("<w:tbl xmlns:w=\"\(w)\"><w:tr><w:tc><w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p></w:tc></w:tr></w:tbl>".utf8))
        var doc = WordDocument()
        doc.body.children = [.table(Table(xmlNode: tree.root))]
        XCTAssertEqual(try doc.cellParagraphTexts(tableIndex: 0, row: 0, col: 0), ["one", "two"])
        XCTAssertThrowsError(try doc.updateCellParagraph(tableIndex: 0, row: 0, col: 0, paragraphIndex: 1, text: "2"))
        XCTAssertEqual(try doc.cellParagraphTexts(tableIndex: 0, row: 0, col: 0), ["one", "two"])
    }
}
