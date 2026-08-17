import XCTest
@testable import OOXMLSwift

/// Round-trip coverage for `<w:tcMar>` (PsychQuant/ooxml-swift #101).
///
/// **Pre-fix bug**: `<w:tcMar>` existed in none of the three layers —
/// `grep -rn "tcMar" Sources/` returned nothing. `TableCellProperties` had no
/// field to hold it, `DocxReader` never parsed it, `toXML()` never emitted it.
/// So any operation that marked the document dirty dropped every cell's
/// margins, including cells that were never touched. A zero-edit
/// `open` → `save` was unaffected because it replays raw bytes without going
/// through the model — the same reason #99 stayed invisible.
///
/// This is the other half of the same `<w:tcPr>`: #99 / PR #100 restored
/// `<w:tcBorders>`; this restores `<w:tcMar>`.
///
/// The type is reused rather than invented: `TableCellMargins` already backs
/// the table-level `<w:tblCellMar>` and has the same four-twips shape.
///
/// Downstream impact: PsychQuant/macdoc#142, PsychQuant/ntu-claude-plugins#16.
final class TcMarRoundTripTests: XCTestCase {

    private var tempDir: URL!

    override func setUp() {
        super.setUp()
        tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("TcMarRoundTripTests-\(UUID().uuidString)")
        try? FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
    }

    override func tearDown() {
        try? FileManager.default.removeItem(at: tempDir)
        super.tearDown()
    }

    // MARK: - Helpers

    private func roundTrip(_ document: WordDocument) throws -> WordDocument {
        let docxURL = tempDir.appendingPathComponent("test-\(UUID().uuidString).docx")
        try DocxWriter.write(document, to: docxURL)
        return try DocxReader.read(from: docxURL)
    }

    private func document(withCellProperties props: TableCellProperties) -> WordDocument {
        let table = Table(rows: [
            TableRow(cells: [
                TableCell(paragraphs: [Paragraph(text: "Cell")], properties: props),
                TableCell(text: "Plain")
            ])
        ])
        var doc = WordDocument()
        doc.body.children = [.table(table)]
        return doc
    }

    private func firstCellProperties(of document: WordDocument) throws -> TableCellProperties? {
        guard case let .table(table)? = document.body.children.first else {
            XCTFail("expected a table as the first body child")
            return nil
        }
        guard let cell = table.rows.first?.cells.first else {
            XCTFail("expected at least one cell")
            return nil
        }
        return cell.properties
    }

    // MARK: - Margins survive (the bug)

    func testCellMarginsSurviveRoundTrip() throws {
        var props = TableCellProperties()
        props.margins = TableCellMargins(top: 20, bottom: 40, left: 108, right: 216)

        let result = try firstCellProperties(of: roundTrip(document(withCellProperties: props)))

        XCTAssertNotNil(result?.margins, "cell margins dropped entirely")
        XCTAssertEqual(result?.margins?.top, 20, "w:top not parsed from <w:tcMar>")
        XCTAssertEqual(result?.margins?.bottom, 40, "w:bottom not parsed from <w:tcMar>")
        XCTAssertEqual(result?.margins?.left, 108, "w:left not parsed from <w:tcMar>")
        XCTAssertEqual(result?.margins?.right, 216, "w:right not parsed from <w:tcMar>")
    }

    /// Per-edge distinctness: real forms differ per edge (the macdoc#142 sample
    /// has `top`/`bottom` at 0 and `left`/`right` at 10). A fix that copied one
    /// value into all four would pass a looser assertion but fail here.
    func testPerEdgeMarginValuesAreNotConflated() throws {
        var props = TableCellProperties()
        props.margins = TableCellMargins(top: 0, bottom: 0, left: 10, right: 10)

        let result = try firstCellProperties(of: roundTrip(document(withCellProperties: props)))

        XCTAssertEqual(result?.margins?.top, 0)
        XCTAssertEqual(result?.margins?.left, 10, "left must stay 10, not be overwritten by top")
        XCTAssertEqual(result?.margins?.right, 10)
    }

    /// A partially-specified `<w:tcMar>` must not have its absent edges filled
    /// in with invented values.
    func testPartialMarginsKeepAbsentEdgesNil() throws {
        var props = TableCellProperties()
        props.margins = TableCellMargins(left: 108)

        let result = try firstCellProperties(of: roundTrip(document(withCellProperties: props)))

        XCTAssertEqual(result?.margins?.left, 108)
        XCTAssertNil(result?.margins?.top, "absent w:top must stay nil, not default to 0")
        XCTAssertNil(result?.margins?.bottom, "absent w:bottom must stay nil")
        XCTAssertNil(result?.margins?.right, "absent w:right must stay nil")
    }

    // MARK: - Absence stays absence

    func testCellWithoutMarginsDoesNotGainThem() throws {
        var doc = WordDocument()
        doc.body.children = [.table(Table(rows: [
            TableRow(cells: [TableCell(text: "A"), TableCell(text: "B")])
        ]))]

        let result = try firstCellProperties(of: roundTrip(doc))

        XCTAssertNil(result?.margins, "a cell with no <w:tcMar> must not gain margins on read")
    }

    /// #99's borders must keep working alongside the new margins — both live in
    /// the same `<w:tcPr>` and an ordering mistake could drop one of them.
    func testBordersAndMarginsCoexist() throws {
        var props = TableCellProperties()
        props.borders = .all(Border(style: .single, size: 4, color: "000000"))
        props.margins = TableCellMargins.all(108)

        let result = try firstCellProperties(of: roundTrip(document(withCellProperties: props)))

        XCTAssertNotNil(result?.borders?.top, "borders lost when margins are present")
        XCTAssertEqual(result?.margins?.top, 108, "margins lost when borders are present")
    }

    // MARK: - Schema order

    /// `CT_TcPr` is a sequence. ECMA-376 order (subset in use here):
    /// `tcW → gridSpan → vMerge → tcBorders → shd → tcMar → vAlign`.
    ///
    /// Pre-fix the writer emitted `vAlign` before `tcBorders`/`shd`, which was
    /// already out of order; `tcMar`'s correct slot sits between `shd` and
    /// `vAlign`, so it cannot be inserted correctly without settling `vAlign`.
    func testTcPrChildOrderFollowsSchemaSequence() {
        var props = TableCellProperties()
        props.width = 1701
        props.widthType = .dxa
        props.gridSpan = 2
        props.verticalMerge = .restart
        props.borders = .all(Border(style: .single, size: 4, color: "000000"))
        props.shading = CellShading(fill: "auto")
        props.margins = TableCellMargins.all(108)
        props.verticalAlignment = .center

        let xml = props.toXML()

        let expectedOrder = ["<w:tcW", "<w:gridSpan", "<w:vMerge", "<w:tcBorders", "<w:shd", "<w:tcMar", "<w:vAlign"]
        var positions: [(String, Int)] = []
        for token in expectedOrder {
            guard let range = xml.range(of: token) else {
                XCTFail("\(token) missing from <w:tcPr>:\n\(xml)")
                return
            }
            positions.append((token, xml.distance(from: xml.startIndex, to: range.lowerBound)))
        }
        for i in 1..<positions.count {
            XCTAssertLessThan(
                positions[i - 1].1, positions[i].1,
                "\(positions[i - 1].0) must precede \(positions[i].0) per CT_TcPr sequence:\n\(xml)"
            )
        }
    }
}
