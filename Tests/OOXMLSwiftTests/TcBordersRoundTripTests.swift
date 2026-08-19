import XCTest
@testable import OOXMLSwift

/// Round-trip coverage for `<w:tcBorders>` (PsychQuant/ooxml-swift #99).
///
/// **Pre-fix bug**: `DocxReader.parseCellProperties` reads only the two
/// *diagonal* borders (`w:tl2br` / `w:tr2bl`) out of `<w:tcBorders>`. The four
/// *edge* borders (`w:top` / `w:bottom` / `w:left` / `w:right`) are never
/// parsed.
///
/// The loss is a **pure reader gap**, not a missing feature:
///   - `CellBorders` already declares `top` / `bottom` / `left` / `right`
///   - `TableCellProperties.toXML()` already emits all six directions
///
/// So any operation that marks the document dirty and re-serialises from the
/// model silently drops **every** cell border in the document — including
/// cells that were never touched. A zero-edit `open` → `save` is unaffected
/// because it replays raw bytes without going through the model, which is
/// exactly what made this invisible: the failure only appears after a real
/// edit, and no tool reports it.
///
/// Same shape as #84 (`DocxReader` never populates `sectionProperties`),
/// #67 and #69.
///
/// Downstream impact: PsychQuant/macdoc#142, PsychQuant/ntu-claude-plugins#16
/// (official submission forms rendered border-less after a single
/// `replace_text`).
final class TcBordersRoundTripTests: XCTestCase {

    private var tempDir: URL!

    override func setUp() {
        super.setUp()
        tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("TcBordersRoundTripTests-\(UUID().uuidString)")
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

    /// A one-cell table whose cell carries `borders`, plus a second bare cell
    /// so the row is not degenerate.
    private func document(withCellBorders borders: CellBorders) -> WordDocument {
        var props = TableCellProperties()
        props.borders = borders
        let table = Table(rows: [
            TableRow(cells: [
                TableCell(paragraphs: [Paragraph(text: "Bordered")], properties: props),
                TableCell(text: "Plain")
            ])
        ])
        var doc = WordDocument()
        doc.body.children = [.table(table)]
        return doc
    }

    private func firstCellBorders(of document: WordDocument) throws -> CellBorders? {
        guard case let .table(table)? = document.body.children.first else {
            XCTFail("expected a table as the first body child")
            return nil
        }
        guard let cell = table.rows.first?.cells.first else {
            XCTFail("expected at least one cell")
            return nil
        }
        return cell.properties.borders
    }

    // MARK: - Edge borders (the bug)

    func testEdgeBordersSurviveRoundTrip() throws {
        let border = Border(style: .single, size: 4, color: "000000")
        let doc = document(withCellBorders: .all(border))

        let borders = try firstCellBorders(of: roundTrip(doc))

        XCTAssertNotNil(borders, "cell borders were dropped entirely by the reader")
        XCTAssertEqual(borders?.top, border, "w:top not parsed from <w:tcBorders>")
        XCTAssertEqual(borders?.bottom, border, "w:bottom not parsed from <w:tcBorders>")
        XCTAssertEqual(borders?.left, border, "w:left not parsed from <w:tcBorders>")
        XCTAssertEqual(borders?.right, border, "w:right not parsed from <w:tcBorders>")
    }

    /// Edges must survive even when no diagonal is present — the pre-fix reader
    /// only allocated `CellBorders` when a diagonal was found, so a document
    /// with edges but no diagonals lost everything.
    func testEdgeBordersSurviveWithoutAnyDiagonal() throws {
        var borders = CellBorders()
        borders.left = Border(style: .double, size: 6, color: "FF0000")
        let doc = document(withCellBorders: borders)

        let result = try firstCellBorders(of: roundTrip(doc))

        XCTAssertEqual(result?.left, borders.left)
        XCTAssertNil(result?.tl2br)
        XCTAssertNil(result?.tr2bl)
    }

    /// Per-edge distinctness: a real Word form typically differs per edge
    /// (the #142 sample has `bottom` = `double` while the others are `single`,
    /// and `right` at a different width). A fix that copies one edge into all
    /// four would pass `testEdgeBordersSurviveRoundTrip` but fail here.
    func testPerEdgeValuesAreNotConflated() throws {
        var borders = CellBorders()
        borders.top = Border(style: .single, size: 4, color: "000000")
        borders.bottom = Border(style: .double, size: 4, color: "000000")
        borders.left = Border(style: .single, size: 4, color: "000000")
        borders.right = Border(style: .single, size: 2, color: "000000")
        let doc = document(withCellBorders: borders)

        let result = try firstCellBorders(of: roundTrip(doc))

        XCTAssertEqual(result?.top, borders.top)
        XCTAssertEqual(result?.bottom, borders.bottom, "style must stay per-edge (double)")
        XCTAssertEqual(result?.left, borders.left)
        XCTAssertEqual(result?.right, borders.right, "size must stay per-edge (2, not 4)")
    }

    // MARK: - Diagonals (control — passed before the fix)

    func testDiagonalBordersStillSurviveRoundTrip() throws {
        var borders = CellBorders()
        borders.tl2br = Border(style: .dashed, size: 8, color: "00FF00")
        borders.tr2bl = Border(style: .dotted, size: 2, color: "0000FF")
        let doc = document(withCellBorders: borders)

        let result = try firstCellBorders(of: roundTrip(doc))

        XCTAssertEqual(result?.tl2br, borders.tl2br)
        XCTAssertEqual(result?.tr2bl, borders.tr2bl)
    }

    // MARK: - Absence stays absence

    /// A cell with no `<w:tcBorders>` must not acquire one. `parseBorder`
    /// substitutes defaults (`single` / 4 / `000000`) for missing attributes,
    /// so an over-eager fix could invent borders where the source had none —
    /// the mirror-image silent infidelity.
    /// `CT_TcBorders` is an `xsd:sequence`: top, start|left, bottom, end|right,
    /// insideH, insideV, tl2br, tr2bl. The writer emitted top, bottom, left,
    /// right — the four edges all survive, but out of schema order, so the
    /// output is schema-invalid even though Word tolerates it in practice.
    ///
    /// Measured through the released MCP face before this test existed:
    ///   before: [top, left, bottom, right, insideH, insideV]
    ///   after:  [top, bottom, left, right]
    func testEdgeBordersAreEmittedInSchemaSequence() throws {
        var borders = CellBorders()
        borders.top = Border(style: .single, size: 8, color: "FF0000")
        borders.left = Border(style: .single, size: 8, color: "00FF00")
        borders.bottom = Border(style: .single, size: 8, color: "0000FF")
        borders.right = Border(style: .single, size: 8, color: "FFFF00")

        let docxURL = tempDir.appendingPathComponent("order-\(UUID().uuidString).docx")
        try DocxWriter.write(document(withCellBorders: borders), to: docxURL)

        let xml = String(decoding: try RawPartChannel.readAllParts(from: docxURL)["word/document.xml"]!,
                         as: UTF8.self)
        guard let range = xml.range(of: "<w:tcBorders>"),
              let end = xml.range(of: "</w:tcBorders>") else {
            return XCTFail("no <w:tcBorders> in the written document")
        }
        let block = String(xml[range.upperBound..<end.lowerBound])
        let order = ["top", "left", "bottom", "right"].map { name -> (String, Int) in
            (name, block.range(of: "<w:\(name)").map { block.distance(from: block.startIndex, to: $0.lowerBound) } ?? -1)
        }
        XCTAssertFalse(order.contains { $0.1 < 0 },
                       "every edge must be present: \(order)")
        let positions = order.map(\.1)
        XCTAssertEqual(positions, positions.sorted(),
                       "tcBorders children must follow the CT_TcBorders sequence "
                       + "(top, left, bottom, right); got \(order.map(\.0)) at \(positions)")
    }

    /// #99 residue: `insideH` / `insideV` had no model field at all, so the
    /// reader could not keep them and the writer could not put them back —
    /// the same three-sided absence #101 had for `<w:tcMar>`. Measured through
    /// the released MCP face: a cell carrying all six children came back with
    /// four.
    func testInsideBordersSurviveRoundTrip() throws {
        var borders = CellBorders()
        borders.top = Border(style: .single, size: 12, color: "111111")
        borders.left = Border(style: .single, size: 12, color: "222222")
        borders.bottom = Border(style: .single, size: 12, color: "333333")
        borders.right = Border(style: .single, size: 12, color: "444444")
        borders.insideH = Border(style: .single, size: 6, color: "555555")
        borders.insideV = Border(style: .single, size: 6, color: "666666")

        let recovered = try XCTUnwrap(
            firstCellBorders(of: try roundTrip(document(withCellBorders: borders))))

        XCTAssertEqual(recovered.insideH?.color, "555555", "insideH must survive the round trip")
        XCTAssertEqual(recovered.insideV?.color, "666666", "insideV must survive the round trip")
        XCTAssertEqual(recovered.insideH?.size, 6)
        XCTAssertEqual(recovered.insideV?.size, 6)
        // The four edges must not be disturbed by carrying the inside pair.
        XCTAssertEqual(recovered.top?.color, "111111")
        XCTAssertEqual(recovered.left?.color, "222222")
        XCTAssertEqual(recovered.bottom?.color, "333333")
        XCTAssertEqual(recovered.right?.color, "444444")
    }

    /// A cell that never carried inside borders must not gain them, for the
    /// same reason the existing no-borders test exists: `parseBorder`
    /// substitutes defaults for absent attributes.
    func testCellWithoutInsideBordersDoesNotGainThem() throws {
        var borders = CellBorders()
        borders.top = Border(style: .single, size: 8, color: "ABCDEF")

        let recovered = try XCTUnwrap(
            firstCellBorders(of: try roundTrip(document(withCellBorders: borders))))
        XCTAssertNil(recovered.insideH, "insideH must not be invented on read")
        XCTAssertNil(recovered.insideV, "insideV must not be invented on read")
    }

    func testCellWithoutBordersDoesNotGainThem() throws {
        var doc = WordDocument()
        doc.body.children = [.table(Table(rows: [
            TableRow(cells: [TableCell(text: "A"), TableCell(text: "B")])
        ]))]

        let result = try firstCellBorders(of: roundTrip(doc))

        XCTAssertNil(result, "a cell with no <w:tcBorders> must not gain borders on read")
    }
}
