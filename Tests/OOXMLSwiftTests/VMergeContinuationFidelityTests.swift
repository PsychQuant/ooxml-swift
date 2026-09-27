import Foundation
import XCTest
@testable import OOXMLSwift

/// Regression suite for PsychQuant/che-word-mcp#185: `save_document` stripped
/// a vertical-merge continuation (`<w:vMerge/>` with no `w:val`) from every
/// untouched row once *any* other cell in the document was edited.
///
/// Root cause lived in `DocxReader.parseTableCellProperties`: the guard
/// chain required `w:val` to be present before recording `verticalMerge` at
/// all, so a bare `<w:vMerge/>` (the common continuation form OOXML allows
/// per ECMA-376 Part 1's `w:vMerge` definition — absence of `w:val` means "continue") was read as
/// `nil`. Because `DocxWriter` re-serializes the *entire* `document.xml`
/// from the typed model as soon as any table mutation marks the part dirty,
/// that "read as nil" silently deleted the continuation from every row in
/// the file, not just the one being edited.
final class VMergeContinuationFidelityTests: XCTestCase {

    // MARK: - Unit-level: writer shape for `.continue`

    /// che-word-mcp#185 follow-up: once `.continue` round-trips at all, decide what byte
    /// shape the writer emits for it. We choose the bare self-closing form
    /// (no `w:val`) rather than `w:val="continue"` because that is the form
    /// real documents already carry (Word, python-docx, REC-P-011-style
    /// government templates); emitting it back unchanged minimizes the byte
    /// diff on every row a save doesn't otherwise touch. `.restart` still
    /// requires an explicit `w:val` — OOXML's implied default is
    /// "continue", so "restart" can never be inferred from absence.
    func testContinueVerticalMergeWritesBareSelfClosingForm() {
        var cellProps = TableCellProperties()
        cellProps.verticalMerge = .continue
        let cell = TableCell(paragraphs: [Paragraph()], properties: cellProps)
        let xml = cell.toXML()
        XCTAssertTrue(xml.contains("<w:vMerge/>"), xml)
        XCTAssertFalse(xml.contains("w:val=\"continue\""), xml)
    }

    func testRestartVerticalMergeStillWritesExplicitValAttribute() {
        var cellProps = TableCellProperties()
        cellProps.verticalMerge = .restart
        let cell = TableCell(paragraphs: [Paragraph()], properties: cellProps)
        let xml = cell.toXML()
        XCTAssertTrue(xml.contains("<w:vMerge w:val=\"restart\"/>"), xml)
    }

    // MARK: - Round-trip: the scenario from the issue

    /// Builds a 3-row, 1-column table where row 0 starts a vertical merge
    /// (`<w:vMerge w:val="restart"/>`) and row 1 continues it with the bare
    /// form (`<w:vMerge/>`, no `w:val` — injected by patching the seed's
    /// raw XML, since the typed writer itself never emitted this form
    /// pre-fix). Row 2 is a plain, unrelated cell. We then edit *only* row 2
    /// and assert row 0/row 1's merge markers survive the save.
    func testEditingUnrelatedRowPreservesBareVerticalMergeContinuation() throws {
        let directory = FileManager.default.temporaryDirectory
            .appendingPathComponent("vmerge-continuation-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: directory, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: directory) }

        var seed = WordDocument()
        seed.appendTable(Table(rows: [
            TableRow(cells: [TableCell(text: "R0")]),
            TableRow(cells: [TableCell(text: "R1")]),
            TableRow(cells: [TableCell(text: "R2")]),
        ]))
        let seedURL = directory.appendingPathComponent("seed.docx")
        try DocxWriter.write(seed, to: seedURL)

        let extracted = try ZipHelper.unzip(seedURL)
        defer { ZipHelper.cleanup(extracted) }
        let documentURL = extracted.appendingPathComponent("word/document.xml")
        var xml = try String(contentsOf: documentURL, encoding: .utf8)

        // Row 0 becomes the merge's restart cell.
        xml = try replacingSoleEmptyTcPr(
            precedingAnchor: ">R0<", in: xml,
            with: "<w:tcPr><w:vMerge w:val=\"restart\"/></w:tcPr>")
        // Row 1 becomes a bare (no w:val) continuation — the form the bug
        // dropped. This mirrors what a real Word / python-docx document
        // carries; the typed writer alone could not have produced it
        // pre-fix, which is exactly why it needs to be injected here.
        xml = try replacingSoleEmptyTcPr(
            precedingAnchor: ">R1<", in: xml,
            with: "<w:tcPr><w:vMerge/></w:tcPr>")
        try xml.write(to: documentURL, atomically: true, encoding: .utf8)

        let inputURL = directory.appendingPathComponent("input.docx")
        try ZipHelper.zip(extracted, to: inputURL)

        var document = try DocxReader.read(from: inputURL)
        defer { document.close() }

        // Sanity check: the reader must have actually picked up both merge
        // markers before we exercise the "does an unrelated edit keep them"
        // question — otherwise a failure below wouldn't tell us which half
        // of the pipeline (read vs. write) is at fault.
        let seededTable = document.getTables()[0]
        XCTAssertEqual(seededTable.rows[0].cells[0].properties.verticalMerge, .restart)
        XCTAssertEqual(seededTable.rows[1].cells[0].properties.verticalMerge, .continue)

        // The only edit in this session — row 2, nowhere near the merge.
        try document.updateCell(tableIndex: 0, row: 2, col: 0, text: "R2-edited")

        let outputURL = directory.appendingPathComponent("output.docx")
        try DocxWriter.write(document, to: outputURL)
        let outputPackage = try ZipHelper.unzip(outputURL)
        defer { ZipHelper.cleanup(outputPackage) }
        let outputXML = try String(
            contentsOf: outputPackage.appendingPathComponent("word/document.xml"),
            encoding: .utf8)

        XCTAssertTrue(
            outputXML.contains("<w:vMerge w:val=\"restart\"/>"),
            "row 0's merge restart must survive an edit made elsewhere in the table")
        XCTAssertTrue(
            outputXML.contains("<w:vMerge/>"),
            "row 1's untouched continuation must survive an edit made elsewhere in the table")
        XCTAssertTrue(outputXML.contains(">R0<"))
        XCTAssertTrue(outputXML.contains(">R1<"))
        XCTAssertTrue(outputXML.contains(">R2-edited<"))
    }

    // MARK: - Helpers

    private enum FixtureError: Error {
        case anchorNotFound(String)
        case emptyTcPrNotFound(String)
    }

    /// Locates the empty `<w:tcPr></w:tcPr>` immediately preceding `anchor`
    /// (i.e. the owning cell's own, still-untouched, properties element) and
    /// replaces just that one occurrence. Cells built by the seed via
    /// `TableCell(text:)` all emit the identical empty-properties shape, so
    /// a global string replace would hit every cell; anchoring on each
    /// cell's own unique text and searching backwards keeps the patch
    /// scoped to exactly one `<w:tc>`.
    private func replacingSoleEmptyTcPr(
        precedingAnchor anchor: String, in xml: String, with replacement: String
    ) throws -> String {
        guard let anchorRange = xml.range(of: anchor) else {
            throw FixtureError.anchorNotFound(anchor)
        }
        guard let tcPrRange = xml.range(
            of: "<w:tcPr></w:tcPr>",
            options: .backwards,
            range: xml.startIndex..<anchorRange.lowerBound
        ) else {
            throw FixtureError.emptyTcPrNotFound(anchor)
        }
        var result = xml
        result.replaceSubrange(tcPrRange, with: replacement)
        return result
    }
}
