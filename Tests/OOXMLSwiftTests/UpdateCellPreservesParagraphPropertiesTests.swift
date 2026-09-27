import Foundation
import XCTest
@testable import OOXMLSwift

/// Regression suite for PsychQuant/che-word-mcp#224, sub-problem (a):
/// `WordDocument.updateCell` dropped the target paragraph's `pPr` whenever
/// the target cell had no existing run — the common shape for an empty,
/// fillable cell in an official form template (checkbox / blank-to-fill
/// column), which still carries a `pPr` (style, spacing, east-Asian font
/// defaults) even though it has zero runs.
///
/// Root cause: `WordDocument.updateCell`'s "cell has no run" branch replaced
/// the whole `cell.paragraphs` array with a brand-new `Paragraph(text:)`,
/// discarding the pre-existing paragraph (and its `properties`/`pPr`)
/// entirely. `updateCellParagraph` (added for che-word-mcp#194) already gets this right
/// — it mutates the existing paragraph's `.runs` in place and leaves
/// `.properties` untouched. This suite locks `updateCell` onto that same
/// shape.
final class UpdateCellPreservesParagraphPropertiesTests: XCTestCase {

    /// Unit-level: exercises `WordDocument.updateCell` directly against an
    /// in-memory document, with no docx round-trip. Isolates the mutation
    /// function itself from the read/write pipeline.
    func testUpdateCellOnRunlessParagraphPreservesExistingProperties() throws {
        var richProperties = ParagraphProperties()
        richProperties.alignment = .both
        richProperties.style = "Standard"
        var runlessParagraph = Paragraph()
        runlessParagraph.properties = richProperties
        XCTAssertTrue(runlessParagraph.runs.isEmpty, "fixture must start with zero runs")

        var document = WordDocument()
        document.appendTable(Table(rows: [
            TableRow(cells: [TableCell(paragraphs: [runlessParagraph])])
        ]))

        try document.updateCell(tableIndex: 0, row: 0, col: 0, text: "V")

        let updatedParagraph = document.getTables()[0].rows[0].cells[0].paragraphs[0]
        XCTAssertEqual(updatedParagraph.properties.alignment, .both,
                       "pPr alignment must survive filling in a previously run-less paragraph")
        XCTAssertEqual(updatedParagraph.properties.style, "Standard")
        XCTAssertEqual(updatedParagraph.getText(), "V")
    }

    /// Round-trip: the scenario from the issue. An empty cell's paragraph
    /// carries a rich `pPr` (style, snapToGrid, exact line spacing,
    /// justification, and an east-Asian run-default font/size inside
    /// `pPr/rPr`) but has zero runs — exactly the shape python-docx and
    /// Word leave behind for an unfilled form field. `update_cell` targets
    /// that cell; the pPr must still be there after save.
    func testUpdateCellOnEmptyFormFieldPreservesRichParagraphPropertiesAcrossSave() throws {
        let directory = FileManager.default.temporaryDirectory
            .appendingPathComponent("update-cell-ppr-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: directory, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: directory) }

        var seed = WordDocument()
        seed.appendTable(Table(rows: [
            TableRow(cells: [TableCell(text: "Label"), TableCell(paragraphs: [Paragraph()])]),
            TableRow(cells: [TableCell(text: "Other"), TableCell(text: "OtherVal")]),
        ]))
        let seedURL = directory.appendingPathComponent("seed.docx")
        try DocxWriter.write(seed, to: seedURL)

        let extracted = try ZipHelper.unzip(seedURL)
        defer { ZipHelper.cleanup(extracted) }
        let documentURL = extracted.appendingPathComponent("word/document.xml")
        var xml = try String(contentsOf: documentURL, encoding: .utf8)

        // The only cell with a run-less paragraph is row 0 / col 1, so the
        // empty-cell shape `<w:tcPr></w:tcPr><w:p></w:p>` appears exactly
        // once in the whole document — safe to replace globally.
        let emptyCellShape = "<w:tcPr></w:tcPr><w:p></w:p>"
        XCTAssertEqual(xml.components(separatedBy: emptyCellShape).count - 1, 1,
                       "fixture must have exactly one run-less cell to patch")
        let richPPr = """
        <w:pPr><w:pStyle w:val="Standard"/><w:snapToGrid w:val="0"/>\
        <w:spacing w:line="280" w:lineRule="exact"/><w:jc w:val="both"/>\
        <w:rPr><w:rFonts w:ascii="Book Antiqua" w:eastAsia="標楷體" w:hAnsi="Book Antiqua"/>\
        <w:sz w:val="22"/></w:rPr></w:pPr>
        """
        xml = xml.replacingOccurrences(
            of: emptyCellShape,
            with: "<w:tcPr></w:tcPr><w:p>\(richPPr)</w:p>")
        try xml.write(to: documentURL, atomically: true, encoding: .utf8)

        let inputURL = directory.appendingPathComponent("input.docx")
        try ZipHelper.zip(extracted, to: inputURL)

        var document = try DocxReader.read(from: inputURL)
        defer { document.close() }

        try document.updateCell(tableIndex: 0, row: 0, col: 1, text: "V")

        let outputURL = directory.appendingPathComponent("output.docx")
        try DocxWriter.write(document, to: outputURL)
        let outputPackage = try ZipHelper.unzip(outputURL)
        defer { ZipHelper.cleanup(outputPackage) }
        let outputXML = try String(
            contentsOf: outputPackage.appendingPathComponent("word/document.xml"),
            encoding: .utf8)

        // Attribute order inside a single element is the writer's own
        // canonical choice (not part of what this fix is about), so check
        // for the presence of each value rather than one exact literal
        // `<w:rFonts .../>` tag.
        for token in [
            "<w:pStyle w:val=\"Standard\"/>",
            "<w:snapToGrid w:val=\"0\"/>",
            "<w:spacing w:line=\"280\" w:lineRule=\"exact\"/>",
            "<w:jc w:val=\"both\"/>",
            "<w:rFonts", "w:ascii=\"Book Antiqua\"", "w:eastAsia=\"標楷體\"", "w:hAnsi=\"Book Antiqua\"",
            "<w:sz w:val=\"22\"/>",
        ] {
            XCTAssertTrue(outputXML.contains(token), "missing paragraph property: \(token)")
        }
        XCTAssertTrue(outputXML.contains(">V<"), "the new run text must still be inserted")
        let pPrRange = try XCTUnwrap(outputXML.range(of: "<w:pPr>"))
        let runRange = try XCTUnwrap(outputXML.range(of: ">V<"))
        XCTAssertLessThan(pPrRange.lowerBound, runRange.lowerBound,
                           "pPr must precede the run inside the paragraph")

        // The other row was never touched by update_cell and must be
        // completely unaffected.
        XCTAssertTrue(outputXML.contains(">Other<"))
        XCTAssertTrue(outputXML.contains(">OtherVal<"))
    }
}
