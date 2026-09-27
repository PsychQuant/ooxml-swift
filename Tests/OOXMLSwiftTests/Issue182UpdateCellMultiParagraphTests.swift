import Foundation
import XCTest
@testable import OOXMLSwift

/// Regression tests for [PsychQuant/ooxml-swift#182](https://github.com/PsychQuant/ooxml-swift/issues/182).
///
/// `WordDocument.updateCell` on a cell with MULTIPLE paragraphs replaced the
/// whole `cell.paragraphs` array with a single freshly-built paragraph,
/// silently deleting every paragraph after the first while reporting
/// success. Decision recorded here (issue's "Option A"): `updateCell` edits
/// only paragraph 0; every other paragraph in the cell survives unchanged.
/// Callers that need to address a specific paragraph use
/// `updateCellParagraph` (che-word-mcp#194).
final class Issue182UpdateCellMultiParagraphTests: XCTestCase {

    /// Unit-level: exercises `updateCell` directly against an in-memory
    /// document, no docx round-trip. Isolates the mutation function itself.
    func testUpdateCellOnMultiParagraphCellPreservesOtherParagraphs() throws {
        var document = WordDocument()
        document.appendTable(Table(rows: [
            TableRow(cells: [
                TableCell(paragraphs: [
                    Paragraph(text: "First"),
                    Paragraph(text: "Second"),
                    Paragraph(text: "Third"),
                ])
            ])
        ]))

        try document.updateCell(tableIndex: 0, row: 0, col: 0, text: "Updated")

        let paragraphs = document.getTables()[0].rows[0].cells[0].paragraphs
        XCTAssertEqual(paragraphs.count, 3,
                       "updateCell must not delete the other paragraphs — got \(paragraphs.count)")
        XCTAssertEqual(paragraphs[0].getText(), "Updated", "paragraph 0 is the one updateCell targets")
        XCTAssertEqual(paragraphs[1].getText(), "Second", "paragraph 1 must survive unchanged")
        XCTAssertEqual(paragraphs[2].getText(), "Third", "paragraph 2 must survive unchanged")
    }

    /// Same guarantee when the target (first) paragraph has no run yet
    /// (che-word-mcp#224a shape) — the "cell has no run" branch must also
    /// only touch index 0.
    func testUpdateCellOnMultiParagraphCellWithRunlessFirstParagraphPreservesOtherParagraphs() throws {
        var document = WordDocument()
        document.appendTable(Table(rows: [
            TableRow(cells: [
                TableCell(paragraphs: [
                    Paragraph(),  // run-less first paragraph
                    Paragraph(text: "Second"),
                ])
            ])
        ]))
        XCTAssertTrue(document.getTables()[0].rows[0].cells[0].paragraphs[0].runs.isEmpty,
                     "fixture must start with a run-less first paragraph")

        try document.updateCell(tableIndex: 0, row: 0, col: 0, text: "Filled")

        let paragraphs = document.getTables()[0].rows[0].cells[0].paragraphs
        XCTAssertEqual(paragraphs.count, 2, "updateCell must not delete paragraph 1")
        XCTAssertEqual(paragraphs[0].getText(), "Filled")
        XCTAssertEqual(paragraphs[1].getText(), "Second", "paragraph 1 must survive unchanged")
    }

    // MARK: - Full-docx regression: the fix survives a save/reload cycle

    func testUpdateCellOnMultiParagraphCellPreservesOtherParagraphsAcrossSave() throws {
        let directory = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue182-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: directory, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: directory) }

        var seed = WordDocument()
        seed.appendTable(Table(rows: [
            TableRow(cells: [
                TableCell(paragraphs: [
                    Paragraph(text: "First"),
                    Paragraph(text: "Second"),
                    Paragraph(text: "Third"),
                ])
            ])
        ]))
        let seedURL = directory.appendingPathComponent("seed.docx")
        try DocxWriter.write(seed, to: seedURL)

        var document = try DocxReader.read(from: seedURL)
        defer { document.close() }
        try document.updateCell(tableIndex: 0, row: 0, col: 0, text: "Updated")

        let outputURL = directory.appendingPathComponent("output.docx")
        try DocxWriter.write(document, to: outputURL)

        let reloaded = try DocxReader.read(from: outputURL)
        let paragraphs = reloaded.getTables()[0].rows[0].cells[0].paragraphs
        XCTAssertEqual(paragraphs.count, 3, "all three paragraphs must survive a save/reload cycle")
        XCTAssertEqual(paragraphs[0].getText(), "Updated")
        XCTAssertEqual(paragraphs[1].getText(), "Second")
        XCTAssertEqual(paragraphs[2].getText(), "Third")
    }
}
