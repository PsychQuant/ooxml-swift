import XCTest
@testable import OOXMLSwift

/// Coverage tests for v0.13.0 `DocxReader.read()` instrumentation
/// (`che-word-mcp-true-byte-preservation` Spectra change):
/// - Reader populates Header/Footer.originalFileName from rels Target
/// - Reader clears modifiedParts to empty as final step
final class ReaderDirtyTrackingTests: XCTestCase {

    /// Minimal in-memory `.docx` fixture (scratch mode) — covers core
    /// dirty-tracking scenarios. Multi-header / multi-footer fixture for
    /// Task 1.7 / 1.9 lives in MultiHeaderFooterFixtureTests.
    private func makeMinimalDocxFixture() throws -> URL {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "Reader dirty-tracking fixture body"))
        let tempURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("reader-dirty-fixture-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: tempURL)
        return tempURL
    }

    private func makeDocxFixtureWithSingleHeader() throws -> URL {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "fixture with header"))
        _ = doc.addHeader(text: "Page header", type: .default)
        let tempURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("reader-header-fixture-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: tempURL)
        return tempURL
    }

    private func makeDocxFixtureWithSingleFooter() throws -> URL {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "fixture with footer"))
        _ = doc.addFooter(text: "Page footer", type: .default)
        let tempURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("reader-footer-fixture-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: tempURL)
        return tempURL
    }

    // MARK: - Reader clears modifiedParts

    func testReaderLoadedDocumentHasEmptyModifiedParts() throws {
        let fixture = try makeMinimalDocxFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }

        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        XCTAssertTrue(doc.modifiedPartsView.isEmpty,
                      "DocxReader.read() must clear modifiedParts to empty as final step")
    }

    // MARK: - Reader populates Header.originalFileName

    func testReaderPopulatesHeaderOriginalFileName() throws {
        let fixture = try makeDocxFixtureWithSingleHeader()
        defer { try? FileManager.default.removeItem(at: fixture) }

        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        XCTAssertEqual(doc.headers.count, 1, "Fixture must contain one header")
        guard let header = doc.headers.first else { return }
        XCTAssertNotNil(header.originalFileName,
                        "DocxReader.read() must populate originalFileName from rels Target")
        XCTAssertEqual(header.originalFileName, "header1.xml",
                       "originalFileName must reflect the actual archive Target attribute")
        XCTAssertEqual(header.fileName, "header1.xml")
    }

    // MARK: - Reader populates Footer.originalFileName

    func testReaderPopulatesFooterOriginalFileName() throws {
        let fixture = try makeDocxFixtureWithSingleFooter()
        defer { try? FileManager.default.removeItem(at: fixture) }

        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        XCTAssertEqual(doc.footers.count, 1, "Fixture must contain one footer")
        guard let footer = doc.footers.first else { return }
        XCTAssertNotNil(footer.originalFileName)
        XCTAssertEqual(footer.originalFileName, "footer1.xml")
        XCTAssertEqual(footer.fileName, "footer1.xml")
    }
}
