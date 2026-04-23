import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Coverage tests for v0.13.0 `DocxWriter` overlay-mode skip-when-not-dirty
/// (`che-word-mcp-true-byte-preservation` Spectra change).
///
/// Validates the core architectural contract: in overlay mode, every typed-part
/// writer is gated by `modifiedParts.contains(<part path>)`. A read+write
/// round-trip with zero edits must leave every typed-managed part byte-equal.
///
/// Pre-v0.13.0 bug: `writeAllParts` unconditionally overwrote document.xml,
/// styles.xml, fontTable.xml, header*.xml, footer*.xml etc. on every save,
/// so a Reader-loaded NTPU thesis lost its custom font declarations after a
/// single `save_document` round-trip even though no typed mutation had occurred.
final class OverlaySkipWhenNotDirtyTests: XCTestCase {

    private func makeBaseFixture() throws -> URL {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "Body content"))
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("overlay-skip-base-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: url)
        return url
    }

    // MARK: - No-op round-trip preserves byte equality across typed parts

    func testNoOpRoundTripPreservesDocumentXMLByteEqual() throws {
        let src = try makeBaseFixture()
        defer { try? FileManager.default.removeItem(at: src) }

        var loaded = try DocxReader.read(from: src)
        defer { loaded.close() }
        XCTAssertTrue(loaded.modifiedPartsView.isEmpty,
                      "Reader must clear modifiedParts; nothing dirty after read")

        let dest = FileManager.default.temporaryDirectory
            .appendingPathComponent("overlay-skip-dest-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: dest) }
        try DocxWriter.write(loaded, to: dest)

        let srcDoc = try unzip(src)
        let destDoc = try unzip(dest)
        defer {
            try? FileManager.default.removeItem(at: srcDoc)
            try? FileManager.default.removeItem(at: destDoc)
        }
        try assertByteEqual(part: "word/document.xml", in: srcDoc, dest: destDoc)
        try assertByteEqual(part: "word/styles.xml", in: srcDoc, dest: destDoc)
        try assertByteEqual(part: "word/fontTable.xml", in: srcDoc, dest: destDoc)
        try assertByteEqual(part: "word/settings.xml", in: srcDoc, dest: destDoc)
        try assertByteEqual(part: "docProps/core.xml", in: srcDoc, dest: destDoc)
        try assertByteEqual(part: "docProps/app.xml", in: srcDoc, dest: destDoc)
    }

    // MARK: - Selective re-emission: dirty part rewritten, others preserved

    func testSingleEditTriggersSelectiveReemissionOnly() throws {
        let src = try makeBaseFixture()
        defer { try? FileManager.default.removeItem(at: src) }

        var loaded = try DocxReader.read(from: src)
        defer { loaded.close() }
        loaded.appendParagraph(Paragraph(text: "Newly inserted"))
        XCTAssertEqual(loaded.modifiedPartsView, ["word/document.xml"],
                       "Only document.xml should be dirty after appendParagraph")

        let dest = FileManager.default.temporaryDirectory
            .appendingPathComponent("overlay-skip-edit-dest-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: dest) }
        try DocxWriter.write(loaded, to: dest)

        let srcDoc = try unzip(src)
        let destDoc = try unzip(dest)
        defer {
            try? FileManager.default.removeItem(at: srcDoc)
            try? FileManager.default.removeItem(at: destDoc)
        }
        // document.xml is rewritten (the new paragraph appears)
        let destDocXML = try String(
            contentsOf: destDoc.appendingPathComponent("word/document.xml"),
            encoding: .utf8
        )
        XCTAssertTrue(destDocXML.contains("Newly inserted"),
                      "Modified document.xml must include the new paragraph")
        // ALL other typed parts preserved byte-equal
        try assertByteEqual(part: "word/styles.xml", in: srcDoc, dest: destDoc)
        try assertByteEqual(part: "word/fontTable.xml", in: srcDoc, dest: destDoc)
        try assertByteEqual(part: "word/settings.xml", in: srcDoc, dest: destDoc)
        try assertByteEqual(part: "docProps/core.xml", in: srcDoc, dest: destDoc)
    }

    // MARK: - Helpers

    private func unzip(_ docx: URL) throws -> URL {
        let dir = FileManager.default.temporaryDirectory
            .appendingPathComponent("overlay-skip-verify-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        try FileManager.default.unzipItem(at: docx, to: dir)
        return dir
    }

    private func assertByteEqual(part: String, in srcDir: URL, dest destDir: URL,
                                 file: StaticString = #file, line: UInt = #line) throws {
        let srcURL = srcDir.appendingPathComponent(part)
        let destURL = destDir.appendingPathComponent(part)
        guard FileManager.default.fileExists(atPath: srcURL.path) else { return }
        XCTAssertTrue(FileManager.default.fileExists(atPath: destURL.path),
                      "\(part) missing from dest archive", file: file, line: line)
        let srcBytes = try Data(contentsOf: srcURL)
        let destBytes = try Data(contentsOf: destURL)
        XCTAssertEqual(srcBytes, destBytes,
                       "\(part) must be byte-equal after no-op round-trip",
                       file: file, line: line)
    }
}
