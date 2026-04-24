import XCTest
@testable import OOXMLSwift

/// Regression test for che-word-mcp#53 (multi-instance Header/Footer fileName collision).
///
/// Pre-fix: `addHeader()` × 2 with default type both produced `Header.fileName == "header1.xml"`.
/// On save: h2 overwrites h1 on disk; in dirty-bit Sets they collapse to one entry.
///
/// Fix: `addHeader*` / `addFooter*` auto-suffix the fileName based on existing
/// typed model state, populating `originalFileName` so the type-based fallback
/// path is no longer hit by programmatic multi-add.
final class MultiInstanceHeaderFooterTests: XCTestCase {

    // MARK: - Headers

    func testTwoDefaultTypeHeadersDoNotCollide() {
        var doc = WordDocument()
        let h1 = doc.addHeader(text: "first")
        let h2 = doc.addHeader(text: "second")

        XCTAssertNotEqual(h1.fileName, h2.fileName,
                          "Two default-type headers SHALL have distinct fileNames; got h1=\(h1.fileName) h2=\(h2.fileName)")
        // Sanity: ids are also distinct (allocator already handles this; not a regression check).
        XCTAssertNotEqual(h1.id, h2.id)
    }

    func testThreeDefaultTypeHeadersGetSequentialFileNames() {
        var doc = WordDocument()
        let h1 = doc.addHeader(text: "1")
        let h2 = doc.addHeader(text: "2")
        let h3 = doc.addHeader(text: "3")

        XCTAssertEqual(h1.fileName, "header1.xml")
        XCTAssertEqual(h2.fileName, "header2.xml")
        XCTAssertEqual(h3.fileName, "header3.xml")
    }

    func testAddHeaderWithPageNumberAlsoAutoSuffixes() {
        var doc = WordDocument()
        let h1 = doc.addHeader(text: "plain")
        let h2 = doc.addHeaderWithPageNumber()
        XCTAssertNotEqual(h1.fileName, h2.fileName,
                          "addHeaderWithPageNumber SHALL auto-suffix like addHeader; got h1=\(h1.fileName) h2=\(h2.fileName)")
    }

    // MARK: - Footers

    func testTwoDefaultTypeFootersDoNotCollide() {
        var doc = WordDocument()
        let f1 = doc.addFooter(text: "first")
        let f2 = doc.addFooter(text: "second")

        XCTAssertNotEqual(f1.fileName, f2.fileName,
                          "Two default-type footers SHALL have distinct fileNames; got f1=\(f1.fileName) f2=\(f2.fileName)")
    }

    func testThreeDefaultTypeFootersGetSequentialFileNames() {
        var doc = WordDocument()
        let f1 = doc.addFooter(text: "1")
        let f2 = doc.addFooter(text: "2")
        let f3 = doc.addFooter(text: "3")

        XCTAssertEqual(f1.fileName, "footer1.xml")
        XCTAssertEqual(f2.fileName, "footer2.xml")
        XCTAssertEqual(f3.fileName, "footer3.xml")
    }

    func testAddFooterWithPageNumberAlsoAutoSuffixes() {
        var doc = WordDocument()
        let f1 = doc.addFooter(text: "plain")
        let f2 = doc.addFooterWithPageNumber()
        XCTAssertNotEqual(f1.fileName, f2.fileName)
    }

    // MARK: - Reader-loaded path unaffected

    func testReaderLoadedHeadersStillUseOriginalFileName() throws {
        let tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("MultiInstanceHF-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: tempDir) }

        var src = WordDocument()
        src.appendParagraph(Paragraph(text: "body"))
        _ = src.addHeader(text: "h1")
        _ = src.addHeader(text: "h2")
        let url = tempDir.appendingPathComponent("test.docx")
        try DocxWriter.write(src, to: url)

        let loaded = try DocxReader.read(from: url)
        let names = loaded.headers.map { $0.fileName }
        XCTAssertEqual(Set(names).count, names.count,
                       "Reader-loaded headers SHALL preserve unique fileNames from source archive; got \(names)")
    }
}
