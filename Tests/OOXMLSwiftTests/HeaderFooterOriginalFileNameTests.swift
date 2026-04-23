import XCTest
@testable import OOXMLSwift

/// Coverage tests for the v0.13.0 originalFileName preservation
/// (`che-word-mcp-true-byte-preservation` Spectra change).
///
/// Validates that `Header` / `Footer` no longer collapse multi-instance
/// same-type files to a single fileName. The pre-v0.13.0 bug: every
/// `.default` header returned `"header1.xml"` regardless of the actual
/// archive path, so an NTPU thesis with 6 default headers (header1.xml
/// through header6.xml) would have all 6 entries lookup the same file.
final class HeaderFooterOriginalFileNameTests: XCTestCase {

    // MARK: - Header.originalFileName

    func testHeaderFreshlyBuiltHasNilOriginalFileName() {
        let header = Header(id: "rId99", paragraphs: [], type: .default)
        XCTAssertNil(header.originalFileName)
    }

    func testHeaderFileNameFallsBackToTypeDefaultWhenOriginalFileNameNil() {
        let header = Header(id: "rId99", paragraphs: [], type: .default)
        XCTAssertEqual(header.fileName, "header1.xml")
    }

    func testHeaderFileNameReturnsOriginalFileNameWhenPresent() {
        var header = Header(id: "rId8", paragraphs: [], type: .default)
        header.originalFileName = "header4.xml"
        XCTAssertEqual(header.fileName, "header4.xml")
    }

    func testHeaderFirstTypeFallsBackToHeaderFirstXml() {
        let header = Header(id: "rId99", paragraphs: [], type: .first)
        XCTAssertEqual(header.fileName, "headerFirst.xml")
    }

    // MARK: - Footer.originalFileName

    func testFooterFreshlyBuiltHasNilOriginalFileName() {
        let footer = Footer(id: "rId99", paragraphs: [], type: .default)
        XCTAssertNil(footer.originalFileName)
    }

    func testFooterFileNameFallsBackToTypeDefaultWhenOriginalFileNameNil() {
        let footer = Footer(id: "rId99", paragraphs: [], type: .default)
        XCTAssertEqual(footer.fileName, "footer1.xml")
    }

    func testFooterFileNameReturnsOriginalFileNameWhenPresent() {
        var footer = Footer(id: "rId10", paragraphs: [], type: .default)
        footer.originalFileName = "footer3.xml"
        XCTAssertEqual(footer.fileName, "footer3.xml")
    }

    // MARK: - Multi-instance distinct fileName

    func testSixDefaultHeadersWithDistinctOriginalFileNamesDoNotCollapse() {
        let originals = ["header1.xml", "header2.xml", "header3.xml",
                         "header4.xml", "header5.xml", "header6.xml"]
        let headers: [Header] = originals.enumerated().map { idx, name in
            var h = Header(id: "rId\(100 + idx)", paragraphs: [], type: .default)
            h.originalFileName = name
            return h
        }
        XCTAssertEqual(Set(headers.map { $0.fileName }), Set(originals))
        XCTAssertEqual(headers.map { $0.fileName }.count, 6)
    }
}
