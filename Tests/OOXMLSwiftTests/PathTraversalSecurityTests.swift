import XCTest
@testable import OOXMLSwift

/// Security regression test for che-word-mcp#55 (path traversal).
///
/// `_rels/document.xml.rels` `Target` attribute flows directly into
/// `URL.appendingPathComponent` (which does NOT normalize `..`) AND into
/// `Header.originalFileName` which is later used at write time. A malicious
/// .docx could read OR write outside the intended `word/` directory.
///
/// Defense-in-depth fix: validate at parse boundary
/// (`isSafeRelativeOOXMLPath`) AND at the property setter (defense for
/// programmatic mutation post-load).
final class PathTraversalSecurityTests: XCTestCase {

    // MARK: - PathValidator unit tests (parse-boundary defense)

    func testValidatorRejectsParentTraversal() {
        XCTAssertFalse(isSafeRelativeOOXMLPath("../etc/passwd"))
        XCTAssertFalse(isSafeRelativeOOXMLPath("../../etc/passwd"))
        XCTAssertFalse(isSafeRelativeOOXMLPath("foo/../../bar.xml"))
        XCTAssertFalse(isSafeRelativeOOXMLPath("../header1.xml"))
    }

    func testValidatorRejectsAbsolutePaths() {
        XCTAssertFalse(isSafeRelativeOOXMLPath("/etc/passwd"))
        XCTAssertFalse(isSafeRelativeOOXMLPath("/Users/victim/.ssh/id_rsa"))
    }

    func testValidatorRejectsURLEncodedTraversal() {
        XCTAssertFalse(isSafeRelativeOOXMLPath("..%2fheader1.xml"),
                       "URL-encoded slash + .. SHALL be rejected")
        XCTAssertFalse(isSafeRelativeOOXMLPath("%2e%2e/header1.xml"),
                       "URL-encoded .. SHALL be rejected")
        XCTAssertFalse(isSafeRelativeOOXMLPath("%2E%2E%2Fheader1.xml"),
                       "URL-encoded .. (uppercase) SHALL be rejected")
    }

    func testValidatorRejectsControlChars() {
        XCTAssertFalse(isSafeRelativeOOXMLPath("header\u{00}.xml"),
                       "NUL byte in path SHALL be rejected")
        XCTAssertFalse(isSafeRelativeOOXMLPath("header\n.xml"),
                       "newline in path SHALL be rejected")
    }

    func testValidatorRejectsOversizedPaths() {
        let oversized = String(repeating: "a", count: 257) + ".xml"
        XCTAssertFalse(isSafeRelativeOOXMLPath(oversized),
                       "Paths > 256 chars SHALL be rejected (DoS guard)")
    }

    func testValidatorAcceptsValidPaths() {
        XCTAssertTrue(isSafeRelativeOOXMLPath("header1.xml"))
        XCTAssertTrue(isSafeRelativeOOXMLPath("header2.xml"))
        XCTAssertTrue(isSafeRelativeOOXMLPath("footer1.xml"))
        XCTAssertTrue(isSafeRelativeOOXMLPath("media/image1.png"))
        XCTAssertTrue(isSafeRelativeOOXMLPath("theme/theme1.xml"))
        XCTAssertTrue(isSafeRelativeOOXMLPath("_rels/document.xml.rels"))
    }

    func testValidatorAcceptsNonASCIIInRange() {
        // CJK-named header (rare but valid in Word installs with CJK locales)
        XCTAssertTrue(isSafeRelativeOOXMLPath("頁首1.xml"),
                      "Printable Unicode in safe range SHALL be accepted")
    }

    // MARK: - Setter defense (programmatic mutation post-load)

    func testHeaderOriginalFileNameSetterRejectsTraversal() {
        var hdr = Header(id: "rId10", paragraphs: [Paragraph(text: "x")])
        hdr.originalFileName = "../../../etc/passwd"
        XCTAssertNil(hdr.originalFileName,
                     "Setter SHALL reject traversal path; expecting nil fallback")
    }

    func testFooterOriginalFileNameSetterRejectsTraversal() {
        var ftr = Footer(id: "rId20", paragraphs: [Paragraph(text: "x")])
        ftr.originalFileName = "/etc/passwd"
        XCTAssertNil(ftr.originalFileName,
                     "Setter SHALL reject absolute path; expecting nil fallback")
    }

    func testHeaderOriginalFileNameSetterAcceptsValid() {
        var hdr = Header(id: "rId10", paragraphs: [Paragraph(text: "x")])
        hdr.originalFileName = "header2.xml"
        XCTAssertEqual(hdr.originalFileName, "header2.xml",
                       "Setter SHALL accept valid relative paths")
    }
}
