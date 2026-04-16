import XCTest
@testable import OOXMLSwift

/// Unit tests for nested revision parsing (rPrChange, pPrChange).
/// Part B of ooxml-swift#1 via the docx-reader-nested-revisions-and-containers change.
final class NestedRevisionTests: XCTestCase {

    // MARK: - Helpers (reuse the RevisionParsingTests pattern)

    private func paragraphElement(inner: String) -> XMLElement {
        let wrapped = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        \(inner)
        </w:p>
        """
        return try! XMLDocument(xmlString: wrapped, options: []).rootElement()!
    }

    private func parse(_ element: XMLElement) throws -> Paragraph {
        try DocxReader.parseParagraph(
            from: element,
            relationships: RelationshipsCollection(),
            styles: [],
            numbering: Numbering()
        )
    }

    // MARK: - rPrChange

    func testParsesRPrChangeFormatRevision() throws {
        let el = paragraphElement(inner: """
        <w:r>
            <w:rPr>
                <w:b/>
                <w:i/>
                <w:rPrChange w:id="10" w:author="Alice" w:date="2026-04-16T14:00:00Z">
                    <w:rPr><w:b/></w:rPr>
                </w:rPrChange>
            </w:rPr>
            <w:t>text</w:t>
        </w:r>
        """)
        let paragraph = try parse(el)

        let formatRevisions = paragraph.revisions.filter { $0.type == .formatChange }
        XCTAssertEqual(formatRevisions.count, 1, "Expected exactly 1 formatChange revision")
        let rev = formatRevisions[0]
        XCTAssertEqual(rev.id, 10)
        XCTAssertEqual(rev.author, "Alice")
        XCTAssertNotNil(rev.previousFormatDescription)
        XCTAssertTrue(rev.previousFormatDescription?.contains("bold") == true,
                       "previousFormatDescription should mention 'bold', got: \(rev.previousFormatDescription ?? "nil")")
    }

    func testNoRPrChangeEmitsNoRevision() throws {
        let el = paragraphElement(inner: """
        <w:r>
            <w:rPr><w:b/></w:rPr>
            <w:t>text</w:t>
        </w:r>
        """)
        let paragraph = try parse(el)

        let formatRevisions = paragraph.revisions.filter { $0.type == .formatChange }
        XCTAssertEqual(formatRevisions.count, 0, "No rPrChange means no formatChange revision")
    }

    // MARK: - pPrChange

    func testParsesPPrChangeParagraphRevision() throws {
        let el = paragraphElement(inner: """
        <w:pPr>
            <w:jc w:val="left"/>
            <w:pPrChange w:id="20" w:author="Bob" w:date="2026-04-16T15:00:00Z">
                <w:pPr><w:jc w:val="center"/></w:pPr>
            </w:pPrChange>
        </w:pPr>
        <w:r><w:t>text</w:t></w:r>
        """)
        let paragraph = try parse(el)

        let pprRevisions = paragraph.revisions.filter { $0.type == .paragraphChange }
        XCTAssertEqual(pprRevisions.count, 1, "Expected exactly 1 paragraphChange revision")
        let rev = pprRevisions[0]
        XCTAssertEqual(rev.id, 20)
        XCTAssertEqual(rev.author, "Bob")
        XCTAssertNotNil(rev.previousFormatDescription)
        XCTAssertTrue(rev.previousFormatDescription?.contains("center") == true,
                       "previousFormatDescription should mention 'center', got: \(rev.previousFormatDescription ?? "nil")")
    }

    func testNoPPrChangeEmitsNoRevision() throws {
        let el = paragraphElement(inner: """
        <w:pPr><w:jc w:val="left"/></w:pPr>
        <w:r><w:t>text</w:t></w:r>
        """)
        let paragraph = try parse(el)

        let pprRevisions = paragraph.revisions.filter { $0.type == .paragraphChange }
        XCTAssertEqual(pprRevisions.count, 0, "No pPrChange means no paragraphChange revision")
    }

    // MARK: - previousFormatDescription nil for non-format revisions

    func testInsertionHasNilPreviousFormatDescription() throws {
        let el = paragraphElement(inner: """
        <w:ins w:id="1" w:author="Alice" w:date="2026-04-16T10:00:00Z">
            <w:r><w:t>hello</w:t></w:r>
        </w:ins>
        """)
        let paragraph = try parse(el)

        XCTAssertEqual(paragraph.revisions.count, 1)
        XCTAssertNil(paragraph.revisions[0].previousFormatDescription,
                      "Insertion revision should have nil previousFormatDescription")
    }
}
