import XCTest
@testable import OOXMLSwift

/// Tests for che-word-mcp-tables-hyperlinks-headers-builtin SDD Phase 3 + 4
/// (Header / Footer / settings extensions). Specs covered:
/// - ooxml-document-part-mutations: WordDocument exposes typed header parts and clone semantics
///
/// Implementation tasks 3.2 + 3.3 + 4.4 will populate these tests; until then
/// they XCTSkip so the suite stays green.
final class HeaderMultiTypeTests: XCTestCase {

    // MARK: - Test fixture helpers

    func makeDocWithHeaders(_ headers: [Header]) -> WordDocument {
        var doc = WordDocument()
        doc.headers = headers
        return doc
    }

    func roundTrip(_ doc: WordDocument) throws -> WordDocument {
        let bytes = try DocxWriter.writeData(doc)
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("hdr-rt-\(UUID().uuidString).docx")
        try bytes.write(to: url)
        defer { try? FileManager.default.removeItem(at: url) }
        return try DocxReader.read(from: url)
    }

    // MARK: - Task 3.3: evenAndOddHeaders flag

    func testEvenAndOddHeadersFlagDefaultsFalseAfterTask33() throws {
        let doc = WordDocument()
        XCTAssertFalse(doc.evenAndOddHeaders, "default should be false")
    }

    // MARK: - Task 4.4: addHeaderOfType / setEvenAndOddHeaders / cloneHeaderForSection

    func testAddHeaderOfTypeFirstAllocatesNewPartAfterTask44() throws {
        var doc = WordDocument()
        let fn = try doc.addHeaderOfType(text: "Cover Page Header", type: .first)
        XCTAssertTrue(fn.hasPrefix("header") && fn.hasSuffix(".xml"), "expected headerN.xml file name, got: \(fn)")
        XCTAssertEqual(doc.headers.count, 1)
        XCTAssertEqual(doc.headers.first?.type, .first)
    }

    func testSetEvenAndOddHeadersTogglesFlagAfterTask44() throws {
        var doc = WordDocument()
        doc.setEvenAndOddHeaders(true)
        XCTAssertTrue(doc.evenAndOddHeaders)
        doc.setEvenAndOddHeaders(false)
        XCTAssertFalse(doc.evenAndOddHeaders)
    }

    func testCloneHeaderForSectionDeepCopiesContentAfterTask44() throws {
        var doc = WordDocument()
        let h1 = Header(id: "rId1", paragraphs: [Paragraph(text: "Original")], type: .default,
                        originalFileName: "header1.xml")
        doc.headers = [h1]
        let cloned = try doc.cloneHeaderForSection(
            sourceFileName: "header1.xml",
            targetSectionIndex: 0,
            type: .default
        )
        XCTAssertNotEqual(cloned, "header1.xml", "clone should have a new file name")
        XCTAssertEqual(doc.headers.count, 2)
        XCTAssertEqual(doc.headers.last?.paragraphs.first?.getText(), "Original",
            "cloned content should match source")
    }

    // MARK: - Pre-existing sanity

    func testFixtureBuilderAcceptsMultipleHeaders() {
        let h1 = Header(id: "rId1", paragraphs: [Paragraph(text: "Default")], type: .default)
        let h2 = Header(id: "rId2", paragraphs: [Paragraph(text: "First Page")], type: .first)
        let doc = makeDocWithHeaders([h1, h2])
        XCTAssertEqual(doc.headers.count, 2)
        XCTAssertEqual(doc.headers[0].type, .default)
        XCTAssertEqual(doc.headers[1].type, .first)
    }
}
