import XCTest
@testable import OOXMLSwift

/// Tests for Phase 5 Section property extensions on WordDocument. Spec
/// covered (see
/// `openspec/changes/che-word-mcp-styles-sections-numbering-foundations/specs/`):
/// - ooxml-document-part-mutations: WordDocument exposes section property extensions
///
/// Implementation tasks 5.1-5.5 will populate these tests; until then they
/// XCTSkip so the suite stays green.
final class SectionPropertiesExtendedTests: XCTestCase {

    // MARK: - Test fixture helpers

    /// Builds a doc with N sections, each separated by a section break.
    /// All sections start with default page geometry.
    func makeDocWithSections(_ count: Int) -> WordDocument {
        var doc = WordDocument()
        for i in 0..<count {
            doc.body.children.append(.paragraph(Paragraph(text: "Section \(i) body")))
        }
        return doc
    }

    func roundTrip(_ doc: WordDocument) throws -> WordDocument {
        let bytes = try DocxWriter.writeData(doc)
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("sect-rt-\(UUID().uuidString).docx")
        try bytes.write(to: url)
        defer { try? FileManager.default.removeItem(at: url) }
        return try DocxReader.read(from: url)
    }

    // MARK: - Task 5.4: section mutators

    func testSetSectionLineNumbersEmitsLnNumTypeAfterTask54() throws {
        var doc = makeDocWithSections(1)
        try doc.setSectionLineNumbers(sectionIndex: 0, countBy: 1, start: 1, restart: .newPage)
        let xml = doc.sectionProperties.toXML()
        XCTAssertTrue(xml.contains("<w:lnNumType w:countBy=\"1\" w:start=\"1\" w:restart=\"newPage\"/>"),
            "expected lnNumType in: \(xml)")
    }

    /// Spec scenario: Set section vertical alignment to center.
    func testSetSectionVerticalAlignmentCenterAfterTask54() throws {
        var doc = makeDocWithSections(1)
        try doc.setSectionVerticalAlignment(sectionIndex: 0, alignment: .center)
        XCTAssertTrue(doc.sectionProperties.toXML().contains("<w:vAlign w:val=\"center\"/>"))
    }

    /// Spec scenario: Set Roman numeral page numbers on section 0.
    func testSetSectionPageNumberFormatLowerRomanAfterTask54() throws {
        var doc = makeDocWithSections(1)
        try doc.setSectionPageNumberFormat(sectionIndex: 0, start: 1, format: .lowerRoman)
        let xml = doc.sectionProperties.toXML()
        XCTAssertTrue(xml.contains("w:start=\"1\""))
        XCTAssertTrue(xml.contains("w:fmt=\"lowerRoman\""))
    }

    func testSetSectionBreakTypeOddPageAfterTask54() throws {
        var doc = makeDocWithSections(1)
        try doc.setSectionBreakType(sectionIndex: 0, type: .oddPage)
        XCTAssertTrue(doc.sectionProperties.toXML().contains("<w:type w:val=\"oddPage\"/>"))
    }

    func testSetTitlePageDistinctTogglesElementAfterTask54() throws {
        var doc = makeDocWithSections(1)
        try doc.setTitlePageDistinct(sectionIndex: 0, enabled: true)
        XCTAssertTrue(doc.sectionProperties.toXML().contains("<w:titlePg/>"))
        try doc.setTitlePageDistinct(sectionIndex: 0, enabled: false)
        XCTAssertFalse(doc.sectionProperties.toXML().contains("<w:titlePg/>"))
    }

    func testSectionMutatorsThrowOnInvalidIndexAfterTask54() throws {
        var doc = makeDocWithSections(1)
        XCTAssertThrowsError(try doc.setTitlePageDistinct(sectionIndex: 5, enabled: true)) { error in
            guard case WordError.invalidIndex(5) = error else { XCTFail("expected invalidIndex(5)"); return }
        }
    }

    // MARK: - Task 5.5: getAllSections

    func testGetAllSectionsReturnsParagraphRangesAfterTask55() throws {
        var doc = makeDocWithSections(3)
        try doc.setSectionPageNumberFormat(sectionIndex: 0, start: 1, format: .lowerRoman)
        let sections = doc.getAllSections()
        XCTAssertEqual(sections.count, 1)
        XCTAssertEqual(sections[0].sectionIndex, 0)
        XCTAssertEqual(sections[0].pageNumberFormat, .lowerRoman)
    }

    // MARK: - Pre-existing sanity

    func testFixtureBuilderProducesMultiSectionDoc() {
        let doc = makeDocWithSections(3)
        XCTAssertEqual(doc.body.children.count, 3)
    }
}
