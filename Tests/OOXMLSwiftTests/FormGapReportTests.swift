// FormGapReportTests.swift
// word-canonical-forms Phase 1 task 1.1 — form-gap measurement names the
// first offending form (`ooxml-script-transcode`, «Form-gap measurement
// names the first offending form»; Decision 1: the report is the work
// queue). A bail records part path + XML path to the first offending
// node/attribute + content class; an upgraded part reports no gaps.

import XCTest
@testable import OOXMLSwift

final class FormGapReportTests: XCTestCase {

    private func parts(documentXML: String) -> [String: Data] {
        ["word/document.xml": Data(documentXML.utf8)]
    }

    // MARK: - Extraction bail (case a): has an XML breadcrumb

    /// Spec scenario: bail names the offending attribute. A paragraph
    /// carrying `w:rsidR` (not the supported w14:paraId-only form) bails; the
    /// gap locates the paragraph and names the offending attribute.
    func testBailNamesOffendingAttributePath() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p w:rsidR="00AB12CD"><w:r><w:t>hi</w:t></w:r></w:p></w:body></w:document>
        """
        let result = try ReverseExtractor.reverse(parts: parts(documentXML: xml))
        XCTAssertFalse(result.dslParts.contains("word/document.xml"))
        let gap = try XCTUnwrap(result.formGaps.first { $0.partPath == "word/document.xml" })
        XCTAssertTrue(gap.xmlPath.contains("w:p"), "gap path should locate the paragraph: \(gap.xmlPath)")
        XCTAssertTrue(gap.xmlPath.contains("rsidR"),
                      "gap path should name the offending attribute: \(gap.xmlPath)")
    }

    /// A pPr carrying an unsupported child element (widowControl) bails
    /// naming that element in the path.
    func testGapNamesUnsupportedPPrElement() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="P1"><w:pPr><w:widowControl/></w:pPr><w:r><w:t>hi</w:t></w:r></w:p></w:body></w:document>
        """
        let result = try ReverseExtractor.reverse(parts: parts(documentXML: xml))
        let gap = try XCTUnwrap(result.formGaps.first)
        XCTAssertTrue(gap.xmlPath.contains("pPr"), gap.xmlPath)
        XCTAssertTrue(gap.xmlPath.contains("widowControl"), gap.xmlPath)
    }

    // MARK: - Upgraded (no gaps)

    /// Spec scenario: upgraded part reports no gaps.
    func testUpgradedDocumentReportsNoGaps() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "第一段", styleId: "Body", paraId: "P1")),
            .appendParagraph(in: nil, paragraph: ParagraphPayload(text: "第二段", paraId: "P2")),
        ])
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("fg-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try doc.writeAuthoringPackage(to: url)
        let parts = try RawPartChannel.readAllParts(from: url)

        let result = try ReverseExtractor.reverse(parts: parts)
        XCTAssertTrue(result.dslParts.contains("word/document.xml"))
        XCTAssertTrue(result.formGaps.isEmpty,
                      "upgraded document must report no gaps; got \(result.formGaps)")
    }

    // MARK: - Byte-mismatch (case b): extraction passes, rebuild bytes differ

    /// A root carrying a namespace declaration the authoring default doesn't
    /// emit passes extraction (root attrs aren't extraction-checked) but the
    /// trial rebuild differs — the gap is class byte-mismatch and the path
    /// carries the byte offset so the diff is locatable.
    func testByteMismatchGapCarriesOffset() throws {
        // Extra xmlns:mc on root — extraction accepts (only localName checked),
        // but emptyAuthoringDocument rebuilds with only w + w14 → byte diff.
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:body><w:p w14:paraId="P1"><w:r><w:t>hi</w:t></w:r></w:p></w:body></w:document>
        """
        let result = try ReverseExtractor.reverse(parts: parts(documentXML: xml))
        XCTAssertFalse(result.dslParts.contains("word/document.xml"))
        let gap = try XCTUnwrap(result.formGaps.first)
        XCTAssertEqual(gap.contentClass, "byte-mismatch", gap.xmlPath)
        XCTAssertTrue(gap.xmlPath.contains("byte@"),
                      "byte-mismatch gap should carry the divergence offset: \(gap.xmlPath)")
    }
}
