import XCTest
@testable import OOXMLSwift

/// Tests for che-word-mcp-tables-hyperlinks-headers-builtin SDD Phase 3 + 4
/// (Hyperlink extensions). Specs covered:
/// - ooxml-document-part-mutations: WordDocument exposes hyperlink type-aware insertion and tooltip mutation
///
/// Implementation tasks 3.1 + 4.3 will populate these tests; until then they
/// XCTSkip so the suite stays green.
final class HyperlinkTypedTests: XCTestCase {

    // MARK: - Test fixture helpers

    func makeDocWithHyperlink(_ link: Hyperlink) -> WordDocument {
        var doc = WordDocument()
        var para = Paragraph()
        para.hyperlinks = [link]
        doc.body.children = [.paragraph(para)]
        return doc
    }

    func roundTrip(_ doc: WordDocument) throws -> WordDocument {
        let bytes = try DocxWriter.writeData(doc)
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("hl-rt-\(UUID().uuidString).docx")
        try bytes.write(to: url)
        defer { try? FileManager.default.removeItem(at: url) }
        return try DocxReader.read(from: url)
    }

    // MARK: - Task 3.1: history field

    func testHyperlinkHistoryFieldRoundTripsAfterTask31() throws {
        let link = Hyperlink(id: "h1", text: "Visit", url: "https://example.com",
                             relationshipId: "rId7", history: false)
        let xml = link.toXML()
        XCTAssertTrue(xml.contains("w:history=\"0\""), "history=false should emit w:history=\"0\": \(xml)")

        let link2 = Hyperlink(id: "h2", text: "Visit", url: "https://example.com",
                              relationshipId: "rId8", history: true)
        XCTAssertFalse(link2.toXML().contains("w:history"), "history=true (default) should NOT emit attribute")
    }

    // MARK: - Task 4.3: setHyperlinkTooltip

    func testSetHyperlinkTooltipMutatesExistingAfterTask43() throws {
        let link = Hyperlink(id: "h-tt", text: "Click", url: "https://e.com", relationshipId: "rId7")
        var doc = makeDocWithHyperlink(link)
        try doc.setHyperlinkTooltip(hyperlinkId: "h-tt", tooltip: "Click here")
        if case .paragraph(let p) = doc.body.children[0] {
            XCTAssertEqual(p.hyperlinks.first?.tooltip, "Click here")
        } else {
            XCTFail("expected paragraph")
        }
    }

    func testSetHyperlinkTooltipThrowsOnNotFoundAfterTask43() throws {
        var doc = WordDocument()
        XCTAssertThrowsError(try doc.setHyperlinkTooltip(hyperlinkId: "missing", tooltip: "x")) { error in
            guard case WordError.hyperlinkNotFound = error else { XCTFail("expected hyperlinkNotFound"); return }
        }
    }

    // MARK: - Pre-existing sanity

    func testFixtureBuilderProducesUrlHyperlink() {
        let link = Hyperlink(
            id: "hl1",
            text: "Example",
            url: "https://example.com",
            relationshipId: "rId7",
            tooltip: "Visit"
        )
        let doc = makeDocWithHyperlink(link)
        if case .paragraph(let p) = doc.body.children[0] {
            XCTAssertEqual(p.hyperlinks.first?.url, "https://example.com")
            XCTAssertEqual(p.hyperlinks.first?.tooltip, "Visit")
        } else {
            XCTFail("expected paragraph")
        }
    }
}
