import XCTest
@testable import OOXMLSwift

/// v0.19.3+ regression tests for the round 2 verify findings on
/// PsychQuant/che-word-mcp#56 (https://github.com/PsychQuant/che-word-mcp/issues/56#issuecomment-4320157395).
/// This batch covers the **Hyperlink suite** (P0-1, P0-2, P0-3, P1-7) — the
/// other batches (sort path / revision / bookmark) ship in follow-up commits
/// inside the same v0.19.3 release.
final class Issue56RoundtripCompletenessTests: XCTestCase {

    // MARK: - P0-1: API path Hyperlink visual styling preserved

    /// `Hyperlink.external` populates a Run via `Run(text:)`. Pre-v0.19.3 the
    /// new `toXML()` walked `runs` directly and emitted a plain `<w:r><w:t>`
    /// without `<w:rStyle w:val="Hyperlink"/>`, `<w:color w:val="0563C1"/>`,
    /// or `<w:u w:val="single"/>` — so every hyperlink built via the 5 MCP
    /// `insert_*hyperlink` tools rendered without blue-underline styling.
    func testExternalHyperlinkAppliesHyperlinkVisualStyling() {
        let hl = Hyperlink.external(
            id: "h1",
            text: "click",
            url: "https://example.com",
            relationshipId: "rId1"
        )
        let xml = hl.toXML()

        XCTAssertTrue(
            xml.contains("w:rStyle w:val=\"Hyperlink\""),
            "API-built external hyperlink must include <w:rStyle w:val=\"Hyperlink\"/>. Output:\n\(xml)"
        )
        XCTAssertTrue(
            xml.contains("w:color w:val=\"0563C1\""),
            "API-built external hyperlink must include <w:color w:val=\"0563C1\"/>. Output:\n\(xml)"
        )
        XCTAssertTrue(
            xml.contains("w:u w:val=\"single\""),
            "API-built external hyperlink must include <w:u w:val=\"single\"/>. Output:\n\(xml)"
        )
    }

    /// Same contract for internal (anchor-based) hyperlinks built via
    /// `Hyperlink.internal(...)`.
    func testInternalHyperlinkAppliesHyperlinkVisualStyling() {
        let hl = Hyperlink.internal(
            id: "h2",
            text: "see section",
            bookmarkName: "section1"
        )
        let xml = hl.toXML()

        XCTAssertTrue(
            xml.contains("w:rStyle w:val=\"Hyperlink\""),
            "API-built internal hyperlink must include <w:rStyle w:val=\"Hyperlink\"/>. Output:\n\(xml)"
        )
        XCTAssertTrue(
            xml.contains("w:color w:val=\"0563C1\""),
            "API-built internal hyperlink must include <w:color w:val=\"0563C1\"/>. Output:\n\(xml)"
        )
        XCTAssertTrue(
            xml.contains("w:u w:val=\"single\""),
            "API-built internal hyperlink must include <w:u w:val=\"single\"/>. Output:\n\(xml)"
        )
    }

    // MARK: - P0-2: Reader preserves w:tgtFrame / w:docLocation via rawAttributes

    /// `parseHyperlink` listed `w:tgtFrame` and `w:docLocation` in
    /// `recognizedAttrs` (so they were skipped from `rawAttributes`), but the
    /// Hyperlink model has no typed field and `toXML()` never emits them.
    /// Net effect: source attributes silently dropped on round-trip. Fix:
    /// remove them from `recognizedAttrs` so they fall into rawAttributes,
    /// where the writer already emits them.
    func testHyperlinkTgtFrameRoundTripsThroughReaderAndWriter() throws {
        let xmlSrc = """
        <w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId5" w:tgtFrame="_blank" w:docLocation="frag1">
            <w:r><w:t>click</w:t></w:r>
        </w:hyperlink>
        """
        let element = try XMLElement(xmlString: xmlSrc)

        let hl = try DocxReader.parseHyperlink(
            from: element,
            relationships: RelationshipsCollection(),
            position: 0
        )

        XCTAssertEqual(
            hl.rawAttributes["w:tgtFrame"],
            "_blank",
            "w:tgtFrame must land in rawAttributes for round-trip. Got rawAttributes=\(hl.rawAttributes)"
        )
        XCTAssertEqual(
            hl.rawAttributes["w:docLocation"],
            "frag1",
            "w:docLocation must land in rawAttributes for round-trip. Got rawAttributes=\(hl.rawAttributes)"
        )

        let outXml = hl.toXML()
        XCTAssertTrue(
            outXml.contains("w:tgtFrame=\"_blank\""),
            "w:tgtFrame must round-trip through Writer. Output:\n\(outXml)"
        )
        XCTAssertTrue(
            outXml.contains("w:docLocation=\"frag1\""),
            "w:docLocation must round-trip through Writer. Output:\n\(outXml)"
        )
    }

    // MARK: - P0-3: Hyperlink internal child order preserved

    /// Source `<w:hyperlink><w:r>A</w:r><w:sdt>X</w:sdt><w:r>B</w:r></w:hyperlink>`
    /// must round-trip with the same A → SDT → B order. Pre-v0.19.3 Reader
    /// split into `runs=[A,B]` and `rawChildren=[<w:sdt>X</w:sdt>]`, then
    /// Writer emitted `<w:r>A</w:r><w:r>B</w:r><w:sdt>X</w:sdt>` — visible
    /// text order changed.
    func testHyperlinkChildOrderPreservedAcrossRunAndNonRunChildren() throws {
        let xmlSrc = """
        <w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId7"><w:r><w:t>A</w:t></w:r><w:sdt><w:sdtContent><w:r><w:t>X</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t>B</w:t></w:r></w:hyperlink>
        """
        let element = try XMLElement(xmlString: xmlSrc)

        let hl = try DocxReader.parseHyperlink(
            from: element,
            relationships: RelationshipsCollection(),
            position: 0
        )

        let outXml = hl.toXML()

        // The output must have A before SDT before B (positions in source order).
        guard let aPos = outXml.range(of: ">A<")?.lowerBound,
              let sdtPos = outXml.range(of: "<w:sdt")?.lowerBound,
              let bPos = outXml.range(of: ">B<")?.lowerBound else {
            XCTFail("Output missing one of A / SDT / B markers. Output:\n\(outXml)")
            return
        }
        XCTAssertLessThan(aPos, sdtPos, "A must precede SDT in round-trip. Output:\n\(outXml)")
        XCTAssertLessThan(sdtPos, bPos, "SDT must precede B in round-trip. Output:\n\(outXml)")
    }

    // MARK: - P1-7: Hyperlink.id unique across duplicate r:id

    /// Two `<w:hyperlink>` elements with the same `r:id` (legitimate when two
    /// links share a relationship target — e.g., two "click here" anchors for
    /// the same URL) must parse into Hyperlink instances with **distinct**
    /// `id` fields. The legacy `id = rId ?? anchor ?? "hl-\(position)"`
    /// returned the same id for both, breaking MCP tools that find/edit/
    /// delete hyperlinks by `id`.
    func testParsedHyperlinksHaveUniqueIdEvenWhenSharingRelationshipId() throws {
        let xml1 = """
        <w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId5"><w:r><w:t>A</w:t></w:r></w:hyperlink>
        """
        let xml2 = """
        <w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId5"><w:r><w:t>B</w:t></w:r></w:hyperlink>
        """
        let el1 = try XMLElement(xmlString: xml1)
        let el2 = try XMLElement(xmlString: xml2)

        let hl1 = try DocxReader.parseHyperlink(
            from: el1,
            relationships: RelationshipsCollection(),
            position: 3
        )
        let hl2 = try DocxReader.parseHyperlink(
            from: el2,
            relationships: RelationshipsCollection(),
            position: 7
        )

        XCTAssertNotEqual(
            hl1.id, hl2.id,
            "Two hyperlinks sharing r:id but at different positions must get distinct ids. Got both = \"\(hl1.id)\""
        )
    }
}
