import Foundation
import XCTest
@testable import OOXMLSwift

/// Regression tests for [PsychQuant/ooxml-swift#179](https://github.com/PsychQuant/ooxml-swift/issues/179).
///
/// #176 taught `DocxReader` to project `<w:pBdr>`/`<w:shd>` into typed
/// `ParagraphProperties.border`/`.shading`, and to re-emit the SOURCE XML
/// verbatim (via `sourceBorder`/`sourceShading`) as long as those fields are
/// never reassigned. But the typed model itself could not express
/// `themeColor`/`themeTint`/`themeShade`/`shadow`/`frame` (on each border
/// side) or the `bar` side, nor the `theme*` shading attributes — so the
/// MOMENT a caller reassigns `border`/`shading` (even to an equal value, or
/// changes only one side), the source XML is discarded and the typed
/// re-emission silently drops every one of those attributes.
///
/// After this fix: `ParagraphBorderStyle` carries `themeColor`/`themeTint`/
/// `themeShade`/`shadow`/`frame`; `ParagraphBorder` carries `bar`;
/// `CellShading` (aliased `ParagraphShading`) carries the six `theme*`
/// shading attributes. `DocxReader` reads them into the typed projection and
/// `toXML()` re-emits them, so "read out, change one side, write back" only
/// changes the side that was actually touched.
final class Issue179ParagraphBorderShadingThemeTests: XCTestCase {

    // MARK: - Section A: unit-level parse/emit round trip (no docx I/O)

    private func parseParagraph(_ xml: String) throws -> Paragraph {
        let element = try XMLElement(xmlString: xml)
        return try DocxReader.parseParagraph(
            from: element, relationships: RelationshipsCollection(), styles: [], numbering: Numbering())
    }

    private static let themedParagraphXML = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\
        <w:pPr><w:pBdr>\
        <w:top w:val="single" w:sz="4" w:space="1" w:color="auto" w:themeColor="accent1" w:themeTint="66" w:shadow="1"/>\
        <w:left w:val="single" w:sz="4" w:space="1" w:color="auto" w:themeColor="accent1" w:themeTint="66"/>\
        <w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto" w:themeColor="accent1" w:themeTint="66"/>\
        <w:right w:val="single" w:sz="4" w:space="1" w:color="auto" w:themeColor="accent1" w:themeTint="66"/>\
        <w:bar w:val="single" w:sz="4" w:space="1" w:color="auto" w:frame="1"/>\
        </w:pBdr>\
        <w:shd w:val="clear" w:color="auto" w:fill="auto" w:themeFill="accent2" w:themeFillTint="33"/>\
        </w:pPr>\
        <w:r><w:t>THEMED_PARAGRAPH</w:t></w:r></w:p>
        """

    func testParseCapturesThemeAndBarAttributes() throws {
        let paragraph = try parseParagraph(Self.themedParagraphXML)
        let top = try XCTUnwrap(paragraph.properties.border?.top)
        XCTAssertEqual(top.themeColor, "accent1")
        XCTAssertEqual(top.themeTint, "66")
        XCTAssertEqual(top.shadow, true)

        let bar = try XCTUnwrap(paragraph.properties.border?.bar)
        XCTAssertEqual(bar.frame, true)

        let shading = try XCTUnwrap(paragraph.properties.shading)
        XCTAssertEqual(shading.themeFill, "accent2")
        XCTAssertEqual(shading.themeFillTint, "33")
    }

    /// The issue's exact scenario 1: read out, set the SAME border back
    /// (`props.border = props.border`) — clears `sourceBorder`, forcing typed
    /// re-emission. Theme attributes must still survive.
    func testReassigningBorderToItselfPreservesThemeAttributes() throws {
        var paragraph = try parseParagraph(Self.themedParagraphXML)
        paragraph.properties.border = paragraph.properties.border
        let xml = paragraph.toXML()
        XCTAssertTrue(xml.contains("w:themeColor=\"accent1\""), "got: \(xml)")
        XCTAssertTrue(xml.contains("w:themeTint=\"66\""), "got: \(xml)")
        XCTAssertTrue(xml.contains("<w:bar"), "bar side must survive — got: \(xml)")
        XCTAssertTrue(xml.contains("w:frame=\"1\""), "got: \(xml)")
    }

    /// The issue's exact scenario 2: change only ONE side's color — the
    /// OTHER sides' theme attributes must still survive (not just the
    /// changed side's own typed fields).
    func testChangingOneSidePreservesOtherSidesThemeAttributes() throws {
        var paragraph = try parseParagraph(Self.themedParagraphXML)
        var border = paragraph.properties.border!
        border.bottom?.color = "FF0000"
        paragraph.properties.border = border
        let xml = paragraph.toXML()

        // All four sides' themeColor survive — bottom's own themeColor was
        // never touched, only its `color` fallback was reassigned.
        let themeColorCount = xml.components(separatedBy: "w:themeColor=\"accent1\"").count - 1
        XCTAssertEqual(themeColorCount, 4, "every side must keep themeColor — got: \(xml)")
        // The changed side's new color must apply.
        XCTAssertTrue(xml.contains("w:color=\"FF0000\""), "got: \(xml)")
    }

    func testReassigningShadingToItselfPreservesThemeFillAttributes() throws {
        var paragraph = try parseParagraph(Self.themedParagraphXML)
        paragraph.properties.shading = paragraph.properties.shading
        let xml = paragraph.toXML()
        XCTAssertTrue(xml.contains("w:themeFill=\"accent2\""), "got: \(xml)")
        XCTAssertTrue(xml.contains("w:themeFillTint=\"33\""), "got: \(xml)")
    }

    // MARK: - Section B: parse -> emit -> reparse round trip

    func testParseThenEmitRoundTripsThemeAttributesVerbatim() throws {
        let paragraph = try parseParagraph(Self.themedParagraphXML)
        let emitted = paragraph.toXML().replacingOccurrences(
            of: "<w:p>",
            with: "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">")
        let reparsed = try parseParagraph(emitted)
        XCTAssertEqual(reparsed.properties.border?.top?.themeColor, "accent1")
        XCTAssertEqual(reparsed.properties.border?.bar?.frame, true)
        XCTAssertEqual(reparsed.properties.shading?.themeFill, "accent2")
    }
}
