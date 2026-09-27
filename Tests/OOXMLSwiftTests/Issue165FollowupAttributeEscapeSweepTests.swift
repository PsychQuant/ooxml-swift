import Foundation
import XCTest
@testable import OOXMLSwift

/// R2 follow-up for [PsychQuant/ooxml-swift#165](https://github.com/PsychQuant/ooxml-swift/issues/165)
/// (independent review finding F2).
///
/// R1 fixed the two functions #165 named as "Technical anchors" plus
/// `Image.escapeXML`. Independent review found the SAME bug class (an
/// attribute-value escaper missing `\n`/`\r`/`\t`) still present in several
/// OTHER existing call sites, none of which #165 named: bookmark names,
/// comment author/initials, tracked-revision author, content-control
/// tag/alias/placeholder/data-binding attributes, and hyperlink
/// tooltip/anchor/rId. Each of these previously duplicated the same
/// 5-character (`& < > " '`) escaper inline; all now delegate to the
/// already-fixed `escapeXMLAttribute`.
///
/// Each test below: build the value with an embedded literal `\n` -> emit
/// the model's own `toXML()`-family method -> wrap in a minimal namespaced
/// document -> parse with a REAL XML parser -> assert the attribute value
/// still contains a real `\n`, not a space. On the pre-fix per-site escaper
/// (literal `\n` written unescaped) this last assertion fails, because a
/// conformant parser normalizes a literal newline in an attribute value to
/// a space (XML 1.0 §3.3.3) — exactly the failure mode #165 itself
/// describes.
final class Issue165FollowupAttributeEscapeSweepTests: XCTestCase {

    private static let wNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    private static let rNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    /// Wraps `fragment` in a namespaced root, parses it, and returns the
    /// value of `attribute` on the first element named `elementLocalName`.
    private func roundTrippedAttribute(
        _ fragment: String, elementLocalName: String, attribute: String
    ) throws -> String? {
        let wrapper = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:root xmlns:w="\(Self.wNS)" xmlns:r="\(Self.rNS)"
                xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
                xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
          \(fragment)
        </w:root>
        """
        let doc = try XMLDocument(xmlString: wrapper)
        guard let root = doc.rootElement() else {
            throw NSError(domain: "test", code: 0, userInfo: [NSLocalizedDescriptionKey: "no root element"])
        }
        func findDescendant(_ node: XMLElement) -> XMLElement? {
            if node.name == elementLocalName { return node }
            for child in node.children ?? [] {
                guard let childElement = child as? XMLElement else { continue }
                if let found = findDescendant(childElement) { return found }
            }
            return nil
        }
        guard let element = findDescendant(root) else {
            throw NSError(domain: "test", code: 0, userInfo: [
                NSLocalizedDescriptionKey: "no <\(elementLocalName)> found in: \(fragment)"])
        }
        return element.attribute(forName: attribute)?.stringValue
    }

    // MARK: - Bookmark.swift — w:name

    func testBookmarkNameNewlineSurvivesRoundTrip() throws {
        let bookmark = Bookmark(id: 1, name: "A\nB")
        let xml = bookmark.toBookmarkStartXML()
        XCTAssertFalse(xml.contains("w:name=\"A\nB\""), "must not write the newline literally — got: \(xml)")
        let value = try roundTrippedAttribute(xml, elementLocalName: "w:bookmarkStart", attribute: "w:name")
        XCTAssertEqual(value, "A\nB")
    }

    // MARK: - Comment.swift — w:author / w:initials

    func testCommentAuthorAndInitialsNewlineSurviveRoundTrip() throws {
        let comment = Comment(id: 1, author: "A\nB", text: "body", paragraphIndex: 0, initials: "C\nD")
        let xml = comment.toXML()
        let author = try roundTrippedAttribute(xml, elementLocalName: "w:comment", attribute: "w:author")
        let initials = try roundTrippedAttribute(xml, elementLocalName: "w:comment", attribute: "w:initials")
        XCTAssertEqual(author, "A\nB")
        XCTAssertEqual(initials, "C\nD")
    }

    // MARK: - Revision.swift — w:author (12 call sites share one function)

    func testRevisionAuthorNewlineSurvivesRoundTrip() throws {
        let revision = Revision(id: 1, type: .insertion, author: "A\nB")
        let xml = revision.toOpeningXML() + revision.toClosingXML()
        let value = try roundTrippedAttribute(xml, elementLocalName: "w:ins", attribute: "w:author")
        XCTAssertEqual(value, "A\nB")
    }

    // MARK: - Field.swift StructuredDocumentTag — w:tag / w:alias / w:docPart val

    func testStructuredDocumentTagAttributesNewlineSurviveRoundTrip() throws {
        let sdt = StructuredDocumentTag(tag: "A\nB", alias: "C\nD", placeholder: "E\nF")
        let xml = sdt.toSdtPrXML()
        let tag = try roundTrippedAttribute(xml, elementLocalName: "w:tag", attribute: "w:val")
        let alias = try roundTrippedAttribute(xml, elementLocalName: "w:alias", attribute: "w:val")
        let placeholder = try roundTrippedAttribute(xml, elementLocalName: "w:docPart", attribute: "w:val")
        XCTAssertEqual(tag, "A\nB")
        XCTAssertEqual(alias, "C\nD")
        XCTAssertEqual(placeholder, "E\nF")
    }

    // MARK: - Field.swift DataBinding — w:xpath / w:storeItemID / w:prefixMappings

    func testDataBindingAttributesNewlineSurviveRoundTrip() throws {
        let binding = DataBinding(xpath: "A\nB", storeItemID: "C\nD", prefixMappings: "E\nF")
        let xml = binding.toXML()
        let xpath = try roundTrippedAttribute(xml, elementLocalName: "w:dataBinding", attribute: "w:xpath")
        let storeItemID = try roundTrippedAttribute(xml, elementLocalName: "w:dataBinding", attribute: "w:storeItemID")
        let prefixMappings = try roundTrippedAttribute(xml, elementLocalName: "w:dataBinding", attribute: "w:prefixMappings")
        XCTAssertEqual(xpath, "A\nB")
        XCTAssertEqual(storeItemID, "C\nD")
        XCTAssertEqual(prefixMappings, "E\nF")
    }

    // MARK: - Hyperlink.swift — w:tooltip / w:anchor (review-list miss, caught by grep sweep)

    func testHyperlinkTooltipAndAnchorNewlineSurviveRoundTrip() throws {
        let hyperlink = Hyperlink(id: "1", text: "link", anchor: "A\nB", tooltip: "C\nD")
        let xml = hyperlink.toXML()
        let anchor = try roundTrippedAttribute(xml, elementLocalName: "w:hyperlink", attribute: "w:anchor")
        let tooltip = try roundTrippedAttribute(xml, elementLocalName: "w:hyperlink", attribute: "w:tooltip")
        XCTAssertEqual(anchor, "A\nB")
        XCTAssertEqual(tooltip, "C\nD")
    }

    // MARK: - ContentTypesOverlay.swift / RelationshipsOverlay.swift — grep sweep, lower-risk but same class

    func testContentTypesOverlayEscapeHandlesNewline() {
        // `escape` is private; exercised indirectly is impractical here (it
        // needs a full archive), so this pins the shared delegation directly
        // — the same guarantee `escapeXMLAttribute`'s own unit test gives,
        // kept here as a sweep-completeness marker for the inventory table.
        XCTAssertEqual(escapeXMLAttribute("A\nB"), "A&#10;B")
    }
}
