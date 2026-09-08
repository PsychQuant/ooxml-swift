import XCTest
@testable import OOXMLSwift

/// PsychQuant/ooxml-swift#129 — a tree-backed paragraph's `runs` must not drop
/// what a run carries.
///
/// The getter built `Run(text:)` for every `<w:r>` child: a text-only stub. A
/// `<w:drawing>`, an `<w:rPr>`, a tab, a break — none of it survived the round
/// trip through the typed view, so any operation that marked
/// `word/document.xml` typed-dirty re-serialised the body without them. 13 of
/// 27 real documents lost images that way (171 -> 3, 36 -> 1, 14 -> 1).
///
/// These are **known-failure guards**, not a fix. The fix needs one primitive
/// this library does not have: serialising a single `XmlNode` back to XML.
/// `XmlTreeWriter.serialize` takes a whole `XmlTree` plus its source bytes, so
/// there is no way to ask a `<w:r>` node for its own markup — which is exactly
/// what the getter would have to carry into `Run.rawXML` for the writer to emit
/// it faithfully.
///
/// #106 is blocked on the same primitive: its guard says "needs node-level XML
/// serialization before resync can re-type this". Two issues, one missing
/// piece — worth knowing before either is scheduled, because building it once
/// unblocks both, and building it twice is how they drift.
///
/// When node-level serialisation lands, these turn green and `XCTExpectFailure`
/// starts failing as an unexpected pass, which is the signal to delete it.
final class Issue129TreeBackedDrawingTests: XCTestCase {

    private func treeBackedParagraph(_ innerXML: String) throws -> Paragraph {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" \
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" \
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" \
        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">\
        <w:body><w:p>\(innerXML)</w:p></w:body></w:document>
        """
        let tree = try XmlTreeReader.parse(Data(xml.utf8))
        // the <w:p> under <w:document>/<w:body>
        let body = try XCTUnwrap(tree.root.children.first { $0.kind == .element && $0.localName == "body" })
        let p = try XCTUnwrap(body.children.first { $0.kind == .element && $0.localName == "p" })
        return Paragraph(xmlNode: p)
    }

    /// The shape from the issue: a run holding a drawing.
    func testATreeBackedRunKeepsItsDrawing() throws {
        XCTExpectFailure("PsychQuant/ooxml-swift#129 — a tree-backed run is a text-only stub; the fix needs node-level XML serialisation, shared with #106")
        let drawingRun = """
        <w:r><w:drawing><wp:inline><wp:extent cx="100" cy="100"/>\
        <a:graphic><a:graphicData uri="urn:pic"><pic:pic xmlns:pic="urn:pic">\
        <pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill></pic:pic>\
        </a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
        """
        let para = try treeBackedParagraph(drawingRun)
        XCTAssertEqual(para.runs.count, 1, "one <w:r> child, one run")
        // What the writer would emit for this paragraph, which is where the
        // loss showed up: `Run.toXML()` prefers `rawXML` when it is present.
        let emitted = para.runs.map { $0.toXML() }.joined()
        XCTAssertTrue(emitted.contains("<w:drawing>"), "the drawing survives the typed view: \(emitted)")
        XCTAssertTrue(emitted.contains("r:embed=\"rId7\""), "the image relationship survives: \(emitted)")
    }

    /// Not only drawings: anything a run holds beyond its text.
    func testATreeBackedRunKeepsFormattingTabsAndBreaks() throws {
        XCTExpectFailure("PsychQuant/ooxml-swift#129 — a tree-backed run is a text-only stub; the fix needs node-level XML serialisation, shared with #106")
        let para = try treeBackedParagraph("""
        <w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>bold red</w:t></w:r>\
        <w:r><w:tab/><w:br/><w:t>after</w:t></w:r>
        """)
        XCTAssertEqual(para.runs.count, 2)
        let emitted = para.runs.map { $0.toXML() }.joined()
        XCTAssertTrue(emitted.contains("<w:b/>"), "bold survives: \(emitted)")
        XCTAssertTrue(emitted.contains("FF0000"), "colour survives: \(emitted)")
        XCTAssertTrue(emitted.contains("<w:tab/>"), "tab survives: \(emitted)")
        XCTAssertTrue(emitted.contains("<w:br/>"), "break survives: \(emitted)")
    }

    /// The text view keeps working — the fix must not change what `text` reads,
    /// which existing callers depend on.
    func testTheTextViewIsUnchanged() throws {
        let para = try treeBackedParagraph("<w:r><w:t>alpha</w:t></w:r><w:r><w:t> beta</w:t></w:r>")
        XCTAssertEqual(para.text, "alpha beta")
        XCTAssertEqual(para.runs.map(\.text), ["alpha", " beta"])
    }

    /// A detached paragraph is untouched by any of this.
    func testDetachedParagraphsAreUnaffected() {
        var para = Paragraph()
        para.runs = [Run(text: "plain")]
        XCTAssertEqual(para.runs.map(\.text), ["plain"])
        XCTAssertNil(para.runs.first?.rawXML, "a detached run carries no raw XML unless it was given some")
    }
}
