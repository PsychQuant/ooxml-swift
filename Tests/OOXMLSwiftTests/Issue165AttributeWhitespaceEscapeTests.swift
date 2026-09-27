import Foundation
import XCTest
@testable import OOXMLSwift

/// Regression tests for [PsychQuant/ooxml-swift#165](https://github.com/PsychQuant/ooxml-swift/issues/165).
///
/// XML 1.0 §3.3.3: a conformant reader normalizes a LITERAL newline / CR /
/// tab inside an attribute value to a plain space; only a character
/// reference (`&#10;` / `&#13;` / `&#9;`) survives intact. Before this fix,
/// `escapeXMLAttribute` (used broadly across `Models/*.swift` typed
/// emission) and `DocxWriter.escapeAttr` (document root tag emission) wrote
/// such characters literally — the source's `&#10;` survived the FIRST
/// parse -> emit round trip (nothing re-escaped it wrong yet), but the
/// SECOND parse silently collapsed it to a space.
final class Issue165AttributeWhitespaceEscapeTests: XCTestCase {

    // MARK: - Section A: direct unit test of the escape function

    func testEscapeXMLAttributeEscapesControlWhitespaceAsCharacterReferences() {
        XCTAssertEqual(escapeXMLAttribute("a\nb\rc\td"), "a&#10;b&#13;c&#9;d")
    }

    // MARK: - Section B: RunProperties.rStyle — typed emission via Run.swift

    private func parseRunFragment(_ xml: String) throws -> Run {
        let wrapper = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          \(xml)
        </w:document>
        """
        let doc = try XMLDocument(xmlString: wrapper)
        guard let root = doc.rootElement(),
              let runElement = root.elements(forName: "w:r").first else {
            throw NSError(domain: "test", code: 0, userInfo: [NSLocalizedDescriptionKey: "no <w:r> found"])
        }
        return try DocxReader.parseRun(from: runElement, relationships: RelationshipsCollection())
    }

    /// Source encodes the newline as a character reference — the form Word
    /// itself, and any spec-conformant writer, uses for embedded whitespace.
    func testRStyleNewlineSurvivesParseEmitReparseRoundTrip() throws {
        let source = "<w:r><w:rPr><w:rStyle w:val=\"A&#10;B\"/></w:rPr><w:t>x</w:t></w:r>"
        let run = try parseRunFragment(source)
        XCTAssertEqual(run.properties.rStyle, "A\nB",
                       "first parse must decode the character reference to an actual newline")

        let emitted = run.toXML()
        XCTAssertFalse(emitted.contains("w:val=\"A\nB\""),
                       "toXML() must not write the newline back literally — got: \(emitted)")
        XCTAssertTrue(emitted.contains("&#10;") || emitted.contains("&#xA;"),
                      "toXML() must re-escape the newline as a character reference — got: \(emitted)")

        let reparsed = try parseRunFragment(emitted)
        XCTAssertEqual(reparsed.properties.rStyle, "A\nB",
                       "SECOND parse must still see the real newline, not a space — this is the bug's failure mode")
    }

    // MARK: - Section C: Drawing description (`wp:docPr/@descr`) — the issue's own named example

    func testDrawingDescriptionNewlineSurvivesEmitReparseRoundTrip() throws {
        let drawing = Drawing(type: .inline, width: 100, height: 100, imageId: "rId1",
                               name: "Pic", description: "line1\nline2")
        let emitted = drawing.toXML()
        XCTAssertFalse(emitted.contains("descr=\"line1\nline2\""),
                       "toXML() must not write the newline back literally — got: \(emitted)")

        // Wrap the fragment so it parses standalone, extract wp:docPr/@descr,
        // and confirm what a conformant XML parser reads back.
        let wrapper = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          \(emitted)
        </w:r>
        """
        let doc = try XMLDocument(xmlString: wrapper)
        guard let drawing = doc.rootElement()?.elements(forName: "w:drawing").first,
              let inline = drawing.elements(forName: "wp:inline").first,
              let docPr = inline.elements(forName: "wp:docPr").first else {
            return XCTFail("no wp:docPr found in emitted XML: \(emitted)")
        }
        XCTAssertEqual(docPr.attribute(forName: "descr")?.stringValue, "line1\nline2",
                       "SECOND parse must still see the real newline, not a space")
    }
}
