import XCTest
@testable import OOXMLSwift

/// Phase A regression test for che-word-mcp#52.
///
/// Spec: `openspec/changes/che-word-mcp-header-footer-raw-element-preservation/specs/ooxml-header-footer-raw-element-preservation/spec.md`
/// Requirement: "Run preserves unknown OOXML child elements via rawElements carrier"
///
/// 3 spec scenarios:
/// 1. VML watermark Run round-trips byte-equal
/// 2. Run with multiple unknown elements preserves source order
/// 3. Equatable conformance treats nil and missing-rawElements as equal
final class RunRawElementPreservationTests: XCTestCase {

    // MARK: - Helper: parse a fragment <w:r>...</w:r> into a Run

    private func parseRunFragment(_ xml: String) throws -> Run {
        let wrapper = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:v="urn:schemas-microsoft-com:vml"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
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

    // MARK: - Scenario 1: VML watermark Run round-trips byte-equal

    func testVMLWatermarkRunPreservesPictElement() throws {
        let source = """
        <w:r>
          <w:rPr><w:noProof/></w:rPr>
          <w:pict>
            <v:shape id="WordPictureWatermark" type="#_x0000_t136" fillcolor="silver">
              <v:textpath style="font-family:&quot;Arial&quot;" string="DRAFT"/>
            </v:shape>
          </w:pict>
        </w:r>
        """
        let run = try parseRunFragment(source)

        XCTAssertEqual(run.text, "", "VML watermark Run has no <w:t>; text SHALL be empty")
        XCTAssertNil(run.drawing, "<w:pict> is NOT <w:drawing>; drawing SHALL be nil")
        guard let raw = run.rawElements, raw.count == 1 else {
            return XCTFail("rawElements SHALL contain exactly one entry; got \(String(describing: run.rawElements))")
        }
        XCTAssertEqual(raw[0].name, "pict",
                       "rawElements[0].name SHALL be 'pict'")
        XCTAssertTrue(raw[0].xml.contains("v:textpath"),
                      "rawElements[0].xml SHALL contain the verbatim VML; got: \(raw[0].xml)")
        XCTAssertTrue(raw[0].xml.contains("string=\"DRAFT\""),
                      "VML attributes SHALL survive round-trip; got: \(raw[0].xml)")

        // Round-trip via toXML
        let emitted = run.toXML()
        XCTAssertTrue(emitted.contains("v:textpath"),
                      "Run.toXML() SHALL emit verbatim <w:pict>; got: \(emitted)")
        XCTAssertTrue(emitted.contains("string=\"DRAFT\""),
                      "VML attribute string=\"DRAFT\" SHALL survive Run.toXML(); got: \(emitted)")
    }

    // MARK: - Scenario 2: multiple unknown elements preserve source order

    func testRunWithMultipleUnknownElementsPreservesSourceOrder() throws {
        let source = """
        <w:r>
          <w:pict><v:shape id="P1"/></w:pict>
          <w:object><v:shape id="O1"/></w:object>
        </w:r>
        """
        let run = try parseRunFragment(source)

        guard let raw = run.rawElements, raw.count == 2 else {
            return XCTFail("rawElements SHALL contain 2 entries; got \(String(describing: run.rawElements))")
        }
        XCTAssertEqual(raw[0].name, "pict",
                       "rawElements[0] SHALL be the first source element 'pict'")
        XCTAssertEqual(raw[1].name, "object",
                       "rawElements[1] SHALL be the second source element 'object'")

        // toXML emits in array order
        let emitted = run.toXML()
        guard let pictRange = emitted.range(of: "<w:pict>"),
              let objectRange = emitted.range(of: "<w:object>") else {
            return XCTFail("Both <w:pict> and <w:object> SHALL appear in emitted XML; got: \(emitted)")
        }
        XCTAssertLessThan(pictRange.lowerBound, objectRange.lowerBound,
                          "<w:pict> SHALL precede <w:object> in toXML() output, matching source order")
    }

    // MARK: - Scenario 3: Equatable conformance treats nil and missing-rawElements as equal

    func testEquatableConformanceWithDefaultRawElementsNil() throws {
        let r1 = Run(text: "hello")  // programmatic, rawElements defaults to nil
        let r2 = try parseRunFragment("<w:r><w:t>hello</w:t></w:r>")  // parsed, no unknowns → nil

        XCTAssertNil(r1.rawElements, "Programmatic Run(text:) SHALL have nil rawElements")
        XCTAssertNil(r2.rawElements, "parseRun on <w:r> with no unknowns SHALL set rawElements to nil (NOT empty array)")
        XCTAssertEqual(r1, r2, "Programmatic Run vs parsed Run with no unknowns SHALL be Equatable-equal")
    }
}
