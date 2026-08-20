import XCTest
@testable import OOXMLSwift

/// che-word-mcp#176 residue — `<w:cols w:space>`.
///
/// `SectionProperties` models the column *count* but not the gutter between
/// columns, and the writer emits `w:space="720"` as a literal. So a document
/// whose sections use any other gutter has it silently replaced the moment
/// anything re-serializes the section — which is any edit that dirties the
/// typed model, `update_cell` and `replace_text` among them.
///
/// The reader already carries a comment naming this as a known narrower gap
/// ("a non-720 `<w:cols w:space>`"). The other half of that comment —
/// `<w:docGrid w:type>` — has since been closed; this is what is left.
final class SectionColumnSpacingTests: XCTestCase {

    private func sectPr(_ inner: String) throws -> XMLElement {
        let xml = """
        <w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\
        \(inner)</w:sectPr>
        """
        return try XMLDocument(xmlString: xml, options: []).rootElement()!
    }

    func testReaderCapturesNonDefaultColumnSpacing() throws {
        let props = DocxReader.parseSectionProperties(try sectPr("<w:cols w:space=\"425\"/>"))
        XCTAssertEqual(props.columnSpacing, 425,
                       "a non-720 w:space SHALL survive the read")
    }

    func testDefaultSpacingIsStill720() throws {
        let props = DocxReader.parseSectionProperties(try sectPr("<w:cols w:num=\"1\"/>"))
        XCTAssertEqual(props.columnSpacing, 720,
                       "absent w:space keeps the OOXML default, so existing output is unchanged")
    }

    func testWriterEmitsTheCapturedSpacing() throws {
        let props = DocxReader.parseSectionProperties(try sectPr("<w:cols w:num=\"2\" w:space=\"425\"/>"))
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:cols w:num=\"2\" w:space=\"425\"/>"),
                      "the emitted gutter SHALL be the one that was read, got: \(xml)")
    }

    /// The writer's attribute order is frozen by the transcoder's canonical
    /// form. Changing the *value* must not change the *shape*, or every
    /// byte-equality fixture moves with it.
    func testDefaultCaseStillEmitsTheFrozenCanonicalForm() throws {
        var props = SectionProperties()
        props.columns = 1
        XCTAssertTrue(props.toXML().contains("<w:cols w:num=\"1\" w:space=\"720\"/>"),
                      "default output SHALL be byte-identical to the pre-change form")
    }
}
