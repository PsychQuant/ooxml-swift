import XCTest
@testable import OOXMLSwift

/// Phase 0 of word-aligned-state-sync: prove `XmlTreeReader` + `XmlTreeWriter`
/// round-trip XML byte-equal on untouched sub-trees.
///
/// Acceptance criterion (spec `ooxml-tree-io`, "Round-trip golden corpus"):
///   read → write with no mutations produces byte-equal output for the
///   four-fixture corpus (multi-section, VML-rich, CJK settings, comments).
///
/// This file starts with synthetic fixtures embedded in the test source so
/// the prototype can be validated before committing real .docx files. The
/// real fixtures land in task 1.8 (and live under `Tests/.../Fixtures/`).
final class TreeRoundTripGoldenTests: XCTestCase {

    // MARK: - Round-trip identity on synthetic fixtures

    func testEmptyElementRoundTrip() throws {
        let input = #"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#
        try assertByteEqualRoundTrip(input)
    }

    func testElementWithAttributesRoundTrip() throws {
        let input = #"<?xml version="1.0" encoding="UTF-8"?><w:p w:rsidR="00ABC123" w14:paraId="0AB7C123" xmlns:w="ns_w" xmlns:w14="ns_w14"/>"#
        try assertByteEqualRoundTrip(input)
    }

    func testNestedElementsRoundTrip() throws {
        let input = #"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="ns_w"><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>"#
        try assertByteEqualRoundTrip(input)
    }

    func testMixedContentRoundTrip() throws {
        // Spec scenario: "Mixed content order is preserved".
        let input = #"<?xml version="1.0" encoding="UTF-8"?><w:r xmlns:w="ns_w"><w:t>foo</w:t><w:tab/><w:t>bar</w:t></w:r>"#
        try assertByteEqualRoundTrip(input)
    }

    func testRsidsListRoundTrip() throws {
        // Spec example: "rsids preservation" (300+ rsid entries shrink-test).
        var rsidEntries = ""
        for i in 0..<50 {
            rsidEntries += String(format: #"<w:rsid w:val="00%06X"/>"#, i * 0x12345)
        }
        let input = #"<?xml version="1.0" encoding="UTF-8"?><w:settings xmlns:w="ns_w"><w:rsids>"# +
                     rsidEntries +
                     #"</w:rsids></w:settings>"#
        try assertByteEqualRoundTrip(input)
    }

    func testCommentAndProcessingInstructionRoundTrip() throws {
        let input = #"<?xml version="1.0" encoding="UTF-8"?><!-- a leading comment --><?someTarget some data?><w:document xmlns:w="ns_w"><!--inline comment--><w:p/></w:document>"#
        try assertByteEqualRoundTrip(input)
    }

    func testSelfClosingAndExpandedFormsRoundTrip() throws {
        // Self-closing <w:tab/> must NOT become <w:tab></w:tab>.
        // Expanded form with explicit children must NOT collapse to self-close.
        let input = #"<?xml version="1.0" encoding="UTF-8"?><w:r xmlns:w="ns_w"><w:tab/><w:t></w:t></w:r>"#
        try assertByteEqualRoundTrip(input)
    }

    func testEntityEncodedAttributeAndTextRoundTrip() throws {
        // Reader decodes entities; writer must re-encode the same set so the
        // bytes match.
        let input = #"<?xml version="1.0" encoding="UTF-8"?><w:r xmlns:w="ns_w" w:val="A &amp; B &lt; C"><w:t>x &amp; y &lt; z</w:t></w:r>"#
        try assertByteEqualRoundTrip(input)
    }

    func testCJKContentRoundTrip() throws {
        let input = #"<?xml version="1.0" encoding="UTF-8"?><w:p xmlns:w="ns_w"><w:r><w:t>標楷體 假設 H₀ 為樣本平均</w:t></w:r></w:p>"#
        try assertByteEqualRoundTrip(input)
    }

    // MARK: - Stable identity

    func testStableIDDerivedFromW14ParaId() throws {
        let input = #"<?xml version="1.0" encoding="UTF-8"?><w:p xmlns:w="ns_w" xmlns:w14="ns_w14" w14:paraId="0AB7C123"/>"#
        let tree = try XmlTreeReader.parse(Data(input.utf8))
        XCTAssertEqual(tree.root.stableID, "w14:paraId=0AB7C123")
    }

    func testStableIDFallsBackToCommentIDWhenParaIdAbsent() throws {
        let input = #"<?xml version="1.0" encoding="UTF-8"?><w:comment xmlns:w="ns_w" w:id="42"/>"#
        let tree = try XmlTreeReader.parse(Data(input.utf8))
        XCTAssertEqual(tree.root.stableID, "w:id=42")
    }

    func testStableIDIsNilForElementsWithoutOOXMLId() throws {
        let input = #"<?xml version="1.0" encoding="UTF-8"?><w:t xmlns:w="ns_w">Hello</w:t>"#
        let tree = try XmlTreeReader.parse(Data(input.utf8))
        XCTAssertNil(tree.root.stableID)
    }

    // MARK: - Mutation invalidates source range

    func testMutatedNodeReSerializesFromTypedFields() throws {
        // Read a doc, mutate one attribute, write. The mutated element must
        // be re-emitted from typed fields; surrounding bytes are byte-equal
        // to source.
        let input = #"<?xml version="1.0" encoding="UTF-8"?><w:r xmlns:w="ns_w"><w:t>before</w:t><w:tab/></w:r>"#
        let tree = try XmlTreeReader.parse(Data(input.utf8))
        // Mutate the inner <w:t> textContent.
        XCTAssertEqual(tree.root.children.count, 2)
        let textElement = tree.root.children[0]
        XCTAssertEqual(textElement.localName, "t")
        XCTAssertEqual(textElement.children.count, 1)
        let textNode = textElement.children[0]
        XCTAssertEqual(textNode.kind, .text)
        XCTAssertEqual(textNode.textContent, "before")
        textNode.textContent = "after"
        XCTAssertTrue(textNode.isDirty)
        let output = try XmlTreeWriter.serialize(tree)
        let outputString = String(decoding: output, as: UTF8.self)
        XCTAssertTrue(outputString.contains("after"), "mutation must appear in output")
        XCTAssertFalse(outputString.contains("before"), "old value must not appear in output")
        // Surrounding sibling <w:tab/> stays byte-equal because it was clean.
        XCTAssertTrue(outputString.contains("<w:tab/>"))
    }

    // MARK: - Helpers

    private func assertByteEqualRoundTrip(_ xml: String, file: StaticString = #file, line: UInt = #line) throws {
        let data = Data(xml.utf8)
        let tree = try XmlTreeReader.parse(data)
        let output = try XmlTreeWriter.serialize(tree)
        if output != data {
            // Provide diagnostic on diff.
            let inputString = String(decoding: data, as: UTF8.self)
            let outputString = String(decoding: output, as: UTF8.self)
            XCTFail(
                "Round-trip not byte-equal.\n--- input ---\n\(inputString)\n--- output ---\n\(outputString)",
                file: file, line: line
            )
        }
    }
}
