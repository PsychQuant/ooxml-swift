// WordCanonicalFormsTests.swift
// word-canonical-forms Phase 2 — Word-canonical form vocabulary tests.
// 2.1 setDocumentRoot (root namespace cloud), 2.2 rsid, 2.3 xml:space,
// 2.4 inline-passthrough markers. Each new form must (a) round-trip the
// JSONL wire additively, (b) stamp back byte-equal via the reducer, and
// (c) let a source carrying it pass the trial-rebuild upgrade gate.

import XCTest
@testable import OOXMLSwift

final class WordCanonicalFormsTests: XCTestCase {

    private func documentRoot(_ doc: WordDocument) throws -> XmlNode {
        try XCTUnwrap(doc.xmlTrees["word/document.xml"]).root
    }

    private func attrPairs(_ node: XmlNode) -> [String] {
        node.attributes.map { ($0.prefix.map { "\($0):" } ?? "") + $0.localName + "=" + $0.value }
    }

    /// Reverse a hand-built document.xml + rebuild; returns (upgraded, rebuilt bytes).
    private func roundTrip(documentXML: String) throws -> (upgraded: Bool, rebuilt: Data?) {
        let parts = ["word/document.xml": Data(documentXML.utf8)]
        let result = try ReverseExtractor.reverse(parts: parts)
        guard result.dslParts.contains("word/document.xml") else { return (false, nil) }
        let parsed = try ScriptImporter.parse(source: ScriptExporter.exportSwift(log: result.log))
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: parsed.entries.map(\.op))
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("wcf-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try doc.writeAuthoringPackage(to: url)
        return (true, try RawPartChannel.readAllParts(from: url)["word/document.xml"])
    }

    // MARK: - 2.1 setDocumentRoot

    func testSetDocumentRootRoundTripsWireOrderPreserved() throws {
        let attrs = [
            RootAttribute(prefix: "xmlns", localName: "w", value: "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
            RootAttribute(prefix: "xmlns", localName: "w14", value: "http://schemas.microsoft.com/office/word/2010/wordml"),
            RootAttribute(prefix: "xmlns", localName: "mc", value: "http://schemas.openxmlformats.org/markup-compatibility/2006"),
            RootAttribute(prefix: "xmlns", localName: "r", value: "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
            RootAttribute(prefix: "mc", localName: "Ignorable", value: "w14 w15"),
        ]
        var log = OperationLog()
        log.append(.setDocumentRoot(attributes: attrs), source: .swift)
        let decoded = try OperationLog.decodeJSONL(log.encodeJSONL())
        guard case .setDocumentRoot(let got) = decoded.entries[0].op else {
            return XCTFail("expected setDocumentRoot")
        }
        XCTAssertEqual(got, attrs, "attributes must round-trip field-for-field in order")
        XCTAssertEqual(log.encodeJSONL(), decoded.encodeJSONL(), "byte-equal re-encode")
    }

    func testSetDocumentRootStampsRootAttributesInOrder() throws {
        let attrs = [
            RootAttribute(prefix: "xmlns", localName: "w", value: "WNS"),
            RootAttribute(prefix: "xmlns", localName: "mc", value: "MCNS"),
            RootAttribute(prefix: "mc", localName: "Ignorable", value: "w14"),
        ]
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: [.setDocumentRoot(attributes: attrs)])
        XCTAssertEqual(attrPairs(try documentRoot(doc)),
                       ["xmlns:w=WNS", "xmlns:mc=MCNS", "mc:Ignorable=w14"],
                       "root attributes replaced wholesale in op order")
    }

    func testAbsentSetDocumentRootKeepsDefaultRoot() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(text: "t", paraId: "P1")),
        ])
        // Default authoring root unchanged (byte-identical to pre-extension).
        XCTAssertEqual(attrPairs(try documentRoot(doc)), [
            "xmlns:w=http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "xmlns:w14=http://schemas.microsoft.com/office/word/2010/wordml",
        ])
    }

    /// A source whose root carries an extra namespace beyond the authoring
    /// default extracts a setDocumentRoot op first and upgrades byte-equal.
    func testRootWithExtraNamespaceUpgradesByteEqual() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body><w:p w14:paraId="P1"><w:r><w:t>hi</w:t></w:r></w:p></w:body></w:document>
        """
        let (upgraded, rebuilt) = try roundTrip(documentXML: xml)
        XCTAssertTrue(upgraded, "root-with-extra-namespace document must upgrade")
        XCTAssertEqual(rebuilt, Data(xml.utf8), "rebuilt document.xml must be byte-equal")
    }

    // MARK: - 2.2 rsid + w14:textId (exact 90_template_ja shapes/order)

    /// Wire round-trip of the extended paragraph/run rsid fields.
    func testExtendedRsidPayloadsRoundTripWire() throws {
        let para = ParagraphPayload(
            text: "t", paraId: "1FE40057", textId: "740FF560",
            rsidR: "00040150", rsidRPr: "007F04E2", rsidRDefault: "00C74BEF", rsidP: "00F32D54")
        let run = RunPayload(text: "字", rsidRPr: "007F04E2", preserveSpace: true)
        var log = OperationLog()
        log.append(.appendParagraph(in: nil, paragraph: para), source: .swift)
        log.append(.setRuns(target: ElementID(rawString: "w14:paraId=1FE40057"), runs: [run]),
                   source: .swift)
        let decoded = try OperationLog.decodeJSONL(log.encodeJSONL())
        guard case .appendParagraph(_, let gotP) = decoded.entries[0].op,
              case .setRuns(_, let gotR) = decoded.entries[1].op else {
            return XCTFail("decode shape")
        }
        XCTAssertEqual(gotP, para)
        XCTAssertEqual(gotR, [run])
    }

    /// A paragraph in 90_template_ja's exact attribute shape/order
    /// (w14:paraId, w14:textId, w:rsidR, w:rsidRPr, w:rsidRDefault, w:rsidP)
    /// with a run carrying w:rsidRPr rebuilds byte-equal.
    func testRealParagraphRsidShapeRoundTripsByteEqual() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="1FE40057" w14:textId="740FF560" w:rsidR="00040150" w:rsidRPr="007F04E2" w:rsidRDefault="00C74BEF" w:rsidP="00F32D54"><w:r w:rsidRPr="007F04E2"><w:t>本文</w:t></w:r></w:p></w:body></w:document>
        """
        let (upgraded, rebuilt) = try roundTrip(documentXML: xml)
        XCTAssertTrue(upgraded, "real rsid paragraph must upgrade")
        XCTAssertEqual(rebuilt, Data(xml.utf8), "rebuilt must be byte-equal")
    }

    /// sectPr carrying w:rsidR + w:rsidSect rebuilds byte-equal.
    func testSectPrRsidRoundTripsByteEqual() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="P1"><w:r><w:t>x</w:t></w:r></w:p><w:sectPr w:rsidR="0081382F" w:rsidSect="00BC5145"><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>
        """
        let (upgraded, rebuilt) = try roundTrip(documentXML: xml)
        XCTAssertTrue(upgraded, "sectPr with rsid must upgrade")
        XCTAssertEqual(rebuilt, Data(xml.utf8))
    }

    // MARK: - 2.3 xml:space="preserve"

    func testPreserveSpaceRoundTripsByteEqual() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="P1"><w:r><w:t xml:space="preserve"> leading and trailing </w:t></w:r></w:p></w:body></w:document>
        """
        let (upgraded, rebuilt) = try roundTrip(documentXML: xml)
        XCTAssertTrue(upgraded, "xml:space paragraph must upgrade")
        XCTAssertEqual(rebuilt, Data(xml.utf8))
    }

    // MARK: - 2.4 inline-passthrough markers

    func testInlineMarkerWireRoundTrips() throws {
        let items: [InlineItem] = [
            .marker(InlineMarker(localName: "bookmarkStart", attributes: [
                RootAttribute(prefix: "w", localName: "id", value: "0"),
                RootAttribute(prefix: "w", localName: "name", value: "_Hlk96608833")])),
            .run(RunPayload(text: "本文")),
            .marker(InlineMarker(localName: "proofErr", attributes: [
                RootAttribute(prefix: "w", localName: "type", value: "gramStart")])),
            .marker(InlineMarker(localName: "bookmarkEnd", attributes: [
                RootAttribute(prefix: "w", localName: "id", value: "0")])),
        ]
        var log = OperationLog()
        log.append(.setParagraphContent(target: ElementID(rawString: "w14:paraId=P1"), items: items),
                   source: .swift)
        let decoded = try OperationLog.decodeJSONL(log.encodeJSONL())
        guard case .setParagraphContent(_, let got) = decoded.entries[0].op else {
            return XCTFail("expected setParagraphContent")
        }
        XCTAssertEqual(got, items, "inline items round-trip field-for-field in order")
        XCTAssertEqual(log.encodeJSONL(), decoded.encodeJSONL(), "byte-equal re-encode")
    }

    /// A paragraph with bookmarkStart/proofErr interleaved between runs
    /// rebuilds byte-equal (the 2.4 structural centerpiece).
    func testInterleavedMarkersRoundTripByteEqual() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="P1"><w:bookmarkStart w:id="0" w:name="_Hlk96608833"/><w:proofErr w:type="gramStart"/><w:r><w:t>本文です</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p></w:body></w:document>
        """
        let (upgraded, rebuilt) = try roundTrip(documentXML: xml)
        XCTAssertTrue(upgraded, "interleaved-marker paragraph must upgrade")
        XCTAssertEqual(rebuilt, Data(xml.utf8), "rebuilt must be byte-equal")
    }

    /// A marker-free paragraph still takes the plain path (no regression:
    /// setParagraphContent only appears when markers are present).
    func testMarkerFreeParagraphKeepsPlainForm() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="P1"><w:r><w:t>plain</w:t></w:r></w:p></w:body></w:document>
        """
        let parts = ["word/document.xml": Data(xml.utf8)]
        let result = try ReverseExtractor.reverse(parts: parts)
        XCTAssertTrue(result.dslParts.contains("word/document.xml"))
        let hasContentOp = result.log.entries.contains {
            if case .setParagraphContent = $0.op { return true }; return false
        }
        XCTAssertFalse(hasContentOp, "marker-free paragraph must not emit setParagraphContent")
    }
}
