// CarryPartOpTests.swift
// format-alignment-engine Phase A task 1.2 (stage 1) — the `carryPart` op that
// backs the all-parts raw channel (`ooxml-script-transcode` capability,
// «All-parts raw channel»; Decision 2). Verifies the op survives the `// @op`
// raw escape byte-exact, including arbitrary XML content (quotes, newlines,
// CJK) — the raw channel's whole promise is byte-exactness.

import XCTest
@testable import OOXMLSwift

final class CarryPartOpTests: XCTestCase {

    func testBinaryPartRoundTripsThroughRawOperation() throws {
        let bytes = Data([0x00, 0xFF, 0x89, 0x50, 0x4E, 0x47])
        var log = OperationLog()
        log.append(.carryBinaryPart(
            partPath: "word/media/image1.png",
            base64: bytes.base64EncodedString()), source: .word)

        let script = ScriptExporter.exportSwift(log: log)
        let reconstructed = try ScriptImporter.parse(source: script)
        XCTAssertEqual(reconstructed.entries.map(\.op), log.entries.map(\.op))

        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: reconstructed.entries.map(\.op))
        XCTAssertEqual(doc.carriedParts["word/media/image1.png"], bytes)
    }

    func testReverseExtractorCarriesBinaryPackageParts() throws {
        let binary = Data([0x89, 0x50, 0x4E, 0x47, 0x00, 0xFF])
        let parts: [String: Data] = [
            "word/document.xml": Data("<w:document xmlns:w=\"urn:w\"><w:body/></w:document>".utf8),
            "word/media/image1.png": binary,
        ]

        let result = try ReverseExtractor.reverse(parts: parts)
        XCTAssertTrue(result.log.entries.contains {
            if case .carryBinaryPart(let path, _) = $0.op {
                return path == "word/media/image1.png"
            }
            return false
        })

        var rebuilt = WordDocument.emptyAuthoringDocument()
        try rebuilt.apply(operations: result.log.entries.map(\.op))
        XCTAssertEqual(rebuilt.carriedParts["word/media/image1.png"], binary)
    }

    func testCarryPartRejectsParentTraversal() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        XCTAssertThrowsError(try doc.apply(operations: [
            .carryPart(partPath: "../../escaped.xml", xml: "<x/>")
        ]))
        XCTAssertThrowsError(try doc.apply(operations: [
            .carryBinaryPart(partPath: "../escaped.bin", base64: "AA==")
        ]))
    }

    func testAddRelationshipRejectsParentTraversal() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        XCTAssertThrowsError(try doc.apply(operations: [
            .addRelationship(
                part: "../../escaped.rels",
                id: "rIdEscape",
                type: "urn:test",
                target: "target.xml",
                targetMode: nil),
        ]))
        XCTAssertNil(doc.xmlTrees["../../escaped.rels"])
        XCTAssertTrue(doc.operationLog.entries.isEmpty)
    }

    func testCarryPartBatchIsAtomicWhenLaterOperationFails() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        let originalParts = doc.carriedParts

        XCTAssertThrowsError(try doc.apply(operations: [
            .carryPart(partPath: "word/custom.xml", xml: "<custom/>"),
            .setText(
                target: ElementID(rawString: "w14:paraId=DOES-NOT-EXIST"),
                text: "boom"),
        ]))

        XCTAssertEqual(doc.carriedParts, originalParts)
        XCTAssertTrue(doc.operationLog.entries.isEmpty)
    }

    /// A carryPart op round-trips through export → parse with its partPath and
    /// xml preserved field-for-field.
    func testCarryPartRoundTripsThroughScript() throws {
        var log = OperationLog()
        log.append(.carryPart(partPath: "word/styles.xml", xml: "<w:styles/>"), source: .word)

        let script = ScriptExporter.exportSwift(log: log)
        let reconstructed = try ScriptImporter.parse(source: script)

        XCTAssertEqual(reconstructed.entries.count, 1)
        guard case let .carryPart(partPath, xml) = reconstructed.entries[0].op else {
            return XCTFail("expected a carryPart op, got \(reconstructed.entries[0].op)")
        }
        XCTAssertEqual(partPath, "word/styles.xml")
        XCTAssertEqual(xml, "<w:styles/>")
    }

    /// The xml field carries arbitrary content byte-exact: double quotes,
    /// backslashes, newlines, and CJK all survive the JSON-escaped `// @op`
    /// wire. A styles.xml part in the wild contains all of these.
    func testCarryPartPreservesArbitraryXMLContent() throws {
        let gnarlyXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:styleId="Title"><w:name w:val="標題 \\ 見出し"/></w:style>
        </w:styles>
        """
        var log = OperationLog()
        log.append(.carryPart(partPath: "word/styles.xml", xml: gnarlyXML), source: .word)

        let script = ScriptExporter.exportSwift(log: log)
        let reconstructed = try ScriptImporter.parse(source: script)

        guard case let .carryPart(_, xml) = reconstructed.entries[0].op else {
            return XCTFail("expected a carryPart op")
        }
        XCTAssertEqual(xml, gnarlyXML, "raw channel must preserve XML byte-exact")
    }

    /// The op is not a DSL-authoring form, so it routes through the `// @op`
    /// raw escape (never a DSL block) — the exporter needs zero changes.
    func testCarryPartEmitsRawEscapeLine() {
        var log = OperationLog()
        log.append(.carryPart(partPath: "word/settings.xml", xml: "<w:settings/>"), source: .word)
        let script = ScriptExporter.exportSwift(log: log)
        XCTAssertTrue(script.contains("// @op "), "carryPart must ride the raw escape channel")
        XCTAssertTrue(script.contains("\"op_type\":\"carryPart\""),
                      "raw line must carry the carryPart discriminator")
        // JSONSerialization escapes `/` as `\/` on the wire; the raw line
        // carries the escaped form. Decode reverses it, so round-trip stays
        // byte-exact (verified by testSiblingPartRoundTripsVerbatim).
        XCTAssertTrue(script.contains(#""partPath":"word\/settings.xml""#))
    }

    /// spec «All-parts raw channel» scenario: a sibling part round-trips
    /// verbatim. GIVEN a styles.xml with 16 styles, docDefaults, and
    /// latentStyles; WHEN the reverse script executes (parse → apply → write);
    /// THEN the rebuilt styles.xml is byte-equal to the source.
    func testSiblingPartRoundTripsVerbatim() throws {
        let styleDefs = (1...16).map {
            "<w:style w:type=\"paragraph\" w:styleId=\"S\($0)\"><w:name w:val=\"樣式 \($0)\"/></w:style>"
        }.joined()
        let stylesXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults><w:latentStyles w:count="16"/>\(styleDefs)</w:styles>
        """

        // Build a rebuild script: a body paragraph (so document.xml exists) plus
        // styles.xml on the raw channel.
        var log = OperationLog()
        log.append(.appendParagraph(in: nil, paragraph: ParagraphPayload(text: "body", paraId: "P1")), source: .swift)
        log.append(.carryPart(partPath: "word/styles.xml", xml: stylesXML), source: .word)
        let script = ScriptExporter.exportSwift(log: log)

        // Execute: parse → apply → writeAuthoringPackage.
        let rebuiltLog = try ScriptImporter.parse(source: script)
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: rebuiltLog.entries.map(\.op))
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("carry-roundtrip-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try doc.writeAuthoringPackage(to: outURL)

        let rebuilt = try CorpusFixtureBuilder.readPart("word/styles.xml", from: outURL)
        XCTAssertEqual(rebuilt, Data(stylesXML.utf8),
                       "styles.xml must round-trip byte-equal via the raw channel")
    }
}
