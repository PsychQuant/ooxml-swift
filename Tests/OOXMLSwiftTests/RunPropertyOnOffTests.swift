import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

final class RunPropertyOnOffTests: XCTestCase {
    func testExplicitFalseBoldEmitsSemanticOff() {
        var properties = RunProperties()
        properties.bold = false

        let xml = Run(text: "plain", properties: properties).toXML()

        XCTAssertTrue(xml.contains(#"<w:b w:val="0"/>"#), "Got: \(xml)")
        XCTAssertFalse(xml.contains("<w:b/>"), "Got: \(xml)")
    }

    func testMergeDistinguishesOmittedFromExplicitFalse() {
        var base = RunProperties(bold: true, italic: true, underline: .single)
        base.merge(with: RunProperties())
        XCTAssertTrue(base.bold)
        XCTAssertTrue(base.italic)
        XCTAssertEqual(base.underline, .single)

        var patch = RunProperties()
        patch.bold = false
        patch.underline = nil
        base.merge(with: patch)

        XCTAssertFalse(base.bold)
        XCTAssertTrue(base.italic)
        XCTAssertNil(base.underline)
    }

    func testReaderAcceptsExplicitOffSpellingsForTypedBooleans() throws {
        let source = try buildFixture(documentXML: documentXML(
            firstRunProperties: #"<w:b w:val="0"/><w:i w:val="false"/><w:strike w:val="off"/><w:noProof w:val="0"/>"#
        ))
        defer { try? FileManager.default.removeItem(at: source) }

        var document = try DocxReader.read(from: source)
        defer { document.close() }
        let properties = try XCTUnwrap(document.getParagraphs().first?.runs.first?.properties)

        XCTAssertFalse(properties.bold)
        XCTAssertFalse(properties.italic)
        XCTAssertFalse(properties.strikethrough)
        XCTAssertFalse(properties.noProof)
    }

    func testReaderAcceptsOnSpellingsForTypedBooleans() throws {
        for value in [nil, "1", "true", "on", "no"] as [String?] {
            let attribute = value.map { " w:val=\"\($0)\"" } ?? ""
            let source = try buildFixture(documentXML: documentXML(
                firstRunProperties: "<w:b\(attribute)/><w:i\(attribute)/><w:strike\(attribute)/><w:noProof\(attribute)/>"
            ))
            defer { try? FileManager.default.removeItem(at: source) }

            var document = try DocxReader.read(from: source)
            let properties = try XCTUnwrap(document.getParagraphs().first?.runs.first?.properties)
            XCTAssertTrue(properties.bold, "value=\(String(describing: value))")
            XCTAssertTrue(properties.italic, "value=\(String(describing: value))")
            XCTAssertTrue(properties.strikethrough, "value=\(String(describing: value))")
            XCTAssertTrue(properties.noProof, "value=\(String(describing: value))")
            document.close()
        }
    }

    func testOMathOnlyFilteringPreservesAbsentAndExplicitFalse() {
        let absent = RunProperties().filteredForOMathSplice(mode: .omathOnly)
        let absentXML = Run(text: "x", properties: absent).toXML()
        XCTAssertFalse(absentXML.contains("<w:b"), "Got: \(absentXML)")
        XCTAssertFalse(absentXML.contains("<w:i"), "Got: \(absentXML)")

        var explicitOff = RunProperties()
        explicitOff.bold = false
        explicitOff.italic = false
        let off = explicitOff.filteredForOMathSplice(mode: .omathOnly)
        let offXML = Run(text: "x", properties: off).toXML()
        XCTAssertTrue(offXML.contains(#"<w:b w:val="0"/>"#), "Got: \(offXML)")
        XCTAssertTrue(offXML.contains(#"<w:i w:val="0"/>"#), "Got: \(offXML)")
    }

    func testUnrelatedMutationPreservesExplicitOffRunProperties() throws {
        let source = try buildFixture(documentXML: documentXML(
            firstRunProperties: #"<w:b w:val="0"/><w:i w:val="false"/><w:strike w:val="off"/><w:noProof w:val="0"/>"#
        ))
        defer { try? FileManager.default.removeItem(at: source) }
        let output = FileManager.default.temporaryDirectory
            .appendingPathComponent("run-on-off-output-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: output) }

        var document = try DocxReader.read(from: source)
        defer { document.close() }
        XCTAssertEqual(try document.replaceText(find: "gamma", with: "GAMMA"), 1)
        try DocxWriter.write(document, to: output)

        let extracted = try ZipHelper.unzip(output)
        defer { ZipHelper.cleanup(extracted) }
        let xml = try String(
            contentsOf: extracted.appendingPathComponent("word/document.xml"),
            encoding: .utf8
        )
        XCTAssertTrue(xml.contains(#"<w:b w:val="0"/>"#), "Got: \(xml)")
        XCTAssertTrue(xml.contains(#"<w:i w:val="0"/>"#), "Got: \(xml)")
        XCTAssertTrue(xml.contains(#"<w:strike w:val="0"/>"#), "Got: \(xml)")
        XCTAssertTrue(xml.contains(#"<w:noProof w:val="0"/>"#), "Got: \(xml)")
        XCTAssertFalse(xml.contains("<w:b/>"), "Got: \(xml)")
        XCTAssertTrue(xml.contains("GAMMA delta"), "Got: \(xml)")
    }

    func testRevisionPreviousFormatEmitsExplicitOffProperties() {
        var previous = RunProperties()
        previous.bold = false
        previous.italic = false
        previous.strikethrough = false
        previous.noProof = false
        previous.underline = nil

        let xml = previous.toChangeXML()

        XCTAssertTrue(xml.contains(#"<w:b w:val="0"/>"#), "Got: \(xml)")
        XCTAssertTrue(xml.contains(#"<w:i w:val="0"/>"#), "Got: \(xml)")
        XCTAssertTrue(xml.contains(#"<w:strike w:val="0"/>"#), "Got: \(xml)")
        XCTAssertTrue(xml.contains(#"<w:noProof w:val="0"/>"#), "Got: \(xml)")
        XCTAssertTrue(xml.contains(#"<w:u w:val="none"/>"#), "Got: \(xml)")
    }

    func testTreeBackedAppendPreservesExplicitOffRunProperties() throws {
        let source = try buildFixture(documentXML: documentXML(firstRunProperties: ""))
        defer { try? FileManager.default.removeItem(at: source) }
        let output = FileManager.default.temporaryDirectory
            .appendingPathComponent("run-on-off-append-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: output) }

        var document = try DocxReader.read(from: source)
        defer { document.close() }
        var properties = RunProperties()
        properties.bold = false
        properties.italic = false
        properties.strikethrough = false
        properties.noProof = false
        properties.underline = nil
        document.appendParagraph(Paragraph(runs: [Run(text: "appended", properties: properties)]))
        try DocxWriter.write(document, to: output)

        let extracted = try ZipHelper.unzip(output)
        defer { ZipHelper.cleanup(extracted) }
        let xml = try String(
            contentsOf: extracted.appendingPathComponent("word/document.xml"),
            encoding: .utf8
        )
        XCTAssertTrue(xml.contains(#"<w:b w:val="0"/>"#), "Got: \(xml)")
        XCTAssertTrue(xml.contains(#"<w:i w:val="0"/>"#), "Got: \(xml)")
        XCTAssertTrue(xml.contains(#"<w:strike w:val="0"/>"#), "Got: \(xml)")
        XCTAssertTrue(xml.contains(#"<w:noProof w:val="0"/>"#), "Got: \(xml)")
        XCTAssertTrue(xml.contains(#"<w:u w:val="none"/>"#), "Got: \(xml)")
        XCTAssertTrue(xml.contains("appended"), "Got: \(xml)")
    }

    func testAlternateWordprocessingMLPrefixParsesExplicitOff() throws {
        let xml = try XMLDocument(xmlString: #"""
        <x:r xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <x:rPr><x:b x:val="0"/><x:i x:val="off"/><x:strike x:val="false"/><x:noProof x:val="0"/></x:rPr>
          <x:t>text</x:t>
        </x:r>
        """#)
        let element = try XCTUnwrap(xml.rootElement())
        let run = try DocxReader.parseRun(
            from: element,
            relationships: RelationshipsCollection()
        )

        XCTAssertFalse(run.properties.bold)
        XCTAssertFalse(run.properties.italic)
        XCTAssertFalse(run.properties.strikethrough)
        XCTAssertFalse(run.properties.noProof)
    }

    private func documentXML(firstRunProperties: String) -> String {
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p><w:r><w:rPr>\(firstRunProperties)</w:rPr><w:t>alpha beta</w:t></w:r></w:p>
            <w:p><w:r><w:t>gamma delta</w:t></w:r></w:p>
          </w:body>
        </w:document>
        """
    }

    private func buildFixture(documentXML: String) throws -> URL {
        let staging = FileManager.default.temporaryDirectory
            .appendingPathComponent("run-on-off-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: staging, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: staging) }

        func write(_ content: String, to relativePath: String) throws {
            let url = staging.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(),
                withIntermediateDirectories: true
            )
            try content.write(to: url, atomically: true, encoding: .utf8)
        }

        try write(#"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>
        """#, to: "[Content_Types].xml")
        try write(#"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """#, to: "_rels/.rels")
        try write(#"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>
        """#, to: "word/_rels/document.xml.rels")
        try write(documentXML, to: "word/document.xml")

        let output = FileManager.default.temporaryDirectory
            .appendingPathComponent("run-on-off-fixture-\(UUID().uuidString).docx")
        let archive = try Archive(url: output, accessMode: .create)
        let base = staging.resolvingSymlinksInPath().path
        let prefixLength = base.count + 1
        let enumerator = FileManager.default.enumerator(
            at: staging,
            includingPropertiesForKeys: [.isDirectoryKey]
        )!
        for case let fileURL as URL in enumerator {
            if (try? fileURL.resourceValues(forKeys: [.isDirectoryKey]).isDirectory) == true {
                continue
            }
            let entry = String(fileURL.resolvingSymlinksInPath().path.dropFirst(prefixLength))
            try archive.addEntry(with: entry, fileURL: fileURL, compressionMethod: .deflate)
        }
        return output
    }
}
