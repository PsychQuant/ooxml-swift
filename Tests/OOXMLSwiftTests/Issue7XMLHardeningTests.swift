import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// XML input hardening tests for [PsychQuant/ooxml-swift#7](https://github.com/PsychQuant/ooxml-swift/issues/7).
///
/// Covers four sub-findings from the post-#56 verification bundle:
/// - F10 — DTD pre-scan reject (`XMLHardeningError.dtdNotAllowed`)
/// - F11 — Foundation `XMLParser` SAX root-attr parser (handles arbitrary prefix variants)
/// - F12 — Attribute-name whitelist (`XMLHardeningError.invalidAttributeName`)
/// - F14 — 64 KB attribute-value byte cap (`XMLHardeningError.attributeValueTooLarge`)
///
/// Each test maps to one `Scenario:` block in the spec
/// `openspec/changes/harden-xml-security/specs/ooxml-input-hardening/spec.md`.
final class Issue7XMLHardeningTests: XCTestCase {

    // MARK: - F10: DTD pre-scan reject

    /// Spec: "Document part with DOCTYPE declaration is rejected"
    func testReadRejectsDocxWithDOCTYPEInDocumentXML() throws {
        let url = try buildHardeningFixture(documentXMLOverride: """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <!DOCTYPE w:document SYSTEM "external.dtd">
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body>
        </w:document>
        """)
        defer { try? FileManager.default.removeItem(at: url) }

        XCTAssertThrowsError(try DocxReader.read(from: url)) { error in
            guard case let XMLHardeningError.dtdNotAllowed(part) = error else {
                XCTFail("Expected XMLHardeningError.dtdNotAllowed, got \(error)")
                return
            }
            XCTAssertEqual(part, "word/document.xml")
        }
    }

    /// Spec: "All container parts share the DTD reject guard"
    /// Header variant + lowercase variant in one test (saves a second fixture round-trip).
    func testReadRejectsDocxWithDOCTYPEInHeaderXML() throws {
        let url = try buildHardeningFixture(headerXMLOverride: """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <!DOCTYPE w:hdr SYSTEM "external.dtd">
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:p><w:r><w:t>hi</w:t></w:r></w:p>
        </w:hdr>
        """)
        defer { try? FileManager.default.removeItem(at: url) }

        XCTAssertThrowsError(try DocxReader.read(from: url)) { error in
            guard case let XMLHardeningError.dtdNotAllowed(part) = error else {
                XCTFail("Expected XMLHardeningError.dtdNotAllowed, got \(error)")
                return
            }
            XCTAssertEqual(part, "word/header1.xml")
        }
    }

    /// Spec example table: case-insensitive DOCTYPE detection.
    func testReadRejectsDocxWithLowercaseDoctypeVariant() throws {
        let url = try buildHardeningFixture(documentXMLOverride: """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <!doctype w:document>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body>
        </w:document>
        """)
        defer { try? FileManager.default.removeItem(at: url) }

        XCTAssertThrowsError(try DocxReader.read(from: url)) { error in
            guard case .dtdNotAllowed = error as? XMLHardeningError else {
                XCTFail("Expected XMLHardeningError.dtdNotAllowed, got \(error)")
                return
            }
        }
    }

    // MARK: - F11: XMLParser SAX root-attr parser

    /// Spec: "Custom-prefix root element is parsed correctly"
    func testParseRootAttrsHandlesCustomPrefix() throws {
        let xml = #"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <wordml:document xmlns:wordml="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <wordml:body/>
        </wordml:document>
        """#
        let data = xml.data(using: .utf8)!
        let attrs = try DocxReader.parseContainerRootAttributes(from: data)
        XCTAssertEqual(attrs["xmlns:wordml"], "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        XCTAssertEqual(attrs["xmlns:r"], "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    }

    /// Spec: "Default-namespace root element is parsed correctly"
    func testParseRootAttrsHandlesDefaultNamespace() throws {
        let xml = #"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <document xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <body/>
        </document>
        """#
        let data = xml.data(using: .utf8)!
        let attrs = try DocxReader.parseContainerRootAttributes(from: data)
        XCTAssertEqual(attrs["xmlns"], "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    }

    /// Spec: "Truly malformed XML returns empty map"
    func testParseRootAttrsReturnsEmptyOnMalformedXML() throws {
        let xml = "<w:document xmlns:w=\"unterminated"
        let data = xml.data(using: .utf8)!
        let attrs = try DocxReader.parseContainerRootAttributes(from: data)
        XCTAssertTrue(attrs.isEmpty, "Malformed XML should return [:] (caller fallback path), got \(attrs)")
    }

    // MARK: - F12: Attribute-name whitelist

    /// Spec: "Attribute name with invalid leading character is rejected on ingest"
    func testSplitAttributesRejectsLeadingDigitInName() throws {
        let attrSlice = #" 0xmlns:w="http://example/""#
        XCTAssertThrowsError(try DocxReader.splitAttributes(attrSlice)) { error in
            guard case let XMLHardeningError.invalidAttributeName(name, context) = error else {
                XCTFail("Expected XMLHardeningError.invalidAttributeName, got \(error)")
                return
            }
            XCTAssertEqual(name, "0xmlns:w")
            XCTAssertEqual(context, "split-attributes")
        }
    }

    /// Spec: "Spec-compliant names pass through unchanged"
    func testSplitAttributesAcceptsConformantNames() throws {
        let attrSlice = #" xmlns:w="A" mc:Ignorable="B" xml:space="C""#
        let attrs = try DocxReader.splitAttributes(attrSlice)
        XCTAssertEqual(attrs["xmlns:w"], "A")
        XCTAssertEqual(attrs["mc:Ignorable"], "B")
        XCTAssertEqual(attrs["xml:space"], "C")
    }

    /// Spec: "Attribute name with embedded whitespace is rejected on emit"
    func testRenderDocumentRootOpenTagRejectsWhitespaceInName() throws {
        XCTAssertThrowsError(try DocxWriter.renderDocumentRootOpenTag(["xmlns w": "http://example/"])) { error in
            guard case let XMLHardeningError.invalidAttributeName(name, context) = error else {
                XCTFail("Expected XMLHardeningError.invalidAttributeName, got \(error)")
                return
            }
            XCTAssertEqual(name, "xmlns w")
            XCTAssertEqual(context, "document root")
        }
    }

    // MARK: - F14: 64 KB attribute-value byte cap

    /// Spec: "Oversized attribute value is rejected"
    func testSplitAttributesRejectsValueOver64KB() throws {
        let big = String(repeating: "a", count: 100_000)
        let attrSlice = #" mc:Ignorable="\#(big)""#
        XCTAssertThrowsError(try DocxReader.splitAttributes(attrSlice)) { error in
            guard case let XMLHardeningError.attributeValueTooLarge(name, byteSize, cap) = error else {
                XCTFail("Expected XMLHardeningError.attributeValueTooLarge, got \(error)")
                return
            }
            XCTAssertEqual(name, "mc:Ignorable")
            XCTAssertEqual(byteSize, 100_000)
            XCTAssertEqual(cap, 65_536)
        }
    }

    /// Spec example table: cap boundary at 65 535 / 65 536 / 65 537 bytes.
    func testSplitAttributesAcceptsValueAt64KBBoundary() throws {
        // 65 536 bytes — exactly at cap, must accept
        let atCap = String(repeating: "a", count: 65_536)
        let atCapSlice = #" foo="\#(atCap)""#
        let attrs = try DocxReader.splitAttributes(atCapSlice)
        XCTAssertEqual(attrs["foo"]?.count, 65_536)

        // 65 537 bytes — one over cap, must throw
        let overCap = String(repeating: "a", count: 65_537)
        let overCapSlice = #" foo="\#(overCap)""#
        XCTAssertThrowsError(try DocxReader.splitAttributes(overCapSlice)) { error in
            guard case let XMLHardeningError.attributeValueTooLarge(_, byteSize, cap) = error else {
                XCTFail("Expected XMLHardeningError.attributeValueTooLarge, got \(error)")
                return
            }
            XCTAssertEqual(byteSize, 65_537)
            XCTAssertEqual(cap, 65_536)
        }
    }

    // MARK: - Fixture builder

    /// Builds a minimal valid `.docx` with overridable `word/document.xml`
    /// and `word/header1.xml` content. Pattern lifted from
    /// `Issue56FollowupTests.buildRevisionFixture`.
    private func buildHardeningFixture(
        documentXMLOverride: String? = nil,
        headerXMLOverride: String? = nil
    ) throws -> URL {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("hardening-fixture-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: stagingDir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: stagingDir) }

        func write(_ s: String, to rel: String) throws {
            let url = stagingDir.appendingPathComponent(rel)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(),
                withIntermediateDirectories: true
            )
            try s.write(to: url, atomically: true, encoding: .utf8)
        }

        try write(#"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
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
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
        </Relationships>
        """#, to: "word/_rels/document.xml.rels")

        let documentXML = documentXMLOverride ?? #"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p><w:r><w:t>hi</w:t></w:r></w:p>
            <w:sectPr>
              <w:headerReference r:id="rId10" w:type="default" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
            </w:sectPr>
          </w:body>
        </w:document>
        """#
        try write(documentXML, to: "word/document.xml")

        let headerXML = headerXMLOverride ?? #"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:p><w:r><w:t>header</w:t></w:r></w:p>
        </w:hdr>
        """#
        try write(headerXML, to: "word/header1.xml")

        let outputURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("hardening-fixture-\(UUID().uuidString).docx")
        let archive = try Archive(url: outputURL, accessMode: .create)
        let normalizedStaging = stagingDir.resolvingSymlinksInPath().path
        let basePathLen = normalizedStaging.count + 1
        let enumerator = FileManager.default.enumerator(at: stagingDir, includingPropertiesForKeys: [.isDirectoryKey])!
        for case let fileURL as URL in enumerator {
            let isDir = (try? fileURL.resourceValues(forKeys: [.isDirectoryKey]))?.isDirectory ?? false
            if isDir { continue }
            let normalizedFile = fileURL.resolvingSymlinksInPath().path
            let entryName = String(normalizedFile.dropFirst(basePathLen))
            try archive.addEntry(with: entryName, fileURL: fileURL, compressionMethod: .deflate)
        }
        return outputURL
    }
}
