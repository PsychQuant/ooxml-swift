import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Regression tests for [PsychQuant/ooxml-swift#112](https://github.com/PsychQuant/ooxml-swift/issues/112).
///
/// `<w:sectPr>` revision-save ids (`w:rsidR`, `w:rsidRPr`, `w:rsidSect`, …)
/// are attributes on `<w:sectPr>` ITSELF, not child elements — the reader
/// never looked at the element's own attributes, only at its children, so
/// they were dropped every time a section was re-serialized from the typed
/// model. No visual effect, but Word's document-compare/merge uses them to
/// attribute edits across sessions.
///
/// After this fix: `SectionProperties.rsidAttributes` carries them opaquely
/// (attribute local name -> value), `DocxReader.parseSectionProperties`
/// captures them from the element's own attributes, and
/// `SectionProperties.toXML()` / `toSectionBreakXML()` re-emit them on the
/// `<w:sectPr>` opening tag in a fixed canonical order.
final class Issue112SectPrRsidAttributesTests: XCTestCase {

    // MARK: - Section A: unit-level parse/emit round trip (no docx I/O)

    private func parseSectPrFragment(_ xml: String) throws -> SectionProperties {
        let wrapper = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          \(xml)
        </w:document>
        """
        let doc = try XMLDocument(xmlString: wrapper)
        guard let root = doc.rootElement(),
              let sectPrElement = root.elements(forName: "w:sectPr").first else {
            throw NSError(domain: "test", code: 0, userInfo: [NSLocalizedDescriptionKey: "no <w:sectPr> found"])
        }
        return DocxReader.parseSectionProperties(sectPrElement)
    }

    func testParseCapturesRsidAttributes() throws {
        let props = try parseSectPrFragment(
            "<w:sectPr w:rsidR=\"00AB12CD\" w:rsidRPr=\"00112233\" w:rsidSect=\"00EF3456\"><w:pgSz w:w=\"11906\" w:h=\"16838\"/></w:sectPr>")
        XCTAssertEqual(props.rsidAttributes["rsidR"], "00AB12CD")
        XCTAssertEqual(props.rsidAttributes["rsidRPr"], "00112233")
        XCTAssertEqual(props.rsidAttributes["rsidSect"], "00EF3456")
    }

    func testToXMLReemitsRsidAttributesInCanonicalOrder() throws {
        var props = try parseSectPrFragment(
            "<w:sectPr w:rsidR=\"00AB12CD\" w:rsidSect=\"00EF3456\"><w:pgSz w:w=\"11906\" w:h=\"16838\"/></w:sectPr>")
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("w:rsidR=\"00AB12CD\""), "got: \(xml)")
        XCTAssertTrue(xml.contains("w:rsidSect=\"00EF3456\""), "got: \(xml)")
        let rIndex = try XCTUnwrap(xml.range(of: "w:rsidR=\""))
        let sectIndex = try XCTUnwrap(xml.range(of: "w:rsidSect=\""))
        XCTAssertLessThan(rIndex.lowerBound, sectIndex.lowerBound,
                          "canonical order is rsidRPr, rsidDel, rsidR, rsidSect")
        // The opening tag itself carries the attributes (not some other element).
        XCTAssertTrue(xml.hasPrefix("<w:sectPr w:rsidR=\"00AB12CD\" w:rsidSect=\"00EF3456\">"), "got: \(xml)")
        _ = props // silence "never mutated" warning path if any
    }

    /// A section with no rsid attributes must emit `<w:sectPr>` exactly as
    /// before this field existed — no regression to the common case.
    func testNoRsidAttributesEmitsPlainOpeningTag() throws {
        let props = try parseSectPrFragment("<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/></w:sectPr>")
        XCTAssertTrue(props.rsidAttributes.isEmpty)
        XCTAssertTrue(props.toXML().hasPrefix("<w:sectPr>"), "got: \(props.toXML())")
    }

    func testParseThenEmitRoundTripsRsidAttributesVerbatim() throws {
        let props = try parseSectPrFragment(
            "<w:sectPr w:rsidR=\"00AB12CD\" w:rsidRPr=\"00112233\" w:rsidSect=\"00EF3456\"><w:pgSz w:w=\"11906\" w:h=\"16838\"/></w:sectPr>")
        let reparsed = try parseSectPrFragment(props.toXML())
        XCTAssertEqual(reparsed.rsidAttributes, props.rsidAttributes)
    }

    // MARK: - Section B: full-docx regression via the issue's exact repro

    /// Body: one plain paragraph (the typed edit's target) + body-level
    /// `<w:sectPr>` carrying the two rsid attributes from the issue body.
    private static let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
        <w:p><w:r><w:t>Intro</w:t></w:r></w:p>\
        <w:sectPr w:rsidR="00AB12CD" w:rsidSect="00EF3456">\
        <w:pgSz w:w="11906" w:h="16838"/>\
        </w:sectPr>\
        </w:body></w:document>
        """

    private func buildFixture() throws -> URL {
        let staging = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue112-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: staging, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: staging) }

        func write(_ content: String, to relativePath: String) throws {
            let url = staging.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(), withIntermediateDirectories: true)
            try content.write(to: url, atomically: true, encoding: .utf8)
        }

        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="xml" ContentType="application/xml"/>
                <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>
            """, to: "[Content_Types].xml")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """, to: "_rels/.rels")
        try write(Self.documentXML, to: "word/document.xml")

        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue112-\(UUID().uuidString).docx")
        do {
            let archive = try Archive(url: docxURL, accessMode: .create)
            let base = staging.resolvingSymlinksInPath().path
            let enumerator = FileManager.default.enumerator(
                at: staging, includingPropertiesForKeys: [.isDirectoryKey])!
            for case let fileURL as URL in enumerator {
                let isDir = (try? fileURL.resourceValues(forKeys: [.isDirectoryKey]))?.isDirectory ?? false
                if isDir { continue }
                let entry = String(fileURL.resolvingSymlinksInPath().path.dropFirst(base.count + 1))
                try archive.addEntry(with: entry, fileURL: fileURL, compressionMethod: .deflate)
            }
        } catch {
            try? FileManager.default.removeItem(at: docxURL)
            throw error
        }
        return docxURL
    }

    private func extractDocumentXML(from docxURL: URL) throws -> String {
        let archive = try Archive(url: docxURL, accessMode: .read)
        let entry = try XCTUnwrap(archive["word/document.xml"])
        var data = Data()
        _ = try archive.extract(entry) { data.append($0) }
        return String(data: data, encoding: .utf8)!
    }

    /// The issue's exact repro: two `update_cell`-equivalent typed edits
    /// (here `updateParagraph` + `replaceText`, both of which mark
    /// `word/document.xml` dirty), then save.
    func testTypedEditThenSavePreservesSectPrRsidAttributes() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        try doc.updateParagraph(at: 0, text: "INTRO")
        _ = try doc.replaceText(find: "INTRO", with: "INTRO2")
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))

        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue112-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(doc, to: outURL)

        let outputXML = try extractDocumentXML(from: outURL)
        XCTAssertTrue(outputXML.contains("w:rsidR=\"00AB12CD\""),
                      "before: <w:sectPr w:rsidR=\"00AB12CD\" ...>; after must still carry it — got: \(outputXML)")
        XCTAssertTrue(outputXML.contains("w:rsidSect=\"00EF3456\""), "got: \(outputXML)")
    }
}
