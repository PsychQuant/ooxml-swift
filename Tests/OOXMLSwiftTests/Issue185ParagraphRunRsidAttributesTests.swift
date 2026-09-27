import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Regression tests for [PsychQuant/ooxml-swift#185](https://github.com/PsychQuant/ooxml-swift/issues/185).
///
/// `w:rsidR`/`w:rsidRPr`/`w:rsidRDefault`/`w:rsidP` on `<w:p>` and
/// `w:rsidR`/`w:rsidRPr`/`w:rsidDel` on `<w:r>` are attributes on the
/// elements THEMSELVES (not child elements), exactly like the `<w:sectPr>`
/// rsid family #112 fixed. Nothing read the paragraph's or run's own
/// attributes for these, so any typed edit that regenerated
/// `word/document.xml` from the typed model dropped them from every
/// paragraph and run in the document. Fix mirrors #112 exactly:
/// `Paragraph.rsidAttributes` / `Run.rsidAttributes` carry them opaquely,
/// captured from the element's own attributes and re-emitted on the opening
/// tag in a fixed canonical order.
final class Issue185ParagraphRunRsidAttributesTests: XCTestCase {

    // MARK: - Section A: unit-level parse/emit round trip (no docx I/O)

    private func parseParagraphFragment(_ xml: String) throws -> Paragraph {
        let wrapper = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>\(xml)</w:body>
        </w:document>
        """
        let doc = try XMLDocument(xmlString: wrapper)
        guard let root = doc.rootElement(),
              let body = root.elements(forName: "w:body").first,
              let p = body.elements(forName: "w:p").first else {
            throw NSError(domain: "test", code: 0, userInfo: [NSLocalizedDescriptionKey: "no <w:p> found"])
        }
        return try DocxReader.parseParagraph(
            from: p, relationships: RelationshipsCollection(), styles: [], numbering: Numbering())
    }

    func testParseCapturesParagraphRsidAttributes() throws {
        let p = try parseParagraphFragment(
            "<w:p w:rsidR=\"00AB12CD\" w:rsidRDefault=\"00112233\" w:rsidP=\"00445566\" w:rsidRPr=\"00778899\">"
            + "<w:r><w:t>x</w:t></w:r></w:p>")
        XCTAssertEqual(p.rsidAttributes["rsidR"], "00AB12CD")
        XCTAssertEqual(p.rsidAttributes["rsidRDefault"], "00112233")
        XCTAssertEqual(p.rsidAttributes["rsidP"], "00445566")
        XCTAssertEqual(p.rsidAttributes["rsidRPr"], "00778899")
    }

    func testToXMLReemitsParagraphRsidAttributesInCanonicalOrder() throws {
        let p = try parseParagraphFragment(
            "<w:p w:rsidR=\"00AB12CD\" w:rsidRDefault=\"00112233\">"
            + "<w:r><w:t>x</w:t></w:r></w:p>")
        let xml = p.toXML()
        XCTAssertTrue(xml.hasPrefix("<w:p w:rsidR=\"00AB12CD\" w:rsidRDefault=\"00112233\">"), "got: \(xml)")
    }

    func testNoParagraphRsidAttributesEmitsPlainOpeningTag() throws {
        let p = try parseParagraphFragment("<w:p><w:r><w:t>x</w:t></w:r></w:p>")
        XCTAssertTrue(p.rsidAttributes.isEmpty)
        XCTAssertTrue(p.toXML().hasPrefix("<w:p>"), "got: \(p.toXML())")
    }

    func testParseCapturesRunRsidAttributes() throws {
        let p = try parseParagraphFragment(
            "<w:p><w:r w:rsidR=\"00AB12CD\" w:rsidRPr=\"00112233\"><w:t>x</w:t></w:r></w:p>")
        let run = try XCTUnwrap(p.runs.first)
        XCTAssertEqual(run.rsidAttributes["rsidR"], "00AB12CD")
        XCTAssertEqual(run.rsidAttributes["rsidRPr"], "00112233")
    }

    func testToXMLReemitsRunRsidAttributesInCanonicalOrder() throws {
        let p = try parseParagraphFragment(
            "<w:p><w:r w:rsidR=\"00AB12CD\" w:rsidRPr=\"00112233\"><w:t>x</w:t></w:r></w:p>")
        let run = try XCTUnwrap(p.runs.first)
        let xml = run.toXML()
        XCTAssertTrue(xml.hasPrefix("<w:r w:rsidRPr=\"00112233\" w:rsidR=\"00AB12CD\">"), "got: \(xml)")
    }

    // MARK: - Section B: full-docx regression via the issue's exact repro

    /// Two paragraphs, each with rsid attributes on the `<w:p>` AND on their
    /// `<w:r>` child. The unrelated edit targets paragraph 0; paragraph 1 and
    /// paragraph 0's OWN attributes (edited via `replaceText`, which keeps
    /// the paragraph node but touches run text) must still keep every rsid.
    private static let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
        <w:p w:rsidR="00905AA1" w:rsidRDefault="00905AA1" w:rsidP="00905AA1" w:rsidRPr="00905AA1">\
        <w:r w:rsidR="00905AA1" w:rsidRPr="00905AA1"><w:t>Intro</w:t></w:r></w:p>\
        <w:p w:rsidR="00AB12CD" w:rsidRDefault="00AB12CD">\
        <w:r w:rsidR="00AB12CD"><w:t>Outro</w:t></w:r></w:p>\
        </w:body></w:document>
        """

    private func buildFixture() throws -> URL {
        let staging = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue185-staging-\(UUID().uuidString)")
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
            .appendingPathComponent("issue185-\(UUID().uuidString).docx")
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

    func testTypedEditThenSavePreservesRsidAttributesOnUntouchedParagraphAndRun() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        // Unrelated edit: touches paragraph 0's run text only.
        _ = try doc.replaceText(find: "Intro", with: "Intro2")
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))

        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue185-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(doc, to: outURL)

        let outputXML = try extractDocumentXML(from: outURL)
        // Paragraph 1 (never touched) must keep all four rsid attributes.
        for needle in ["w:rsidR=\"00AB12CD\"", "w:rsidRDefault=\"00AB12CD\""] {
            XCTAssertTrue(outputXML.contains(needle), "got: \(outputXML)")
        }
        // Paragraph 0's own <w:p> rsid attributes must survive despite its
        // run's TEXT being edited.
        for needle in ["w:rsidR=\"00905AA1\"", "w:rsidRDefault=\"00905AA1\"",
                       "w:rsidP=\"00905AA1\"", "w:rsidRPr=\"00905AA1\""] {
            XCTAssertTrue(outputXML.contains(needle), "got: \(outputXML)")
        }
    }
}
