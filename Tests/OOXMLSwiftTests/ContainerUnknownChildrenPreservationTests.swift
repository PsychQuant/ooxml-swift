import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// word-aligned-state-sync Phase 1 task 2.7 — `docx-container-parsing`
/// Requirements: "Reader does not silently drop element classes" and
/// "Container parts preserve unknown children identically to body".
///
/// Pins the preservation contract for element classes outside the typed
/// model (`<mc:AlternateContent>`, `<w:pict>` VML, `w16cid:*` extensions):
/// 1. `DocxReader.read` parses parts containing them without throwing, and
///    the elements are present in the part's `XmlTree`.
/// 2. Round-trip with no mutations is byte-equal on every such part.
/// 3. A body mutation (making the save real) leaves *untouched* container
///    parts (header, footnotes, endnotes, comments) byte-equal.
final class ContainerUnknownChildrenPreservationTests: XCTestCase {

    // MARK: - Fixture

    private static let headerXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:pict w14:anchorId="5A1B2C3D" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" path="m@7,l@8,m@5,21600l@6,21600e"><v:formulas><v:f eqn="sum #0 0 10800"/></v:formulas></v:shapetype><v:shape id="PowerPlusWaterMarkObject" o:spid="_x0000_s2049" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:494.9pt;height:164.95pt;rotation:315;z-index:-251656192" o:allowincell="f" fillcolor="silver" stroked="f"><v:fill opacity=".5"/><v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="DRAFT"/></v:shape></w:pict></w:r></w:p></w:hdr>
        """

    private static let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" mc:Ignorable="wps"><w:body><w:p><w:r><w:t>Before shapes</w:t></w:r></w:p><w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" behindDoc="0"><wp:extent cx="914400" cy="914400"/></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:rect id="fallback-rect" style="width:72pt;height:72pt"/></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHdr1"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>
        """

    private static let footnotesXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w16cid:commentsExtensible w16cid:durableId="123ABC"/><w:r><w:t>Footnote with extension child</w:t></w:r></w:p></w:footnote></w:footnotes>
        """

    private static let endnotesXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"><w:endnote w:id="1"><w:p><w16cid:commentsExtensible w16cid:durableId="DEF456"/><w:r><w:t>Endnote with extension child</w:t></w:r></w:p></w:endnote></w:endnotes>
        """

    private func buildFixture() throws -> URL {
        let staging = FileManager.default.temporaryDirectory
            .appendingPathComponent("container-unknown-staging-\(UUID().uuidString)")
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
                <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
                <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
            </Types>
            """, to: "[Content_Types].xml")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """, to: "_rels/.rels")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rIdHdr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                <Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
                <Relationship Id="rIdEn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
            </Relationships>
            """, to: "word/_rels/document.xml.rels")
        try write(Self.documentXML, to: "word/document.xml")
        try write(Self.headerXML, to: "word/header1.xml")
        try write(Self.footnotesXML, to: "word/footnotes.xml")
        try write(Self.endnotesXML, to: "word/endnotes.xml")

        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("container-unknown-\(UUID().uuidString).docx")
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
        return docxURL
    }

    private func extractPart(_ partPath: String, from docxURL: URL) throws -> Data {
        let archive = try Archive(url: docxURL, accessMode: .read)
        guard let entry = archive[partPath] else {
            XCTFail("\(partPath) missing from \(docxURL.lastPathComponent)")
            return Data()
        }
        var data = Data()
        _ = try archive.extract(entry) { data.append($0) }
        return data
    }

    private func save(_ doc: WordDocument) throws -> URL {
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("container-unknown-out-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: outURL)
        return outURL
    }

    // MARK: - Reader does not silently drop element classes

    func testUnknownElementClassesAreInTheTree() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        let doc = try DocxReader.read(from: fixture)

        func containsDescendant(_ node: XmlNode, prefix: String?, localName: String) -> Bool {
            if node.kind == .element && node.localName == localName && node.prefix == prefix {
                return true
            }
            return node.children.contains { containsDescendant($0, prefix: prefix, localName: localName) }
        }

        guard let docRoot = doc.xmlTrees["word/document.xml"]?.root,
              let hdrRoot = doc.xmlTrees["word/header1.xml"]?.root,
              let fnRoot = doc.xmlTrees["word/footnotes.xml"]?.root else {
            return XCTFail("expected document/header/footnotes trees")
        }
        XCTAssertTrue(containsDescendant(docRoot, prefix: "mc", localName: "AlternateContent"),
                      "mc:AlternateContent must appear in the document tree")
        XCTAssertTrue(containsDescendant(hdrRoot, prefix: "w", localName: "pict"),
                      "header VML <w:pict> must appear in the header tree")
        XCTAssertTrue(containsDescendant(fnRoot, prefix: "w16cid", localName: "commentsExtensible"),
                      "w16cid extension child must appear in the footnotes tree")
    }

    // MARK: - Spec scenario: no-mutation round-trip is byte-equal everywhere

    func testNoOpRoundTripByteEqualOnAllParts() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        let doc = try DocxReader.read(from: fixture)
        let outURL = try save(doc)
        defer { try? FileManager.default.removeItem(at: outURL) }

        for part in ["word/document.xml", "word/header1.xml",
                     "word/footnotes.xml", "word/endnotes.xml"] {
            XCTAssertEqual(
                try extractPart(part, from: fixture),
                try extractPart(part, from: outURL),
                "\(part) must round-trip byte-equal with no mutations")
        }
    }

    // MARK: - Spec scenario: untouched containers stay byte-equal under a real save

    func testUntouchedContainersByteEqualUnderBodyMutation() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        doc.appendParagraph(Paragraph(text: "new body paragraph"))

        let outURL = try save(doc)
        defer { try? FileManager.default.removeItem(at: outURL) }

        // document.xml is touched (canonicalization allowed); the containers
        // are not and must remain byte-equal — including the VML watermark.
        for part in ["word/header1.xml", "word/footnotes.xml", "word/endnotes.xml"] {
            XCTAssertEqual(
                try extractPart(part, from: fixture),
                try extractPart(part, from: outURL),
                "untouched \(part) must remain byte-equal when only the body was mutated")
        }
        let outDoc = String(decoding: try extractPart("word/document.xml", from: outURL), as: UTF8.self)
        XCTAssertTrue(outDoc.contains("new body paragraph"),
                      "body mutation must be present in the saved document")
    }

    // MARK: - Typed model still reads normal content from these parts

    func testTypedModelStillParsesKnownContent() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        let doc = try DocxReader.read(from: fixture)

        XCTAssertFalse(doc.headers.isEmpty, "typed header parse must still work")
        XCTAssertTrue(doc.footnotes.footnotes.contains { $0.id == 1 },
                      "typed footnote parse must still surface footnote id 1")
        XCTAssertTrue(doc.endnotes.endnotes.contains { $0.id == 1 },
                      "typed endnote parse must still surface endnote id 1")
    }
}
