import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Regression tests for [PsychQuant/ooxml-swift#191](https://github.com/PsychQuant/ooxml-swift/issues/191).
///
/// `SectionProperties` had no field for page borders (`<w:pgBorders>`) at
/// all, and `<w:cols>` only carried column count + gutter — `w:equalWidth`,
/// `w:sep`, and individual `<w:col>` widths had no typed representation.
/// Reading a document with either of these, then making any unrelated typed
/// edit (which regenerates `word/document.xml` from the typed model), lost
/// them silently.
final class Issue191PageBordersAndColumnsTests: XCTestCase {

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

    func testParseCapturesPageBorders() throws {
        let props = try parseSectPrFragment(
            "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/>"
            + "<w:pgBorders w:offsetFrom=\"page\" w:display=\"firstPage\" w:zOrder=\"front\">"
            + "<w:top w:val=\"single\" w:sz=\"24\" w:space=\"24\" w:color=\"FF0000\"/>"
            + "<w:left w:val=\"single\" w:sz=\"24\" w:space=\"24\" w:color=\"FF0000\"/>"
            + "<w:bottom w:val=\"single\" w:sz=\"24\" w:space=\"24\" w:color=\"FF0000\"/>"
            + "<w:right w:val=\"single\" w:sz=\"24\" w:space=\"24\" w:color=\"FF0000\"/>"
            + "</w:pgBorders></w:sectPr>")
        let pb = try XCTUnwrap(props.pageBorders)
        XCTAssertEqual(pb.offsetFrom, "page")
        XCTAssertEqual(pb.display, "firstPage")
        XCTAssertEqual(pb.zOrder, "front")
        let top = try XCTUnwrap(pb.top)
        XCTAssertEqual(top.style, "single")
        XCTAssertEqual(top.size, 24)
        XCTAssertEqual(top.space, 24)
        XCTAssertEqual(top.color, "FF0000")
        XCTAssertNotNil(pb.left)
        XCTAssertNotNil(pb.bottom)
        XCTAssertNotNil(pb.right)
    }

    func testToXMLReemitsPageBordersBetweenPgMarAndPgNumType() throws {
        var props = try parseSectPrFragment(
            "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/>"
            + "<w:pgBorders w:offsetFrom=\"page\">"
            + "<w:top w:val=\"single\" w:sz=\"24\" w:space=\"24\" w:color=\"FF0000\"/>"
            + "</w:pgBorders></w:sectPr>")
        props.pageNumberFormat = .decimal
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:pgBorders w:offsetFrom=\"page\">"), "got: \(xml)")
        XCTAssertTrue(xml.contains("<w:top w:val=\"single\" w:sz=\"24\" w:space=\"24\" w:color=\"FF0000\"/>"), "got: \(xml)")
        let pgMarIdx = xml.range(of: "<w:pgSz")!.upperBound
        let pgBordersIdx = try XCTUnwrap(xml.range(of: "<w:pgBorders")).lowerBound
        let pgNumIdx = try XCTUnwrap(xml.range(of: "<w:pgNumType")).lowerBound
        XCTAssertLessThan(pgMarIdx, pgBordersIdx, "got: \(xml)")
        XCTAssertLessThan(pgBordersIdx, pgNumIdx, "got: \(xml)")
    }

    func testNoPageBordersEmitsUnchangedSectPr() throws {
        let props = try parseSectPrFragment("<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/></w:sectPr>")
        XCTAssertNil(props.pageBorders)
        XCTAssertFalse(props.toXML().contains("pgBorders"))
    }

    func testParseCapturesUnequalColumnWidths() throws {
        let props = try parseSectPrFragment(
            "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/>"
            + "<w:cols w:num=\"2\" w:space=\"720\" w:equalWidth=\"0\" w:sep=\"1\">"
            + "<w:col w:w=\"3600\" w:space=\"720\"/><w:col w:w=\"4680\"/>"
            + "</w:cols></w:sectPr>")
        XCTAssertFalse(props.columnsEqualWidth)
        XCTAssertTrue(props.columnSeparator)
        XCTAssertEqual(props.individualColumnWidths.count, 2)
        XCTAssertEqual(props.individualColumnWidths[0].width, 3600)
        XCTAssertEqual(props.individualColumnWidths[0].spacing, 720)
        XCTAssertEqual(props.individualColumnWidths[1].width, 4680)
        XCTAssertNil(props.individualColumnWidths[1].spacing)
    }

    func testToXMLReemitsUnequalColumnWidths() throws {
        let props = try parseSectPrFragment(
            "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/>"
            + "<w:cols w:num=\"2\" w:space=\"720\" w:equalWidth=\"0\" w:sep=\"1\">"
            + "<w:col w:w=\"3600\" w:space=\"720\"/><w:col w:w=\"4680\"/>"
            + "</w:cols></w:sectPr>")
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:cols w:num=\"2\" w:space=\"720\" w:equalWidth=\"0\" w:sep=\"1\">"), "got: \(xml)")
        XCTAssertTrue(xml.contains("<w:col w:w=\"3600\" w:space=\"720\"/>"), "got: \(xml)")
        XCTAssertTrue(xml.contains("<w:col w:w=\"4680\"/>"), "got: \(xml)")
        XCTAssertTrue(xml.contains("</w:cols>"), "got: \(xml)")
    }

    func testDefaultEqualWidthColumnsEmitUnchangedColsTag() throws {
        let props = try parseSectPrFragment(
            "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:cols w:num=\"2\" w:space=\"720\"/></w:sectPr>")
        XCTAssertTrue(props.columnsEqualWidth)
        XCTAssertFalse(props.columnSeparator)
        XCTAssertTrue(props.individualColumnWidths.isEmpty)
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:cols w:num=\"2\" w:space=\"720\"/>"), "got: \(xml)")
    }

    // MARK: - Section B: full-docx regression

    private static let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
        <w:p><w:r><w:t>Intro</w:t></w:r></w:p>\
        <w:sectPr w:rsidR="00AB12CD">\
        <w:pgSz w:w="11906" w:h="16838"/>\
        <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>\
        <w:pgBorders w:offsetFrom="page">\
        <w:top w:val="single" w:sz="24" w:space="24" w:color="FF0000"/>\
        <w:left w:val="single" w:sz="24" w:space="24" w:color="FF0000"/>\
        <w:bottom w:val="single" w:sz="24" w:space="24" w:color="FF0000"/>\
        <w:right w:val="single" w:sz="24" w:space="24" w:color="FF0000"/>\
        </w:pgBorders>\
        <w:cols w:num="2" w:space="720" w:equalWidth="0"><w:col w:w="3600" w:space="720"/><w:col w:w="4680"/></w:cols>\
        <w:docGrid w:linePitch="360"/>\
        </w:sectPr>\
        </w:body></w:document>
        """

    private func buildFixture() throws -> URL {
        let staging = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue191-staging-\(UUID().uuidString)")
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
            .appendingPathComponent("issue191-\(UUID().uuidString).docx")
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

    func testTypedEditThenSavePreservesPageBordersAndColumnWidths() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        // Unrelated edit: touches the paragraph, not the section properties.
        try doc.updateParagraph(at: 0, text: "INTRO2")
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))

        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue191-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(doc, to: outURL)

        let outputXML = try extractDocumentXML(from: outURL)
        for needle in ["<w:pgBorders w:offsetFrom=\"page\">",
                       "<w:top w:val=\"single\" w:sz=\"24\" w:space=\"24\" w:color=\"FF0000\"/>",
                       "<w:left w:val=\"single\" w:sz=\"24\" w:space=\"24\" w:color=\"FF0000\"/>",
                       "<w:bottom w:val=\"single\" w:sz=\"24\" w:space=\"24\" w:color=\"FF0000\"/>",
                       "<w:right w:val=\"single\" w:sz=\"24\" w:space=\"24\" w:color=\"FF0000\"/>",
                       "w:equalWidth=\"0\"",
                       "<w:col w:w=\"3600\" w:space=\"720\"/>",
                       "<w:col w:w=\"4680\"/>"] {
            XCTAssertTrue(outputXML.contains(needle),
                          "expected \(needle) to survive an unrelated typed edit — got: \(outputXML)")
        }
    }
}
