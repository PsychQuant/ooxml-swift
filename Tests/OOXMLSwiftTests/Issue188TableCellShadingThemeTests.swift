import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Regression tests for [PsychQuant/ooxml-swift#188](https://github.com/PsychQuant/ooxml-swift/issues/188).
///
/// #179 taught `CellShading` (shared by `ParagraphShading` and table-cell
/// shading) to carry the six `<w:shd>` theme attributes, but only the
/// paragraph-shading read path (`DocxReader.parseParagraphShading`) was ever
/// wired to populate them — the table-cell read path
/// (`DocxReader.parseTableCellProperties`) still only read `fill`/`color`/
/// `pattern`, so a cell whose background used a theme color lost the theme
/// attributes on the very next typed edit anywhere in the document (whole
/// `word/document.xml` regenerated from the typed model).
final class Issue188TableCellShadingThemeTests: XCTestCase {

    // MARK: - Section A: unit-level parse/emit round trip (no docx I/O)

    private func parseTcPrFragment(_ shdXML: String) throws -> TableCellProperties {
        let wrapper = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:tc><w:tcPr>\(shdXML)</w:tcPr><w:p/></w:tc>
        </w:document>
        """
        let doc = try XMLDocument(xmlString: wrapper)
        guard let root = doc.rootElement(),
              let tc = root.elements(forName: "w:tc").first,
              let tcPr = tc.elements(forName: "w:tcPr").first else {
            throw NSError(domain: "test", code: 0, userInfo: [NSLocalizedDescriptionKey: "no <w:tcPr> found"])
        }
        return DocxReader.parseTableCellProperties(from: tcPr)
    }

    func testParseCapturesCellShadingThemeAttributes() throws {
        let props = try parseTcPrFragment(
            "<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"auto\" w:themeFill=\"accent2\" "
            + "w:themeFillTint=\"33\" w:themeFillShade=\"CC\" w:themeColor=\"text1\" "
            + "w:themeTint=\"80\" w:themeShade=\"40\"/>")
        let shading = try XCTUnwrap(props.shading)
        XCTAssertEqual(shading.themeFill, "accent2")
        XCTAssertEqual(shading.themeFillTint, "33")
        XCTAssertEqual(shading.themeFillShade, "CC")
        XCTAssertEqual(shading.themeColor, "text1")
        XCTAssertEqual(shading.themeTint, "80")
        XCTAssertEqual(shading.themeShade, "40")
    }

    func testToXMLReemitsCellShadingThemeAttributes() throws {
        let props = try parseTcPrFragment(
            "<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"auto\" w:themeFill=\"accent2\" w:themeFillTint=\"33\"/>")
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("w:themeFill=\"accent2\""), "got: \(xml)")
        XCTAssertTrue(xml.contains("w:themeFillTint=\"33\""), "got: \(xml)")
    }

    func testCellShadingWithoutThemeAttributesUnaffected() throws {
        let props = try parseTcPrFragment("<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"FFFF00\"/>")
        let shading = try XCTUnwrap(props.shading)
        XCTAssertNil(shading.themeFill)
        XCTAssertEqual(shading.fill, "FFFF00")
        XCTAssertTrue(props.toXML().contains("w:fill=\"FFFF00\""))
    }

    // MARK: - Section B: full-docx regression

    /// Table with two rows. Row 0's cell has a theme-colored `<w:shd>`; row 1
    /// is a plain untouched row. Paragraph outside the table is the target
    /// of the unrelated typed edit.
    private static let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
        <w:p><w:r><w:t>Intro</w:t></w:r></w:p>\
        <w:tbl>\
        <w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr>\
        <w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/>\
        <w:shd w:val="clear" w:color="auto" w:fill="auto" w:themeFill="accent2" w:themeFillTint="33"/>\
        </w:tcPr><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>\
        <w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Cell2</w:t></w:r></w:p></w:tc></w:tr>\
        </w:tbl>\
        <w:p><w:r><w:t>Outro</w:t></w:r></w:p>\
        </w:body></w:document>
        """

    private func buildFixture() throws -> URL {
        let staging = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue188-staging-\(UUID().uuidString)")
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
            .appendingPathComponent("issue188-\(UUID().uuidString).docx")
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

    func testTypedEditThenSavePreservesCellShadingThemeOnUntouchedRow() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        try doc.updateParagraph(at: 1, text: "OUTRO2")
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))

        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue188-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(doc, to: outURL)

        let outputXML = try extractDocumentXML(from: outURL)
        XCTAssertTrue(outputXML.contains("w:themeFill=\"accent2\""), "got: \(outputXML)")
        XCTAssertTrue(outputXML.contains("w:themeFillTint=\"33\""), "got: \(outputXML)")
    }
}
