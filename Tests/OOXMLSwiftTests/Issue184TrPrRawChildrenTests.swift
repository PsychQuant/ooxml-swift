import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Regression tests for [PsychQuant/ooxml-swift#184](https://github.com/PsychQuant/ooxml-swift/issues/184).
///
/// `<w:trPr>` child elements the typed model has no field for — `w:cnfStyle`,
/// `w:divId`, `w:gridBefore`/`w:gridAfter`, `w:wBefore`/`w:wAfter`,
/// `w:tblCellSpacing`, `w:ins`/`w:del`, `w:trPrChange` — were silently dropped
/// from EVERY row in the document the moment ANY typed edit marked
/// `word/document.xml` dirty (the writer regenerates the whole part from the
/// typed model). Fix: `TableRowProperties.rawChildren` captures them verbatim
/// (same "if not typed, preserve as raw" principle as `ParagraphProperties
/// .rawChildren` / `Run.rawElements`), and `TableRow.toXML()` re-emits them
/// interleaved with the typed fields in `CT_TrPr` schema order.
final class Issue184TrPrRawChildrenTests: XCTestCase {

    // MARK: - Section A: unit-level parse/emit round trip (no docx I/O)

    private func parseTrPrFragment(_ xml: String) throws -> TableRowProperties {
        let wrapper = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:tr>\(xml)</w:tr>
        </w:document>
        """
        let doc = try XMLDocument(xmlString: wrapper)
        guard let root = doc.rootElement(),
              let tr = root.elements(forName: "w:tr").first,
              let trPrElement = tr.elements(forName: "w:trPr").first else {
            throw NSError(domain: "test", code: 0, userInfo: [NSLocalizedDescriptionKey: "no <w:trPr> found"])
        }
        return DocxReader.parseTableRowProperties(from: trPrElement)
    }

    func testParseCapturesUnmodeledChildrenAsRawElements() throws {
        let props = try parseTrPrFragment(
            "<w:trPr><w:cnfStyle w:val=\"100000000000\"/><w:gridBefore w:val=\"1\"/>"
            + "<w:wBefore w:w=\"200\" w:type=\"dxa\"/><w:tblCellSpacing w:w=\"20\" w:type=\"dxa\"/>"
            + "<w:jc w:val=\"center\"/><w:ins w:id=\"3\" w:author=\"Che\" w:date=\"2026-01-01T00:00:00Z\"/></w:trPr>")
        let names = props.rawChildren.map { $0.name }
        XCTAssertEqual(names, ["cnfStyle", "gridBefore", "wBefore", "tblCellSpacing", "ins"], "got: \(names)")
        // Typed field still parsed alongside raw capture.
        XCTAssertEqual(props.justification, .center)
    }

    func testToXMLReemitsRawChildrenInCanonicalSchemaOrder() throws {
        // Source order deliberately scrambled (ins before cnfStyle before
        // gridBefore) to prove toXML() re-sorts by CT_TrPr schema position,
        // not source/array order.
        let props = try parseTrPrFragment(
            "<w:trPr><w:ins w:id=\"3\" w:author=\"Che\" w:date=\"2026-01-01T00:00:00Z\"/>"
            + "<w:cnfStyle w:val=\"100000000000\"/><w:gridBefore w:val=\"1\"/>"
            + "<w:cantSplit/><w:hidden/></w:trPr>")
        let xml = TableRow(properties: props).toXML()
        func index(_ needle: String) throws -> String.Index {
            try XCTUnwrap(xml.range(of: needle), "got: \(xml)").lowerBound
        }
        let cnfIdx = try index("<w:cnfStyle")
        let gridIdx = try index("<w:gridBefore")
        let cantSplitIdx = try index("<w:cantSplit")
        let hiddenIdx = try index("<w:hidden")
        let insIdx = try index("<w:ins")
        // Schema order: cnfStyle, gridBefore, cantSplit, hidden(before ins is
        // wrong — hidden precedes ins in EG_TrPrBase, ins is the trailing
        // CT_TrPr sequence element), ins.
        XCTAssertLessThan(cnfIdx, gridIdx, "got: \(xml)")
        XCTAssertLessThan(gridIdx, cantSplitIdx, "got: \(xml)")
        XCTAssertLessThan(cantSplitIdx, hiddenIdx, "got: \(xml)")
        XCTAssertLessThan(hiddenIdx, insIdx, "got: \(xml)")
    }

    func testNoUnmodeledChildrenEmitsUnchangedTrPr() throws {
        let props = try parseTrPrFragment("<w:trPr><w:jc w:val=\"center\"/></w:trPr>")
        XCTAssertTrue(props.rawChildren.isEmpty)
        let xml = TableRow(properties: props).toXML()
        XCTAssertTrue(xml.contains("<w:trPr><w:jc w:val=\"center\"/></w:trPr>"), "got: \(xml)")
    }

    // MARK: - Section B: full-docx regression via the issue's exact repro

    /// Body: table with two rows. Row 0 carries every unmodeled `<w:trPr>`
    /// child the issue lists; row 1 is a plain untouched row. A second
    /// paragraph outside the table is the target of the unrelated typed edit.
    private static let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
        <w:p><w:r><w:t>Intro</w:t></w:r></w:p>\
        <w:tbl>\
        <w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr>\
        <w:tr>\
        <w:trPr><w:cnfStyle w:val="100000000000"/><w:divId w:val="123456"/>\
        <w:gridBefore w:val="1"/><w:gridAfter w:val="1"/>\
        <w:wBefore w:w="100" w:type="dxa"/><w:wAfter w:w="100" w:type="dxa"/>\
        <w:tblCellSpacing w:w="20" w:type="dxa"/>\
        <w:ins w:id="7" w:author="Che" w:date="2026-01-01T00:00:00Z"/></w:trPr>\
        <w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc>\
        </w:tr>\
        <w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Cell2</w:t></w:r></w:p></w:tc></w:tr>\
        </w:tbl>\
        <w:p><w:r><w:t>Outro</w:t></w:r></w:p>\
        </w:body></w:document>
        """

    private func buildFixture() throws -> URL {
        let staging = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue184-staging-\(UUID().uuidString)")
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
            .appendingPathComponent("issue184-\(UUID().uuidString).docx")
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

    func testTypedEditThenSavePreservesTrPrUnmodeledChildrenOnUntouchedRow() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        // Unrelated edit: touches the paragraph AFTER the table, not the table itself.
        try doc.updateParagraph(at: 1, text: "OUTRO2")
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))

        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue184-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(doc, to: outURL)

        let outputXML = try extractDocumentXML(from: outURL)
        for needle in ["<w:cnfStyle w:val=\"100000000000\"/>",
                       "<w:divId w:val=\"123456\"/>",
                       "<w:gridBefore w:val=\"1\"/>",
                       "<w:gridAfter w:val=\"1\"/>",
                       "<w:wBefore w:w=\"100\" w:type=\"dxa\"/>",
                       "<w:wAfter w:w=\"100\" w:type=\"dxa\"/>",
                       "<w:tblCellSpacing w:w=\"20\" w:type=\"dxa\"/>",
                       "<w:ins w:id=\"7\" w:author=\"Che\" w:date=\"2026-01-01T00:00:00Z\"/>"] {
            XCTAssertTrue(outputXML.contains(needle),
                          "expected \(needle) to survive an unrelated typed edit — got: \(outputXML)")
        }
    }
}
