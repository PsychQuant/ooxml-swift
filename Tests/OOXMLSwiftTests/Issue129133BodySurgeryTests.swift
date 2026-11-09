import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// PsychQuant/ooxml-swift#129 + #133 (Decision (B), 2026-11-09): a typed-dirty
/// save rewrites only the body blocks the edit touched. Every other block, and
/// everything outside `<w:body>`, keeps its source bytes.
///
/// The marker for "this block was rewritten from the typed model" is a
/// paragraph-level section break (`w:pPr/w:sectPr`): the typed model drops it
/// (#133 comment, 2026-09-25), so it survives a save only when the block's
/// source bytes were kept. All fixtures are synthetic.
final class Issue129133BodySurgeryTests: XCTestCase {

    // MARK: - Fixture plumbing

    private static let fullNamespaces = [
        #"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#,
        #"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#,
        #"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#,
        #"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006""#,
        #"mc:Ignorable="w14""#,
    ].joined(separator: " ")

    private static let minimalNamespaces = #"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#

    /// A block the typed model cannot reproduce: its paragraph-level
    /// section break is dropped by a typed re-serialization.
    private let keptBlock = #"<w:p><w:pPr><w:sectPr><w:pgSz w:w="10000" w:h="14000"/></w:sectPr></w:pPr><w:r><w:t>乙段</w:t></w:r></w:p>"#
    private let firstBlock = "<w:p><w:r><w:t>甲段</w:t></w:r></w:p>"

    private var cleanup: [URL] = []

    override func tearDown() {
        for url in cleanup { try? FileManager.default.removeItem(at: url) }
        cleanup = []
        super.tearDown()
    }

    private func documentXML(body: String, beforeBody: String = "", sectPr: String = "<w:sectPr/>",
                             namespaces: String = fullNamespaces) -> String {
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
            + "<w:document \(namespaces)>\(beforeBody)<w:body>\(body)\(sectPr)</w:body></w:document>"
    }

    private func makeDocx(_ documentXML: String) throws -> URL {
        let parts: [String: Data] = [
            "[Content_Types].xml": Data(#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.utf8),
            "_rels/.rels": Data(#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.utf8),
            "word/document.xml": Data(documentXML.utf8),
            "word/_rels/document.xml.rels": Data(#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#.utf8),
        ]
        let archive = try Archive(accessMode: .create)
        for name in parts.keys.sorted() {
            let data = parts[name]!
            try archive.addEntry(with: name, type: .file, uncompressedSize: Int64(data.count), compressionMethod: .deflate) { position, size in
                let start = data.startIndex.advanced(by: Int(position))
                return data.subdata(in: start..<start.advanced(by: size))
            }
        }
        let url = FileManager.default.temporaryDirectory.appendingPathComponent("i129133-\(UUID().uuidString).docx")
        try XCTUnwrap(archive.data).write(to: url)
        cleanup.append(url)
        return url
    }

    private func documentPart(ofArchiveData data: Data) throws -> String {
        let archive = try Archive(data: data, accessMode: .read)
        let entry = try XCTUnwrap(archive["word/document.xml"])
        var out = Data()
        _ = try archive.extract(entry) { out.append($0) }
        return String(decoding: out, as: UTF8.self)
    }

    /// Save in overlay mode (the path che-word-mcp uses) and return the
    /// saved `word/document.xml`.
    private func save(_ doc: WordDocument) throws -> String {
        let url = FileManager.default.temporaryDirectory.appendingPathComponent("i129133-out-\(UUID().uuidString).docx")
        cleanup.append(url)
        try DocxWriter.write(doc, to: url)
        return try documentPart(ofArchiveData: Data(contentsOf: url))
    }

    private func read(_ xml: String) throws -> WordDocument {
        try DocxReader.read(from: makeDocx(xml))
    }

    // MARK: - Untouched blocks keep their bytes

    func testUpdateParagraphKeepsTheOtherBlockByteForByte() throws {
        var doc = try read(documentXML(body: firstBlock + keptBlock))
        defer { doc.close() }
        try doc.updateParagraph(at: 0, text: "新甲")
        let xml = try save(doc)
        XCTAssertTrue(xml.contains(keptBlock), "the untouched block must keep its source bytes:\n\(xml)")
        XCTAssertTrue(xml.contains("新甲"))
        XCTAssertFalse(xml.contains("甲段"))
    }

    func testDeleteParagraphKeepsTheOtherBlockByteForByte() throws {
        var doc = try read(documentXML(body: firstBlock + keptBlock))
        defer { doc.close() }
        try doc.deleteParagraph(at: 0)
        let xml = try save(doc)
        XCTAssertTrue(xml.contains(keptBlock), xml)
        XCTAssertFalse(xml.contains("甲段"))
    }

    func testInsertTableKeepsTheOtherBlocksByteForByte() throws {
        var doc = try read(documentXML(body: firstBlock + keptBlock))
        defer { doc.close() }
        doc.insertTable(Table(rowCount: 1, columnCount: 1), at: 1)
        let xml = try save(doc)
        XCTAssertTrue(xml.contains(firstBlock + "<w:tbl"), "the table goes between the two blocks:\n\(xml)")
        XCTAssertTrue(xml.contains(keptBlock), xml)
    }

    func testSetPageSizeReplacesOnlyTheSectionProperties() throws {
        let source = documentXML(body: firstBlock + keptBlock,
                                 sectPr: #"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>"#)
        var doc = try read(source)
        defer { doc.close() }
        doc.setPageSize(.a4)
        let xml = try save(doc)
        XCTAssertTrue(xml.contains(firstBlock + keptBlock), "every block keeps its bytes:\n\(xml)")
        XCTAssertFalse(xml.contains(#"w:w="12240""#), "the body-level section properties are rewritten")
        XCTAssertTrue(xml.contains(#"w:w="11906""#))
    }

    // MARK: - Content outside the blocks

    func testContentOutsideTheBodySurvivesATypedEdit() throws {
        var doc = try read(documentXML(body: firstBlock, beforeBody: #"<w:background w:color="FF0000"/>"#))
        defer { doc.close() }
        doc.insertParagraph(Paragraph(text: "新"), at: 0)
        let xml = try save(doc)
        XCTAssertTrue(xml.contains(#"<w:background w:color="FF0000"/>"#), xml)
        XCTAssertTrue(xml.hasPrefix("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"),
                      "the source prolog (CRLF) is kept")
    }

    func testUnchangedSectionPropertiesAreNotAddedWhenTheSourceHadNone() throws {
        var doc = try read(documentXML(body: firstBlock, sectPr: ""))
        defer { doc.close() }
        doc.insertParagraph(Paragraph(text: "新"), at: 0)
        let xml = try save(doc)
        XCTAssertFalse(xml.contains("<w:sectPr"), "nothing changed the section properties:\n\(xml)")
    }

    func testInsertedParagraphDeclaresW14WhenTheRootDoesNot() throws {
        var doc = try read(documentXML(body: keptBlock, namespaces: Self.minimalNamespaces))
        defer { doc.close() }
        doc.insertParagraph(Paragraph(text: "新"), at: 0)
        let xml = try save(doc)
        XCTAssertTrue(xml.contains(keptBlock), xml)
        XCTAssertNoThrow(try XMLDocument(xmlString: xml, options: []), "the saved part must be well-formed")
        // Every w14 prefix in use resolves: either on the new paragraph or the root.
        if xml.contains("w14:") {
            XCTAssertTrue(xml.contains(#"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#), xml)
        }
    }

    // MARK: - Saves without an edit, repeated saves, fallback

    func testWriteDataOfAnUnmodifiedDocumentKeepsDocumentXMLByteIdentical() throws {
        let source = documentXML(body: firstBlock + keptBlock)
        let doc = try read(source)
        defer { var d = doc; d.close() }
        let xml = try documentPart(ofArchiveData: DocxWriter.writeData(doc))
        XCTAssertEqual(xml, source)
    }

    func testRepeatedSavesAreIdentical() throws {
        var doc = try read(documentXML(body: firstBlock + keptBlock))
        defer { doc.close() }
        doc.insertParagraph(Paragraph(text: "新"), at: 1)
        XCTAssertEqual(try save(doc), try save(doc))
    }

    /// An op-log edit makes the live tree the authority and the read-time
    /// baseline stale; a later typed edit must fall back to the full typed
    /// serialization and keep both edits.
    func testTypedEditAfterAnOpLogEditKeepsBothEdits() throws {
        var doc = try read(documentXML(body: firstBlock + keptBlock))
        defer { doc.close() }
        doc.appendParagraph(Paragraph(text: "尾段"))
        doc.insertParagraph(Paragraph(text: "首段"), at: 0)
        let xml = try save(doc)
        for text in ["首段", "甲段", "乙段", "尾段"] {
            XCTAssertTrue(xml.contains(text), "\(text) missing:\n\(xml)")
        }
        XCTAssertNoThrow(try XMLDocument(xmlString: xml, options: []))
    }
}
