import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// PsychQuant/ooxml-swift#133 (Decision (B), 2026-11-09): a typed-dirty
/// save rewrites only the body blocks the edit touched. Every other block, and
/// everything outside `<w:body>`, keeps its source bytes.
///
/// The marker for "this block was rewritten from the typed model" is a
/// paragraph-level section break (`w:pPr/w:sectPr`): the typed model drops it
/// (#133 comment, 2026-09-25), so it survives a save only when the block's
/// source bytes were kept. All fixtures are synthetic.
final class Issue133BodySurgeryTests: XCTestCase {

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
        let url = FileManager.default.temporaryDirectory.appendingPathComponent("i133-\(UUID().uuidString).docx")
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
        let url = FileManager.default.temporaryDirectory.appendingPathComponent("i133-out-\(UUID().uuidString).docx")
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

    /// Two blocks the typed model serializes identically but whose source
    /// bytes differ: the deleted one's bytes must not survive, and the one
    /// that remains keeps its own (verify R1: content equality alone cannot
    /// tell them apart, so each kept block must come from its own source).
    func testDeletingOneOfTwoTypedEqualBlocksDoesNotKeepItsBytes() throws {
        let first = #"<w:p><w:pPr><w:sectPr><w:pgSz w:w="10001" w:h="14000"/></w:sectPr></w:pPr><w:r><w:t>同文</w:t></w:r></w:p>"#
        let second = #"<w:p><w:pPr><w:sectPr><w:pgSz w:w="10002" w:h="14000"/></w:sectPr></w:pPr><w:r><w:t>同文</w:t></w:r></w:p>"#
        var doc = try read(documentXML(body: first + second))
        defer { doc.close() }
        try doc.deleteParagraph(at: 0)
        let xml = try save(doc)
        XCTAssertFalse(xml.contains(#"w:w="10001""#), "the deleted block's bytes must not survive:\n\(xml)")
        XCTAssertTrue(xml.contains(second), "the remaining block keeps its own bytes:\n\(xml)")
        XCTAssertEqual(xml.components(separatedBy: "同文").count - 1, 1)
    }

    /// The same typed-equal pair, but the edit only inserts a new block (a new
    /// paragraph gets a fresh paraId, so it can never equal either): the pair
    /// keeps its order and count, so each keeps its own bytes.
    func testInsertingNextToTypedEqualBlocksKeepsBoth() throws {
        let first = #"<w:p><w:pPr><w:sectPr><w:pgSz w:w="10001" w:h="14000"/></w:sectPr></w:pPr></w:p>"#
        let second = #"<w:p><w:pPr><w:sectPr><w:pgSz w:w="10002" w:h="14000"/></w:sectPr></w:pPr></w:p>"#
        var doc = try read(documentXML(body: first + second))
        defer { doc.close() }
        doc.insertParagraph(Paragraph(text: "新"), at: 0)
        let xml = try save(doc)
        XCTAssertTrue(xml.contains(first + second), "both blocks keep their bytes, in order:\n\(xml)")
    }

    // MARK: - A kept block's bytes come from that block (verify R1)

    private let thirdBlock = "<w:p><w:r><w:t>丙段</w:t></w:r></w:p>"

    /// Delete the first block and add one at the end: the counts match but
    /// no position lines up. The untouched blocks keep their bytes.
    func testDeletingTheFirstAndAddingAtTheEndKeepsTheUntouchedBlocks() throws {
        var doc = try read(documentXML(body: firstBlock + keptBlock + thirdBlock))
        defer { doc.close() }
        try doc.deleteParagraph(at: 0)
        doc.insertParagraph(Paragraph(text: "新尾"), at: 2)
        let xml = try save(doc)
        XCTAssertTrue(xml.contains(keptBlock + thirdBlock), xml)
        XCTAssertFalse(xml.contains("甲段"))
        XCTAssertTrue(xml.contains("新尾"))
    }

    /// Two typed-equal paragraphs whose section breaks differ (no paraId);
    /// delete the first, then change a third paragraph to the same text.
    /// The deleted block's section break must not survive on the edited one.
    func testAnEditedBlockNeverBorrowsADeletedTwinsBytes() throws {
        let first = #"<w:p><w:pPr><w:sectPr><w:pgSz w:w="10001" w:h="14000"/></w:sectPr></w:pPr><w:r><w:t>同文</w:t></w:r></w:p>"#
        let second = #"<w:p><w:pPr><w:sectPr><w:pgSz w:w="10002" w:h="14000"/></w:sectPr></w:pPr><w:r><w:t>同文</w:t></w:r></w:p>"#
        var doc = try read(documentXML(body: first + second + "<w:p><w:r><w:t>其他</w:t></w:r></w:p>"))
        defer { doc.close() }
        try doc.deleteParagraph(at: 0)
        try doc.updateParagraph(at: 1, text: "同文")
        let xml = try save(doc)
        XCTAssertFalse(xml.contains(#"w:w="10001""#), "the deleted block's section break survived:\n\(xml)")
        XCTAssertTrue(xml.contains(second), "the untouched twin keeps its own bytes:\n\(xml)")
        XCTAssertEqual(xml.components(separatedBy: "同文").count - 1, 2)
    }

    /// Without paraIds, empty paragraphs and section-break paragraphs are
    /// typed-equal. Deleting an ordinary empty paragraph must not cost the
    /// section breaks around it.
    func testDeletingAnEmptyParagraphKeepsTheSectionBreaksAroundIt() throws {
        let secA = #"<w:p><w:pPr><w:sectPr><w:pgSz w:w="10001" w:h="14000"/></w:sectPr></w:pPr></w:p>"#
        let secB = #"<w:p><w:pPr><w:sectPr><w:pgSz w:w="10002" w:h="14000"/></w:sectPr></w:pPr></w:p>"#
        var doc = try read(documentXML(body: firstBlock + secA + "<w:p/><w:p/>" + secB + thirdBlock))
        defer { doc.close() }
        try doc.deleteParagraph(at: 2)
        let xml = try save(doc)
        XCTAssertTrue(xml.contains(secA), xml)
        XCTAssertTrue(xml.contains(secB), xml)
        let paragraphStart = try NSRegularExpression(pattern: "<w:p[ />]")
        XCTAssertEqual(paragraphStart.numberOfMatches(in: xml, range: NSRange(xml.startIndex..., in: xml)), 5,
                       "one paragraph fewer:\n\(xml)")
    }

    /// Two deletions far apart in a long body: every block between them is
    /// untouched and keeps its bytes, however long the stretch.
    func testTwoDistantDeletionsInALongBodyKeepTheUntouchedStretch() throws {
        var body = ""
        for i in 0..<2500 {
            body += i == 1200 ? keptBlock : "<w:p><w:r><w:t>段\(i)</w:t></w:r></w:p>"
        }
        var doc = try read(documentXML(body: body))
        defer { doc.close() }
        try doc.deleteParagraph(at: 2398)
        try doc.deleteParagraph(at: 0)
        let xml = try save(doc)
        XCTAssertTrue(xml.contains(keptBlock), "a block between the two edits lost its bytes")
        XCTAssertFalse(xml.contains(">段0<"))
        XCTAssertFalse(xml.contains(">段2398<"))
        XCTAssertTrue(xml.contains(">段2399<"))
    }

    /// A copy gets a fresh paraId on insert, so it is a different block: it
    /// must not gain the original's section break, and the original keeps it.
    func testACopiedBlockDoesNotDuplicateUnmodelledContent() throws {
        var doc = try read(documentXML(body: keptBlock + firstBlock))
        defer { doc.close() }
        guard case .paragraph(let original) = doc.body.children[0] else { return XCTFail("fixture") }
        doc.insertParagraph(original, at: 2)
        let xml = try save(doc)
        XCTAssertTrue(xml.contains(keptBlock), xml)
        XCTAssertEqual(xml.components(separatedBy: #"w:w="10000""#).count - 1, 1,
                       "the section break exists once, on the original:\n\(xml)")
        XCTAssertEqual(xml.components(separatedBy: "乙段").count - 1, 2)
    }

    /// A block value taken from another document never borrows bytes here,
    /// even when it lands where a typed-equal block was deleted.
    func testABlockFromAnotherDocumentNeverBorrowsBytes() throws {
        let k1 = #"<w:p><w:pPr><w:sectPr><w:pgSz w:w="10001" w:h="14000"/></w:sectPr></w:pPr><w:r><w:t>同文</w:t></w:r></w:p>"#
        let k2 = #"<w:p><w:pPr><w:sectPr><w:pgSz w:w="10002" w:h="14000"/></w:sectPr></w:pPr><w:r><w:t>同文</w:t></w:r></w:p>"#
        var other = try read(documentXML(body: k1 + firstBlock))
        defer { other.close() }
        var doc = try read(documentXML(body: k2 + firstBlock))
        defer { doc.close() }
        let foreign = other.body.children[0]
        try doc.deleteParagraph(at: 0)
        // A value that reaches the body without paraId stamping.
        doc.body.children.insert(foreign, at: 0)
        doc.markTypedDirty("word/document.xml")
        let xml = try save(doc)
        XCTAssertFalse(xml.contains(#"w:w="10002""#), "the deleted block's bytes survived:\n\(xml)")
        XCTAssertFalse(xml.contains(#"w:w="10001""#), "the other document's bytes leaked in:\n\(xml)")
        XCTAssertTrue(xml.contains("同文"))
    }

    /// The root start tag is rebuilt from `documentRootAttributes` when they
    /// changed after reading (the old writer did; surgery must not drop it).
    func testRootAttributesChangedAfterReadAreWritten() throws {
        var doc = try read(documentXML(body: firstBlock + keptBlock))
        defer { doc.close() }
        doc.documentRootAttributes["xmlns:foo"] = "urn:example:foo"
        doc.insertParagraph(Paragraph(text: "新"), at: 0)
        let xml = try save(doc)
        XCTAssertTrue(xml.contains(#"xmlns:foo="urn:example:foo""#), xml)
        XCTAssertTrue(xml.contains("新"))
    }

    /// A rewritten block is written where it stood, so a sibling the typed
    /// model does not see (here a comment) keeps its place after it.
    func testARewrittenBlockStaysBeforeTheSiblingThatFollowedIt() throws {
        var doc = try read(documentXML(body: firstBlock + "<!-- note -->" + keptBlock))
        defer { doc.close() }
        try doc.updateParagraph(at: 0, text: "新甲")
        let xml = try save(doc)
        let edited = try XCTUnwrap(xml.range(of: "新甲"))
        let note = try XCTUnwrap(xml.range(of: "<!-- note -->"))
        XCTAssertLessThan(edited.lowerBound, note.lowerBound, xml)
        XCTAssertTrue(xml.contains("<!-- note -->" + keptBlock), xml)
    }

    /// The reader takes the last body-level `sectPr` and the splice assumes
    /// it ends the body; a body with two gets no baseline (full writer).
    func testABodyWithTwoSectionPropertiesGetsNoBaseline() throws {
        let doc = try read(documentXML(body: firstBlock + "<w:sectPr/>" + keptBlock))
        defer { var d = doc; d.close() }
        XCTAssertNil(doc.bodySourceBaseline)
    }

    // MARK: - Namespace repair of a rewritten block (verify R1)

    private let mainURI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    private let w14URI = "http://schemas.microsoft.com/office/word/2010/wordml"

    func testRepairDoesNotRedeclareAPrefixTheBlockDeclaresItself() throws {
        let block = #"<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="12345678"/>"#
        let repaired = try XCTUnwrap(BodySurgery.repairNamespaces(block, rootBindings: ["w": mainURI]))
        XCTAssertEqual(repaired.components(separatedBy: "xmlns:w14=").count - 1, 1, repaired)
        XCTAssertNoThrow(try XMLDocument(xmlString: "<root xmlns:w=\"\(mainURI)\">\(repaired)</root>", options: []))
    }

    func testRepairAcceptsAPrefixDeclaredInsideTheBlock() {
        let block = #"<w:p><w:r><x:mark xmlns:x="urn:example:x"/></w:r></w:p>"#
        XCTAssertEqual(BodySurgery.repairNamespaces(block, rootBindings: ["w": mainURI]), block)
    }

    /// A Strict document binds `w` to another URI; the rewritten block stays
    /// in the root's namespace (as the full typed writer would write it).
    func testRepairKeepsTheRootsBindingOfAStandardPrefix() {
        let strict = "http://purl.oclc.org/ooxml/wordprocessingml/main"
        let block = "<w:p><w:r><w:t>x</w:t></w:r></w:p>"
        XCTAssertEqual(BodySurgery.repairNamespaces(block, rootBindings: ["w": strict]), block)
    }

    func testRepairDeclaresAStandardPrefixNobodyBinds() throws {
        let block = #"<w:p w14:paraId="12345678"/>"#
        let repaired = try XCTUnwrap(BodySurgery.repairNamespaces(block, rootBindings: ["w": mainURI]))
        XCTAssertTrue(repaired.contains("xmlns:w14=\"\(w14URI)\""), repaired)
    }

    func testRepairGivesUpOnAPrefixWithNoKnownURI() {
        XCTAssertNil(BodySurgery.repairNamespaces(#"<w:p><zz:x/></w:p>"#, rootBindings: ["w": mainURI]))
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

    /// With `word/document.xml` already typed-dirty, `appendParagraph` stays on
    /// the typed path, so surgery still applies and untouched blocks keep
    /// their bytes (the order the CHANGELOG says is unaffected).
    func testAppendAfterATypedEditStillKeepsUntouchedBlocks() throws {
        var doc = try read(documentXML(body: firstBlock + keptBlock))
        defer { doc.close() }
        doc.insertParagraph(Paragraph(text: "首段"), at: 0)
        doc.appendParagraph(Paragraph(text: "尾段"))
        let xml = try save(doc)
        XCTAssertTrue(xml.contains(firstBlock + keptBlock), xml)
        XCTAssertTrue(xml.contains("首段"))
        XCTAssertTrue(xml.contains("尾段"))
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
        // Known limit (CHANGELOG): the read-time baseline does not follow the
        // tree, so in this order the untouched block is rewritten and loses
        // its section break. When this starts failing, the limit is gone:
        // update the CHANGELOG.
        XCTAssertFalse(xml.contains(keptBlock), "known limit changed:\n\(xml)")
    }
}
