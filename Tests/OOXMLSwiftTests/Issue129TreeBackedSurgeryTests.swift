import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// PsychQuant/ooxml-swift#129 (Decision (B), 2026-11-09): a document read with
/// `wireTreeBackedViews: true` gets the same body surgery as the default read
/// mode (#133). A typed-dirty save then keeps the source bytes of the blocks
/// the edit did not touch, so their `<w:drawing>` survives, although a
/// tree-backed paragraph's typed runs are a text-only stub. All fixtures are
/// synthetic.
final class Issue129TreeBackedSurgeryTests: XCTestCase {

    // MARK: - Fixture plumbing

    private static let namespaces = [
        #"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#,
        #"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#,
        #"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#,
        #"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006""#,
        #"mc:Ignorable="w14""#,
    ].joined(separator: " ")

    /// A paragraph whose only run is an inline picture.
    private let drawingParagraph = #"<w:p w14:paraId="11111111"><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill><a:blip r:embed="rId4"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#
    /// A second picture, on its own relationship.
    private var drawingParagraphB: String {
        drawingParagraph.replacingOccurrences(of: "11111111", with: "33333333").replacingOccurrences(of: "rId4", with: "rId5")
    }
    private let textParagraph = #"<w:p w14:paraId="22222222"><w:r><w:t>文字段</w:t></w:r></w:p>"#

    private var cleanup: [URL] = []

    override func tearDown() {
        for url in cleanup { try? FileManager.default.removeItem(at: url) }
        cleanup = []
        super.tearDown()
    }

    private func makeDocx(body: String) throws -> URL {
        let documentXML = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
            + "<w:document \(Self.namespaces)><w:body>\(body)<w:sectPr/></w:body></w:document>"
        let parts: [String: Data] = [
            "[Content_Types].xml": Data(#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.utf8),
            "_rels/.rels": Data(#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.utf8),
            "word/document.xml": Data(documentXML.utf8),
            "word/_rels/document.xml.rels": Data(#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.png"/></Relationships>"#.utf8),
            "word/media/image1.png": Data("png".utf8),
            "word/media/image2.png": Data("png2".utf8),
        ]
        let archive = try Archive(accessMode: .create)
        for name in parts.keys.sorted() {
            let data = parts[name]!
            try archive.addEntry(with: name, type: .file, uncompressedSize: Int64(data.count), compressionMethod: .deflate) { position, size in
                let start = data.startIndex.advanced(by: Int(position))
                return data.subdata(in: start..<start.advanced(by: size))
            }
        }
        let url = FileManager.default.temporaryDirectory.appendingPathComponent("i129-\(UUID().uuidString).docx")
        try XCTUnwrap(archive.data).write(to: url)
        cleanup.append(url)
        return url
    }

    private func readTreeBacked(body: String) throws -> WordDocument {
        try DocxReader.read(from: makeDocx(body: body), wireTreeBackedViews: true)
    }

    /// Save in overlay mode (the path che-word-mcp uses) and return the
    /// saved `word/document.xml`.
    private func save(_ doc: WordDocument) throws -> String {
        let url = FileManager.default.temporaryDirectory.appendingPathComponent("i129-out-\(UUID().uuidString).docx")
        cleanup.append(url)
        try DocxWriter.write(doc, to: url)
        let archive = try Archive(url: url, accessMode: .read)
        var out = Data()
        _ = try archive.extract(try XCTUnwrap(archive["word/document.xml"])) { out.append($0) }
        return String(decoding: out, as: UTF8.self)
    }

    private func drawingCount(_ xml: String) -> Int { xml.components(separatedBy: "<w:drawing>").count - 1 }

    // MARK: - The reported triggers

    func testInsertParagraphKeepsAnUntouchedDrawing() throws {
        var doc = try readTreeBacked(body: drawingParagraph + textParagraph)
        defer { doc.close() }
        doc.insertParagraph(Paragraph(text: "新段"), at: 0)
        let xml = try save(doc)
        XCTAssertEqual(drawingCount(xml), 1, xml)
        XCTAssertTrue(xml.contains(drawingParagraph + textParagraph), "untouched blocks keep their bytes:\n\(xml)")
        XCTAssertTrue(xml.contains("新段"))
    }

    /// On a freshly read document, a paragraph the op payload cannot carry
    /// (here: a page break) is grafted into the tree (#132), so the save
    /// serializes the tree and never reaches surgery. Pinned as it behaves;
    /// the typed fallback is the next test.
    func testAppendOfANonRepresentableParagraphIsGrafted() throws {
        var doc = try readTreeBacked(body: drawingParagraph + textParagraph)
        defer { doc.close() }
        var paragraph = Paragraph(text: "尾段")
        paragraph.hasPageBreak = true
        doc.appendParagraph(paragraph)
        let xml = try save(doc)
        XCTAssertEqual(drawingCount(xml), 1, xml)
        XCTAssertTrue(xml.contains(drawingParagraph + textParagraph), xml)
        XCTAssertTrue(xml.contains("尾段"))
    }

    /// A paragraph the op payload can carry goes through the op log and the
    /// tree; the drawing was never at risk there. Pinned so both append
    /// paths are covered.
    func testRepresentableAppendKeepsTheDrawing() throws {
        var doc = try readTreeBacked(body: drawingParagraph + textParagraph)
        defer { doc.close() }
        doc.appendParagraph(Paragraph(text: "尾段"))
        let xml = try save(doc)
        XCTAssertEqual(drawingCount(xml), 1, xml)
        XCTAssertTrue(xml.contains("尾段"))
    }

    func testDeleteParagraphKeepsAnUntouchedDrawing() throws {
        var doc = try readTreeBacked(body: textParagraph + drawingParagraph)
        defer { doc.close() }
        try doc.deleteParagraph(at: 0)
        let xml = try save(doc)
        XCTAssertEqual(drawingCount(xml), 1, xml)
        XCTAssertFalse(xml.contains("文字段"))
    }

    /// The tree-backed wiring misbinds paragraphs after a body-level bookmark
    /// (#242). Surgery pairs blocks by its own rules, so the untouched blocks
    /// still keep their own bytes.
    func testABodyLevelBookmarkDoesNotMisplaceKeptBytes() throws {
        let body = #"<w:bookmarkStart w:id="0" w:name="b"/>"# + drawingParagraph + textParagraph + #"<w:bookmarkEnd w:id="0"/>"#
        var doc = try readTreeBacked(body: body)
        defer { doc.close() }
        doc.insertParagraph(Paragraph(text: "新段"), at: 0)
        let xml = try save(doc)
        XCTAssertEqual(drawingCount(xml), 1, xml)
        XCTAssertTrue(xml.contains(body), "the bookmarked stretch keeps its bytes:\n\(xml)")
    }

    /// The append trigger the issue reports: after a typed edit the graft
    /// refuses (the part is typed-dirty), so `appendParagraph` falls back to
    /// the typed path (non-representable paragraph) or materializes the typed
    /// part through surgery first (representable paragraph). Both keep the
    /// untouched drawing.
    func testAppendAfterATypedEditKeepsAnUntouchedDrawing() throws {
        for representable in [false, true] {
            var doc = try readTreeBacked(body: drawingParagraph + textParagraph)
            defer { doc.close() }
            doc.insertParagraph(Paragraph(text: "新段"), at: 0)
            var paragraph = Paragraph(text: "尾段")
            paragraph.hasPageBreak = !representable
            doc.appendParagraph(paragraph)
            let xml = try save(doc)
            XCTAssertEqual(drawingCount(xml), 1, "representable=\(representable):\n\(xml)")
            XCTAssertTrue(xml.contains(drawingParagraph + textParagraph), "representable=\(representable):\n\(xml)")
            XCTAssertTrue(xml.contains("新段") && xml.contains("尾段"), "representable=\(representable)")
        }
    }

    // MARK: - Resources removed since read (verify R1)

    /// `deleteImage` drops the relationship, but on a tree-backed paragraph it
    /// cannot remove the drawing from the typed runs (a text-only stub). The
    /// paragraph that shows the deleted picture must not keep its source
    /// bytes, or the saved file refers to a relationship that is gone. The
    /// other picture's paragraph is untouched and keeps its bytes.
    func testDeletingAnImageNeverLeavesADanglingReference() throws {
        var doc = try readTreeBacked(body: drawingParagraph + drawingParagraphB + textParagraph)
        defer { doc.close() }
        XCTAssertTrue(doc.images.contains { $0.id == "rId4" }, "fixture: the image must be typed")
        try doc.deleteImage(imageId: "rId4")
        let xml = try save(doc)
        XCTAssertFalse(xml.contains(#"r:embed="rId4""#), "a dangling reference to the deleted image:\n\(xml)")
        XCTAssertTrue(xml.contains(drawingParagraphB + textParagraph), "the other picture keeps its bytes:\n\(xml)")
        XCTAssertEqual(drawingCount(xml), 1, xml)
    }

    /// The scan that finds such references, for every kind of resource.
    func testTheReferenceScanCoversEveryResourceKind() throws {
        let parts: [String] = [
            #"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
            #"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" "#,
            #"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#,
            #"<w:hyperlink r:id="rId7"><w:r><w:t>連結</w:t></w:r></w:hyperlink>"#,
            #"<w:r><w:drawing><a:blip r:embed="rId4"/></w:drawing></w:r>"#,
            #"<w:r><w:footnoteReference w:id="3"/></w:r><w:r><w:endnoteReference w:id="5"/></w:r>"#,
            #"<w:commentRangeStart w:id="9"/><w:r><w:commentReference w:id="9"/></w:r></w:p>"#,
        ]
        let root = try XmlTreeReader.parse(Data(parts.joined().utf8)).root
        func removed(rels: Set<String> = [], footnotes: Set<Int> = [], endnotes: Set<Int> = [], comments: Set<Int> = []) -> BodyResourceIds {
            BodyResourceIds(relationships: rels, footnotes: footnotes, endnotes: endnotes, comments: comments)
        }
        XCTAssertTrue(BodySurgery.refers(root, to: removed(rels: ["rId4"])))
        XCTAssertTrue(BodySurgery.refers(root, to: removed(rels: ["rId7"])))
        XCTAssertTrue(BodySurgery.refers(root, to: removed(footnotes: [3])))
        XCTAssertTrue(BodySurgery.refers(root, to: removed(endnotes: [5])))
        XCTAssertTrue(BodySurgery.refers(root, to: removed(comments: [9])))
        XCTAssertFalse(BodySurgery.refers(root, to: removed(rels: ["rId5"], footnotes: [5], endnotes: [3], comments: [3])))
    }

    /// The dirty mark does not propagate to ancestors: editing a text node
    /// inside a paragraph leaves the paragraph node clean, with its source
    /// range intact. Only the recursive scan sees it; without it the save
    /// would copy the paragraph's read-time bytes over the edit.
    func testAnEditedTextNodeInsideAParagraphIsNeverOverwritten() throws {
        var doc = try readTreeBacked(body: drawingParagraph + textParagraph)
        defer { doc.close() }
        guard case .paragraph(let paragraph) = doc.body.children[1],
              let run = paragraph.xmlNode?.children.first(where: { $0.localName == "r" }),
              let text = run.children.first(where: { $0.localName == "t" }),
              let leaf = text.children.first
        else { return XCTFail("fixture") }
        leaf.textContent = "葉改"
        XCTAssertNotNil(paragraph.xmlNode?.sourceRange, "the paragraph node itself stays clean")
        doc.insertParagraph(Paragraph(text: "新段"), at: 0)
        let xml = try save(doc)
        XCTAssertTrue(xml.contains("葉改"), xml)
        XCTAssertFalse(xml.contains("文字段"), xml)
    }

    // MARK: - Tree edits and known limits

    /// Edit a tree-backed paragraph through its tree (the `text` setter writes
    /// the node and marks it dirty).
    private func editTextThroughTheTree(_ doc: inout WordDocument, at index: Int, to text: String) throws {
        guard case .paragraph(var paragraph) = doc.body.children[index] else { return XCTFail("fixture") }
        XCTAssertNotNil(paragraph.xmlNode, "the fixture must be tree-backed")
        paragraph.text = text
        doc.body.children[index] = .paragraph(paragraph)
    }

    /// A dirty tree means the read-time bytes may be stale, so a typed edit
    /// after a tree edit falls back to the full typed writer and an untouched
    /// paragraph's drawing is still lost. When this starts failing, the limit
    /// is gone: update the CHANGELOG.
    func testATypedEditAfterATreeEditFallsBack() throws {
        var doc = try readTreeBacked(body: drawingParagraph + textParagraph)
        defer { doc.close() }
        try editTextThroughTheTree(&doc, at: 1, to: "改過")
        doc.insertParagraph(Paragraph(text: "新段"), at: 0)
        let xml = try save(doc)
        XCTAssertTrue(xml.contains("改過"), xml)
        XCTAssertTrue(xml.contains("新段"), xml)
        XCTAssertEqual(drawingCount(xml), 0, "known limit changed:\n\(xml)")
    }

    /// A tree edit stays visible to later saves: the dirty mark is never
    /// cleared, so a typed edit and a second save never bring back the
    /// read-time bytes of the edited paragraph.
    func testATreeEditSurvivesALaterTypedEditAndSecondSave() throws {
        var doc = try readTreeBacked(body: drawingParagraph + textParagraph)
        defer { doc.close() }
        try editTextThroughTheTree(&doc, at: 1, to: "改過")
        doc.markPartDirty("word/document.xml")
        let first = try save(doc)
        XCTAssertTrue(first.contains("改過"), first)
        doc.insertParagraph(Paragraph(text: "新段"), at: 0)
        let xml = try save(doc)
        XCTAssertTrue(xml.contains("改過"), xml)
        XCTAssertFalse(xml.contains("文字段"), xml)
        XCTAssertTrue(xml.contains("新段"), xml)
    }

    /// #244: `updateParagraph` sets `runs`, which a tree-backed paragraph
    /// stores where its getter never reads, so the edit is lost (with or
    /// without surgery). Strict, so the fix shows up as an unexpected pass.
    func testUpdateParagraphOnATreeBackedParagraphKeepsTheEdit_knownBug() throws {
        var doc = try readTreeBacked(body: drawingParagraph + textParagraph)
        defer { doc.close() }
        try doc.updateParagraph(at: 1, text: "改過")
        let xml = try save(doc)
        XCTExpectFailure("PsychQuant/ooxml-swift#244 — tree-backed runs setter does not reach the tree", strict: true) {
            XCTAssertTrue(xml.contains("改過"), xml)
        }
        XCTAssertEqual(drawingCount(xml), 1, "the untouched drawing survives either way:\n\(xml)")
    }
}
