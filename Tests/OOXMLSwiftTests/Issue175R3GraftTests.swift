import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// macdoc#175 verify round 3: non-representable appends are grafted into the
/// live document.xml tree instead of re-serializing the whole part from the
/// (lossy) typed model; the run layer gets the same tree-backed guard as the
/// paragraph layer; PackageInspector handles commented declarations, '>' in
/// quoted values and nested parts.
final class Issue175R3GraftTests: XCTestCase {

    private let wNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    private let w14NS = "http://schemas.microsoft.com/office/word/2010/wordml"
    private let rNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    private let imageType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    private let onePixelPNG = Data(base64Encoded: "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==")!

    private func zip(_ parts: [String: Data]) throws -> Data {
        let archive = try Archive(accessMode: .create)
        for name in parts.keys.sorted() {
            let data = parts[name]!
            try archive.addEntry(with: name, type: .file, uncompressedSize: Int64(data.count), compressionMethod: .deflate) { position, size in
                let start = data.startIndex.advanced(by: Int(position)); return data.subdata(in: start..<start.advanced(by: size))
            }
        }
        return try XCTUnwrap(archive.data)
    }

    /// A document whose body carries an odd but legal shape the typed model
    /// does not represent: a body-level `<w:customXml>` block wrapping a
    /// paragraph, plus rsid-stamped runs. Byte preservation of this region is
    /// what the graft path guarantees and the typed-dirty path cannot.
    private let oddBody = #"<w:customXml w:element="oddblock"><w:p w14:paraId="0A0A0A0A"><w:r w:rsidR="00AB12CD"><w:t>kept-verbatim</w:t></w:r></w:p></w:customXml><w:p w14:paraId="11111111" w14:textId="11111111"><w:r><w:t>One</w:t></w:r></w:p>"#

    private func docx(body: String, docRels: String = "", withImage: Bool = false, minimalRoot: Bool = false, crlfProlog: Bool = false, foreignW: Bool = false) throws -> URL {
        var parts: [String: Data] = [
            "[Content_Types].xml": Data("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>
            """.utf8),
            "_rels/.rels": Data(#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.utf8),
            "word/document.xml": Data(("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + (crlfProlog ? "\r\n" : "\n") + "<w:document xmlns:w=\"\(foreignW ? "urn:vendor:not-wordprocessingml" : wNS)\"" + (minimalRoot ? "" : " xmlns:w14=\"\(w14NS)\" xmlns:r=\"\(rNS)\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"w14\"") + "><w:body>\(body)</w:body></w:document>" + (crlfProlog ? "\r\n" : "")).utf8),
            "word/_rels/document.xml.rels": Data(#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#.utf8 + Data(docRels.utf8) + Data("</Relationships>".utf8)),
        ]
        if withImage { parts["word/media/image1.png"] = onePixelPNG }
        let url = FileManager.default.temporaryDirectory.appendingPathComponent("i175r3-\(UUID().uuidString).docx")
        try zip(parts).write(to: url)
        return url
    }

    private func part(_ name: String, of data: Data) throws -> String {
        let archive = try Archive(data: data, accessMode: .read)
        let entry = try XCTUnwrap(archive[name]); var d = Data(); _ = try archive.extract(entry) { d.append($0) }
        return String(decoding: d, as: UTF8.self)
    }

    private func pngPath() throws -> String {
        let url = FileManager.default.temporaryDirectory.appendingPathComponent("i175r3-\(UUID().uuidString).png")
        try onePixelPNG.write(to: url); return url.path
    }

    // MARK: - Graft path

    func testAppendImageKeepsDocumentTreeFreshAndBytesVerbatim() throws {
        let url = try docx(body: oddBody); defer { try? FileManager.default.removeItem(at: url) }
        let png = try pngPath(); defer { try? FileManager.default.removeItem(atPath: png) }
        var doc = try DocxReader.read(from: url)
        _ = try doc.insertImage(path: png, widthPx: 10, heightPx: 10, at: nil)
        XCTAssertTrue(doc.treeFreshParts.contains("word/document.xml"), "append-image must not fall back to typed re-serialization of the whole part")
        let xml = try part("word/document.xml", of: try DocxWriter.writeData(doc))
        XCTAssertTrue(xml.contains(#"<w:customXml w:element="oddblock"><w:p w14:paraId="0A0A0A0A"><w:r w:rsidR="00AB12CD"><w:t>kept-verbatim</w:t></w:r></w:p></w:customXml>"#), "pre-existing bytes must survive verbatim")
        XCTAssertEqual(xml.components(separatedBy: "<w:drawing").count - 1, 1, "the appended image must be in the body")
        XCTAssertTrue(xml.contains("<w:sectPr") == false || xml.range(of: "<w:drawing")!.lowerBound < xml.range(of: "<w:sectPr")!.lowerBound)
        XCTAssertTrue(try PackageInspector.imageConsistencyReport(of: try DocxWriter.writeData(doc)).isConsistent)
    }

    func testAppendImageOnTreeBackedDocumentKeepsExistingDrawing() throws {
        // #129's append case: with the graft path there is no typed re-serialization, so the existing drawing survives.
        let existing = #"<w:p w14:paraId="22222222"><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill><a:blip r:embed="rId4"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#
        let url = try docx(body: existing, docRels: #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#, withImage: true)
        defer { try? FileManager.default.removeItem(at: url) }
        let png = try pngPath(); defer { try? FileManager.default.removeItem(atPath: png) }
        var doc = try DocxReader.read(from: url, wireTreeBackedViews: true)
        _ = try doc.insertImage(path: png, widthPx: 10, heightPx: 10, at: nil)
        let saved = try DocxWriter.writeData(doc)
        XCTAssertEqual(try part("word/document.xml", of: saved).components(separatedBy: "<w:drawing").count - 1, 2)
        XCTAssertTrue(try PackageInspector.imageConsistencyReport(of: saved).isConsistent)
    }

    func testTreeBackedRunIsNotFastPathed() throws {
        let url = try docx(body: #"<w:p w14:paraId="11111111"><w:r><w:t>One</w:t></w:r></w:p>"#); defer { try? FileManager.default.removeItem(at: url) }
        var doc = try DocxReader.read(from: url)
        let runTree = try XmlTreeReader.parse(Data("<w:r xmlns:w=\"\(wNS)\"><w:rPr><w:b/></w:rPr><w:t>tree-run</w:t></w:r>".utf8))
        let p = Paragraph(runs: [Run(xmlNode: runTree.root)])
        let before = doc.operationLog.entries.count
        doc.appendParagraph(p)
        XCTAssertEqual(doc.operationLog.entries.count, before, "a tree-backed run must not take the op fast path (R2 logic N1)")
        let xml = try part("word/document.xml", of: try DocxWriter.writeData(doc))
        XCTAssertTrue(xml.contains("tree-run"))
    }

    /// R3 codex F1: a document whose root declares only `xmlns:w` must still
    /// produce well-formed XML after a grafted image paragraph (w14/wp/a/pic).
    func testGraftDeclaresMissingNamespacePrefixesOnRoot() throws {
        let url = try docx(body: #"<w:p><w:r><w:t>One</w:t></w:r></w:p>"#, minimalRoot: true); defer { try? FileManager.default.removeItem(at: url) }
        let png = try pngPath(); defer { try? FileManager.default.removeItem(atPath: png) }
        var doc = try DocxReader.read(from: url)
        _ = try doc.insertImage(path: png, widthPx: 10, heightPx: 10, at: nil)
        let data = try DocxWriter.writeData(doc)
        let xmlData = Data(try part("word/document.xml", of: data).utf8)
        XCTAssertNoThrow(try XMLDocument(data: xmlData, options: []), "grafted output must be namespace-well-formed")
        let xml = String(decoding: xmlData, as: UTF8.self)
        for p in ["w14", "wp", "a", "pic"] { XCTAssertTrue(xml.contains("xmlns:\(p)=\""), "grafted node must declare \(p)") }
        XCTAssertTrue(doc.treeFreshParts.contains("word/document.xml"), "still the graft path, not typed fallback")
        // R3 logic N1: the ROOT must stay clean — a dirty root rewrites the prolog.
        let root = try XCTUnwrap(doc.xmlTrees["word/document.xml"]?.root)
        XCTAssertFalse(root.isDirty, "root must not be mutated by the graft")
        XCTAssertTrue(xml.hasPrefix("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:document xmlns:w=\""), "root open tag byte-identical: \(xml.prefix(120))")
    }

    /// R3 logic N1 (regression in 3.6.3): a CRLF prolog must survive a graft byte-for-byte.
    func testGraftPreservesCRLFPrologAndEpilog() throws {
        let url = try docx(body: #"<w:p><w:r><w:t>One</w:t></w:r></w:p>"#, minimalRoot: true, crlfProlog: true); defer { try? FileManager.default.removeItem(at: url) }
        let png = try pngPath(); defer { try? FileManager.default.removeItem(atPath: png) }
        let original = Data(try part("word/document.xml", of: Data(contentsOf: url)).utf8)
        var doc = try DocxReader.read(from: url)
        _ = try doc.insertImage(path: png, widthPx: 10, heightPx: 10, at: nil)
        let out = Data(try part("word/document.xml", of: try DocxWriter.writeData(doc)).utf8)
        XCTAssertEqual(out.prefix(70), original.prefix(70), "prolog (CRLF) and root open tag must be untouched")
        XCTAssertEqual(out.suffix(30), original.suffix(30), "epilog must be untouched")
    }

    // MARK: - Inspector hardening

    private func pkg(_ parts: [String: String]) throws -> Data {
        var d: [String: Data] = [:]; for (k, v) in parts { d[k] = Data(v.utf8) }; d["word/media/image1.png"] = onePixelPNG
        return try zip(d)
    }

    func testCommentedOutDeclarationIsNotADeclaration() throws {
        let data = try pkg([
            "word/document.xml": "<w:document><w:body><w:p/></w:body></w:document>",
            "word/_rels/document.xml.rels": #"<Relationships><!-- <Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/> --></Relationships>"#,
        ])
        let r = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(r.imageRelationshipCount, 0); XCTAssertTrue(r.isConsistent)
    }

    func testGreaterThanInsideQuotedTargetDoesNotHideTheRelationship() throws {
        let data = try pkg([
            "word/document.xml": "<w:document><w:body><w:p/></w:body></w:document>",
            "word/_rels/document.xml.rels": #"<Relationships><Relationship Target="media/a>b.png" Type="\#(imageType)" Id="rId4"/></Relationships>"#,
        ])
        let r = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(r.orphanImageRelationshipIds, ["rId4"], "R2 logic: a '>' in a quoted value used to make the whole element invisible (fail-open)")
    }

    func testNestedChartPartOrphanIsDetected() throws {
        let data = try pkg([
            "word/document.xml": "<w:document><w:body><w:p/></w:body></w:document>",
            "word/_rels/document.xml.rels": "<Relationships/>",
            "word/charts/chart1.xml": "<c:chartSpace/>",
            "word/charts/_rels/chart1.xml.rels": #"<Relationships><Relationship Id="rId2" Type="\#(imageType)" Target="../media/image1.png"/></Relationships>"#,   // OPC: <dir>/_rels/<name>.rels
        ])
        let r = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(r.orphanImageRelationshipRefs, [ImageRelationshipRef(part: "word/charts/chart1.xml", id: "rId2")])
    }
}
