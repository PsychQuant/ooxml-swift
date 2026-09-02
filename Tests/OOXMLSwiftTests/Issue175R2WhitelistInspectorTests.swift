import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// macdoc#175 verify round 2: the R1 6-AI verify proved the 3.6.0 whitelist
/// was not exhaustive at the `Paragraph` layer and that `PackageInspector`
/// scoped relationship ids globally. Each probe here appends a paragraph
/// carrying one previously-unguarded field, saves, and asserts the content
/// is in `word/document.xml` — on 3.6.0 every one of these was LOST.
final class Issue175R2WhitelistInspectorTests: XCTestCase {

    // MARK: - Fixture

    private func zip(_ parts: [String: String]) throws -> Data {
        let archive = try Archive(accessMode: .create)
        for name in parts.keys.sorted() {
            let data = Data(parts[name]!.utf8)
            try archive.addEntry(with: name, type: .file,
                                 uncompressedSize: Int64(data.count),
                                 compressionMethod: .deflate) { position, size in
                let start = data.startIndex.advanced(by: Int(position))
                return data.subdata(in: start..<start.advanced(by: size))
            }
        }
        return try XCTUnwrap(archive.data)
    }

    private let wNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    private let w14NS = "http://schemas.microsoft.com/office/word/2010/wordml"
    private let rNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    private func minimalDocx(body: String, docRels: String = "") throws -> URL {
        let parts: [String: String] = [
            "[Content_Types].xml": """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\
            <Default Extension="xml" ContentType="application/xml"/>\
            <Default Extension="png" ContentType="image/png"/>\
            <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\
            </Types>
            """,
            "_rels/.rels": """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>\
            </Relationships>
            """,
            "word/document.xml": """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="\(wNS)" xmlns:w14="\(w14NS)" xmlns:r="\(rNS)" \
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body>\(body)</w:body></w:document>
            """,
            "word/_rels/document.xml.rels": """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\(docRels)</Relationships>
            """,
        ]
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue175r2-\(UUID().uuidString).docx")
        try zip(parts).write(to: url)
        return url
    }

    private func seedDocx() throws -> URL {
        try minimalDocx(body: #"<w:p w14:paraId="11111111" w14:textId="11111111"><w:r><w:t>One</w:t></w:r></w:p>"#)
    }

    private func documentXML(of saved: Data) throws -> String {
        let archive = try Archive(data: saved, accessMode: .read)
        let entry = try XCTUnwrap(archive["word/document.xml"])
        var data = Data()
        _ = try archive.extract(entry) { data.append($0) }
        return String(decoding: data, as: UTF8.self)
    }

    /// Append `paragraph` to a disk-opened document, save, return document.xml.
    private func appendAndSave(_ paragraph: Paragraph) throws -> String {
        let url = try seedDocx()
        defer { try? FileManager.default.removeItem(at: url) }
        var doc = try DocxReader.read(from: url)
        doc.appendParagraph(paragraph)
        return try documentXML(of: try DocxWriter.writeData(doc))
    }

    // MARK: - Whitelist probes (each LOST on 3.6.0)

    func testUnrecognizedChildSurvivesAppend() throws {
        var p = Paragraph(runs: [Run(text: "x")])
        p.unrecognizedChildren = [UnrecognizedChild(name: "w:probe", rawXML: #"<w:probe w:val="probe-unrecognized"/>"#)]
        XCTAssertTrue(try appendAndSave(p).contains("probe-unrecognized"))
    }

    func testSmartTagTextSurvivesAppend() throws {
        var p = Paragraph(runs: [Run(text: "x")])
        p.smartTags = [SmartTagBlock(rawXML: #"<w:smartTag w:element="probe"><w:r><w:t>probe-smarttag</w:t></w:r></w:smartTag>"#)]
        XCTAssertTrue(try appendAndSave(p).contains("probe-smarttag"))
    }

    func testCustomXmlBlockTextSurvivesAppend() throws {
        var p = Paragraph(runs: [Run(text: "x")])
        p.customXmlBlocks = [CustomXmlBlock(rawXML: #"<w:customXml w:element="probe"><w:r><w:t>probe-customxml</w:t></w:r></w:customXml>"#)]
        XCTAssertTrue(try appendAndSave(p).contains("probe-customxml"))
    }

    func testCommentRangeMarkerSurvivesAppend() throws {
        var p = Paragraph(runs: [Run(text: "x")])
        p.commentRangeMarkers = [CommentRangeMarker(kind: .start, id: 42), CommentRangeMarker(kind: .end, id: 42)]
        let xml = try appendAndSave(p)
        XCTAssertTrue(xml.contains("commentRangeStart"), xml.suffix(300).description)
    }

    func testProofErrorMarkerSurvivesAppend() throws {
        var p = Paragraph(runs: [Run(text: "x")])
        p.proofErrorMarkers = [ProofErrorMarker(type: .spellStart), ProofErrorMarker(type: .spellEnd)]
        XCTAssertTrue(try appendAndSave(p).contains("proofErr"))
    }

    func testTextIdIsProjectedNotDropped() throws {
        var p = Paragraph(runs: [Run(text: "x")])
        p.w14TextId = "5EADBEEF"
        XCTAssertTrue(try appendAndSave(p).contains(#"w14:textId="5EADBEEF""#))
    }

    func testLeadingWhitespaceKeepsXmlSpacePreserve() throws {
        let p = Paragraph(runs: [Run(text: "  padded")])
        let xml = try appendAndSave(p)
        XCTAssertTrue(xml.contains(#"<w:t xml:space="preserve">  padded</w:t>"#), String(xml.suffix(300)))
    }

    func testComplexScriptFontSurvivesAppend() throws {
        var rp = RunProperties()
        rp.rFonts = RFontsProperties(ascii: "Calibri", cs: "Probe CS Font")
        let p = Paragraph(runs: [Run(text: "x", properties: rp)])
        XCTAssertTrue(try appendAndSave(p).contains("Probe CS Font"))
    }

    /// Plain text still takes the op fast path (the whole point of the whitelist).
    func testPlainAppendStillEmitsOp() throws {
        let url = try seedDocx()
        defer { try? FileManager.default.removeItem(at: url) }
        var doc = try DocxReader.read(from: url)
        let before = doc.operationLog.entries.count
        doc.appendParagraph(Paragraph(runs: [Run(text: "plain")]))
        XCTAssertGreaterThan(doc.operationLog.entries.count, before, "plain text append must still be op-emitted")
    }

    // MARK: - Known limitation, documented not hidden (PsychQuant/ooxml-swift#129)

    /// Tree-backed views + typed re-serialization drop existing `<w:drawing>`
    /// (regression lens, 13/27 real documents). Not introduced by #128 but
    /// re-routed onto by it. Expected-failure until #129 lands; non-strict
    /// because a minimal fixture may not reproduce what real Word files do.
    func testTreeBackedTypedDirtyKeepsExistingDrawing_knownLimitation() throws {
        let drawing = #"<w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill><a:blip r:embed="rId4"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"#
        let url = try minimalDocx(
            body: #"<w:p w14:paraId="11111111" w14:textId="11111111"><w:r>"# + drawing + #"</w:r></w:p>"#,
            docRels: #"<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>"#)
        defer { try? FileManager.default.removeItem(at: url) }
        var doc = try DocxReader.read(from: url, wireTreeBackedViews: true)
        try doc.insertParagraph(Paragraph(runs: [Run(text: "typed-dirty edit")]), at: 0)
        let xml = try documentXML(of: try DocxWriter.writeData(doc))
        XCTExpectFailure("PsychQuant/ooxml-swift#129 — tree-backed typed re-serialization drops existing drawings", strict: true) {
            XCTAssertEqual(xml.components(separatedBy: "<w:drawing").count - 1, 1)
        }
    }

    // MARK: - PackageInspector: per-part relationship scoping

    private let imageType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

    private func package(document: String, docRels: String, header: String? = nil, headerRels: String? = nil) throws -> Data {
        var parts: [String: String] = [
            "word/document.xml": document,
            "word/_rels/document.xml.rels": #"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"# + docRels + "</Relationships>",
            "word/media/image1.png": "png",
        ]
        if let header { parts["word/header1.xml"] = header }
        if let headerRels { parts["word/_rels/header1.xml.rels"] = #"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"# + headerRels + "</Relationships>" }
        return try zip(parts)
    }

    private func rel(_ id: String, type: String? = nil) -> String {
        #"<Relationship Id="\#(id)" Type="\#(type ?? imageType)" Target="media/image1.png"/>"#
    }

    func testHeaderOrphanIsDetected() throws {
        let data = try package(
            document: #"<w:document xmlns:r="\#(rNS)"><w:body><w:p><w:r><w:drawing><a:blip r:embed="rId4"/></w:drawing></w:r></w:p></w:body></w:document>"#,
            docRels: rel("rId4"),
            header: #"<w:hdr xmlns:r="\#(rNS)"><w:p><w:r><w:t>no image here</w:t></w:r></w:p></w:hdr>"#,
            headerRels: rel("rId9"))
        let r = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertFalse(r.isConsistent)
        XCTAssertEqual(r.orphanImageRelationshipRefs, [ImageRelationshipRef(part: "word/header1.xml", id: "rId9")])
        XCTAssertEqual(r.orphanImageRelationshipIds, [], "document-part view must not include header orphans")
        XCTAssertEqual(r.imageRelationshipCount, 2)
    }

    func testCrossPartAliasNoLongerMasksDocumentOrphan() throws {
        // R1 security PoC A: header references its OWN rId4; document's rId4 is an orphan.
        let data = try package(
            document: #"<w:document xmlns:r="\#(rNS)"><w:body><w:p><w:r><w:t>text only</w:t></w:r></w:p></w:body></w:document>"#,
            docRels: rel("rId4"),
            header: #"<w:hdr xmlns:r="\#(rNS)"><w:p><w:r><w:drawing><a:blip r:embed="rId4"/></w:drawing></w:r></w:p></w:hdr>"#,
            headerRels: rel("rId4"))
        let r = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(r.orphanImageRelationshipRefs, [ImageRelationshipRef(part: "word/document.xml", id: "rId4")])
        XCTAssertEqual(r.orphanImageRelationshipIds, ["rId4"])
    }

    func testCommentedOutReferenceDoesNotCount() throws {
        // R1 security PoC B.
        let data = try package(
            document: #"<w:document xmlns:r="\#(rNS)"><w:body><!-- r:embed="rId4" --><w:p/></w:body></w:document>"#,
            docRels: rel("rId4"))
        XCTAssertFalse(try PackageInspector.imageConsistencyReport(of: data).isConsistent)
    }

    func testSingleQuotesAndForeignPrefixAreAccepted() throws {
        let data = try package(
            document: "<w:document xmlns:rel='\(rNS)'><w:body><w:p><w:r><w:drawing><a:blip rel:embed='rId4'/></w:drawing></w:r></w:p></w:body></w:document>",
            docRels: "<Relationship Id='rId4' Type='\(imageType)' Target='media/image1.png'/>")
        let r = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertTrue(r.isConsistent, "orphans: \(r.orphanImageRelationshipRefs)")
        XCTAssertEqual(r.imageRelationshipCount, 1)
    }

    func testTypeMatchIsSuffixNotSubstring() throws {
        let data = try package(
            document: #"<w:document><w:body><w:p/></w:body></w:document>"#,
            docRels: rel("rId4", type: imageType + "Extended"))
        let r = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(r.imageRelationshipCount, 0)
        XCTAssertTrue(r.isConsistent)
    }
}
