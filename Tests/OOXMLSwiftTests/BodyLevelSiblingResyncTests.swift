import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// #106 — `resyncBodyFromDocumentTree()` rebuilds `body.children` from the XML
/// tree but only re-types `<w:p>` and `<w:tbl>`. Everything else hits
/// `default: continue` and vanishes from the typed view.
///
/// The XML bytes are never lost — they stay in `xmlTrees` — so any test that
/// checks byte fidelity passes while the defect is present. What is lost is the
/// typed projection, and that is what downstream indexes against: che-word-mcp
/// reports insertion positions as offsets into `body.children` (its #61). Drop
/// two markers and every index after them is off by two.
///
/// This is the same shape as #104, at the call sites #104's fix did not touch.
/// #104 fixed `appendParagraph` by making it stop depending on the lossy
/// rebuild; the eight remaining callers still call it.
final class BodyLevelSiblingResyncTests: XCTestCase {

    // MARK: - Fixtures

    /// A docx whose `<w:body>` holds, in order: paragraph, bookmarkStart,
    /// bookmarkEnd, paragraph. The bookmark pair is a *sibling* of the
    /// paragraphs, not nested inside one — that is what makes it a body child.
    /// Real documents get this shape from a TOC anchor or any bookmark that
    /// spans more than one paragraph.
    private func bodyLevelSiblingDocx(extraBodyXML: String = "") throws -> URL {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" \
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" \
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body>
        <w:p w14:paraId="11111111" w14:textId="11111111"><w:r><w:t>First</w:t></w:r></w:p>
        <w:bookmarkStart w:id="1" w:name="bodyLevelMark"/>
        <w:bookmarkEnd w:id="1"/>
        \(extraBodyXML)
        <w:p w14:paraId="22222222" w14:textId="22222222"><w:r><w:t>Second</w:t></w:r></w:p>
        <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
        </w:body></w:document>
        """
        let contentTypes = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\
        <Default Extension="xml" ContentType="application/xml"/>\
        <Override PartName="/word/document.xml" \
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>
        """
        let rootRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\
        <Relationship Id="rId1" \
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" \
        Target="word/document.xml"/></Relationships>
        """
        let docRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
        """
        let parts: [String: Data] = [
            "[Content_Types].xml": Data(contentTypes.utf8),
            "_rels/.rels": Data(rootRels.utf8),
            "word/document.xml": Data(documentXML.utf8),
            "word/_rels/document.xml.rels": Data(docRels.utf8),
        ]
        let archive = try Archive(accessMode: .create)
        for name in parts.keys.sorted() {
            let data = parts[name]!
            try archive.addEntry(with: name, type: .file,
                                 uncompressedSize: Int64(data.count),
                                 compressionMethod: .deflate) { position, size in
                let start = data.startIndex.advanced(by: Int(position))
                return data.subdata(in: start..<start.advanced(by: size))
            }
        }
        guard let bytes = archive.data else {
            XCTFail("could not obtain in-memory archive bytes")
            return URL(fileURLWithPath: "/dev/null")
        }
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("body-level-siblings-\(UUID().uuidString).docx")
        try bytes.write(to: url)
        return url
    }

    private func kinds(_ doc: WordDocument) -> [String] {
        doc.body.children.map {
            switch $0 {
            case .paragraph:       return "p"
            case .table:           return "tbl"
            case .bookmarkMarker:  return "bookmarkMarker"
            case .rawBlockElement: return "rawBlockElement"
            default:               return "other"
            }
        }
    }

    private func firstAddressableParagraph(_ doc: WordDocument) -> ElementID? {
        for child in doc.body.children {
            if case .paragraph(let p) = child, let id = p.elementID { return id }
        }
        return nil
    }

    // MARK: - Reader baseline

    /// Establishes that the reader is NOT the lossy step. If this fails, the
    /// defect below would be a reading problem, and the diagnosis changes.
    func testReaderProducesBodyLevelBookmarkMarkers() throws {
        let url = try bodyLevelSiblingDocx()
        defer { try? FileManager.default.removeItem(at: url) }

        let doc = try DocxReader.read(from: url, wireTreeBackedViews: true)
        XCTAssertEqual(kinds(doc), ["p", "bookmarkMarker", "bookmarkMarker", "p"],
                       "the reader SHALL surface body-level bookmark markers as typed body children")
    }

    // MARK: - The defect

    /// A public typed setter that ends in `resyncBodyFromDocumentTree()` must
    /// not shrink the typed view. `setParagraphText` edits one paragraph's
    /// text; nothing about that operation licenses deleting the document's
    /// bookmarks from the caller's view of the body.
    func testBodyLevelBookmarkMarkersSurviveTypedSetter() throws {
        let url = try bodyLevelSiblingDocx()
        defer { try? FileManager.default.removeItem(at: url) }

        var doc = try DocxReader.read(from: url, wireTreeBackedViews: true)
        let before = kinds(doc)
        XCTAssertEqual(before, ["p", "bookmarkMarker", "bookmarkMarker", "p"],
                       "precondition: the markers are present before the setter runs")

        let id = try XCTUnwrap(firstAddressableParagraph(doc),
                              "fixture must expose an addressable paragraph (w14:paraId)")
        try doc.setParagraphText(id: id, "probe text")

        XCTAssertEqual(kinds(doc), before,
                       "setParagraphText SHALL NOT drop non-paragraph body children (#106)")
    }

    /// The bytes were never the casualty. Pinning this separates "the typed
    /// view lost it" from "the file lost it" — without this, a future reader of
    /// the test above could conclude the document was being corrupted on disk.
    func testBookmarkBytesSurviveEvenWhileTypedViewLosesThem() throws {
        let url = try bodyLevelSiblingDocx()
        defer { try? FileManager.default.removeItem(at: url) }

        var doc = try DocxReader.read(from: url, wireTreeBackedViews: true)
        let id = try XCTUnwrap(firstAddressableParagraph(doc))
        try doc.setParagraphText(id: id, "probe text")

        let out = FileManager.default.temporaryDirectory
            .appendingPathComponent("resync-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: out) }
        try DocxWriter.write(doc, to: out)

        let xml = String(decoding: try CorpusFixtureBuilder.readPart("word/document.xml", from: out),
                         as: UTF8.self)
        XCTAssertTrue(xml.contains("bookmarkStart"), "the saved part SHALL still carry the bookmark")
        XCTAssertTrue(xml.contains("bodyLevelMark"), "including its name")
    }

    /// The part of #106 that is NOT fixed: body children with no typed
    /// representation (vendor extensions, `<w:sdt>`, other EG_BlockLevelElts)
    /// still disappear, because re-typing them needs a single `XmlNode`
    /// serialized back to XML — and `XmlTreeWriter.emitElement` is private and
    /// wants `sourceBytes`/`dirtyMap`. The reader can do it (it holds a
    /// Foundation `XMLElement` and calls `.xmlString`); resync cannot.
    ///
    /// Written as an expected failure rather than an assertion of the broken
    /// behaviour: it states the contract we want, stays quiet while the gap is
    /// open, and fails loudly the moment someone closes it — so #106 cannot be
    /// marked done on the strength of the bookmark half alone.
    func testUnknownBodyLevelElementSurvivesTypedSetter() throws {
        XCTExpectFailure("#106 residue: needs node-level XML serialization before resync can re-type this")

        let url = try bodyLevelSiblingDocx(
            extraBodyXML: "<w:moveFromRangeStart w:id=\"7\" w:name=\"mv\"/>")
        defer { try? FileManager.default.removeItem(at: url) }

        var doc = try DocxReader.read(from: url, wireTreeBackedViews: true)
        let before = kinds(doc)
        let id = try XCTUnwrap(firstAddressableParagraph(doc))
        try doc.setParagraphText(id: id, "probe text")

        XCTAssertEqual(kinds(doc), before,
                       "resync SHALL preserve body children it has no typed case for")
    }
}
