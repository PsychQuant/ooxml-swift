import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// #137 / #138 / #139 — `PackageInspector` scans with `XMLParser` (the parser
/// the reader already uses) instead of attribute regexes, and no serialization
/// path can trap on a duplicate relationship id.
///
/// The three defects shared one cause: the inspector answered questions about
/// XML without parsing XML. Downstream (PsychQuant/che-word-mcp#199) spent
/// three verify rounds trying to re-implement libxml2's attribute rules and a
/// regex's backtracking behaviour in a consumer, and each round a new shape got
/// through — entity, then zero-padded reference and literal whitespace, then
/// CRLF folding; count-balanced comment payloads, then nested openers. These
/// tests pin the properties that make that emulation unnecessary.
final class Issue137to139InspectorParserTests: XCTestCase {

    // MARK: - Fixtures

    private let imageType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    private let rNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    private let pkgNS = "http://schemas.openxmlformats.org/package/2006/relationships"

    private func zip(_ parts: [String: String]) throws -> Data {
        let archive = try Archive(accessMode: .create)
        for name in parts.keys.sorted() {
            let data = Data(parts[name]!.utf8)
            try archive.addEntry(with: name, type: .file, uncompressedSize: Int64(data.count),
                                 compressionMethod: .deflate) { position, size in
                let start = data.startIndex.advanced(by: Int(position))
                return data.subdata(in: start..<start.advanced(by: size))
            }
        }
        return try XCTUnwrap(archive.data)
    }

    private func rels(_ body: String) -> String {
        #"<Relationships xmlns="\#(pkgNS)">"# + body + "</Relationships>"
    }

    private func package(document: String, docRels: String, extra: [String: String] = [:], media: Bool = true) throws -> Data {
        var parts: [String: String] = [
            "word/document.xml": document,
            "word/_rels/document.xml.rels": rels(docRels),
        ]
        if media { parts["word/media/image1.png"] = "png" }
        for (k, v) in extra { parts[k] = v }
        return try zip(parts)
    }

    private func body(referencing id: String? = nil) -> String {
        let run = id.map { #"<w:p><w:r><w:drawing><a:blip r:embed="\#($0)"/></w:drawing></w:r></w:p>"# } ?? "<w:p/>"
        return #"<w:document xmlns:r="\#(rNS)"><w:body>"# + run + "</w:body></w:document>"
    }

    /// What `DocxReader` sees for the same attribute: NSXML, i.e. the same
    /// libxml2, reached through the API the reader uses.
    private func readerValue(ofAttribute name: String, inRels relsXML: String) throws -> String? {
        let doc = try XMLDocument(data: Data(relsXML.utf8))
        let element = try XCTUnwrap(doc.rootElement()?.elements(forName: "Relationship").first)
        return element.attribute(forName: name)?.stringValue
    }

    // MARK: - #137 · ids are what the reader's parser delivers

    func testDeclaredIdsEqualWhatTheReadersParserDelivers() throws {
        // Every spelling the downstream emulation failed on, in one table. The
        // expectation is never a literal: it is whatever NSXML says, so this
        // asserts equivalence rather than my guess about XML's rules.
        let spellings = [
            "plain":              "rId6",
            "decimal entity":     "rId&#54;",
            "zero-padded hex":    "rId&#x00000036;",
            "very long decimal":  "rId&#000000000000000000000054;",
            "literal tab":        "rId\t6",
            "literal LF":         "rId\n6",
            "literal CRLF":       "rId\r\n6",
            "predefined entity":  "rId&amp;6",
            "double encoded":     "rId&amp;#54;",
        ]
        for (label, spelling) in spellings {
            let relsXML = rels(#"<Relationship Id="\#(spelling)" Type="\#(imageType)" Target="media/image1.png"/>"#)
            let expected = try XCTUnwrap(readerValue(ofAttribute: "Id", inRels: relsXML), label)
            XCTAssertEqual(PackageInspector.imageRelationshipIds(inRels: relsXML), [expected], label)
        }
    }

    func testEntityEncodedOrphanIsReportedWithTheDecodedId() throws {
        let data = try package(document: body(), docRels: #"<Relationship Id="rId&#54;" Type="\#(imageType)" Target="media/image1.png"/>"#)
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(report.orphanImageRelationshipRefs, [ImageRelationshipRef(part: "word/document.xml", id: "rId6")])
        XCTAssertEqual(report.orphanImageRelationshipIds, ["rId6"])
    }

    func testEntityEncodedReferenceSatisfiesAPlainDeclaration() throws {
        // The reference side must be decoded too: before #137 only declarations
        // were compared, so this document reported a phantom orphan.
        let document = #"<w:document xmlns:r="\#(rNS)"><w:body><w:p><w:r><w:drawing><a:blip r:embed="rId&#x36;"/></w:drawing></w:r></w:p></w:body></w:document>"#
        let data = try package(document: document, docRels: #"<Relationship Id="rId6" Type="\#(imageType)" Target="media/image1.png"/>"#)
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertTrue(report.isConsistent, "orphans: \(report.orphanImageRelationshipRefs)")
        XCTAssertEqual(report.bodyDrawingCount, 1)
    }

    // MARK: - #138 · comments and CDATA are structure, and scanning is linear

    func testPathologicalCommentPayloadsFinishImmediatelyAndClaimNoOrphan() throws {
        let n = 20_000
        let payloads: [String: String] = [
            "unterminated openers":  String(repeating: "<!--", count: n),
            "balanced wrong order":  String(repeating: "-->", count: n) + String(repeating: "<!--", count: n),
            "nested then newline":   String(repeating: "<!--", count: n) + "\n-->",
        ]
        for (label, payload) in payloads {
            let data = try package(
                document: body(),
                docRels: #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#,
                extra: ["word/charts/chart1.xml": "<c:chartSpace/>",
                        "word/charts/_rels/chart1.xml.rels": rels(payload)])
            let started = Date()
            let report = try PackageInspector.imageConsistencyReport(of: data)
            XCTAssertLessThan(Date().timeIntervalSince(started), 1.0,
                              "\(label): pre-3.7.0 this took 35–60 s on a 2 KB package")
            // Whatever the parser makes of the payload, it must not invent a
            // chart-part orphan out of a part it could not read.
            XCTAssertFalse(report.orphanImageRelationshipRefs.contains { $0.part.hasPrefix("word/charts/") }, label)
        }
    }

    func testMultiLineCommentedOutRelationshipIsNotADeclaration() throws {
        let data = try package(document: body(), docRels: "<!--\n" + #"<Relationship Id="rId9" Type="\#(imageType)" Target="media/fake.png"/>"# + "\n-->")
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(report.imageRelationshipCount, 0)
        XCTAssertEqual(report.declaredImageRelationshipRefs, [])
        XCTAssertTrue(report.isConsistent)
    }

    func testCDATAIsTextNotMarkup() throws {
        // A `<Relationship>` inside CDATA is character data, and a literal
        // `<!--` inside CDATA opens nothing. The pre-3.7.0 regex saw both as
        // markup: one invented a declaration, the other disabled the scan.
        let cdata = "<![CDATA[" + #"<Relationship Id="rId9" Type="\#(imageType)" Target="media/fake.png"/> <!-- "# + "]]>"
        let document = #"<w:document xmlns:r="\#(rNS)"><w:body><w:p><w:t>\#(cdata)</w:t></w:p><w:p><w:r><w:drawing><a:blip r:embed="rId4"/></w:drawing></w:r></w:p></w:body></w:document>"#
        let data = try package(document: document, docRels: #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#)
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(report.imageRelationshipCount, 1)
        XCTAssertTrue(report.isConsistent, "orphans: \(report.orphanImageRelationshipRefs)")
        XCTAssertEqual(report.unparsableParts, [])
    }

    func testUnparsablePartProducesNoOrphansAndIsNamed() throws {
        // Unknown is not missing: a part XML rejects must not make its declared
        // relationships look like the #175 signature (the shape that refused
        // saves on legitimate files downstream).
        let data = try package(
            document: body(referencing: "rId4"),
            docRels: #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#,
            extra: ["word/header1.xml": "<w:hdr><w:p>",   // never closed
                    "word/_rels/header1.xml.rels": rels(#"<Relationship Id="rId9" Type="\#(imageType)" Target="media/image1.png"/>"#)])
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(report.unparsableParts, ["word/header1.xml"])
        XCTAssertEqual(report.orphanImageRelationshipRefs, [], "an unreadable part yields no verdict, not a guilty one")
        XCTAssertTrue(report.isConsistent)
        XCTAssertEqual(report.declaredImageRelationshipRefs.count, 2, "its declarations are still visible")
    }

    func testMissingPartStillYieldsOrphansAndIsNotCalledUnparsable() throws {
        // A part that is absent is not a part that could not be read: pre-3.7.0
        // reported its relationships as orphans, and that verdict is right —
        // the images are gone with the part. Only unreadable XML gets "no verdict".
        let data = try package(
            document: body(referencing: "rId4"),
            docRels: #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#,
            extra: ["word/_rels/header1.xml.rels": rels(#"<Relationship Id="rId9" Type="\#(imageType)" Target="media/image1.png"/>"#)])
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(report.unparsableParts, [], "absent is not unparsable")
        XCTAssertEqual(report.orphanImageRelationshipRefs, [ImageRelationshipRef(part: "word/header1.xml", id: "rId9")])
    }

    func testUnparsableRelsIsNamedAndDeclaresNothing() throws {
        let data = try package(document: body(), docRels: #"<Relationship Id="rId4" "#)   // truncated element
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(report.unparsableParts, ["word/_rels/document.xml.rels"])
        XCTAssertEqual(report.imageRelationshipCount, 0)
    }

    func testPartsDeclaringADocumentTypeAreRefused() throws {
        // OPC forbids DTDs in package parts. Switching to a parser made "what
        // does an entity mean inside an attribute value" answerable in more
        // than one way; refusing document types keeps the answer structural.
        let dtd = #"<?xml version="1.0"?><!DOCTYPE Relationships [<!ENTITY x "rId9">]>"#
        let data = try package(
            document: body(referencing: "rId4"),
            docRels: "",
            extra: ["word/header1.xml": #"<w:hdr xmlns:r="\#(rNS)"><w:p/></w:hdr>"#,
                    "word/_rels/header1.xml.rels": dtd + rels(#"<Relationship Id="&x;" Type="\#(imageType)" Target="media/image1.png"/>"#)])
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(report.unparsableParts, ["word/_rels/header1.xml.rels"])
        XCTAssertEqual(report.declaredImageRelationshipRefs, [], "a refused part declares nothing")
        XCTAssertTrue(report.isConsistent)
    }

    func testEntityExpansionBombIsRefusedImmediately() throws {
        var dtd = #"<?xml version="1.0"?><!DOCTYPE Relationships [<!ENTITY a "aaaaaaaaaa">"#
        for i in 1...9 {
            let prev = i == 1 ? "a" : "e\(i - 1)"
            dtd += "<!ENTITY e\(i) \"" + String(repeating: "&\(prev);", count: 10) + "\">"
        }
        dtd += "]>"
        let data = try package(document: body(), docRels: "", extra: ["word/header1.xml": "<w:hdr/>",
                                                                     "word/_rels/header1.xml.rels": dtd + rels("<!-- &e9; -->")])
        let started = Date()
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertLessThan(Date().timeIntervalSince(started), 1.0)
        XCTAssertEqual(report.unparsableParts, ["word/_rels/header1.xml.rels"])
    }

    // MARK: - #139 · a duplicate relationship id is refused, never fatal

    private func documentWithImages(_ ids: [String]) -> WordDocument {
        var doc = WordDocument()
        doc.images = ids.enumerated().map {
            ImageReference(id: $1, fileName: "image\($0 + 1).png", contentType: "image/png", data: Data("png".utf8))
        }
        return doc
    }

    func testDuplicateImageIdsThrowInsteadOfTrapping() throws {
        var thrown: Error?
        XCTAssertThrowsError(try DocxWriter.writeData(documentWithImages(["rId5", "rId5"]))) { thrown = $0 }
        let message = (thrown as? LocalizedError)?.errorDescription ?? String(describing: thrown)
        XCTAssertTrue(message.contains("rId5"), message)
        XCTAssertTrue(message.contains("twice"), message)
    }

    func testImageIdCollidingWithAFixedSlotThrowsAndNamesTheCause() throws {
        // A legitimate package may number an image `rId1`; this writer reserves
        // rId1–rId3 for styles / settings / fontTable. Until #140 that document
        // cannot be serialized — but it must say so, not trap.
        var thrown: Error?
        XCTAssertThrowsError(try DocxWriter.writeData(documentWithImages(["rId1"]))) { thrown = $0 }
        let message = (thrown as? LocalizedError)?.errorDescription ?? String(describing: thrown)
        XCTAssertTrue(message.contains("rId1"), message)
        XCTAssertTrue(message.contains("#140"), message)
    }

    func testDistinctIdsStillSerialize() throws {
        XCTAssertNoThrow(try DocxWriter.writeData(documentWithImages(["rId5", "rId6"])))
        XCTAssertNoThrow(try DocxWriter.writeData(WordDocument()))
    }

    func testMergeItselfCannotTrapOnDuplicates() throws {
        // Defence in depth: even called directly, the merge must return.
        let overlay = RelationshipsOverlay(originalRelsXML: rels(#"<Relationship Id="rId5" Type="\#(imageType)" Target="media/image1.png"/>"#))
        let dupes = [
            RelationshipDescriptor(id: "rId5", type: imageType, target: "media/a.png", targetMode: nil),
            RelationshipDescriptor(id: "rId5", type: imageType, target: "media/b.png", targetMode: nil),
        ]
        let xml = overlay.merge(typedRels: dupes, typedManagedTypes: [imageType])
        XCTAssertTrue(xml.contains("media/a.png"), "first declaration wins: \(xml)")
        XCTAssertFalse(xml.contains("media/b.png"), xml)
    }

    func testDuplicateDeclarationsAreNamedInTheReport() throws {
        let twice = #"<Relationship Id="rId5" Type="\#(imageType)" Target="media/image1.png"/>"#
            + #"<Relationship Id="rId&#53;" Type="\#(imageType)" Target="media/image2.png"/>"#
        let data = try package(document: body(referencing: "rId5"), docRels: twice)
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(report.duplicateRelationshipRefs, [ImageRelationshipRef(part: "word/document.xml", id: "rId5")],
                       "`rId5` and `rId&#53;` are one id once parsed")
        XCTAssertTrue(report.isConsistent)
    }

    // MARK: - declaredImageRelationshipRefs

    func testDeclaredRefsIncludeRelationshipsNoDocumentModelCanCarry() throws {
        // Missing media and external targets never reach `WordDocument.images`;
        // a consumer reconciling a listing against the package needs them named.
        let data = try package(
            document: body(referencing: "rId4"),
            docRels: #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#
                + #"<Relationship Id="rId77" Type="\#(imageType)" Target="media/missing.png"/>"#
                + #"<Relationship Id="rId88" Type="\#(imageType)" TargetMode="External" Target="https://example.com/x.png"/>"#)
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(report.declaredImageRelationshipRefs.map(\.id), ["rId4", "rId77", "rId88"])
        XCTAssertEqual(report.imageRelationshipCount, 3)
        XCTAssertEqual(report.orphanImageRelationshipRefs.map(\.id), ["rId77", "rId88"])
    }

    func testDrawingCountIgnoresNonImageDrawings() throws {
        // A chart is a `<w:drawing>` and not an image: the count is
        // informational, and a consumer must not read it as "has images".
        let document = #"<w:document xmlns:r="\#(rNS)"><w:body><w:p><w:r><w:drawing><wp:inline><a:graphic/></wp:inline></w:drawing></w:r></w:p></w:body></w:document>"#
        let data = try package(document: document, docRels: "", media: false)
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(report.bodyDrawingCount, 1)
        XCTAssertEqual(report.imageRelationshipCount, 0)
        XCTAssertEqual(report.mediaEntryCount, 0)
        XCTAssertTrue(report.isConsistent)
    }
}
