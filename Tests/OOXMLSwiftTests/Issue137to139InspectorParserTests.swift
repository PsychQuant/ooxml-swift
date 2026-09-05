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
    private let wNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    private let aNS = "http://schemas.openxmlformats.org/drawingml/2006/main"
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

    /// Namespaces declared like real Word output: with namespace processing
    /// on, an undeclared prefix is a parse error for the inspector exactly as
    /// it is for the reader (verify R2 DA).
    private func body(referencing id: String? = nil) -> String {
        let run = id.map { #"<w:p><w:r><w:drawing><a:blip r:embed="\#($0)"/></w:drawing></w:r></w:p>"# } ?? "<w:p/>"
        return #"<w:document xmlns:w="\#(wNS)" xmlns:a="\#(aNS)" xmlns:r="\#(rNS)"><w:body>"# + run + "</w:body></w:document>"
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

    func testAttributeNamesAreMatchedExactlyLikeTheReader() throws {
        // verify R1 logic F2 / codex 1: `xmlns:Id`, `r:Id`, `p:Type` are not the
        // attributes DocxReader reads with attribute(forName:). With `Id` and
        // `r:Id` both present the rc answered by dictionary order — run it
        // several times so a flaky right answer cannot pass.
        let fakeOnly = rels(#"<Relationship xmlns:Id="urn:x" r:Id="rIdFAKE" xmlns:r="\#(rNS)" Type="\#(imageType)" Target="media/image1.png"/>"#)
        XCTAssertEqual(PackageInspector.imageRelationshipIds(inRels: fakeOnly), [])
        XCTAssertEqual(try readerValue(ofAttribute: "Id", inRels: fakeOnly), nil)
        let fakeType = rels(#"<Relationship Id="rId4" p:Type="\#(imageType)" xmlns:p="urn:p" Target="media/image1.png"/>"#)
        XCTAssertEqual(PackageInspector.imageRelationshipIds(inRels: fakeType), [], "p:Type is not Type")
        let both = rels(#"<Relationship Id="rIdREAL" r:Id="rIdFAKE" xmlns:r="\#(rNS)" Type="\#(imageType)" Target="media/image1.png"/>"#)
        for _ in 0..<20 {
            XCTAssertEqual(PackageInspector.imageRelationshipIds(inRels: both), ["rIdREAL"])
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
        let document = #"<w:document xmlns:w="\#(wNS)" xmlns:a="\#(aNS)" xmlns:r="\#(rNS)"><w:body><w:p><w:r><w:drawing><a:blip r:embed="rId&#x36;"/></w:drawing></w:r></w:p></w:body></w:document>"#
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
                extra: ["word/charts/chart1.xml": #"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#,
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

    func testCommentShapesThatDegradeLibxml2AreRefusedBeforeParsing() throws {
        // verify R1 requirements R1: replacing the regex moved the quadratic
        // into libxml2's error recovery — `--` inside a comment. 4.6 KB of
        // package, 82 s. These are refused by the linear pre-check instead.
        let n = 800_000
        let payloads: [String: String] = [
            "nested openers, newline, one close": String(repeating: "<!--", count: n) + "\n-->",
            "one comment full of --":             "<!--" + String(repeating: "--", count: n) + "\n-->",
            "unterminated CDATA":                 "<![CDATA[" + String(repeating: "x", count: n),
        ]
        for (label, payload) in payloads {
            let started = Date()
            XCTAssertNotNil(PackageInspector.linearPrecheckFailure(Data(payload.utf8)), label)
            XCTAssertLessThan(Date().timeIntervalSince(started), 0.5, label)
            let data = try package(document: body(referencing: "rId4"),
                                   docRels: #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#,
                                   extra: ["word/charts/chart1.xml": #"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#, "word/charts/_rels/chart1.xml.rels": rels(payload)])
            let t0 = Date()
            let report = try PackageInspector.imageConsistencyReport(of: data)
            XCTAssertLessThan(Date().timeIntervalSince(t0), 1.0, label)
            XCTAssertEqual(report.unparsableParts, ["word/charts/_rels/chart1.xml.rels"], label)
            XCTAssertFalse(report.isConsistent, label)
        }
        // …and a benign comment of the same size is parsed normally.
        let benign = "<!-- " + String(repeating: "x", count: n) + " -->"
        XCTAssertNil(PackageInspector.linearPrecheckFailure(Data((benign + rels("")).utf8)))
        XCTAssertEqual(PackageInspector.scanRels(Data((benign + rels(#"<Relationship Id="rId4" Type="\#(imageType)" Target="t"/>"#)).utf8)).imageIds, ["rId4"])
    }

    func testOverWideStartTagIsRefusedAndOrdinaryOnesAreNot() throws {
        // verify R1 security S2: libxml2 is quadratic in per-element attribute count.
        let wide = "<r " + (1...(PackageInspector.maxAttributesPerElement + 1)).map { "a\($0)=\"v\"" }.joined(separator: " ") + "/>"
        XCTAssertNotNil(PackageInspector.linearPrecheckFailure(Data(wide.utf8)))
        let ordinary = "<r " + (1...200).map { "a\($0)=\"v\"" }.joined(separator: " ") + "/>"
        XCTAssertNil(PackageInspector.linearPrecheckFailure(Data(ordinary.utf8)))
        // many elements, same total attribute count: linear, allowed
        let many = "<r>" + String(repeating: "<c " + (1...10).map { "a\($0)=\"v\"" }.joined(separator: " ") + "/>", count: 2_000) + "</r>"
        XCTAssertNil(PackageInspector.linearPrecheckFailure(Data(many.utf8)))
        // an attribute VALUE may contain `=` and `>` freely
        XCTAssertNil(PackageInspector.linearPrecheckFailure(Data(#"<r a="x=y>z" b='p=q'/>"#.utf8)))
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
        let document = #"<w:document xmlns:w="\#(wNS)" xmlns:a="\#(aNS)" xmlns:r="\#(rNS)"><w:body><w:p><w:t>\#(cdata)</w:t></w:p><w:p><w:r><w:drawing><a:blip r:embed="rId4"/></w:drawing></w:r></w:p></w:body></w:document>"#
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
            extra: ["word/header1.xml": #"<w:hdr xmlns:w="\#(wNS)"><w:p>"#,   // never closed
                    "word/_rels/header1.xml.rels": rels(#"<Relationship Id="rId9" Type="\#(imageType)" Target="media/image1.png"/>"#)])
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(report.unparsableParts, ["word/header1.xml"])
        XCTAssertEqual(report.orphanImageRelationshipRefs, [], "an unreadable part yields no verdict, not a guilty one")
        XCTAssertFalse(report.isConsistent, "…and no verdict is not a verdict of consistency (verify R1 security S1)")
        XCTAssertEqual(report.declaredImageRelationshipRefs.count, 2, "its declarations are still visible")
    }

    func testCorruptingAnUnrelatedPartCannotHideARealOrphan() throws {
        // verify R1 security S1 / codex 3: one appended `<` in a chart part
        // made 3.7.0-rc report isConsistent == true while 3.6.4 reported the
        // document-part orphan. Unreadable must fail closed.
        let data = try package(
            document: body(),   // rId4 declared, never referenced → real orphan
            docRels: #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#,
            extra: ["word/charts/chart1.xml": #"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/><"#,
                    "word/charts/_rels/chart1.xml.rels": rels("")])
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(report.orphanImageRelationshipRefs, [ImageRelationshipRef(part: "word/document.xml", id: "rId4")], "the readable part's orphan is still reported")
        XCTAssertEqual(report.unparsableParts, ["word/charts/chart1.xml"])
        XCTAssertFalse(report.isConsistent)
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
        XCTAssertFalse(report.isConsistent)
    }

    func testRelsTruncatedAfterACompleteDeclarationDiscardsThePrefix() throws {
        // verify R1 logic F1: the parser delivers rId4 before it fails on the
        // second element; a prefix of a declaration list is not a declaration
        // list and must not become an orphan.
        let data = try package(document: body(), docRels: #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/><Relationship Id="rId5" "#)
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(report.unparsableParts, ["word/_rels/document.xml.rels"])
        XCTAssertEqual(report.declaredImageRelationshipRefs, [])
        XCTAssertEqual(report.orphanImageRelationshipRefs, [])
        XCTAssertFalse(report.isConsistent)
    }

    func testPartsDeclaringADocumentTypeAreRefused() throws {
        // OPC forbids DTDs in package parts. Switching to a parser made "what
        // does an entity mean inside an attribute value" answerable in more
        // than one way; refusing document types keeps the answer structural.
        let dtd = #"<?xml version="1.0"?><!DOCTYPE Relationships [<!ENTITY x "rId9">]>"#
        let data = try package(
            document: body(referencing: "rId4"),
            docRels: "",
            extra: ["word/header1.xml": #"<w:hdr xmlns:w="\#(wNS)" xmlns:r="\#(rNS)"><w:p/></w:hdr>"#,
                    "word/_rels/header1.xml.rels": dtd + rels(#"<Relationship Id="&x;" Type="\#(imageType)" Target="media/image1.png"/>"#)])
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(report.unparsableParts, ["word/_rels/header1.xml.rels"])
        XCTAssertEqual(report.declaredImageRelationshipRefs, [], "a refused part declares nothing")
        XCTAssertFalse(report.isConsistent)
        // The policy is DocxReader.rejectDTD — byte-level, so a document type
        // with no declarations at all is refused too (verify R1 security).
        for bare in [#"<!DOCTYPE Relationships>"#, #"<!DOCTYPE Relationships []>"#, #"<!DOCTYPE Relationships SYSTEM "x.dtd">"#] {
            XCTAssertEqual(PackageInspector.scanRels(Data((bare + rels("")).utf8), part: "p").parsed, false, bare)
        }
    }

    func testEntityExpansionBombIsRefusedImmediately() throws {
        var dtd = #"<?xml version="1.0"?><!DOCTYPE Relationships [<!ENTITY a "aaaaaaaaaa">"#
        for i in 1...9 {
            let prev = i == 1 ? "a" : "e\(i - 1)"
            dtd += "<!ENTITY e\(i) \"" + String(repeating: "&\(prev);", count: 10) + "\">"
        }
        dtd += "]>"
        let data = try package(document: body(), docRels: "", extra: ["word/header1.xml": #"<w:hdr xmlns:w="\#(wNS)"/>"#,
                                                                     "word/_rels/header1.xml.rels": dtd + rels("<!-- &e9; -->")])
        let started = Date()
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertLessThan(Date().timeIntervalSince(started), 1.0)
        XCTAssertEqual(report.unparsableParts, ["word/_rels/header1.xml.rels"])
    }

    // MARK: - verify R2: the reader's other refusals, mirrored

    func testUndeclaredPrefixIsAParseErrorHereAsInTheReader() throws {
        // verify R2 DA: 7 bytes of `<zz:x/>` in a part the reader cannot open
        // (namespace error) must not read as consistent here.
        let data = try package(document: body(referencing: "rId4") + "<zz:x/>", docRels: #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#)
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(report.unparsableParts, ["word/document.xml"])
        XCTAssertFalse(report.isConsistent)
        XCTAssertThrowsError(try XMLDocument(data: Data((body(referencing: "rId4") + "<zz:x/>").utf8)), "the reader's DOM refuses the same bytes")
    }

    func testReferenceMustBeInTheRelationshipsNamespace() throws {
        // verify R2 codex N3: `fake:embed` in another namespace is not a
        // reference and cannot satisfy a declaration; a foreign PREFIX bound
        // to the relationships namespace is (strict OOXML namespace too).
        let fake = #"<w:document xmlns:w="\#(wNS)" xmlns:fake="urn:not-a-relationship"><w:body><w:p fake:embed="rId4"/></w:body></w:document>"#
        let fakeReport = try PackageInspector.imageConsistencyReport(of: try package(document: fake, docRels: #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#))
        XCTAssertEqual(fakeReport.orphanImageRelationshipRefs.map(\.id), ["rId4"], "a same-named attribute in another namespace hides nothing")
        let strict = #"<w:document xmlns:w="\#(wNS)" xmlns:a="\#(aNS)" xmlns:rel="http://purl.oclc.org/ooxml/officeDocument/relationships"><w:body><w:p><w:r><w:drawing><a:blip rel:embed="rId4"/></w:drawing></w:r></w:p></w:body></w:document>"#
        let strictReport = try PackageInspector.imageConsistencyReport(of: try package(document: strict, docRels: #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#))
        XCTAssertTrue(strictReport.isConsistent, "orphans: \(strictReport.orphanImageRelationshipRefs)")
    }

    func testPartNamesAreCaseInsensitive() throws {
        // verify R2 DA: `Word/` is `word/` to OPC and to a case-insensitive file
        // system the reader extracts onto; it must not be a mute switch here.
        let parts: [String: String] = [
            "Word/document.xml": body(),
            "Word/_rels/document.xml.rels": rels(#"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#),
            "Word/media/image1.png": "png",
        ]
        let report = try PackageInspector.imageConsistencyReport(of: try zip(parts))
        XCTAssertEqual(report.orphanImageRelationshipRefs, [ImageRelationshipRef(part: "word/document.xml", id: "rId4")])
        XCTAssertEqual(report.mediaEntryCount, 1)
        XCTAssertFalse(report.isConsistent)
    }

    func testNonUTF8PartsAreRefusedLikeTheReaderAndUTF8BOMIsFine() throws {
        let relsXML = rels(#"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#)
        for (label, data) in [
            ("UTF-16LE BOM", Data([0xFF, 0xFE]) + relsXML.data(using: .utf16LittleEndian)!),
            ("UTF-16BE BOM", Data([0xFE, 0xFF]) + relsXML.data(using: .utf16BigEndian)!),
            ("UTF-16LE no BOM", relsXML.data(using: .utf16LittleEndian)!),
        ] {
            XCTAssertNotNil(PackageInspector.linearPrecheckFailure(data), label)
            XCTAssertFalse(PackageInspector.scanRels(data, part: "p").parsed, label)
            XCTAssertThrowsError(try XmlTreeReader.parse(data), "\(label): the reader refuses it too")
        }
        let bom = Data([0xEF, 0xBB, 0xBF]) + Data(relsXML.utf8)
        XCTAssertEqual(PackageInspector.scanRels(bom, part: "p").imageIds, ["rId4"])
        XCTAssertNoThrow(try XmlTreeReader.parse(bom))
    }

    func testNestingDeeperThanTheReadersLimitIsRefusedBeforeParsing() throws {
        // verify R2 security N1: depth × xmlns is quadratic in libxml2; the
        // reader already stops at 1024 (XmlTreeReader.maxElementDepth).
        let limit = PackageInspector.maxElementDepth
        let deep = "<r>" + String(repeating: "<a xmlns:x=\"urn:x\">", count: limit + 1) + String(repeating: "</a>", count: limit + 1) + "</r>"
        let started = Date()
        XCTAssertNotNil(PackageInspector.linearPrecheckFailure(Data(deep.utf8)))
        XCTAssertLessThan(Date().timeIntervalSince(started), 0.5)
        let ok = "<r>" + String(repeating: "<a>", count: limit - 1) + String(repeating: "</a>", count: limit - 1) + "</r>"
        XCTAssertNil(PackageInspector.linearPrecheckFailure(Data(ok.utf8)))
        XCTAssertNil(PackageInspector.linearPrecheckFailure(Data("<r><a/><a/><a/></r>".utf8)), "self-closing tags do not nest")
    }

    func testDuplicateDeclarationsMakeThePackageInconsistent() throws {
        // verify R2 security N3: a package the writer refuses must not read as consistent.
        let twice = #"<Relationship Id="rId5" Type="\#(imageType)" Target="media/image1.png"/>"# + #"<Relationship Id="rId5" Type="\#(imageType)" Target="media/image2.png"/>"#
        let report = try PackageInspector.imageConsistencyReport(of: try package(document: body(referencing: "rId5"), docRels: twice))
        XCTAssertEqual(report.duplicateRelationshipRefs.map(\.id), ["rId5"])
        XCTAssertEqual(report.orphanImageRelationshipRefs, [])
        XCTAssertFalse(report.isConsistent)
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

    func testMergeEmitsATypedDuplicateOnceInPassTwo() throws {
        // verify R1 codex 6 / logic: original has no rId5; typed carries it twice.
        let overlay = RelationshipsOverlay(originalRelsXML: rels(""))
        let dupes = [
            RelationshipDescriptor(id: "rId5", type: imageType, target: "media/a.png", targetMode: nil),
            RelationshipDescriptor(id: "rId5", type: imageType, target: "media/b.png", targetMode: nil),
        ]
        let xml = overlay.merge(typedRels: dupes, typedManagedTypes: [imageType])
        XCTAssertEqual(xml.components(separatedBy: "Id=\"rId5\"").count - 1, 1, xml)
    }

    func testDuplicateInTheOriginalRelsIsRefusedOnSave() throws {
        // verify R1 requirements R14: a clean model over a package whose own
        // rels declare an id twice — first-wins would drop one and save.
        var doc = WordDocument()
        doc.body.children.append(.paragraph(Paragraph(runs: [Run(text: "x")])))
        let url = FileManager.default.temporaryDirectory.appendingPathComponent("i139-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: url); defer { try? FileManager.default.removeItem(at: url) }
        let dir = try ZipHelper.unzip(url); defer { ZipHelper.cleanup(dir) }
        let relsURL = dir.appendingPathComponent("word/_rels/document.xml.rels")
        var relsXML = try String(contentsOf: relsURL, encoding: .utf8)
        relsXML = relsXML.replacingOccurrences(of: "</Relationships>", with: #"<Relationship Id="rId9" Type="\#(imageType)" Target="media/a.png"/><Relationship Id="rId&#57;" Type="\#(imageType)" Target="media/b.png"/></Relationships>"#)
        try relsXML.write(to: relsURL, atomically: true, encoding: .utf8)
        let damaged = FileManager.default.temporaryDirectory.appendingPathComponent("i139-dup-\(UUID().uuidString).docx")
        try ZipHelper.zip(dir, to: damaged); defer { try? FileManager.default.removeItem(at: damaged) }
        var read = try DocxReader.read(from: damaged); defer { read.close() }
        var thrown: Error?
        XCTAssertThrowsError(try DocxWriter.writeData(read)) { thrown = $0 }
        let message = (thrown as? LocalizedError)?.errorDescription ?? String(describing: thrown)
        XCTAssertTrue(message.contains("rId9"), message)
        XCTAssertTrue(message.contains("#139"), message)
    }

    func testUnreadableOriginalRelsIsRefusedNotMergedByRegex() throws {
        // verify R2 codex N2 / security N2 / logic N1: a rels the strict scan
        // cannot read must not fall through to the regex merge.
        var doc = WordDocument()
        doc.body.children.append(.paragraph(Paragraph(runs: [Run(text: "x")])))
        let url = FileManager.default.temporaryDirectory.appendingPathComponent("i139-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: url); defer { try? FileManager.default.removeItem(at: url) }
        let dir = try ZipHelper.unzip(url); defer { ZipHelper.cleanup(dir) }
        let relsURL = dir.appendingPathComponent("word/_rels/document.xml.rels")
        var relsXML = try String(contentsOf: relsURL, encoding: .utf8)
        let wide = (1...(PackageInspector.maxAttributesPerElement + 1)).map { "a\($0)=\"v\"" }.joined(separator: " ")
        relsXML = relsXML.replacingOccurrences(of: "</Relationships>", with: "<Relationship \(wide) Id=\"rId9\" Type=\"\(imageType)\" Target=\"media/a.png\"/><Relationship Id=\"rId9\" Type=\"\(imageType)\" Target=\"media/b.png\"/></Relationships>")
        try relsXML.write(to: relsURL, atomically: true, encoding: .utf8)
        let damaged = FileManager.default.temporaryDirectory.appendingPathComponent("i139-wide-\(UUID().uuidString).docx")
        try ZipHelper.zip(dir, to: damaged); defer { try? FileManager.default.removeItem(at: damaged) }
        var read = try DocxReader.read(from: damaged); defer { read.close() }
        var thrown: Error?
        XCTAssertThrowsError(try DocxWriter.writeData(read)) { thrown = $0 }
        let message = (thrown as? LocalizedError)?.errorDescription ?? String(describing: thrown)
        XCTAssertTrue(message.contains("could not be scanned"), message)
    }

    func testOriginalIdWrittenWithACharacterReferenceIsRefused() throws {
        // verify R2 codex N4 / logic N1: the overlay indexes raw ids, the model
        // holds decoded ones; refuse rather than emit two views of one id.
        var doc = WordDocument()
        doc.body.children.append(.paragraph(Paragraph(runs: [Run(text: "x")])))
        let url = FileManager.default.temporaryDirectory.appendingPathComponent("i139-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: url); defer { try? FileManager.default.removeItem(at: url) }
        let dir = try ZipHelper.unzip(url); defer { ZipHelper.cleanup(dir) }
        let relsURL = dir.appendingPathComponent("word/_rels/document.xml.rels")
        var relsXML = try String(contentsOf: relsURL, encoding: .utf8)
        relsXML = relsXML.replacingOccurrences(of: "</Relationships>", with: #"<Relationship Id="rId&#57;" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>"#)
        try relsXML.write(to: relsURL, atomically: true, encoding: .utf8)
        let damaged = FileManager.default.temporaryDirectory.appendingPathComponent("i139-ent-\(UUID().uuidString).docx")
        try ZipHelper.zip(dir, to: damaged); defer { try? FileManager.default.removeItem(at: damaged) }
        var read = try DocxReader.read(from: damaged); defer { read.close() }
        var thrown: Error?
        XCTAssertThrowsError(try DocxWriter.writeData(read)) { thrown = $0 }
        let message = (thrown as? LocalizedError)?.errorDescription ?? String(describing: thrown)
        XCTAssertTrue(message.contains("rId&#57;"), message)
        XCTAssertTrue(message.contains("#142"), message)
    }

    func testDuplicateDeclarationsAreNamedInTheReport() throws {
        let twice = #"<Relationship Id="rId5" Type="\#(imageType)" Target="media/image1.png"/>"#
            + #"<Relationship Id="rId&#53;" Type="\#(imageType)" Target="media/image2.png"/>"#
        let data = try package(document: body(referencing: "rId5"), docRels: twice)
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(report.duplicateRelationshipRefs, [ImageRelationshipRef(part: "word/document.xml", id: "rId5")],
                       "`rId5` and `rId&#53;` are one id once parsed")
        XCTAssertFalse(report.isConsistent, "a package the writer refuses is not consistent")
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
        let document = #"<w:document xmlns:w="\#(wNS)" xmlns:a="\#(aNS)" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:r="\#(rNS)"><w:body><w:p><w:r><w:drawing><wp:inline><a:graphic/></wp:inline></w:drawing></w:r></w:p></w:body></w:document>"#
        let data = try package(document: document, docRels: "", media: false)
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(report.bodyDrawingCount, 1)
        XCTAssertEqual(report.imageRelationshipCount, 0)
        XCTAssertEqual(report.mediaEntryCount, 0)
        XCTAssertTrue(report.isConsistent)
    }
}
