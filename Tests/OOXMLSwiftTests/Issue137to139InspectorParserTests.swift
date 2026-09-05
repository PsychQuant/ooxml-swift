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

    /// A package with arbitrary entry names, in order — including names that
    /// collide or repeat (ZIPFoundation writes what it is given).
    private func zipEntries(_ entries: [(String, String)]) throws -> Data {
        let archive = try Archive(accessMode: .create)
        for (name, text) in entries {
            let data = Data(text.utf8)
            try archive.addEntry(with: name, type: name.hasSuffix("/") ? .directory : .file, uncompressedSize: Int64(data.count),
                                 compressionMethod: .deflate) { position, size in
                let start = data.startIndex.advanced(by: Int(position))
                return data.subdata(in: start..<start.advanced(by: size))
            }
        }
        return try XCTUnwrap(archive.data)
    }

    /// What `DocxReader` does with the same bytes: the number of images it
    /// loads, or the error it throws. The by-construction oracle for every
    /// "as the reader does" claim below.
    private func readerImageCount(_ data: Data) throws -> Int {
        let url = FileManager.default.temporaryDirectory.appendingPathComponent("i137-\(UUID().uuidString).docx")
        try data.write(to: url); defer { try? FileManager.default.removeItem(at: url) }
        var document = try DocxReader.read(from: url); defer { document.close() }
        return document.images.count
    }

    /// Writer refusal for a package whose `word/_rels/document.xml.rels` was
    /// rewritten by `mutate` — read back through `DocxReader` first, so the
    /// shape is one the reader accepts.
    private func writerRefusal(mutatingRels mutate: (URL) throws -> Void, file: StaticString = #filePath, line: UInt = #line) throws -> String {
        var doc = WordDocument()
        doc.body.children.append(.paragraph(Paragraph(runs: [Run(text: "x")])))
        let url = FileManager.default.temporaryDirectory.appendingPathComponent("i139-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: url); defer { try? FileManager.default.removeItem(at: url) }
        let dir = try ZipHelper.unzip(url); defer { ZipHelper.cleanup(dir) }
        try mutate(dir.appendingPathComponent("word/_rels/document.xml.rels"))
        let damaged = FileManager.default.temporaryDirectory.appendingPathComponent("i139-shape-\(UUID().uuidString).docx")
        try ZipHelper.zip(dir, to: damaged); defer { try? FileManager.default.removeItem(at: damaged) }
        var read = try DocxReader.read(from: damaged); defer { read.close() }
        var thrown: Error?
        XCTAssertThrowsError(try DocxWriter.writeData(read), file: file, line: line) { thrown = $0 }
        return (thrown as? LocalizedError)?.errorDescription ?? String(describing: thrown)
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
            "very long decimal":  "rId&#54;",
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

    func testUndeclaredPrefixIsRefusedHereAsInTheReader() throws {
        // verify R2 DA / R3 B7: 7 bytes of `<zz:x/>` in a part the reader
        // cannot open (namespace error 201) must not read as consistent here.
        // libxml2's SAX path only records the error, so the scanner refuses
        // it itself. The R2 fixture put the bytes AFTER the root element and
        // passed for an unrelated reason ("extra content"); these are inside.
        // This is the ONE namespace class the inspector re-implements — see
        // the leniency test below for the eleven it deliberately does not.
        let rel = #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#
        let head = #"<w:document xmlns:w="\#(wNS)" xmlns:a="\#(aNS)" xmlns:r="\#(rNS)"><w:body><w:p><w:r><w:drawing><a:blip r:embed="rId4"/></w:drawing></w:r></w:p>"#
        for (label, document) in [
            ("undeclared element prefix", head + "<zz:x/></w:body></w:document>"),
            ("undeclared attribute prefix", head + #"<w:p zz:a="1"/></w:body></w:document>"#),
            ("prefix declared only in a sibling", head + #"<w:p xmlns:zz="urn:zz"/><zz:x/></w:body></w:document>"#),
        ] {
            let data = try package(document: document, docRels: rel)
            let report = try PackageInspector.imageConsistencyReport(of: data)
            XCTAssertEqual(report.unparsableParts, ["word/document.xml"], label)
            XCTAssertFalse(report.isConsistent, label)
            XCTAssertThrowsError(try XMLDocument(data: Data(document.utf8)), "\(label): the reader's DOM refuses the same bytes")
            XCTAssertThrowsError(try readerImageCount(data), "\(label): the reader refuses the package")
        }
        // …and a declared prefix, or `xml:`, is not refused (no over-refusal).
        let fine = head + #"<zz:x xmlns:zz="urn:zz"/><w:p xml:space="preserve"/></w:body></w:document>"#
        let report = try PackageInspector.imageConsistencyReport(of: try package(document: fine, docRels: rel))
        XCTAssertTrue(report.isConsistent, "\(report.unparsableParts) \(report.orphanImageRelationshipRefs)")
        XCTAssertEqual(try readerImageCount(try package(document: fine, docRels: rel)), 1)
    }

    func testNamespaceIllFormednessBeyondUndeclaredPrefixesIsDocumentedLeniencyNotParity() throws {
        // verify R3 DA D1: XMLDocument (the reader) refuses twelve classes of
        // namespace ill-formedness; XMLParser refuses none of them, and the
        // inspector re-implements exactly one (an undeclared prefix, the one
        // shape a real writer produces). The rest are documented leniency —
        // this test pins that boundary so nobody claims reader parity again,
        // and so nobody starts emulating libxml2's namespace layer in a
        // delegate either (the anti-pattern #137 removed from the consumer).
        // The property that IS promised: the report never hides a
        // relationship on these shapes — declarations and references are
        // what the parser delivered.
        let rel = #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#
        let head = #"<w:document xmlns:w="\#(wNS)" xmlns:a="\#(aNS)" xmlns:r="\#(rNS)"><w:body><w:p><w:r><w:drawing><a:blip r:embed="rId4"/></w:drawing></w:r></w:p>"#
        // `refused` is what XMLParser + the undeclared-prefix rule happen to do
        // today, pinned so a change is a conscious one: a trailing colon
        // arrives with no namespace URI and falls under the prefix rule; the
        // rest parse. None of this is reader parity.
        let shapes: [(String, String, Bool)] = [
            ("QName with two colons",            head + "<w:p:q/></w:body></w:document>", false),
            ("QName with a trailing colon",      head + "<w:/></w:body></w:document>", true),
            ("leading colon under a default ns", head + #"<w:p xmlns="urn:d"><:b/></w:p></w:body></w:document>"#, false),
            ("prefix bound to an empty URI",     head + #"<w:p xmlns:zz=""/></w:body></w:document>"#, false),
            ("prefix bound to an invalid URI",   head + #"<w:p xmlns:zz="urn: bad"/></w:body></w:document>"#, false),
            ("xml prefix rebound",               head + #"<w:p xmlns:xml="urn:wrong"/></w:body></w:document>"#, false),
            ("expanded duplicate attribute",     head + #"<w:p xmlns:p1="urn:u" xmlns:p2="urn:u" p1:k="1" p2:k="2"/></w:body></w:document>"#, false),
        ]
        for (label, document, refused) in shapes {
            XCTAssertThrowsError(try XMLDocument(data: Data(document.utf8)), "\(label): the reader's DOM refuses it")
            let report = try PackageInspector.imageConsistencyReport(of: try package(document: document, docRels: rel))
            XCTAssertEqual(report.unparsableParts, refused ? ["word/document.xml"] : [], "\(label): documented leniency — not emulated")
            // Whatever the verdict, the report never hides a relationship.
            XCTAssertEqual(report.declaredImageRelationshipRefs, [ImageRelationshipRef(part: "word/document.xml", id: "rId4")], label)
            XCTAssertEqual(report.orphanImageRelationshipRefs, [], label)
            if !refused { XCTAssertEqual(report.bodyDrawingCount, 1, label) }
        }
    }

    func testMediaEntryCountCountsFilesNotTheDirectoryEntryAndCollidingMediaRefusesLikeTheReader() throws {
        // verify R3 DA D3 / M9: media counting had no test at all. Files
        // directly under word/media count; the `word/media/` directory entry
        // some writers store does not (3.6.4 counted it — 6/740 real docs
        // differ by exactly 1). A second entry that lands on the same media
        // file makes extraction fail, so the report is refused exactly when
        // the reader is (there is no count to "flatten" any more).
        let rel = #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#
        let withDirEntry = try zipEntries([("word/document.xml", body(referencing: "rId4")), ("word/_rels/document.xml.rels", rels(rel)),
                                           ("word/media/", ""), ("word/media/image1.png", "png"), ("word/media/image2.png", "png")])
        XCTAssertEqual(try PackageInspector.imageConsistencyReport(of: withDirEntry).mediaEntryCount, 2)
        let colliding = try zipEntries([("word/document.xml", body(referencing: "rId4")), ("word/_rels/document.xml.rels", rels(rel)),
                                        ("word/media/image1.png", "png"), ("Word/media/image1.png", "png"), ("word/media/image1.png", "png")])
        let readerRefuses = (try? readerImageCount(colliding)) == nil
        let inspectorRefuses = (try? PackageInspector.imageConsistencyReport(of: colliding)) == nil
        XCTAssertEqual(inspectorRefuses, readerRefuses, "the inspector refuses colliding media exactly when the reader does")
        XCTAssertTrue(readerRefuses, "an exact duplicate entry cannot be extracted on any file system")
    }

    func testCollidingEntryNamesRefuseTheReportExactlyWhenTheyRefuseTheReader() throws {
        // verify R3 security S-R3-1/2: the reader extracts to a file system;
        // a second entry that lands on an existing file makes extraction fail
        // (NSCocoaError 516) and the reader never opens the package. A
        // lowercased first-wins index chose by archive order instead and
        // reported such packages consistent — the opposite of the reader.
        // The inspector now extracts the same way, so it fails when the
        // reader fails, whatever the file system decides about case.
        let rel = #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#
        let cases: [(String, [(String, String)], Bool)] = [
            ("case-colliding document parts, the referencing one first",
             [("WORD/DOCUMENT.XML", body(referencing: "rId4")), ("word/document.xml", body()),
              ("word/_rels/document.xml.rels", rels(rel)), ("word/media/image1.png", "png")], false),
            ("case-colliding rels, the empty one first",
             [("WORD/_RELS/DOCUMENT.XML.RELS", rels("")), ("word/_rels/document.xml.rels", rels(rel)),
              ("word/document.xml", body()), ("word/media/image1.png", "png")], false),
            ("the same entry name twice",
             [("word/document.xml", body(referencing: "rId4")), ("word/document.xml", body()),
              ("word/_rels/document.xml.rels", rels(rel)), ("word/media/image1.png", "png")], true),
        ]
        for (label, entries, refusalIsFileSystemIndependent) in cases {
            let data = try zipEntries(entries)
            let readerRefuses = (try? readerImageCount(data)) == nil
            var inspectorError: Error?
            let report = try? { () throws -> ImageConsistencyReport in
                do { return try PackageInspector.imageConsistencyReport(of: data) } catch { inspectorError = error; throw error }
            }()
            XCTAssertEqual(report == nil, readerRefuses, "\(label): the inspector must refuse exactly when the reader does")
            if refusalIsFileSystemIndependent { XCTAssertTrue(readerRefuses, label) }
            if let inspectorError {
                XCTAssertTrue(String(describing: inspectorError).contains("no consistency verdict"), "\(label): \(inspectorError)")
            }
        }
    }

    func testRelsPathSpellingsMeanWhatTheyMeanToTheReader() throws {
        // verify R3 security S-R3-3: `word/_rels/./document.xml.rels`,
        // `./word/_rels/document.xml.rels` and a U+017F (long s) in the path
        // were invisible to the archive index while the file system served
        // them to the reader (APFS collapses `.` and folds U+017F to `s`;
        // `lowercased()` does neither). Extracting like the reader makes the
        // question disappear: whatever the file system does, both see it.
        let rel = #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#
        for relsName in ["word/_rels/./document.xml.rels", "./word/_rels/document.xml.rels", "word/_rel\u{017F}/document.xml.rel\u{017F}"] {
            let data = try zipEntries([("word/document.xml", body()), (relsName, rels(rel)), ("word/media/image1.png", "png")])
            let readerImages = try? readerImageCount(data)
            let report = try? PackageInspector.imageConsistencyReport(of: data)
            XCTAssertEqual(report == nil, readerImages == nil, relsName)
            guard let report, let readerImages else { continue }
            XCTAssertEqual(report.declaredImageRelationshipRefs.count, readerImages, "\(relsName): the inspector sees the rels iff the reader loads its image")
            if readerImages == 1 {
                XCTAssertEqual(report.orphanImageRelationshipRefs, [ImageRelationshipRef(part: "word/document.xml", id: "rId4")], relsName)
            }
        }
    }

    func testInvalidUTF8IsRefusedBeforeParsing() throws {
        // verify R3 requirements: a bad byte is not a character the reader and
        // the inspector could agree on (the reader substitutes U+FFFD), so it
        // is refused here — stricter than the reader, documented as such.
        var bytes = Data(body().utf8)
        let marker = Data("<w:p/>".utf8)
        let range = try XCTUnwrap(bytes.range(of: marker))
        bytes.replaceSubrange(range, with: Data("<w:p>".utf8) + Data([0xC3, 0x28]) + Data("</w:p>".utf8))
        XCTAssertEqual(PackageInspector.linearPrecheckFailure(bytes), "not valid UTF-8")
        XCTAssertFalse(PackageInspector.scanPart(bytes, part: "p").parsed)
        let archive = try Archive(accessMode: .create)
        for (name, data) in [("word/document.xml", bytes), ("word/_rels/document.xml.rels", Data(rels(#"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#).utf8))] {
            try archive.addEntry(with: name, type: .file, uncompressedSize: Int64(data.count), compressionMethod: .deflate) { position, size in
                let start = data.startIndex.advanced(by: Int(position))
                return data.subdata(in: start..<start.advanced(by: size))
            }
        }
        let report = try PackageInspector.imageConsistencyReport(of: try XCTUnwrap(archive.data))
        XCTAssertEqual(report.unparsableParts, ["word/document.xml"])
        XCTAssertFalse(report.isConsistent)
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
        // verify R2 DA: `Word/` is `word/` to OPC and to the case-insensitive
        // file system the reader extracts onto; it must not be a mute switch
        // here. (Since R3 the inspector extracts the same way, so this holds
        // by construction — and on a case-sensitive volume both would refuse.)
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
        // verify R3 logic G5: the reader enters a self-closing element too, so
        // one AT the limit trips it; mirror that exactly, both ways.
        let selfClosingOver = "<r>" + String(repeating: "<a>", count: limit - 1) + "<b/>" + String(repeating: "</a>", count: limit - 1) + "</r>"
        XCTAssertNotNil(PackageInspector.linearPrecheckFailure(Data(selfClosingOver.utf8)))
        XCTAssertThrowsError(try XmlTreeReader.parse(Data(selfClosingOver.utf8)), "the reader refuses the same depth")
        let selfClosingAt = "<r>" + String(repeating: "<a>", count: limit - 2) + "<b/>" + String(repeating: "</a>", count: limit - 2) + "</r>"
        XCTAssertNil(PackageInspector.linearPrecheckFailure(Data(selfClosingAt.utf8)))
        XCTAssertNoThrow(try XmlTreeReader.parse(Data(selfClosingAt.utf8)), "the reader accepts the same depth")
        XCTAssertThrowsError(try XmlTreeReader.parse(Data(deep.utf8)))
        XCTAssertNoThrow(try XmlTreeReader.parse(Data(ok.utf8)))
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
        relsXML = relsXML.replacingOccurrences(of: "</Relationships>", with: #"<Relationship Id="rId9" Type="\#(imageType)" Target="media/a.png"/><Relationship Id="rId&#211;" Type="\#(imageType)" Target="media/b.png"/></Relationships>"#)
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
        relsXML = relsXML.replacingOccurrences(of: "</Relationships>", with: #"<Relationship Id="rId&#211;" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>"#)
        try relsXML.write(to: relsURL, atomically: true, encoding: .utf8)
        let damaged = FileManager.default.temporaryDirectory.appendingPathComponent("i139-ent-\(UUID().uuidString).docx")
        try ZipHelper.zip(dir, to: damaged); defer { try? FileManager.default.removeItem(at: damaged) }
        var read = try DocxReader.read(from: damaged); defer { read.close() }
        var thrown: Error?
        XCTAssertThrowsError(try DocxWriter.writeData(read)) { thrown = $0 }
        let message = (thrown as? LocalizedError)?.errorDescription ?? String(describing: thrown)
        XCTAssertTrue(message.contains("rId&#211;"), message)
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

    // MARK: - #139 · shapes the regex merge cannot index are refused, and named

    func testRelsShapesTheMergeCannotIndexAreRefusedByShape() throws {
        // verify R3 security S-R3-5 / logic G2-G3 / codex N4: five legal rels
        // spellings that 3.6.4 merged anyway — silently dropping the
        // relationship (W1/W2/W3/W5) or writing a commented-out one as live
        // (W4). Each is refused, and the message names its own cause rather
        // than "character references" for all of them.
        let theme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
        func appending(_ fragment: String) -> (URL) throws -> Void {
            { url in
                let xml = try String(contentsOf: url, encoding: .utf8)
                try xml.replacingOccurrences(of: "</Relationships>", with: fragment + "</Relationships>").write(to: url, atomically: true, encoding: .utf8)
            }
        }
        let cases: [(String, (URL) throws -> Void, [String])] = [
            ("comment", appending(#"<!-- <Relationship Id="rId99" Type="\#(imageType)" Target="media/gone.png"/> -->"#), ["an XML comment", "#142"]),
            ("CDATA", appending("<![CDATA[x]]>"), ["a CDATA section", "#142"]),
            ("prefixed element", { url in
                let xml = try String(contentsOf: url, encoding: .utf8)
                    .replacingOccurrences(of: "</Relationships>", with: #"<pkg:Relationship xmlns:pkg="\#(self.pkgNS)" Id="rId9" Type="\#(theme)" Target="theme/theme1.xml"/></Relationships>"#)
                try xml.write(to: url, atomically: true, encoding: .utf8)
            }, ["namespace-prefixed", "#142"]),
            ("not self-closing", appending(#"<Relationship Id="rId9" Type="\#(theme)" Target="theme/theme1.xml"></Relationship>"#), ["not self-closing", "rId9", "#142"]),
            ("single quotes", appending("<Relationship Id='rId9' Type='\(theme)' Target='theme/theme1.xml'/>"), ["single-quoted", "rId9", "#142"]),
            ("whitespace around =", appending(#"<Relationship Id = "rId9" Type="\#(theme)" Target="theme/theme1.xml"/>"#), ["whitespace around", "rId9", "#142"]),
        ]
        for (label, mutate, expected) in cases {
            let message = try writerRefusal(mutatingRels: mutate)
            for phrase in expected { XCTAssertTrue(message.contains(phrase), "\(label): \(message)") }
            XCTAssertFalse(message.contains("(reads as"), "\(label): no mis-pairing of unequal lists: \(message)")
        }
    }

    func testRelsFileThatExistsButIsEmptyOrNotUTF8IsRefusedNotTreatedAsAbsent() throws {
        // verify R3 codex N3 / logic G6: the gate keyed on "the string is not
        // empty" — an empty rels file, or one that is not UTF-8, read as ""
        // and took the scratch path, which drops every relationship the
        // typed model does not manage. The gate is now keyed on existence.
        var doc = WordDocument()
        doc.body.children.append(.paragraph(Paragraph(runs: [Run(text: "x")])))
        let url = FileManager.default.temporaryDirectory.appendingPathComponent("i139-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: url); defer { try? FileManager.default.removeItem(at: url) }
        for (label, bytes, phrase) in [("empty", Data(), "could not be scanned"), ("Latin-1", Data([0x3C, 0xE9, 0x3E]), "not readable as UTF-8")] {
            var read = try DocxReader.read(from: url); defer { read.close() }
            let relsURL = try XCTUnwrap(read.archiveTempDir).appendingPathComponent("word/_rels/document.xml.rels")
            try bytes.write(to: relsURL)
            var thrown: Error?
            XCTAssertThrowsError(try DocxWriter.writeData(read), label) { thrown = $0 }
            let message = (thrown as? LocalizedError)?.errorDescription ?? String(describing: thrown)
            XCTAssertTrue(message.contains(phrase), "\(label): \(message)")
        }
    }
}
