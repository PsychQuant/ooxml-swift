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
                    // The chart's rels declares something, so the chart IS in the
                    // scan set (a part whose rels declares nothing is not read at all — R5 L4).
                    "word/charts/_rels/chart1.xml.rels": rels(#"<Relationship Id="rId2" Type="\#(imageType)" Target="../media/image1.png"/>"#)])
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(report.orphanImageRelationshipRefs, [ImageRelationshipRef(part: "word/document.xml", id: "rId4")], "the readable part's orphan is still reported; the unreadable part yields none")
        XCTAssertEqual(report.unparsableParts, ["word/charts/chart1.xml"])
        XCTAssertEqual(report.declaredImageRelationshipRefs.count, 2, "the unreadable part's declarations stay listed")
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
        // Of the reader's twelve namespace classes the inspector refuses two —
        // this one and a prefix bound to an empty URI (libxml2 reports no
        // mapping for it) — and deliberately judges none of the other ten; see
        // the leniency test below.
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
        // inspector re-implements exactly two (an undeclared prefix — the one
        // shape a real writer produces — and a prefix bound to an empty URI:
        // one rule, "no reported mapping", two of the reader's classes). The
        // other ten are documented leniency —
        // this test pins that boundary so nobody claims reader parity again,
        // and so nobody starts emulating libxml2's namespace layer in a
        // delegate either (the anti-pattern #137 removed from the consumer).
        // The property that IS promised: the report never hides a
        // relationship on these shapes — declarations and references are
        // what the parser delivered.
        let rel = #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#
        let head = #"<w:document xmlns:w="\#(wNS)" xmlns:a="\#(aNS)" xmlns:r="\#(rNS)"><w:body><w:p><w:r><w:drawing><a:blip r:embed="rId4"/></w:drawing></w:r></w:p>"#
        // Every shape here is refused by the reader. `refused` says what the
        // inspector does, and matches the class doc's two closed lists: a
        // prefix bound to an EMPTY URI is refused — libxml2 reports no
        // mapping for `xmlns:zz=""`, so to the undeclared-prefix rule `zz`
        // is undeclared (not emulation: the parser's own report). Everything
        // else parses (codex R4-B1: the earlier rule judged "no namespace
        // URI" and so refused a trailing colon on a declared prefix; it now
        // asks only whether the prefix is declared). Seven in the document,
        // three in the rels — the DA's ten shapes.
        let shapes: [(String, String, Bool)] = [
            ("QName with two colons",            head + "<w:p:q/></w:body></w:document>", false),
            ("QName with a trailing colon",      head + "<w:/></w:body></w:document>", false),
            ("leading colon under a default ns", head + #"<w:p xmlns="urn:d"><:b/></w:p></w:body></w:document>"#, false),
            ("prefix bound to an empty URI, used on an element and an attribute", head + #"<w:p xmlns:zz=""><zz:x zz:a="1"/></w:p></w:body></w:document>"#, true),
            ("prefix bound to an invalid URI",   head + #"<w:p xmlns:zz="urn: bad"/></w:body></w:document>"#, false),
            ("xml prefix rebound",               head + #"<w:p xmlns:xml="urn:wrong"/></w:body></w:document>"#, false),
            ("expanded duplicate attribute",     head + #"<w:p xmlns:p1="urn:u" xmlns:p2="urn:u" p1:k="1" p2:k="2"/></w:body></w:document>"#, false),
        ]
        for (label, document, refused) in shapes {
            XCTAssertThrowsError(try XMLDocument(data: Data(document.utf8)), "\(label): the reader's DOM refuses it")
            let report = try PackageInspector.imageConsistencyReport(of: try package(document: document, docRels: rel))
            XCTAssertEqual(report.unparsableParts, refused ? ["word/document.xml"] : [], label)
            // Leniency never hides a relationship: the report is exact.
            XCTAssertEqual(report.declaredImageRelationshipRefs, [ImageRelationshipRef(part: "word/document.xml", id: "rId4")], label)
            XCTAssertEqual(report.orphanImageRelationshipRefs, [], label)
            if !refused { XCTAssertEqual(report.bodyDrawingCount, 1, label) }
        }
        let relsShapes: [(String, String, Bool)] = [
            ("rels: QName with two colons",       #"<Relationships xmlns="\#(pkgNS)" xmlns:a="urn:a">"# + rel + "<a:b:c/></Relationships>", false),
            ("rels: prefix bound to an empty URI", #"<Relationships xmlns="\#(pkgNS)" xmlns:zz="">"# + rel + "<zz:x/></Relationships>", true),
            ("rels: trailing colon",              #"<Relationships xmlns="\#(pkgNS)" xmlns:x="urn:x">"# + rel + "<x:/></Relationships>", false),
        ]
        for (label, relsXML, refused) in relsShapes {
            XCTAssertThrowsError(try XMLDocument(data: Data(relsXML.utf8)), "\(label): the reader's DOM refuses it")
            let report = try PackageInspector.imageConsistencyReport(of: try zip(["word/document.xml": body(referencing: "rId4"), "word/_rels/document.xml.rels": relsXML, "word/media/image1.png": "png"]))
            XCTAssertEqual(report.unparsableParts, refused ? ["word/_rels/document.xml.rels"] : [], label)
            if !refused {
                XCTAssertEqual(report.declaredImageRelationshipRefs, [ImageRelationshipRef(part: "word/document.xml", id: "rId4")], label)
                XCTAssertTrue(report.isConsistent, label)
            }
        }
    }

    func testMediaEntryCountCountsFilesNotTheDirectoryEntryAndCollidingMediaRefusesLikeTheReader() throws {
        // verify R3 DA D3 / M9: media counting had no test at all. Files
        // directly under word/media count; the `word/media/` directory entry
        // some writers store does not (3.6.4 counted it — 7/740 real docs
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
        // Not pinned here, pinned elsewhere: that the document dedupe compares
        // file identity rather than folding the name. On this (case-insensitive)
        // volume the two are indistinguishable for an ASCII name (verify R5
        // DA-R5-3); verify R6 security pinned it on a case-sensitive APFS volume
        // (`word/Document.xml` beside `word/document.xml` is scanned as its own
        // part and its orphan reported — a case-folding dedupe would hide it).
        let parts: [String: String] = [
            "Word/document.xml": body(),
            "Word/_rels/document.xml.rels": rels(#"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#),
            "Word/media/image1.png": "png",
        ]
        let report = try PackageInspector.imageConsistencyReport(of: try zip(parts))
        XCTAssertEqual(report.orphanImageRelationshipRefs, [ImageRelationshipRef(part: "word/document.xml", id: "rId4")])
        XCTAssertEqual(report.declaredImageRelationshipRefs.count, 1, "the listed `Word/document.xml` IS the document (same file), not a second part")
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
        // verify R5 (codex W-R5-1 / req R5-7 / logic L3): the cause is named as
        // what it is — the raw spelling that decodes to the parsed id.
        XCTAssertTrue(message.contains("character or entity reference"), message)
        XCTAssertFalse(message.contains("does not recognise"), message)
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
            ("plain comment", appending("<!-- note -->"), ["an XML comment", "#142"]),
            ("CDATA", appending("<![CDATA[x]]>"), ["a CDATA section", "#142"]),
            ("processing instruction", appending(#"<?audit <p:Relationship ?>"#), ["a processing instruction", "#142"]),
            ("prefixed root element, unprefixed children", { url in
                let xml = try String(contentsOf: url, encoding: .utf8)
                    .replacingOccurrences(of: "<Relationships xmlns=\"\(self.pkgNS)\">", with: "<pkg:Relationships xmlns:pkg=\"\(self.pkgNS)\" xmlns=\"\(self.pkgNS)\">")
                    .replacingOccurrences(of: "</Relationships>", with: "</pkg:Relationships>")
                try xml.write(to: url, atomically: true, encoding: .utf8)
            }, ["namespace-prefixed", "pkg:Relationships", "#142"]),
            ("prefixed element", { url in
                let xml = try String(contentsOf: url, encoding: .utf8)
                    .replacingOccurrences(of: "</Relationships>", with: #"<pkg:Relationship xmlns:pkg="\#(self.pkgNS)" Id="rId9" Type="\#(theme)" Target="theme/theme1.xml"/></Relationships>"#)
                try xml.write(to: url, atomically: true, encoding: .utf8)
            }, ["namespace-prefixed", "#142"]),
            ("not self-closing", appending(#"<Relationship Id="rId9" Type="\#(theme)" Target="theme/theme1.xml"></Relationship>"#), ["rId9: the <Relationship> element is not self-closing", "#142"]),
            ("single quotes", appending("<Relationship Id='rId9' Type='\(theme)' Target='theme/theme1.xml'/>"), ["rId9: single-quoted", "#142"]),
            ("whitespace around =", appending(#"<Relationship Id = "rId9" Type="\#(theme)" Target="theme/theme1.xml"/>"#), ["rId9: whitespace around", "#142"]),
        ]
        // The cause named is THE cause (codex R4-B7): a message that lists
        // every possible spelling names none.
        let otherCauses = ["single-quoted", "whitespace around", "not self-closing"]
        for (label, mutate, expected) in cases {
            let message = try writerRefusal(mutatingRels: mutate)
            for phrase in expected { XCTAssertTrue(message.contains(phrase), "\(label): \(message)") }
            let named = otherCauses.filter { message.contains($0) }
            XCTAssertLessThanOrEqual(named.count, 1, "\(label): names one cause, not a list: \(message)")
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

    // MARK: - verify R4 (codex B1–B8)

    func testAPartWithoutARelsIsNotReadSoItCannotRefuseAPackageTheReaderOpens() throws {
        // codex R4-B2: a part that declares nothing has nothing to reconcile.
        // Reading it anyway would let `word/unused.xml` (any bytes; the
        // reader never opens it) make a readable package "inconsistent".
        let rel = #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#
        let data = try package(document: body(referencing: "rId4"), docRels: rel, extra: ["word/unused.xml": "<not xml at all"])
        XCTAssertEqual(try readerImageCount(data), 1, "the reader opens it")
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(report.unparsableParts, [])
        XCTAssertTrue(report.isConsistent)
    }

    func testAPartWhoseRelsDeclaresNothingIsNotReadEither() throws {
        // verify R5 logic L4: an EMPTY rels declares nothing to reconcile, so
        // the part is not read — a rels the reader never reads must not make
        // a readable package inconsistent through bytes nobody opens.
        let rel = #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#
        let data = try package(document: body(referencing: "rId4"), docRels: rel,
                               extra: ["word/junk.xml": "<not xml at all", "word/_rels/junk.xml.rels": rels("")])
        XCTAssertEqual(try readerImageCount(data), 1, "the reader opens it")
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(report.unparsableParts, [])
        XCTAssertTrue(report.isConsistent)
    }

    func testARelsForAnAbsentPartIsFoundByTheFileSystemsRulesNotOurs() throws {
        // codex R4-B3: the absent-owner rels was still found by a suffix
        // match (`lowercased().hasSuffix(".rels")`), which U+017F defeats
        // where the file system does not. Now it is recognised by identity:
        // the file IS `<name>.rels` inside a directory that IS `_rels` to the
        // file system. The oracle is the file system itself: the orphan is
        // reported exactly when a lookup of the folded name finds the file.
        let rel = #"<Relationship Id="rId2" Type="\#(imageType)" Target="../media/image1.png"/>"#
        for relsName in ["word/charts/_rels/missing.xml.rels", "word/charts/_REL\u{017F}/MISSING.XML.REL\u{017F}", "word/charts/_rel\u{017F}/missing.xml.rel\u{017F}"] {
            let data = try zipEntries([("word/document.xml", body()), ("word/_rels/document.xml.rels", rels("")), (relsName, rels(rel)), ("word/media/image1.png", "png")])
            let url = FileManager.default.temporaryDirectory.appendingPathComponent("i137-\(UUID().uuidString).docx")
            try data.write(to: url); defer { try? FileManager.default.removeItem(at: url) }
            let dir = try ZipHelper.unzip(url); defer { ZipHelper.cleanup(dir) }
            let servedAtFoldedName = FileManager.default.fileExists(atPath: dir.appendingPathComponent("word/charts/_rels/missing.xml.rels").path)
            let report = try PackageInspector.imageConsistencyReport(of: data)
            // The owner does not exist, so it is named `<stem>.xml` from the rels
            // file's stem (identity-checked against `<stem>.xml.rels`).
            let ownerStem = String(relsName.split(separator: "/").last!.dropLast(9))
            let expected = servedAtFoldedName ? [ImageRelationshipRef(part: "word/charts/" + ownerStem + ".xml", id: "rId2")] : []
            XCTAssertEqual(report.orphanImageRelationshipRefs, expected, "\(relsName): served at the folded name = \(servedAtFoldedName)")
        }
    }

    func testSymlinkAndTraversalEntriesAreRefusedByReaderAndInspectorAlike() throws {
        // codex R4-B6: ZIPFoundation extracts a contained symlink, and every
        // later read follows it — `word/alias.xml → document.xml` would be a
        // second document. `ZipHelper.unzip` (the reader's own call) now
        // refuses any archive with a symlink entry; a traversal entry is
        // refused by ZIPFoundation itself. Both sides refuse both.
        let rel = #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#
        let target = Data("document.xml".utf8)
        let archive = try Archive(accessMode: .create)
        for (name, text) in [("word/document.xml", body(referencing: "rId4")), ("word/_rels/document.xml.rels", rels(rel)), ("word/media/image1.png", "png")] {
            let d = Data(text.utf8)
            try archive.addEntry(with: name, type: .file, uncompressedSize: Int64(d.count), compressionMethod: .deflate) { p, n in d.subdata(in: Int(p)..<Int(p) + n) }
        }
        try archive.addEntry(with: "word/alias.xml", type: .symlink, uncompressedSize: Int64(target.count)) { p, n in target.subdata(in: Int(p)..<Int(p) + n) }
        let withLink = try XCTUnwrap(archive.data)
        XCTAssertThrowsError(try readerImageCount(withLink), "the reader refuses a symlink entry")
        XCTAssertThrowsError(try PackageInspector.imageConsistencyReport(of: withLink)) { XCTAssertTrue(String(describing: $0).contains("symbolic-link"), "\($0)") }
        let traversal = try zipEntries([("word/document.xml", body(referencing: "rId4")), ("word/_rels/document.xml.rels", rels(rel)), ("../evil.xml", "<x/>"), ("word/media/image1.png", "png")])
        let readerRefuses = (try? readerImageCount(traversal)) == nil
        let inspectorRefuses = (try? PackageInspector.imageConsistencyReport(of: traversal)) == nil
        XCTAssertEqual(inspectorRefuses, readerRefuses, "traversal: same verdict as the reader")
        XCTAssertTrue(readerRefuses, "an entry that escapes the destination is refused")
        // verify R4 security S-R4-1: the conjunction ZIPFoundation 0.9.20 does
        // not catch — a link to `.` makes later `..` components resolve in
        // place while the containment check collapses them lexically, so
        // `word/a/a/../../../x` lands OUTSIDE the temporary directory. Both
        // halves are refused up front; nothing is written.
        let dot = Data(".".utf8)
        let chain = try Archive(accessMode: .create)
        for (name, text) in [("word/document.xml", body(referencing: "rId4")), ("word/_rels/document.xml.rels", rels(rel)), ("word/media/image1.png", "png")] {
            let d = Data(text.utf8)
            try chain.addEntry(with: name, type: .file, uncompressedSize: Int64(d.count), compressionMethod: .deflate) { p, n in d.subdata(in: Int(p)..<Int(p) + n) }
        }
        try chain.addEntry(with: "word/a", type: .symlink, uncompressedSize: Int64(dot.count)) { p, n in dot.subdata(in: Int(p)..<Int(p) + n) }
        let canaryName = "i137-canary-\(UUID().uuidString).txt"
        let canary = Data("written outside".utf8)
        try chain.addEntry(with: "word/a/a/a/../../../../\(canaryName)", type: .file, uncompressedSize: Int64(canary.count), compressionMethod: .deflate) { p, n in canary.subdata(in: Int(p)..<Int(p) + n) }
        let escaping = try XCTUnwrap(chain.data)
        XCTAssertThrowsError(try readerImageCount(escaping), "the reader refuses the chain")
        XCTAssertThrowsError(try PackageInspector.imageConsistencyReport(of: escaping), "the inspector refuses the chain") {
            // Ours, not ZIPFoundation's isContained (verify R5 DA-R5-2): the
            // symlink clause fires first; delete it and the `..` clause must.
            XCTAssertTrue(String(describing: $0).contains("symbolic-link") || String(describing: $0).contains("leaves its own directory"), "\($0)")
        }
        let temp = FileManager.default.temporaryDirectory
        for dir in [temp, temp.appendingPathComponent("che-word-mcp"), temp.deletingLastPathComponent()] {
            XCTAssertFalse(FileManager.default.fileExists(atPath: dir.appendingPathComponent(canaryName).path), "nothing written at \(dir.path)")
        }
    }

    func testFixedSlotCollisionIsReportedEvenWhenTheSameIdIsAlsoAModelDuplicate() throws {
        // codex R4-B8: the two causes are not a partition. Two images both on
        // rId1 are a model duplicate AND a collision with the styles slot.
        var doc = WordDocument()
        doc.body.children.append(.paragraph(Paragraph(runs: [Run(text: "x")])))
        doc.images = [ImageReference(id: "rId1", fileName: "a.png", contentType: "image/png", data: Data([0x89, 0x50])),
                      ImageReference(id: "rId1", fileName: "b.png", contentType: "image/png", data: Data([0x89, 0x50]))]
        var thrown: Error?
        XCTAssertThrowsError(try DocxWriter.writeData(doc)) { thrown = $0 }
        let message = (thrown as? LocalizedError)?.errorDescription ?? String(describing: thrown)
        XCTAssertTrue(message.contains("more than once"), message)
        XCTAssertTrue(message.contains("#140"), message)
        XCTAssertFalse(message.contains("well-formed"), "a document that carries the id twice is not called well-formed: \(message)")
    }

    func testAbsoluteNulAndDirectoryTraversalEntriesAreRefusedBeforeAnythingIsWritten() throws {
        // verify R5 (codex Z-R5-1): the absolute-path branch of the policy had
        // no test; a `..` directory entry is refused too. Empty and NUL paths
        // have their own test (testEmptyAndNulEntryPathsAreRefusedBeforeExtraction).
        let rel = #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#
        let base: [(String, String)] = [("word/document.xml", body(referencing: "rId4")), ("word/_rels/document.xml.rels", rels(rel)), ("word/media/image1.png", "png")]
        for (label, extra) in [("absolute file", ("/tmp/i137-abs-\(UUID().uuidString).xml", "<x/>")), ("absolute directory", ("/tmp/i137-absdir-\(UUID().uuidString)/", "")), ("dot-dot directory", ("word/../i137-dd-\(UUID().uuidString)/", ""))] {
            let data = try zipEntries(base + [extra])
            XCTAssertThrowsError(try readerImageCount(data), "\(label): the reader refuses")
            XCTAssertThrowsError(try PackageInspector.imageConsistencyReport(of: data), "\(label): the inspector refuses") {
                XCTAssertTrue(String(describing: $0).contains("leaves its own directory"), "\(label): \($0)")
            }
            let name = extra.0.trimmingCharacters(in: CharacterSet(charactersIn: "/")).components(separatedBy: "/").last!
            XCTAssertFalse(FileManager.default.fileExists(atPath: "/tmp/" + name), "\(label): nothing written at /tmp")
        }
    }

    func testTheDataAndURLEntryPointsExtractTheSameBytesAndAgree() throws {
        // verify R5 (codex S-R5-1 / security S-R5-1): the policy pre-scan and the
        // extraction must see one immutable byte sequence. Both entry points now
        // go through `ZipHelper.unzip(data:)`; their reports are identical, and
        // the Data overload writes nothing before the policy scan passes.
        let rel = #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#
        let data = try package(document: body(), docRels: rel)
        let url = FileManager.default.temporaryDirectory.appendingPathComponent("i137-\(UUID().uuidString).docx")
        try data.write(to: url); defer { try? FileManager.default.removeItem(at: url) }
        XCTAssertEqual(try PackageInspector.imageConsistencyReport(of: data), try PackageInspector.imageConsistencyReport(ofPackageAt: url))
        // A refused package leaves no extraction directory behind. Observed in
        // a namespace nobody else writes to (verify R6 logic N-L6-R6: counting
        // entries of the SHARED inspector namespace re-created the very
        // non-hermetic snapshot N-R5-1 was about); the internal entry point
        // is the same code path with the namespace as its only difference.
        let own = "i137-\(UUID().uuidString)"
        let ownDir = FileManager.default.temporaryDirectory.appendingPathComponent(own)
        defer { try? FileManager.default.removeItem(at: ownDir) }
        let link = Data("document.xml".utf8)
        let bad = try Archive(accessMode: .create)
        try bad.addEntry(with: "word/document.xml", type: .file, uncompressedSize: Int64(data.count), compressionMethod: .deflate) { p, n in data.subdata(in: Int(p)..<Int(p) + n) }
        try bad.addEntry(with: "word/alias.xml", type: .symlink, uncompressedSize: Int64(link.count)) { p, n in link.subdata(in: Int(p)..<Int(p) + n) }
        XCTAssertThrowsError(try ZipHelper.unzip(data: try XCTUnwrap(bad.data), namespace: own))
        XCTAssertEqual((try? FileManager.default.contentsOfDirectory(atPath: ownDir.path))?.count ?? 0, 0, "a refused package leaves no extraction directory behind")
        let good = try ZipHelper.unzip(data: data, namespace: own)
        XCTAssertTrue(good.path.hasPrefix(ownDir.path + "/"), "extraction lands in the requested namespace")
        ZipHelper.cleanup(good)
        XCTAssertEqual((try? FileManager.default.contentsOfDirectory(atPath: ownDir.path))?.count ?? 0, 0, "cleanup removes it")
        // The inspector's namespace is not the reader's (verify R5 regression
        // N-R5-1) — a fact of the constants, not of a shared directory count.
        XCTAssertNotEqual(ZipHelper.inspectorNamespace, ZipHelper.readerNamespace)
    }

    func testStructureNotesStayBoundedOnManyDistinctPrefixedElements() throws {
        // verify R5 logic L1: dedup by kind, not by an ever-growing list.
        let many = (1...20000).map { #"<p:e\#($0) xmlns:p="urn:p"/>"# }.joined()
        let relsXML = #"<Relationships xmlns="\#(pkgNS)">"# + many + "</Relationships>"
        let started = Date()
        let scan = PackageInspector.scanRels(Data(relsXML.utf8), part: "p")
        XCTAssertLessThan(Date().timeIntervalSince(started), 2.0)
        XCTAssertEqual(scan.structure.count, 1)
        XCTAssertTrue(scan.structure[0].hasPrefix("a namespace-prefixed <p:e1>"), scan.structure[0])
    }

    func testAFailedDirectoryListingIsNoVerdictNotAnEmptyPackage() throws {
        // verify R5 DA-R5-1: codex R4-B5's rule (a listing that fails throws
        // instead of reading as an empty package) had no test; reverting it
        // to `try? … ?? []` left the whole suite green. Inject the failure.
        struct ListingFailed: Error {}
        let rel = #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#
        let data = try package(document: body(), docRels: rel, extra: ["word/header1.xml": body(), "word/_rels/header1.xml.rels": rels(rel)])
        for (label, failSubpaths, failMedia) in [("word/ listing", true, false), ("word/media listing", false, true)] {
            var thrown: Error?
            XCTAssertThrowsError(try PackageInspector.imageConsistencyReport(
                extracting: { try ZipHelper.unzip(data: data, namespace: ZipHelper.inspectorNamespace) },
                listSubpaths: { url in if failSubpaths { throw ListingFailed() }; return try FileManager.default.subpathsOfDirectory(atPath: url.path) },
                listDirectory: { url in if failMedia { throw ListingFailed() }; return try FileManager.default.contentsOfDirectory(atPath: url.path) }
            ), label) { thrown = $0 }
            XCTAssertTrue(String(describing: thrown).contains("no consistency verdict"), "\(label): \(String(describing: thrown))")
        }
        // …and with working listings the same package reports its orphans.
        let report = try PackageInspector.imageConsistencyReport(of: data)
        XCTAssertEqual(report.orphanImageRelationshipRefs.count, 2)
    }

    // MARK: - verify R6 (codex R6-1 … R6-10)

    func testExtractionIsOwnerOnlyFromTheRootDownWhateverTheArchiveSays() throws {
        // codex R6-2 / R6-3: the root, every directory and every file end up
        // 0700 / 0600 even when the archive stores setuid or world-writable
        // bits; the private copy is gone; the namespace directory is 0700.
        let rel = #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#
        let archive = try Archive(accessMode: .create)
        // verify R6 security N4: a directory the archive stores as 0300 (writable,
        // not listable) holding a setuid file — the walk must reach inside it,
        // which it can only do by resetting the parent before descending.
        try archive.addEntry(with: "word/locked/", type: .directory, uncompressedSize: 0, permissions: 0o300, provider: { (_: Int64, _: Int) -> Data in Data() })
        let parts: [(String, String, UInt16)] = [("word/document.xml", body(referencing: "rId4"), 0o4755), ("word/_rels/document.xml.rels", rels(rel), 0o666), ("word/media/image1.png", "png", 0o777), ("word/locked/inner.bin", "x", 0o4755)]
        for (name, text, perms) in parts {
            let d = Data(text.utf8)
            try archive.addEntry(with: name, type: .file, uncompressedSize: Int64(d.count), permissions: perms, compressionMethod: .deflate, provider: { (p: Int64, n: Int) -> Data in d.subdata(in: Int(p)..<Int(p) + n) })
        }
        try archive.addEntry(with: "word/media/", type: .directory, uncompressedSize: 0, permissions: 0o1777, provider: { (_: Int64, _: Int) -> Data in Data() })
        let data = try XCTUnwrap(archive.data)
        let dir = try ZipHelper.unzip(data: data, namespace: ZipHelper.inspectorNamespace); defer { ZipHelper.cleanup(dir) }
        func mode(_ url: URL) throws -> Int { try XCTUnwrap(FileManager.default.attributesOfItem(atPath: url.path)[.posixPermissions] as? Int) }
        XCTAssertEqual(try mode(dir), 0o700, "root")
        XCTAssertEqual(try mode(dir.deletingLastPathComponent()), 0o700, "namespace directory")
        XCTAssertEqual(try mode(dir.appendingPathComponent("word")), 0o700, "directory")
        XCTAssertEqual(try mode(dir.appendingPathComponent("word/media")), 0o700, "directory with sticky+world-writable in the archive")
        XCTAssertEqual(try mode(dir.appendingPathComponent("word/document.xml")), 0o600, "setuid file")
        XCTAssertEqual(try mode(dir.appendingPathComponent("word/media/image1.png")), 0o600, "world-writable file")
        XCTAssertEqual(try mode(dir.appendingPathComponent("word/locked")), 0o700, "directory stored 0300 (unlistable)")
        XCTAssertEqual(try mode(dir.appendingPathComponent("word/locked/inner.bin")), 0o600, "setuid file inside the unlistable directory")
        XCTAssertFalse(try FileManager.default.contentsOfDirectory(atPath: dir.path).contains { $0.hasSuffix(".zip") }, "the private copy is removed")
        // Real work on such a package is unaffected.
        XCTAssertTrue(try PackageInspector.imageConsistencyReport(of: data).isConsistent)
    }

    func testEmptyAndNulEntryPathsAreRefusedBeforeExtraction() throws {
        // codex R6-10: the two policy members that had branches but no fixtures.
        let rel = #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#
        var executedCases = 0
        defer { XCTAssertEqual(executedCases, 2, "codex R7-8: both fixtures must actually be built and refused; a fixture ZIPFoundation will not write is not coverage") }
        for (label, name) in [("NUL in the path", "word/a\u{0}.xml"), ("empty path", "")] {
            let archive = try Archive(accessMode: .create)
            for (n, text) in [("word/document.xml", body(referencing: "rId4")), ("word/_rels/document.xml.rels", rels(rel)), ("word/media/image1.png", "png")] {
                let d = Data(text.utf8)
                try archive.addEntry(with: n, type: .file, uncompressedSize: Int64(d.count), compressionMethod: .deflate, provider: { (p: Int64, m: Int) -> Data in d.subdata(in: Int(p)..<Int(p) + m) })
            }
            let x = Data("<x/>".utf8)
            do {
                try archive.addEntry(with: name, type: .file, uncompressedSize: Int64(x.count), compressionMethod: .deflate, provider: { (p: Int64, m: Int) -> Data in x.subdata(in: Int(p)..<Int(p) + m) })
            } catch {
                continue   // ZIPFoundation will not even write such an entry; nothing to refuse
            }
            let data = try XCTUnwrap(archive.data)
            executedCases += 1
            XCTAssertThrowsError(try ZipHelper.unzip(data: data, namespace: ZipHelper.inspectorNamespace), label) {
                XCTAssertTrue(String(describing: $0).contains("empty or NUL-containing"), "\(label): \($0)")
            }
            XCTAssertThrowsError(try readerImageCount(data), "\(label): the reader refuses too")
        }
    }

    func testCharacterReferenceCauseIsNamedForQuotedAndSpacedSpellingsAndStaysLinear() throws {
        // codex R6-7: `Id = 'rId&#57;'` decodes to rId57 like `Id="rId&#57;"` does;
        // codex R6-1: the diagnosis is one pass over the text — 20 000 such ids
        // must not make the refusal quadratic; the message is capped.
        let theme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
        let message = try writerRefusal(mutatingRels: { url in
            let xml = try String(contentsOf: url, encoding: .utf8)
            try xml.replacingOccurrences(of: "</Relationships>", with: "<Relationship Id = 'rId&#57;' Type='\(theme)' Target='theme/theme1.xml'/></Relationships>").write(to: url, atomically: true, encoding: .utf8)
        })
        XCTAssertTrue(message.contains("rId9: written with a character or entity reference (`rId&#57;` in the file)"), message)
        // `rId&#49;<n>` decodes to `rId1<n>`: 20 000 distinct ids, every one spelled with a reference.
        let many = (1...20000).map { #"<Relationship Id="rId&#49;\#($0)" Type="\#(theme)" Target="theme/t\#($0).xml"/>"# }.joined()
        let started = Date()
        let big = try writerRefusal(mutatingRels: { url in
            let xml = try String(contentsOf: url, encoding: .utf8)
            try xml.replacingOccurrences(of: "</Relationships>", with: many + "</Relationships>").write(to: url, atomically: true, encoding: .utf8)
        })
        XCTAssertLessThan(Date().timeIntervalSince(started), 5.0, "one pass, not one pass per id")
        XCTAssertTrue(big.contains("…and"), "the message is capped: \(big.count) characters — \(big)")
        XCTAssertLessThan(big.count, 8000, "the message is capped")
    }

    func testAnyModelDuplicateMeansTheDocumentIsNotCalledWellFormed() throws {
        // codex R6-6: the model duplicate (rId5) and the slot collision (rId1)
        // are different ids; the document is still not well-formed.
        var doc = WordDocument()
        doc.body.children.append(.paragraph(Paragraph(runs: [Run(text: "x")])))
        doc.images = [ImageReference(id: "rId5", fileName: "a.png", contentType: "image/png", data: Data([0x89, 0x50])),
                      ImageReference(id: "rId5", fileName: "b.png", contentType: "image/png", data: Data([0x89, 0x50])),
                      ImageReference(id: "rId1", fileName: "c.png", contentType: "image/png", data: Data([0x89, 0x50]))]
        var thrown: Error?
        XCTAssertThrowsError(try DocxWriter.writeData(doc)) { thrown = $0 }
        let message = (thrown as? LocalizedError)?.errorDescription ?? String(describing: thrown)
        XCTAssertTrue(message.contains("rId5") && message.contains("rId1") && message.contains("#140"), message)
        XCTAssertFalse(message.contains("well-formed"), message)
        XCTAssertFalse(message.contains("RId"), "an id is never capitalized by sentence-casing (logic N-L2-R6): \(message)")
        XCTAssertTrue(message.contains(". The document model carries"), "a cause sentence starts capitalized (logic N-L5-R7): \(message)")
    }

    func testThousandsOfMismatchedIdsAreRefusedInLinearTime() throws {
        // verify R6 DA: an N-id rels whose every Id the text scan cannot read was
        // O(N²) in the R6 snapshot (400 → 1.0 s, 800 → 3.9 s, 1600 → 202 s as the
        // DA measured); the cause pass is one scan, capped at 20, and nothing is
        // computed past the cap. Two spellings: all character references (the
        // single-pass map answers) and single quotes (the per-id regex path).
        let spellings: [(String, (Int) -> String)] = [
            ("character references", { n in "\"" + "rId\(n)".unicodeScalars.map { "&#\($0.value);" }.joined() + "\"" }),
            ("single quotes", { n in "'rId\(n)'" }),
        ]
        for (label, spell) in spellings {
            let relsXML = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + (0..<3200).map { "<Relationship Id=\(spell($0)) Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com/\($0)\" TargetMode=\"External\"/>" }.joined()
                + "</Relationships>"
            let start = Date()
            let message = try writerRefusal { try relsXML.write(to: $0, atomically: true, encoding: .utf8) }
            let elapsed = Date().timeIntervalSince(start)
            XCTAssertTrue(message.contains("does not match"), "\(label): \(message.prefix(200))")
                XCTAssertTrue(message.range(of: #"…and [0-9]+ more"#, options: .regularExpression) != nil, "\(label): capped at 20 causes: \(message.suffix(160))")
                XCTAssertLessThan(message.count, 6000, "\(label): capped message, got \(message.count) characters")
            XCTAssertLessThan(elapsed, 10, "\(label): 3200 mismatched ids must be refused in linear time (took \(elapsed) s; the bound is load-insensitive — the R6 snapshot took 60+ s here; linearity itself is the release probe in the CHANGELOG)")
        }
    }

    func testAPlantedFileOrLinkAtTheNamespaceIsRefusedAndOurOwnWrongModeIsRepaired() throws {
        // codex R7-1: the namespace is created with mkdir(0700) or found, opened
        // O_NOFOLLOW|O_DIRECTORY, and the DESCRIPTOR is verified — a file or a
        // link planted at the name is refused; our own directory with another
        // mode is reset to 0700.
        let data = try zipEntries([("word/document.xml", body())])
        let tmp = FileManager.default.temporaryDirectory
        let plants: [(String, (URL) throws -> Void)] = [
            ("regular file", { try Data("x".utf8).write(to: $0) }),
            ("symbolic link", { try FileManager.default.createSymbolicLink(at: $0, withDestinationURL: tmp) }),
        ]
        for (label, plant) in plants {
            let ns = "i137-ns-\(UUID().uuidString)"
            let planted = tmp.appendingPathComponent(ns)
            try plant(planted); defer { try? FileManager.default.removeItem(at: planted) }
            XCTAssertThrowsError(try ZipHelper.unzip(data: data, namespace: ns), label) {
                XCTAssertTrue(String(describing: $0).contains("refused") || String(describing: $0).contains("not a directory"), "\(label): \($0)")
            }
            XCTAssertTrue(FileManager.default.fileExists(atPath: planted.path), "\(label): the planted item is left alone")
        }
        let ns = "i137-ns-\(UUID().uuidString)"
        let ours = tmp.appendingPathComponent(ns)
        try FileManager.default.createDirectory(at: ours, withIntermediateDirectories: false, attributes: [.posixPermissions: 0o755])
        defer { try? FileManager.default.removeItem(at: ours) }
        let out = try ZipHelper.unzip(data: data, namespace: ns); ZipHelper.cleanup(out)
        XCTAssertEqual(try FileManager.default.attributesOfItem(atPath: ours.path)[.posixPermissions] as? Int, 0o700, "our own namespace is reset to 0700")
    }

    func testOtherAttributesNamedLikeIdAreNotTheRelationshipIdAndTheCauseStaysRight() throws {
        // codex R7-2 / R7-7: 3000 Relationship tags each carrying a `data-Id`
        // attribute spelled as a distinct reference form of "rId9" (an XML Name
        // may contain `-`; the old `(?<![:\w])Id` boundary read them all as Id and
        // grew one bucket quadratically), plus one real single-quoted Id='rId9'.
        // The real cause is the single quote, not the references; and it is linear.
        // (`data-Id` sits after `Id` here because the text scan's own attribute
        // regex — #142 — takes the first `\bId="` in a tag; that looseness is
        // #142's, and the message must not compound it by naming a reference.)
        var relationships = (0..<3000).map { i -> String in
            let zeros = String(repeating: "0", count: i % 50)
            let spell = "rId9".unicodeScalars.map { "&#\(zeros)\($0.value);" }.joined() + "&#\(String(repeating: "0", count: i / 50 + 1))59;"
            return "<Relationship Id=\"rIdA\(i)\" data-Id=\"\(spell)\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com/\(i)\" TargetMode=\"External\"/>"
        }
        relationships.append("<Relationship Id='rId9' Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com/9\" TargetMode=\"External\"/>")
        let relsXML = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" + relationships.joined() + "</Relationships>"
        let start = Date()
        let message = try writerRefusal { try relsXML.write(to: $0, atomically: true, encoding: .utf8) }
        let elapsed = Date().timeIntervalSince(start)
        XCTAssertTrue(message.contains("rId9: single-quoted attribute values"), message.suffix(300).description)
        XCTAssertFalse(message.contains("reference"), "the data-Id spellings are not causes: \(message.suffix(300))")
        XCTAssertLessThan(elapsed, 10, "took \(elapsed) s (load-insensitive bound; quadratic would be minutes)")
        XCTAssertEqual(DocxWriter.rawSpellingsByDecodedId(inRaw: #"<x foo.Id="rId&#57;" data-Id="rId&#57;"/><Relationship r:Id="rId&#57;" Id="rId&#57;"/>"#), ["rId9": ["rId&#57;"]])
    }

    func testSameMultisetInAnotherOrderIsRefusedWithACappedMessage() throws {
        // codex R7-3: when the text scan and the parser see the same ids the
        // same number of times but in another order, the pairwise fallback is
        // capped like the causes are.
        let n = 60
        let relationships = (0..<n).map { i -> String in
            let mine = "rId\(i)", other = "rId\((i + 1) % n)"
            return "<Relationship data-Id=\"\(other)\" Id='\(mine)' Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com/\(i)\" TargetMode=\"External\"/>"
        }
        let relsXML = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" + relationships.joined() + "</Relationships>"
        let message = try writerRefusal { try relsXML.write(to: $0, atomically: true, encoding: .utf8) }
        XCTAssertTrue(message.contains("does not match"), message.prefix(200).description)
        XCTAssertTrue(message.contains("…and \(n - 20) more"), message.suffix(200).description)
        XCTAssertLessThan(message.count, 3000, "capped: \(message.count) characters")
    }

    func testAHugeSpellingCannotInflateTheMessage() throws {
        // codex R7-3: one reference may carry any number of leading zeros — the
        // displayed spelling is truncated; a spelling past 4 KB is not even decoded.
        for zeros in [3000, 100_000] {
            let spell = "&#\(String(repeating: "0", count: zeros))114;Id9"
            let relsXML = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"\(spell)\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com/\" TargetMode=\"External\"/></Relationships>"
            let message = try writerRefusal { try relsXML.write(to: $0, atomically: true, encoding: .utf8) }
            XCTAssertLessThan(message.count, 1200, "\(zeros) zeros: \(message.count) characters")
            if zeros == 3000 { XCTAssertTrue(message.contains("more characters)"), message.suffix(200).description) }
        }
    }

    func testAnUnremovableDirectoryEntryLeavesNothingBehindAndTheErrorNamesNoPath() throws {
        // verify R7 logic N-L1-R7 / N-L2-R7: a directory entry stored 0400 or 0500
        // ahead of its file makes ZIPFoundation's extraction fail; the partial
        // tree (with the private copy of the whole package) used to stay behind
        // because `removeItem` cannot empty such a directory. Now the tree is
        // made ours again before removal, the error is ours, and it names no path.
        let rel = #"<Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/>"#
        for mode: UInt16 in [0o400, 0o500] {
            let archive = try Archive(accessMode: .create)
            try archive.addEntry(with: "word/media/", type: .directory, uncompressedSize: 0, permissions: mode, provider: { (_: Int64, _: Int) -> Data in Data() })
            for (name, text) in [("word/document.xml", body(referencing: "rId4")), ("word/_rels/document.xml.rels", rels(rel)), ("word/media/image1.png", "SECRET-BODY-BYTES")] {
                let d = Data(text.utf8)
                try archive.addEntry(with: name, type: .file, uncompressedSize: Int64(d.count), compressionMethod: .deflate, provider: { (p: Int64, n: Int) -> Data in d.subdata(in: Int(p)..<Int(p) + n) })
            }
            let data = try XCTUnwrap(archive.data)
            let ns = "i137-ns-\(UUID().uuidString)"
            let nsDir = FileManager.default.temporaryDirectory.appendingPathComponent(ns)
            defer { try? FileManager.default.removeItem(at: nsDir) }
            XCTAssertThrowsError(try ZipHelper.unzip(data: data, namespace: ns), "mode \(String(mode, radix: 8))") { error in
                let message = String(describing: error)
                XCTAssertTrue(message.contains("could not be extracted"), "mode \(String(mode, radix: 8)): \(message)")
                XCTAssertFalse(message.contains(nsDir.path) || message.contains("/var/") || message.contains("/private/"), "no path in the message: \(message)")
            }
            let leftovers = (try? FileManager.default.contentsOfDirectory(atPath: nsDir.path)) ?? []
            XCTAssertEqual(leftovers, [], "mode \(String(mode, radix: 8)): nothing stays behind (private copy included)")
            XCTAssertThrowsError(try readerImageCount(data), "the reader refuses too") { XCTAssertTrue($0 is WordError, "the reader's error is ours: \($0)") }
        }
    }

    func testADecoyInsideAnotherAttributeValueDoesNotChangeTheCause() throws {
        // verify R7 logic N-L3-R7: only the tag's own Id attribute is a cause;
        // an `Id='…'` inside a Target value is data.
        let base = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
        let cases: [(String, String, String)] = [
            ("single quote, decoy reference in Target",
             "<Relationship Id='rId9' Type=\"\(base)\" Target=\"https://example.com/?q= Id='rId&#57;'\" TargetMode=\"External\"/>",
             "rId9: single-quoted attribute values"),
            ("whitespace around =, decoy plain Id in Target",
             "<Relationship Id = \"rId9\" Type=\"\(base)\" Target=\"https://example.com/? Id='rId9'\" TargetMode=\"External\"/>",
             "rId9: whitespace around `=`"),
            ("> inside a value ends the text scan's tag early",
             "<Relationship Type=\"\(base)\" Target=\"https://example.com/a>b\" Id=\"rId9\" TargetMode=\"External\"/>",
             "rId9: an attribute value containing `>`"),
        ]
        for (label, relationship, expected) in cases {
            let relsXML = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" + relationship + "</Relationships>"
            let message = try writerRefusal { try relsXML.write(to: $0, atomically: true, encoding: .utf8) }
            XCTAssertTrue(message.contains(expected), "\(label): \(message.suffix(300))")
            if label.hasPrefix("single") { XCTAssertFalse(message.contains("reference"), "\(label): the decoy is not a cause: \(message.suffix(300))") }
        }
    }

    func testANamespaceUriIsNotReportedAsARelationshipId() throws {
        // verify R7 logic N-L4-R7: the overlay's own attribute regex reads names
        // at attribute-name position only, so `xmlns:Id="urn:zz"` is not an Id.
        let relsXML = #"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship xmlns:Id="urn:zz" Id="rId&#57;" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/></Relationships>"#
        let message = try writerRefusal { try relsXML.write(to: $0, atomically: true, encoding: .utf8) }
        XCTAssertFalse(message.contains("urn:zz"), message.suffix(300).description)
        XCTAssertTrue(message.contains("both see 1 relationship,"), "singular (logic N-L6-R7): \(message.suffix(300))")
        XCTAssertTrue(message.contains("rId9: written with a character or entity reference"), message.suffix(300).description)
    }

    func testAPrefixedAttributeBeforeTheRealIdSaves() throws {
        // verify R7 requirements N-R7-1: `xmlns:Id="urn:x" Id="rId4"` and
        // `r:Id="zzz" Id="rId4"` are well-formed, the reader opens them, and the
        // Id is the plainest spelling there is — the text scan used to read the
        // prefixed attribute as the Id and refuse. Now both scanners agree.
        let base = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
        for shape in ["<Relationship xmlns:Id=\"urn:x\" Id=\"rId4\" Type=\"\(base)\" Target=\"media/image1.png\"/>",
                      "<Relationship xmlns:r=\"urn:r\" r:Id=\"zzz\" Id=\"rId4\" Type=\"\(base)\" Target=\"media/image1.png\"/>"] {
            let relsXML = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" + shape + "</Relationships>"
            let data = try package(document: body(referencing: "rId4"), docRels: relsXML)
            let url = FileManager.default.temporaryDirectory.appendingPathComponent("i137-tenth-\(UUID().uuidString).docx")
            try data.write(to: url); defer { try? FileManager.default.removeItem(at: url) }
            var read = try DocxReader.read(from: url); defer { read.close() }
            XCTAssertNoThrow(try DocxWriter.writeData(read), shape)
        }
    }

    func testAnUnlistableDirectoryWithContentIsRemovedAfterALaterFailure() throws {
        // verify R7 security S-R7-3: `keep/` stored 0300 (writable, unlistable)
        // and already holding a file, then `boom/` stored 0400 whose file cannot
        // be written — the failure leaves a tree `removeItem` cannot empty
        // unless every directory is made ours again first.
        let archive = try Archive(accessMode: .create)
        try archive.addEntry(with: "keep/", type: .directory, uncompressedSize: 0, permissions: 0o300, provider: { (_: Int64, _: Int) -> Data in Data() })
        let payload = Data("payload".utf8)
        try archive.addEntry(with: "keep/payload.bin", type: .file, uncompressedSize: Int64(payload.count), compressionMethod: .deflate, provider: { (p: Int64, n: Int) -> Data in payload.subdata(in: Int(p)..<Int(p) + n) })
        try archive.addEntry(with: "boom/", type: .directory, uncompressedSize: 0, permissions: 0o400, provider: { (_: Int64, _: Int) -> Data in Data() })
        try archive.addEntry(with: "boom/x.bin", type: .file, uncompressedSize: Int64(payload.count), compressionMethod: .deflate, provider: { (p: Int64, n: Int) -> Data in payload.subdata(in: Int(p)..<Int(p) + n) })
        let data = try XCTUnwrap(archive.data)
        let ns = "i137-ns-\(UUID().uuidString)"
        let nsDir = FileManager.default.temporaryDirectory.appendingPathComponent(ns)
        defer { try? FileManager.default.removeItem(at: nsDir) }
        XCTAssertThrowsError(try ZipHelper.unzip(data: data, namespace: ns)) { XCTAssertTrue($0 is WordError, "\($0)") }
        XCTAssertEqual((try? FileManager.default.contentsOfDirectory(atPath: nsDir.path)) ?? ["(missing)"], [], "nothing stays behind")
    }

    func testCleanupNeverFollowsASymbolicLink() throws {
        // verify R8 security/logic: `removeTreeForcibly` used `chmod`, which
        // follows a link — a link planted inside a failed extraction changed
        // the mode of a file and a directory OUTSIDE the tree. Now lstat +
        // fchmodat(AT_SYMLINK_NOFOLLOW): the link is removed with the tree,
        // never followed.
        let tmp = FileManager.default.temporaryDirectory
        let outside = tmp.appendingPathComponent("i137-outside-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: outside, withIntermediateDirectories: false, attributes: [.posixPermissions: 0o755])
        defer { try? FileManager.default.removeItem(at: outside) }
        let outsideFile = outside.appendingPathComponent("target.txt")
        try Data("t".utf8).write(to: outsideFile); try FileManager.default.setAttributes([.posixPermissions: 0o644], ofItemAtPath: outsideFile.path)
        let root = tmp.appendingPathComponent("i137-tree-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: root.appendingPathComponent("sub"), withIntermediateDirectories: true)
        try FileManager.default.createSymbolicLink(at: root.appendingPathComponent("sub/link-file"), withDestinationURL: outsideFile)
        try FileManager.default.createSymbolicLink(at: root.appendingPathComponent("link-dir"), withDestinationURL: outside)
        try FileManager.default.setAttributes([.posixPermissions: 0o500], ofItemAtPath: root.appendingPathComponent("sub").path)   // unremovable until reset
        ZipHelper.removeTreeForcibly(root)
        func mode(_ url: URL) throws -> Int { try XCTUnwrap(FileManager.default.attributesOfItem(atPath: url.path)[.posixPermissions] as? Int) }
        XCTAssertFalse(FileManager.default.fileExists(atPath: root.path), "the tree (0500 directory included) is removed")
        XCTAssertEqual(try mode(outside), 0o755, "the directory behind a link is untouched")
        XCTAssertEqual(try mode(outsideFile), 0o644, "the file behind a link is untouched")
    }

    func testACreationFailureNamesNoPathAndLeavesNoReport() throws {
        // verify R8 logic: with the namespace made immutable, `mkdirat` fails —
        // the error must not carry the temporary path, and nothing is removed
        // (nor reported) because nothing was created.
        let data = try zipEntries([("word/document.xml", body())])
        let ns = "i137-ns-\(UUID().uuidString)"
        let nsDir = FileManager.default.temporaryDirectory.appendingPathComponent(ns)
        try FileManager.default.createDirectory(at: nsDir, withIntermediateDirectories: false, attributes: [.posixPermissions: 0o700])
        guard chflags(nsDir.path, UInt32(UF_IMMUTABLE)) == 0 else { throw XCTSkip("cannot set UF_IMMUTABLE here") }
        defer { chflags(nsDir.path, 0); try? FileManager.default.removeItem(at: nsDir) }
        XCTAssertThrowsError(try ZipHelper.unzip(data: data, namespace: ns)) { error in
            let message = String(describing: error)
            XCTAssertTrue(message.contains("could not create the extraction directory"), message)
            XCTAssertFalse(message.contains("/"), "no path in the message: \(message)")
        }
    }

    func testErrorDescriptionsAreBuiltFromCodesNotText() {
        // verify R8: a token filter on the error text kept quoted paths and dropped
        // words that merely began with `/`; descriptions now come from the code.
        let posix = NSError(domain: NSPOSIXErrorDomain, code: 13, userInfo: [NSFilePathErrorKey: "/var/x/y.zip"])
        XCTAssertEqual(ZipHelper.describeWithoutPaths(posix), "Permission denied")
        let cocoa = NSError(domain: NSCocoaErrorDomain, code: 513, userInfo: [NSLocalizedDescriptionKey: "“/var/x/y.zip” couldn’t be removed.", NSFilePathErrorKey: "/var/x/y.zip"])
        XCTAssertEqual(ZipHelper.describeWithoutPaths(cocoa), "permission denied")
        let wrapped = NSError(domain: NSCocoaErrorDomain, code: 4, userInfo: [NSUnderlyingErrorKey: NSError(domain: NSPOSIXErrorDomain, code: 2, userInfo: nil), NSFilePathErrorKey: "/var/x"])
        XCTAssertEqual(ZipHelper.describeWithoutPaths(wrapped), "No such file or directory")
        XCTAssertFalse(ZipHelper.describeWithoutPaths(NSError(domain: "Other", code: 7, userInfo: [NSLocalizedDescriptionKey: "at /tmp/z"])).contains("/"))
    }

    func testAnEntryNameCannotForgeOrFloodTheRefusalMessage() throws {
        // verify R8 security N-S8-1: an entry name is attacker-controlled text.
        let esc = "\u{1B}"
        let names = ["word/../\nooxml-swift: everything is fine, extraction succeeded\n../x.xml",
                     "word/../" + esc + "[31mred" + esc + "[0m\r../y.xml",
                     "word/../" + String(repeating: "a", count: 9000) + "/z.xml"]
        for name in names {
            let data = try zipEntries([("word/document.xml", body()), (name, "<x/>")])
            XCTAssertThrowsError(try ZipHelper.unzip(data: data, namespace: ZipHelper.inspectorNamespace)) { error in
                let message = (error as? LocalizedError)?.errorDescription ?? String(describing: error)
                XCTAssertFalse(message.contains("\n") || message.contains("\r") || message.contains(esc), "escaped: \(message.prefix(200))")
                XCTAssertLessThan(message.count, 400, "truncated: \(message.count) characters")
                XCTAssertTrue(message.contains("leaves its own directory"), message.prefix(200).description)
            }
        }
    }

    func testTheTextScanNeverReadsInsideAnotherAttributesValue() throws {
        // verify R8 requirements N-R8-2: `Target='x Id="HIJACK"' Id="rId4"` — the
        // text scan used to read HIJACK; now it walks attribute by attribute and
        // reads rId4, so both scans agree and the package saves.
        let base = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
        XCTAssertEqual(RelationshipsOverlay.attribute(#"Target='x Id="HIJACK"' Id="rId4" Type="t""#, name: "Id"), "rId4")
        XCTAssertNil(RelationshipsOverlay.attribute(#"Id='rId4' Type="t""#, name: "Id"), "single-quoted: present but unreadable")
        XCTAssertNil(RelationshipsOverlay.attribute(#"Id = "rId4" Type="t""#, name: "Id"), "spaced: present but unreadable")
        XCTAssertNil(RelationshipsOverlay.attribute(#"xmlns:Id="urn:x" data-Id="d" Type="t""#, name: "Id"), "prefixed names are other attributes")
        XCTAssertEqual(RelationshipsOverlay.attribute(#"Type="t" Target="a>b" Id="rId4""#, name: "Target"), "a>b")
        // verify R8 logic NEW-L3: `<Relationship-2>` is another element to both scans.
        XCTAssertEqual(RelationshipsOverlay.rawIds(inRelsXML: #"<Relationships><Relationship-2 Id="rIdH" Type="t" Target="x"/><Relationship Id="rId4" Type="t" Target="y"/></Relationships>"#), ["rId4"])
        let relsXML = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Target='media/image1.png?q= Id=\"HIJACK\"' Id=\"rId4\" Type=\"\(base)\"/></Relationships>"
        let data = try package(document: body(referencing: "rId4"), docRels: relsXML)
        let url = FileManager.default.temporaryDirectory.appendingPathComponent("i137-hijack-\(UUID().uuidString).docx")
        try data.write(to: url); defer { try? FileManager.default.removeItem(at: url) }
        var read = try DocxReader.read(from: url); defer { read.close() }
        // The reader keeps the single-quoted Target, and the text scan refuses
        // the tag as unreadable — the Id agrees, the Target does not: refused
        // with a cause that names the sibling attribute, never HIJACK.
        XCTAssertThrowsError(try DocxWriter.writeData(read)) { error in
            let message = String(describing: error)
            XCTAssertFalse(message.contains("HIJACK"), message.suffix(300).description)
            XCTAssertTrue(message.contains("`Target` attribute is single-quoted"), message.suffix(300).description)
        }
    }

    func testAValueContainingASelfClosingSequenceIsRefusedNotTruncated() throws {
        // verify R9 logic / requirements N-R9-1 (a fix-round-9 regression): the tag
        // regex is lazy up to `/>`, so `/>` inside the LAST attribute's value cuts
        // the attribute text mid-value. The tokenizer once returned the prefix and
        // the merge wrote a truncated Target (and dropped TargetMode) with exit 0.
        // Now an unterminated value makes the whole tag untrusted → refused.
        XCTAssertNil(RelationshipsOverlay.tokenize(#"Id="rId9" Type="t" Target="theme/a/"#), "unterminated value: no tokens at all")
        XCTAssertNil(RelationshipsOverlay.attribute(#"Id="rId9" Type="t" Target="theme/a/"#, name: "Id"), "even an earlier, intact Id is not read from a cut tag")
        let base = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
        for relationship in ["<Relationship Id=\"rId9\" Type=\"\(base)\" Target=\"settings.xml?a=1/>2\"/>",
                             "<Relationship Id=\"rId9\" Type=\"\(base)\" Target=\"https://example.com/q?a=1/>2\" TargetMode=\"External\"/>",
                             "<Relationship Id=\"rId9\" Type=\"\(base)/>evil\" Target=\"x\"/>"] {
            let relsXML = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" + relationship + "</Relationships>"
            let message = try writerRefusal { try relsXML.write(to: $0, atomically: true, encoding: .utf8) }
            XCTAssertTrue(message.contains("does not match"), message.prefix(200).description)
            XCTAssertTrue(message.contains("rId9: an attribute value containing `>`"), message.suffix(300).description)
        }
    }

    func testTheRelativeNameIsComputedUnderTheResolvedRoot() throws {
        // verify R8 logic NEW-L1 / R9 NEW-R9-2: `temporaryDirectory` is `/var/…`, an
        // enumerator returns `/private/var/…`, and Foundation's resolver maps both
        // to the `/var/…` form — the previous prefix set never matched.
        let root = FileManager.default.temporaryDirectory.appendingPathComponent("i137-rel-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: root.appendingPathComponent("word/_rels"), withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: root) }
        try Data("x".utf8).write(to: root.appendingPathComponent("word/_rels/document.xml.rels"))
        let walker = try XCTUnwrap(FileManager.default.enumerator(at: root, includingPropertiesForKeys: nil))
        var names: [String] = []
        for case let item as URL in walker { names.append(ZipHelper.relativeName(of: item, under: root)) }
        XCTAssertEqual(names.sorted(), ["word", "word/_rels", "word/_rels/document.xml.rels"])
    }

    func testTheSiblingCauseIsFoundByTokenizingNotByRegex() throws {
        // verify R9 logic: `Target="http://x/?a='b'"` must not name `a`; the real
        // unreadable sibling is `Type`.
        XCTAssertEqual(DocxWriter.firstUnreadableAttribute(in: #" Target="http://x/?a='b'" Type='t'"#), "Type")
        XCTAssertNil(DocxWriter.firstUnreadableAttribute(in: #" Target="http://x/?a='b'" Type="t""#))
        XCTAssertNil(DocxWriter.firstUnreadableAttribute(in: #" Target="cut/"#), "a cut fragment names nothing")
    }

    func testDisplayNameEscapesLineSeparatorsAndCapsTheRenderedText() {
        // verify R9 security N-S9-2: U+2028 / U+2029 are line breaks to a renderer;
        // R9 logic NEW-R9-4: escapes expand, so the rendered text has its own cap.
        let shown = ZipHelper.displayName("a\u{2028}b\u{2029}c\u{200B}d\u{202E}e\u{FEFF}f")
        XCTAssertFalse(shown.contains("\u{2028}") || shown.contains("\u{2029}") || shown.contains("\u{200B}") || shown.contains("\u{202E}") || shown.contains("\u{FEFF}"), shown)
        XCTAssertTrue(shown.contains("\\u{2028}"), shown)
        XCTAssertLessThanOrEqual(ZipHelper.displayName(String(repeating: "\u{01}", count: 9000)).count, 481)
    }

    func testDisplayNameEscapesByUnicodePropertyNotByAHandWrittenList() {
        // verify R10 requirements N-R10-2 / security N-S10-3: the R9 hand-written
        // list missed U+00AD, U+061C, U+180E, U+FFF9 …; the criterion is now a
        // Unicode property (Cc / Cf / Zl / Zp / Default_Ignorable_Code_Point).
        for scalar in ["\u{00AD}", "\u{061C}", "\u{180E}", "\u{FFF9}", "\u{2060}", "\u{3164}", "\u{FE0F}", "\u{E0001}", "\u{2028}", "\u{200B}", "\u{202E}", "\u{FEFF}"] {
            let shown = ZipHelper.displayName("a\(scalar)b")
            XCTAssertFalse(shown.contains(scalar), "escaped: \(shown)")
            XCTAssertTrue(shown.hasPrefix("a\\u{") && shown.hasSuffix("}b"), shown)
        }
        XCTAssertEqual(ZipHelper.displayName("標楷體 😀 café rId9"), "標楷體 😀 café rId9", "visible text with no invisible joiner passes through untouched")
        // Not "all visible text": an emoji sequence's variation selector and ZWJ
        // are Cf / Default_Ignorable, so they ARE escaped — deliberately, in an
        // error message (verify R11 logic N-L11-4, regression N-REG11-4).
        XCTAssertEqual(ZipHelper.displayName("\u{2764}\u{FE0F}"), "\u{2764}\\u{FE0F}")
        // The escape character and the delimiter are escaped too, so a rendering
        // maps back to one input (R11 security N-S11-1 / N-S11-3).
        XCTAssertNotEqual(ZipHelper.displayName("a\\u{202E}b"), ZipHelper.displayName("a\u{202E}b"), "a literal backslash is not mistaken for an escape")
        XCTAssertFalse(ZipHelper.displayName("rId`9; ignore the rest `x").contains("`"), "a backtick cannot close the message's code span")
    }

    func testAnAttackersOwnAttributeNameNeverReachesTheRefusalMessage() throws {
        // verify R10 security NEW-S10-1 asked for the sibling attribute name to be
        // capped; verify R11 DA N-DA11-3 removed the need for a cap at this site:
        // only the four attributes the merge READS can make it skip a tag, so only
        // those may be named. An attacker's own 9000-character attribute name is
        // therefore not merely truncated — it never appears, and the tag it sits
        // on is not refused for its sake either.
        let base = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
        let longName = String(repeating: "a", count: 9000)
        // (a) a long unreadable NON-gated attribute alongside an unreadable gated one:
        //     the message names the gated one, and carries none of the attacker's name.
        let mixed = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId9\" \(longName)='v' Type='\(base)' Target=\"x\" TargetMode=\"External\"/></Relationships>"
        let message = try writerRefusal { try mixed.write(to: $0, atomically: true, encoding: .utf8) }
        XCTAssertTrue(message.contains("`Type` attribute is single-quoted or spaced"), message.prefix(400).description)
        XCTAssertFalse(message.contains("aaaaaaaaaa"), "the attacker's attribute name is absent, not truncated")
        XCTAssertLessThan(message.count, 1200, "message length: \(message.count)")
        // (b) the same long attribute with every gated attribute readable: saved, so
        //     the tag was never skipped for it — which is why naming it was a lie.
        XCTAssertNil(DocxWriter.firstUnreadableAttribute(in: " Id=\"rId9\" \(longName)='v' Type=\"t\" Target=\"x\""))
    }

    func testADecodedIdCannotInjectALineIntoTheMessage() throws {
        // verify R9 security N-S9-1: the parser decodes `&#10;` inside an Id to a
        // real newline; the message shows ids as escaped text.
        let base = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
        let relsXML = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId&#10;ooxml-swift: all good&#10;9\" Type=\"\(base)\" Target=\"x\" TargetMode=\"External\"/></Relationships>"
        let message = try writerRefusal { try relsXML.write(to: $0, atomically: true, encoding: .utf8) }
        XCTAssertFalse(message.contains("\n"), "no real newline in the message: \(message.prefix(300))")
        XCTAssertTrue(message.contains("\\n"), message.suffix(300).description)
    }

    func testAPrefixedIdAttributeIsNotTheRelationshipId() throws {
        // logic N-L4-R6: `r:Id="…"` / `xmlns:Id="…"` are not the Id attribute; the
        // spelling map must not attribute a cause to them.
        let raw = #"<Relationships xmlns="urn:p"><Relationship xmlns:Id="urn:x" r:Id="rId&#57;" Id="rId9" Type="t" Target="a"/></Relationships>"#
        XCTAssertEqual(DocxWriter.rawSpellingsByDecodedId(inRaw: raw), [:], "the only real Id is spelled plainly")
        let raw2 = #"<Relationship r:Id="rId9" Id="rId&#57;" Type="t" Target="a"/>"#
        XCTAssertEqual(DocxWriter.rawSpellingsByDecodedId(inRaw: raw2), ["rId9": ["rId&#57;"]])
    }

    // MARK: - verify R10 Devil's Advocate: four mutation probes stayed green (N-DA10-2..5) + N-DA10-1

    func testADecoyElementNamedLikeRelationshipCannotSupplyTheCause() throws {
        // N-DA10-4 (mutation M5): with `\b` as the tag boundary, a `<Relationship-2 …/>`
        // decoy placed BEFORE the real tag became an occurrence — the cause named its
        // `-2` "attribute" and a spelling that is not in any real tag was reported.
        let raw = #"<Relationships xmlns="urn:p"><Relationship-2 Id="rId&#57;" Type="t" Target="a"/><Relationship Id = "rId9" Type="t" Target="b"/></Relationships>"#
        let occurrences = DocxWriter.relationshipIdOccurrences(inRaw: raw)
        XCTAssertEqual(occurrences["rId9"]?.count, 1, "\(occurrences)")
        XCTAssertEqual(occurrences["rId9"]?.first?.whitespaceAroundEquals, true)
        XCTAssertNil(occurrences["rId9"]?.first?.unreadableSiblingAttribute, "the decoy's `-2` is not a sibling of the real tag")
        XCTAssertEqual(DocxWriter.rawSpellingsByDecodedId(inRaw: raw), [:], "no spelling that only a decoy carries")
        XCTAssertEqual(DocxWriter.relsSpellingCause(forParsedId: "rId9", occurrences: occurrences), "whitespace around `=`")
    }

    func testTheDecodeBudgetBoundsHowManyReferenceSpellingsAreDecoded() {
        // N-DA10-5 (mutation M6): 4000 distinct reference spellings — only `decodeBudget`
        // of them are handed to the parser; the rest are skipped, not decoded later.
        var raw = #"<Relationships xmlns="urn:p">"#
        for i in 0..<4000 { raw += "<Relationship Id=\"rId&#49;\(i)\" Type=\"t\" Target=\"a\"/>" }
        raw += "</Relationships>"
        let occurrences = DocxWriter.relationshipIdOccurrences(inRaw: raw)
        XCTAssertEqual(occurrences.count, DocxWriter.decodeBudget, "decoded exactly `decodeBudget` spellings, not \(occurrences.count)")
        XCTAssertLessThan(occurrences.count, 4000)
    }

    func testThePrivateCopyIsOwnerOnlyForAsLongAsItExists() throws {
        // N-DA10-2 (mutation M2): the copy is created by openat(O_CREAT|O_EXCL|O_NOFOLLOW, 0o600);
        // a `createFile(attributes: nil)` copy is 0644 for its whole lifetime. The copy lives
        // for the length of the extraction, so a poller sees it: every sample must be 0600.
        let ns = "i137-copy-\(UUID().uuidString)"
        let nsDir = FileManager.default.temporaryDirectory.appendingPathComponent(ns)
        defer { try? FileManager.default.removeItem(at: nsDir) }
        let data = try zipEntries([("word/document.xml", body()), ("word/media/blob.bin", String(repeating: "0123456789abcdef", count: 3_000_000))])   // 48 MB to extract
        final class Samples { var modes: [mode_t] = []; let lock = NSLock() }
        let samples = Samples()
        let poller = Thread {
            while !Thread.current.isCancelled {
                for uuid in (try? FileManager.default.contentsOfDirectory(atPath: nsDir.path)) ?? [] {
                    let dir = nsDir.appendingPathComponent(uuid)
                    for name in (try? FileManager.default.contentsOfDirectory(atPath: dir.path)) ?? [] where name.hasSuffix(".zip") {
                        var st = stat()
                        if lstat(dir.appendingPathComponent(name).path, &st) == 0 { samples.lock.lock(); samples.modes.append(st.st_mode & 0o777); samples.lock.unlock() }
                    }
                }
                usleep(50)
            }
        }
        poller.start()
        let out = try ZipHelper.unzip(data: data, namespace: ns)
        poller.cancel()
        ZipHelper.cleanup(out)
        samples.lock.lock(); let modes = samples.modes; samples.lock.unlock()
        XCTAssertFalse(modes.isEmpty, "the poller observed the private copy at least once")
        XCTAssertEqual(Set(modes), [0o600], "every sample of the private copy's mode is 0600: \(Set(modes).map { String($0, radix: 8) })")
    }

    func testTheExtractionWalkNeverFollowsALinkSwappedInDuringTheWalk() throws {
        // N-DA10-3 (mutation M4): `chmod` follows a link swapped in between the walk's
        // lstat and its mode change; fchmodat(AT_SYMLINK_NOFOLLOW) never does. Only a
        // same-uid race can plant that link — this test IS that race, repeated: a
        // swapper replaces extracted files with links to a file outside the tree while
        // the walk runs. On the real code the outside file's mode cannot change.
        //
        // DETECTION POWER IS PROBABILISTIC — do not read a green run as proof.
        // Measured against the M4 mutant: 3 catches in 17 runs (verify R11 logic
        // N-L11-1) and 10 in 80 pooled across two reviewers (R11 DA N-DA11-2,
        // ~12.5% — the lower, better-sampled figure is the one to quote);
        // against the real code 8/8 green, no false alarm. So this
        // narrows DA's M4 gap from "nothing guards it" to "something guards it
        // about one time in six", and it must NOT be used as a CI gate on its own.
        // A deterministic guard would need the link present before the walk starts,
        // which the archive cannot deliver (a symlink ENTRY is refused in the
        // pre-scan) — closing it properly means making the walk callable on a
        // prepared tree, which is a refactor, not a test change.
        let tmp = FileManager.default.temporaryDirectory
        let outside = tmp.appendingPathComponent("i137-outside-\(UUID().uuidString).txt")
        try Data("o".utf8).write(to: outside); defer { try? FileManager.default.removeItem(at: outside) }
        try FileManager.default.setAttributes([.posixPermissions: 0o640], ofItemAtPath: outside.path)   // 0640: a followed chmod to 0600 or 0700 is visible
        let ns = "i137-walk-\(UUID().uuidString)"
        let nsDir = tmp.appendingPathComponent(ns)
        // The race can leave an extraction tree the archive made unremovable, so
        // clean up the way the library does, not with removeItem (verify R11 DA
        // N-DA11-5: this test intermittently left its own tree behind).
        defer { ZipHelper.removeTreeForcibly(nsDir) }
        var entries = [("word/document.xml", body())]
        for i in 0..<300 { entries.append(("word/d/f\(i).txt", "x")) }
        let data = try zipEntries(entries)
        // The swapper starts only once the LAST entry exists, and never touches
        // the last hundred: ZIPFoundation applies each entry's archive mode with
        // a following `setAttributes` right after writing it, so a link swapped
        // in DURING extraction is followed by ZIPFoundation, not by the walk —
        // a same-uid race the library documents as outside its model. The race
        // this test runs is confined to the walk that follows extraction.
        let swapper = Thread {
            while !Thread.current.isCancelled {
                for uuid in (try? FileManager.default.contentsOfDirectory(atPath: nsDir.path)) ?? [] {
                    let dir = nsDir.appendingPathComponent(uuid)
                    guard FileManager.default.fileExists(atPath: dir.appendingPathComponent("word/d/f299.txt").path) else { continue }
                    for i in 0..<200 {
                        let p = dir.appendingPathComponent("word/d/f\(i).txt").path
                        unlink(p); symlink(outside.path, p)
                    }
                }
            }
        }
        swapper.start()
        for _ in 0..<25 {
            if let out = try? ZipHelper.unzip(data: data, namespace: ns) { ZipHelper.cleanup(out) }   // refusing (a link seen by lstat) is fine
            var st = stat(); XCTAssertEqual(stat(outside.path, &st), 0)
            XCTAssertEqual(st.st_mode & 0o777, 0o640, "the file behind a swapped-in link keeps its mode")
        }
        swapper.cancel()
    }

    /// Both reads of a part, and the description they carry.
    ///
    /// verify R12 re-ran the DA's mutations on the tag candidate: the fix round
    /// 12 test covered only `fileData` (M14 RED) — the pass-1 read (M14b) and
    /// the five `describeWithoutPaths` call sites (M16) were STILL green on all
    /// 1554 tests, and M16 is the one that reopens the temporary-path leak of
    /// #146 that verify R8 spent two rounds closing.
    ///
    /// The locale-independent difference M16 must be caught by: an NSError's
    /// `localizedDescription` names the FILE ("The file “header1.xml” couldn't
    /// be opened…"), while a description built from the error code cannot.
    /// Asserting the file name is absent therefore fails under M16 wherever the
    /// system text names the file — measured in two locales (verify R12b
    /// N-R12b-2 notes that is two samples, not a proof for every locale).
    /// Asserting "no `/`" does not fail under M16 at all, which is why fix
    /// round 12's test let it through.
    /// The other three places an inspector error is described.
    ///
    /// verify R12b reverted `describeWithoutPaths` at each of the five call
    /// sites one at a time and found only two of them guarded — the two the
    /// part-read test covers. Listing `word/`, listing `word/media` and
    /// stat-ing a path could all go back to `localizedDescription`, which names
    /// the file, with all 1554 tests green. Two of the three have injectable
    /// seams; the third is reached by putting the part inside a directory that
    /// cannot be searched.
    func testEveryInspectorErrorDescriptionComesFromTheCodeNotTheSystemText() throws {
        func rootWithDocument() throws -> URL {
            let root = FileManager.default.temporaryDirectory.appendingPathComponent("i137-desc-\(UUID().uuidString)")
            try FileManager.default.createDirectory(at: root.appendingPathComponent("word"), withIntermediateDirectories: true)
            try Data(body().utf8).write(to: root.appendingPathComponent("word/document.xml"))
            return root
        }
        func assertNoSystemText(_ label: String, _ expectedPhrase: String, _ operation: () throws -> Void) {
            var captured = ""
            XCTAssertThrowsError(try operation(), label) { error in
                guard case WordError.invalidDocx(let message) = error else { return XCTFail("\(label): \(error)") }
                captured = message
            }
            XCTAssertTrue(captured.contains(expectedPhrase), "\(label): \(captured)")
            XCTAssertTrue(captured.contains("no consistency verdict"), "\(label): \(captured)")
            // The fixed English of the message mentions `word/` on purpose; what
            // must carry nothing from the system is the DESCRIPTION in
            // parentheses. `localizedDescription` names the file or directory it
            // failed on there; a description built from the error code cannot.
            guard let open = captured.lastIndex(of: "("), let close = captured.lastIndex(of: ")"), open < close else {
                return XCTFail("\(label): no parenthesised description in \(captured)")
            }
            let description = String(captured[captured.index(after: open)..<close])
            XCTAssertFalse(description.isEmpty, "\(label): \(captured)")
            XCTAssertFalse(description.contains("/"), "\(label) describes without a path: \(description)")
            XCTAssertFalse(description.contains("word"), "\(label) describes without naming a file or directory: \(description)")
            XCTAssertFalse(description.contains("absent"), "\(label): \(description)")
            XCTAssertFalse(description.contains("part.xml"), "\(label): \(description)")
        }
        struct Boom: Error {}
        // (a) listing word/ fails — injectable.
        let r1 = try rootWithDocument(); defer { ZipHelper.removeTreeForcibly(r1) }
        assertNoSystemText("listing word/", "could not list the package's word/ directory") {
            _ = try PackageInspector.imageConsistencyReport(
                extracting: { r1 },
                listSubpaths: { _ in try FileManager.default.attributesOfItem(atPath: r1.appendingPathComponent("word/absent.xml").path); return [] })
        }
        // (b) listing word/media fails — injectable, and only reached when the
        //     directory exists, so create it.
        let r2 = try rootWithDocument(); defer { ZipHelper.removeTreeForcibly(r2) }
        try FileManager.default.createDirectory(at: r2.appendingPathComponent("word/media"), withIntermediateDirectories: true)
        assertNoSystemText("listing word/media", "could not list the package's word/media directory") {
            _ = try PackageInspector.imageConsistencyReport(
                extracting: { r2 },
                listDirectory: { _ in try FileManager.default.attributesOfItem(atPath: r2.appendingPathComponent("word/media/absent.png").path); return [] })
        }
        // (c) stat-ing a path fails — no seam, so make the directory holding it
        //     unsearchable and hand the scan that path.
        let r3 = try rootWithDocument(); defer { ZipHelper.removeTreeForcibly(r3) }
        let closed = r3.appendingPathComponent("word/closed")
        try FileManager.default.createDirectory(at: closed, withIntermediateDirectories: true)
        try Data(body().utf8).write(to: closed.appendingPathComponent("part.xml"))
        try FileManager.default.setAttributes([.posixPermissions: 0o000], ofItemAtPath: closed.path)
        assertNoSystemText("stat under the package", "could not read file attributes under the extracted package") {
            _ = try PackageInspector.imageConsistencyReport(extracting: { r3 }, listSubpaths: { _ in ["closed/part.xml", "document.xml"] })
        }
    }

    func testNeitherReadOfAPartCanCarryTheFileNameOrAPathIntoTheMessage() throws {
        func report(unreadable part: String, alsoWrite extra: [String: String] = [:]) throws -> String {
            let root = FileManager.default.temporaryDirectory.appendingPathComponent("i137-unreadable-\(UUID().uuidString)")
            defer { ZipHelper.removeTreeForcibly(root) }
            try FileManager.default.createDirectory(at: root.appendingPathComponent("word/_rels"), withIntermediateDirectories: true)
            try Data(body().utf8).write(to: root.appendingPathComponent("word/document.xml"))
            for (name, contents) in extra {
                try Data(contents.utf8).write(to: root.appendingPathComponent(name))
            }
            let target = root.appendingPathComponent(part)
            if !FileManager.default.fileExists(atPath: target.path) { try Data(body().utf8).write(to: target) }
            try FileManager.default.setAttributes([.posixPermissions: 0o000], ofItemAtPath: target.path)
            var captured = ""
            XCTAssertThrowsError(try PackageInspector.imageConsistencyReport(extracting: { root })) { error in
                guard case WordError.invalidDocx(let message) = error else {
                    return XCTFail("a part that cannot be read is refused as invalidDocx, not \(error)")
                }
                captured = message
            }
            XCTAssertTrue(captured.contains("could not read a part"), captured)
            XCTAssertTrue(captured.contains("no consistency verdict"), captured)
            XCTAssertFalse(captured.contains("/"), "no file-system path: \(captured)")
            XCTAssertFalse(captured.contains(root.lastPathComponent), "no temporary directory name: \(captured)")
            let fileName = (part as NSString).lastPathComponent
            XCTAssertFalse(captured.contains(fileName), "the description comes from the error CODE, so it cannot name \(fileName): \(captured)")
            return captured
        }
        // (a) the `fileData` read — document.xml is fetched through it.
        _ = try report(unreadable: "word/document.xml")
        // (b) the pass-1 read — a part is scanned only when its own rels
        //     declares at least one relationship, so give header1.xml one.
        let headerRels = #"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>"#
        _ = try report(unreadable: "word/header1.xml", alsoWrite: ["word/_rels/header1.xml.rels": headerRels])
    }

    func testAnUnreadableTargetModeIsRefusedNotReadAsAbsent() throws {
        // N-DA10-1: `TargetMode` is optional, so an unreadable one used to be read as
        // ABSENT — a legal edit silently turned an external link into an internal part
        // path. It now refuses by name, exactly like Id / Type / Target.
        let base = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate"
        for spelling in ["TargetMode='External'", "TargetMode = \"External\""] {
            let relsXML = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId9\" Type=\"\(base)\" Target=\"https://example.com/t.dotx\" \(spelling)/></Relationships>"
            let message = try writerRefusal { try relsXML.write(to: $0, atomically: true, encoding: .utf8) }
            XCTAssertTrue(message.contains("`TargetMode` attribute is single-quoted or spaced"), "\(spelling): \(message.prefix(400))")
        }
        XCTAssertFalse(RelationshipsOverlay.isPresentButUnreadable(#" Id="rId9" Type="t" Target="a""#, name: "TargetMode"), "absent is not unreadable")
        XCTAssertTrue(RelationshipsOverlay.isPresentButUnreadable(#" Id="rId9" TargetMode='External'"#, name: "TargetMode"))
    }
}
