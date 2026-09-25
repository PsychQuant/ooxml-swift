import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Regression tests for [PsychQuant/ooxml-swift#168](https://github.com/PsychQuant/ooxml-swift/issues/168).
///
/// Before this fix: any typed edit — `updateCell`, `updateCellParagraph`,
/// `insertParagraph`, and every other API that calls `markTypedDirty("word/document.xml")`
/// — forces the whole `word/document.xml` to be regenerated from the typed
/// model. `pPr` children the typed model has no field for (`<w:kinsoku>`,
/// `<w:snapToGrid>`, `<w:widowControl>`, `<w:wordWrap>`, …) silently vanished
/// from EVERY paragraph in the document, not just the one the edit touched.
///
/// After this fix: `DocxReader.parseParagraphProperties` captures those
/// children into `ParagraphProperties.rawChildren` (same "if not typed,
/// preserve as raw" pattern as `Run.rawElements`), and `ParagraphProperties.toXML()`
/// re-emits them. `sectPr` and `pPrChange` are deliberately excluded from raw
/// capture — see `DocxReader.recognizedPPrChildNames` — because both already
/// have dedicated (if separately incomplete) handling elsewhere and naively
/// carrying either into this slot would misplace a position-sensitive element.
/// That gap (mid-body section breaks are not parsed into a typed field at
/// all) predates #168 and is intentionally out of scope here.
final class Issue168PPrRawChildrenTests: XCTestCase {

    // MARK: - Section A: unit-level parse/emit round trip (no docx I/O)

    private func parseParagraph(_ xml: String) throws -> Paragraph {
        let element = try XMLElement(xmlString: xml)
        return try DocxReader.parseParagraph(
            from: element,
            relationships: RelationshipsCollection(),
            styles: [],
            numbering: Numbering()
        )
    }

    /// A pPr mixing modeled fields (pStyle, jc) with four elements the typed
    /// model does not recognize — the exact shape macdoc#156 tripped over
    /// (kinsoku/snapToGrid) plus two more from CT_PPrBase's long tail.
    private static let watchedParagraphXML = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\
        <w:pPr><w:pStyle w:val="Normal"/><w:kinsoku w:val="0"/><w:snapToGrid w:val="0"/>\
        <w:widowControl w:val="0"/><w:wordWrap w:val="0"/><w:jc w:val="center"/></w:pPr>\
        <w:r><w:t>WATCHED_PARAGRAPH</w:t></w:r></w:p>
        """

    func testParseCapturesUnmodeledPPrChildren() throws {
        let paragraph = try parseParagraph(Self.watchedParagraphXML)
        let names = paragraph.properties.rawChildren.map(\.name)
        XCTAssertEqual(names, ["kinsoku", "snapToGrid", "widowControl", "wordWrap"],
                       "unmodeled pPr children must be captured, in source order")
        // Modeled fields are unaffected — still populated the old way.
        XCTAssertEqual(paragraph.properties.style, "Normal")
        XCTAssertEqual(paragraph.properties.alignment, .center)
    }

    func testToXMLReemitsRawChildrenBeforeParagraphMarkRPr() {
        var props = ParagraphProperties()
        props.markRunProperties = RunProperties() // forces a <w:rPr> in the output
        props.markRunProperties?.bold = true
        props.rawChildren = [
            RawElement(name: "kinsoku", xml: "<w:kinsoku w:val=\"0\"/>"),
            RawElement(name: "snapToGrid", xml: "<w:snapToGrid w:val=\"0\"/>"),
        ]
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:kinsoku w:val=\"0\"/>"))
        XCTAssertTrue(xml.contains("<w:snapToGrid w:val=\"0\"/>"))
        let kinsokuRange = try! XCTUnwrap(xml.range(of: "<w:kinsoku"))
        let rPrRange = try! XCTUnwrap(xml.range(of: "<w:rPr>"))
        XCTAssertLessThan(kinsokuRange.lowerBound, rPrRange.lowerBound,
                          "raw pPr children must be emitted before the paragraph-mark <w:rPr>")
    }

    /// Codex round-2 review, MEDIUM finding #2: the original version of this
    /// test compared only `.map(\.name)`, which would still pass if `toXML()`
    /// silently dropped every attribute while keeping the element names. This
    /// version also asserts each raw child's `w:val` attribute survives the
    /// parse -> emit -> re-parse round trip, not just its element name.
    func testParseThenEmitRoundTripsUnmodeledChildrenVerbatim() throws {
        let paragraph = try parseParagraph(Self.watchedParagraphXML)
        // toXML() emits a bare <w:p> fragment (no xmlns declaration — that's
        // the document root's job); re-add it so XMLElement(xmlString:) can
        // parse the fragment standalone, same as `buildFixture()` does below.
        let emitted = paragraph.toXML().replacingOccurrences(
            of: "<w:p>",
            with: "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">")
        let reparsed = try parseParagraph(emitted)
        // #175：toXML() 現在依 CT_PPr schema 順序輸出（widowControl 6、kinsoku 13、
        // wordWrap 14、snapToGrid 21），不再照來源 fixture 的非 schema 順序；
        // 這裡要驗的是「四個都還在」，順序改以 schema 順序為準。
        XCTAssertEqual(reparsed.properties.rawChildren.map(\.name),
                       ["widowControl", "kinsoku", "wordWrap", "snapToGrid"],
                       "re-parsing the emitted XML must still find all four unmodeled children")
        for raw in reparsed.properties.rawChildren {
            XCTAssertTrue(raw.xml.contains("w:val=\"0\""),
                          "\(raw.name)'s w:val must survive the parse -> emit -> re-parse round trip, not just its element name")
        }
    }

    /// An all-default paragraph (no rawChildren, no other pPr field) must
    /// still drop its empty <w:pPr> — this fix must not regress the
    /// established "no synthetic empty <w:pPr/>" gate
    /// (Issue4PPrRegressionGuardTests.testEmptyPPrSelfClosingProducesNoUnrecognizedAndDropsEmptyBlock).
    func testAllDefaultParagraphStillEmitsNoPPr() {
        let props = ParagraphProperties()
        XCTAssertTrue(props.toXML().isEmpty)
    }

    // MARK: - Section B: full-docx regression via common typed-edit entry points

    /// `word/document.xml` for the fixture:
    ///   0: "Intro"                          — plain
    ///   1: "WATCHED_PARAGRAPH"               — kinsoku/snapToGrid/widowControl/wordWrap + pStyle/jc
    ///   2: <w:tbl> one row, one cell, two paragraphs ("Cell A", "Cell B")
    ///   3: "Outro"                           — plain (insertParagraph target neighbor)
    private static let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
        <w:p><w:r><w:t>Intro</w:t></w:r></w:p>\
        \(watchedParagraphXML.replacingOccurrences(
            of: " xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"", with: ""))\
        <w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>\
        <w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>\
        <w:p><w:r><w:t>Cell A</w:t></w:r></w:p>\
        <w:p><w:r><w:t>Cell B</w:t></w:r></w:p>\
        </w:tc></w:tr></w:tbl>\
        <w:p><w:r><w:t>Outro</w:t></w:r></w:p>\
        <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>\
        </w:body></w:document>
        """

    /// - Parameter documentXML: overrides `Self.documentXML` — used by tests
    ///   that need a different body shape (e.g. the shading-collision fixture
    ///   below) without duplicating the ZIP-assembly boilerplate.
    /// Codex round-2 review, LOW finding #5: cleans up its own partial output
    /// on throw (archive population failing after the file was created)
    /// instead of relying on a caller `defer` that never gets installed
    /// because the function never returns a URL in that case.
    private func buildFixture(documentXML: String? = nil) throws -> URL {
        let staging = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue168-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: staging, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: staging) }

        func write(_ content: String, to relativePath: String) throws {
            let url = staging.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(), withIntermediateDirectories: true)
            try content.write(to: url, atomically: true, encoding: .utf8)
        }

        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="xml" ContentType="application/xml"/>
                <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>
            """, to: "[Content_Types].xml")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """, to: "_rels/.rels")
        try write(documentXML ?? Self.documentXML, to: "word/document.xml")

        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue168-\(UUID().uuidString).docx")
        do {
            let archive = try Archive(url: docxURL, accessMode: .create)
            let base = staging.resolvingSymlinksInPath().path
            let enumerator = FileManager.default.enumerator(
                at: staging, includingPropertiesForKeys: [.isDirectoryKey])!
            for case let fileURL as URL in enumerator {
                let isDir = (try? fileURL.resourceValues(forKeys: [.isDirectoryKey]))?.isDirectory ?? false
                if isDir { continue }
                let entry = String(fileURL.resolvingSymlinksInPath().path.dropFirst(base.count + 1))
                try archive.addEntry(with: entry, fileURL: fileURL, compressionMethod: .deflate)
            }
        } catch {
            try? FileManager.default.removeItem(at: docxURL)
            throw error
        }
        return docxURL
    }

    private func extractDocumentXML(from docxURL: URL) throws -> Data {
        let archive = try Archive(url: docxURL, accessMode: .read)
        guard let entry = archive["word/document.xml"] else {
            XCTFail("word/document.xml missing from \(docxURL.lastPathComponent)")
            return Data()
        }
        var data = Data()
        _ = try archive.extract(entry) { data.append($0) }
        return data
    }

    private func save(_ doc: WordDocument) throws -> URL {
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue168-out-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: outURL)
        return outURL
    }

    /// Locates the `<w:p>` containing `marker` in its flattened text and
    /// returns its `<w:pPr>` node (nil if the paragraph has none).
    private func pPrNode(in documentXML: Data, paragraphContaining marker: String) throws -> XmlNode? {
        let tree = try XmlTreeReader.parse(documentXML)

        func containsMarker(_ node: XmlNode) -> Bool {
            if node.kind == .text, node.textContent.contains(marker) { return true }
            return node.children.contains(where: containsMarker)
        }
        func findParagraph(_ node: XmlNode) -> XmlNode? {
            if node.kind == .element, node.localName == "p", containsMarker(node) {
                return node
            }
            for child in node.children {
                if let found = findParagraph(child) { return found }
            }
            return nil
        }
        guard let paragraph = findParagraph(tree.root) else {
            XCTFail("no <w:p> containing '\(marker)' found")
            return nil
        }
        return paragraph.children.first { $0.kind == .element && $0.localName == "pPr" }
    }

    private func assertWatchedParagraphPPrSurvived(
        in outputDocumentXML: Data,
        file: StaticString = #filePath, line: UInt = #line
    ) throws {
        let pPr = try XCTUnwrap(
            pPrNode(in: outputDocumentXML, paragraphContaining: "WATCHED_PARAGRAPH"),
            "WATCHED_PARAGRAPH must still have a <w:pPr>", file: file, line: line)
        let children = pPr.children.filter { $0.kind == .element }
        let names = children.map(\.localName)

        for expected in ["kinsoku", "snapToGrid", "widowControl", "wordWrap"] {
            XCTAssertEqual(names.filter { $0 == expected }.count, 1,
                           "\(expected) must appear exactly once", file: file, line: line)
        }
        // Modeled fields on the same paragraph must survive too.
        XCTAssertTrue(names.contains("pStyle"), file: file, line: line)
        XCTAssertTrue(names.contains("jc"), file: file, line: line)

        for child in children where ["kinsoku", "snapToGrid", "widowControl", "wordWrap"].contains(child.localName) {
            let val = child.attributes.first { $0.prefix == "w" && $0.localName == "val" }?.value
            XCTAssertEqual(val, "0", "\(child.localName)'s w:val must survive unchanged", file: file, line: line)
        }
    }

    func testUpdateCellPreservesUnmodeledPPrChildrenOnUnrelatedParagraph() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        try doc.updateCell(tableIndex: 0, row: 0, col: 0, text: "Cell A edited")

        let outURL = try save(doc)
        defer { try? FileManager.default.removeItem(at: outURL) }
        try assertWatchedParagraphPPrSurvived(in: try extractDocumentXML(from: outURL))
    }

    func testUpdateCellParagraphPreservesUnmodeledPPrChildrenOnUnrelatedParagraph() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        try doc.updateCellParagraph(tableIndex: 0, row: 0, col: 0, paragraphIndex: 1, text: "Cell B edited")

        let outURL = try save(doc)
        defer { try? FileManager.default.removeItem(at: outURL) }
        try assertWatchedParagraphPPrSurvived(in: try extractDocumentXML(from: outURL))
    }

    func testInsertParagraphPreservesUnmodeledPPrChildrenOnUnrelatedParagraph() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        // body.children is [paragraph(Intro), paragraph(WATCHED), table, paragraph(Outro)];
        // insert before "Outro" (index 3) so neither the table nor the watched
        // paragraph is the node touched by this edit.
        doc.insertParagraph(Paragraph(text: "INSERTED"), at: 3)

        let outURL = try save(doc)
        defer { try? FileManager.default.removeItem(at: outURL) }
        let outXML = try extractDocumentXML(from: outURL)
        try assertWatchedParagraphPPrSurvived(in: outXML)
        XCTAssertTrue(String(decoding: outXML, as: UTF8.self).contains("INSERTED"),
                     "the edit itself must actually be present in the saved document")
    }

    /// A DETERMINISM check, not a preservation check (Codex round-2 review,
    /// MEDIUM finding #2 — the original doc comment overclaimed this as "the
    /// strongest form" of byte-for-byte preservation; it is not: a writer
    /// that dropped the same unmodeled children under BOTH edits would still
    /// pass this comparison, since `normalizedFingerprint()` compares the two
    /// OUTPUTS to each other, not either output to the source). What this
    /// test actually proves: two INDEPENDENT unrelated edits — one via
    /// `updateCell`, one via `insertParagraph` — regenerate a
    /// structurally-identical `<w:pPr>` (rsid/prefix/attribute-order noise
    /// aside, per `normalizedFingerprint()`'s own normalization rules) for
    /// the untouched watched paragraph. Its typed properties are identical in
    /// both runs, so a deterministic writer must reproduce the same content
    /// regardless of which unrelated edit triggered the rewrite. The actual
    /// presence/value regression guard is `assertWatchedParagraphPPrSurvived`,
    /// exercised by the three tests above this one.
    func testWatchedParagraphPPrIsDeterministicAcrossDifferentUnrelatedEdits() throws {
        let fixtureA = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixtureA) }
        var docA = try DocxReader.read(from: fixtureA)
        defer { docA.close() }
        try docA.updateCell(tableIndex: 0, row: 0, col: 0, text: "via updateCell")
        let outA = try save(docA)
        defer { try? FileManager.default.removeItem(at: outA) }

        let fixtureB = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixtureB) }
        var docB = try DocxReader.read(from: fixtureB)
        defer { docB.close() }
        docB.insertParagraph(Paragraph(text: "INSERTED"), at: 3)
        let outB = try save(docB)
        defer { try? FileManager.default.removeItem(at: outB) }

        let pPrA = try XCTUnwrap(pPrNode(in: try extractDocumentXML(from: outA), paragraphContaining: "WATCHED_PARAGRAPH"))
        let pPrB = try XCTUnwrap(pPrNode(in: try extractDocumentXML(from: outB), paragraphContaining: "WATCHED_PARAGRAPH"))
        XCTAssertEqual(pPrA.normalizedFingerprint(), pPrB.normalizedFingerprint(),
                       "the untouched watched paragraph's <w:pPr> must be structurally identical regardless of which unrelated edit forced the rewrite")
    }

    // MARK: - Section B.1: raw capture must not collide with typed setters
    // (Codex round-2 review, HIGH finding #1)
    //
    // `ParagraphProperties` already declares `border` / `shading` fields, and
    // `setParagraphBorder(at:border:)` / `setParagraphShading(at:fill:pattern:)`
    // are public typed setters that write them — but `parseParagraphProperties`
    // has never read `<w:pBdr>` / `<w:shd>` INTO those fields (a pre-existing,
    // separately-tracked reader gap, not something #168 fixes). If the raw
    // capture added for #168 treated `pBdr`/`shd` as "unrecognized", a
    // paragraph loaded with a `<w:shd>` already present, then given a NEW
    // shading via `setParagraphShading`, would emit BOTH the raw captured
    // `<w:shd>` and the newly-set typed one — two singleton elements where
    // OOXML expects at most one. `recognizedPPrChildNames` therefore also
    // excludes `pBdr`/`shd` from raw capture (same treatment as `sectPr`):
    // they stay silently dropped on an unrelated typed edit, exactly like
    // before #168 — not improved for these two elements, but not corrupted
    // either.

    private static let shadedParagraphDocumentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
        <w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/></w:pPr><w:r><w:t>Shaded</w:t></w:r></w:p>\
        <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>\
        </w:body></w:document>
        """

    func testSettingShadingOnAParagraphThatAlreadyHasRawShadingDoesNotDuplicateIt() throws {
        let fixture = try buildFixture(documentXML: Self.shadedParagraphDocumentXML)
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        try doc.setParagraphShading(at: 0, fill: "00FF00")

        let outURL = try save(doc)
        defer { try? FileManager.default.removeItem(at: outURL) }
        let pPr = try XCTUnwrap(pPrNode(in: try extractDocumentXML(from: outURL), paragraphContaining: "Shaded"))
        let shdNodes = pPr.children.filter { $0.kind == .element && $0.localName == "shd" }
        XCTAssertEqual(shdNodes.count, 1,
                       "exactly one <w:shd> must survive — a raw-captured old one plus a typed new one would both be schema-invalid and ambiguous to Word")
        // Codex round-3 review, LOW finding #2: the `if let` form let a
        // `<w:shd>` with no `w:fill` attribute at all pass this assertion —
        // the count check above guarded the duplicate-emission bug, but not
        // the "new value actually wins" claim in this test's name.
        // `XCTUnwrap` makes a missing `w:fill` an explicit failure.
        let fill = try XCTUnwrap(
            shdNodes.first?.attributes.first { $0.prefix == "w" && $0.localName == "fill" }?.value,
            "the surviving <w:shd> must carry a w:fill attribute")
        XCTAssertEqual(fill, "00FF00", "the typed override must win, not the stale raw-captured value")
    }

    // MARK: - Section C: real-world template (MACDOC_TEMPLATE_DIR-gated)
    //
    //   MACDOC_TEMPLATE_DIR=/path/to/private/docx swift test --filter Issue168PPrRawChildrenTests

    /// 90_template_ja.docx's own pPr content (checked by hand, 2026-09-24) is
    /// entirely within the already-modeled vocabulary (ind/jc/spacing/rPr;
    /// plus sectPr, which stays a pre-existing, separately-tracked gap) — it
    /// has no kinsoku/snapToGrid/etc. of its own. So this test cannot exercise
    /// the new rawChildren capture path directly; it instead verifies the
    /// weaker but still meaningful claim #168 is actually about: an unrelated
    /// typed edit must not change what a real document's untouched paragraph
    /// *means* (compared as typed `ParagraphProperties`, not raw bytes — the
    /// writer's fixed pPr child order already differs from this template's
    /// source order for reasons unrelated to #168, so raw byte-equality would
    /// fail even on a correct implementation).
    func testRealTemplateUnrelatedEditDoesNotChangeUntouchedParagraphProperties() throws {
        let url = try TemplateFixtureGate.requireTemplate(TemplateFixtureGate.baselineTemplateName)
        let before = try DocxReader.read(from: url)
        // A paragraph with distinctive text, far from the paragraph this test
        // edits (index 0), so the edit cannot possibly touch it.
        guard let watchedIndex = before.body.children.firstIndex(where: {
            if case .paragraph(let p) = $0 { return p.getText().contains("研究背景") }
            return false
        }), case .paragraph(let watchedBefore) = before.body.children[watchedIndex] else {
            throw XCTSkip("template fixture no longer contains the expected watched paragraph text")
        }

        var doc = before
        doc.insertParagraph(Paragraph(text: "issue168-probe"), at: 0)
        let outURL = try save(doc)
        defer { try? FileManager.default.removeItem(at: outURL) }

        var after = try DocxReader.read(from: outURL)
        defer { after.close() }
        guard case .paragraph(let watchedAfter) = after.body.children[watchedIndex + 1] else {
            return XCTFail("watched paragraph shifted by more than the one inserted paragraph")
        }
        XCTAssertEqual(watchedAfter.getText(), watchedBefore.getText())
        XCTAssertEqual(watchedAfter.properties, watchedBefore.properties,
                       "an unrelated typed edit must not change the untouched paragraph's properties")
    }
}
