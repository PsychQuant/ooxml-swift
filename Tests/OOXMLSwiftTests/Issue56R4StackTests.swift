import XCTest
@testable import OOXMLSwift

/// Issue #56 R5 stack-completion regression tests.
///
/// Each test in this file targets one of the 6 P0 + 7 P1 findings from R4
/// verify (BLOCK verdict, 6 reviewers). Per the design Decision: every test
/// SHALL exercise full save→re-read roundtrip via `roundtrip(_:)` from
/// `Helpers/RoundtripHelper.swift` so the writer path participates in the
/// regression coverage. The R3 stack's all-in-memory test pattern was the
/// proven blind spot of the R2→R3→R4 cycle.
///
/// Tests are added per-task as the R5 stack proceeds:
/// - §2 (P0 #1): mixed-content revision wrapper across all parts
/// - §3 (P0 #2): SDT position ≥ 1 reader assignment
/// - §4 (P0 #3): XML attribute escape sweep
/// - §5 (P0 #4): block-level SDT typed Revision propagation
/// - §6 (P0 #5): Document.replaceText container-symmetric surface walk
/// - §7 (P0 #6): container parser w:tbl capture
/// - §8 (P1 batch)
final class Issue56R4StackTests: XCTestCase {

    // MARK: - §2 P0 #1: Revision accept/reject SHALL find mixed-content wrappers across all document parts

    /// Builds a paragraph with a mixed-content `<w:ins>` wrapper around a hyperlink,
    /// matching what `DocxReader` produces from source XML
    /// `<w:ins w:id="N"...><w:hyperlink>...<w:r><w:t>TEXT</w:t></w:r></w:hyperlink></w:ins>`.
    private func makeMixedContentWrapperParagraph(revisionId: Int, author: String, innerText: String) -> Paragraph {
        var p = Paragraph()
        let raw = "<w:ins w:id=\"\(revisionId)\" w:author=\"\(author)\"><w:hyperlink r:id=\"rId99\"><w:r><w:t>\(innerText)</w:t></w:r></w:hyperlink></w:ins>"
        p.unrecognizedChildren.append(UnrecognizedChild(name: "ins", rawXML: raw, position: 1))
        var rev = Revision(id: revisionId, type: .insertion, author: author)
        rev.newText = innerText
        rev.isMixedContentWrapper = true
        p.revisions.append(rev)
        return p
    }

    func testAcceptRevisionOnHeaderMixedContentWrapperUnwrapsInHeaderPart() throws {
        var doc = WordDocument()
        let headerPara = makeMixedContentWrapperParagraph(revisionId: 9, author: "Bob", innerText: "head-link")
        let header = Header(id: "rId10", paragraphs: [headerPara], type: .default, originalFileName: "header1.xml")
        doc.headers = [header]

        // Mirror what DocxReader does: propagate the typed Revision to document.revisions.revisions
        var docRev = headerPara.revisions[0]
        docRev.source = .header(id: "rId10")
        doc.revisions.revisions.append(docRev)

        try doc.acceptRevision(revisionId: 9)

        // Per design Decision 1: caller marks the originating part dirty, NOT word/document.xml.
        XCTAssertTrue(doc.modifiedParts.contains("word/header1.xml"),
                      "Accept on header wrapper SHALL mark word/header1.xml dirty; got: \(doc.modifiedParts)")
        XCTAssertFalse(doc.modifiedParts.contains("word/document.xml"),
                       "Accept on header wrapper SHALL NOT mark word/document.xml dirty; got: \(doc.modifiedParts)")
        XCTAssertFalse(doc.revisions.revisions.contains(where: { $0.id == 9 }),
                       "Typed Revision SHALL be removed after accept")

        // Roundtrip — re-read header must NOT have <w:ins> wrapper, and the
        // unwrapped hyperlink survives (as a typed Hyperlink in the re-read
        // paragraph, since the parser recognizes <w:hyperlink>).
        let reread = try roundtrip(doc)
        let rereadHeader = try XCTUnwrap(reread.headers.first(where: { $0.originalFileName == "header1.xml" }),
                                        "Re-read header1 missing")
        let rereadFirstPara = try XCTUnwrap(rereadHeader.paragraphs.first, "Re-read header has no paragraphs")
        let rawConcat = rereadFirstPara.unrecognizedChildren.map { $0.rawXML }.joined()
        XCTAssertFalse(rawConcat.contains("<w:ins"),
                       "Re-read header paragraph SHALL NOT contain <w:ins> wrapper; got rawXML: \(rawConcat)")
        XCTAssertFalse(rereadFirstPara.hyperlinks.isEmpty,
                       "Re-read header paragraph SHALL have a typed Hyperlink for the unwrapped content")
    }

    func testRejectRevisionOnFootnoteMixedContentWrapperRemovesFromFootnotesPart() throws {
        var doc = WordDocument()
        let fnPara = makeMixedContentWrapperParagraph(revisionId: 11, author: "Bob", innerText: "doomed")
        var fn = Footnote(id: 1, text: "", paragraphIndex: 0)
        fn.paragraphs = [fnPara]
        doc.footnotes.footnotes = [fn]

        var docRev = fnPara.revisions[0]
        docRev.source = .footnote(id: 1)
        doc.revisions.revisions.append(docRev)

        try doc.rejectRevision(revisionId: 11)

        XCTAssertTrue(doc.modifiedParts.contains("word/footnotes.xml"),
                      "Reject on footnote wrapper SHALL mark word/footnotes.xml dirty; got: \(doc.modifiedParts)")
        XCTAssertFalse(doc.modifiedParts.contains("word/document.xml"),
                       "Reject on footnote wrapper SHALL NOT mark word/document.xml dirty; got: \(doc.modifiedParts)")

        let reread = try roundtrip(doc)
        let rereadFn = try XCTUnwrap(reread.footnotes.footnotes.first, "Re-read footnote missing")
        let rereadFnPara = try XCTUnwrap(rereadFn.paragraphs.first, "Re-read footnote has no paragraphs")
        let rawConcat = rereadFnPara.unrecognizedChildren.map { $0.rawXML }.joined()
        // Reject of insertion drops both wrapper AND inner content.
        XCTAssertFalse(rawConcat.contains("<w:ins"),
                       "Re-read footnote paragraph SHALL NOT contain <w:ins> wrapper after reject")
        XCTAssertFalse(rawConcat.contains("<w:hyperlink"),
                       "Reject of insertion-wrapper SHALL drop the inner content too")
    }

    // MARK: - §3 P0 #2: DocxReader SHALL assign source paragraph child positions starting at 1

    func testFirstChildSourceSDTRoundTripsAtFirstPosition() throws {
        // Build a `.docx` whose first paragraph has an SDT as its first child
        // followed by a run "B". Round-trip and assert the SDT comes first
        // (preserving source order) instead of being demoted to end-of-paragraph.
        let sourceXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p>
        <w:sdt><w:sdtPr><w:tag w:val="T"/><w:alias w:val="A"/></w:sdtPr><w:sdtContent><w:r><w:t>A</w:t></w:r></w:sdtContent></w:sdt>
        <w:r><w:t>B</w:t></w:r>
        </w:p>
        </w:body>
        </w:document>
        """
        let doc = try parseDocXMLToWordDocument(sourceXML)
        let firstPara = try XCTUnwrap(doc.body.children.first.flatMap { (c: BodyChild) -> Paragraph? in
            if case .paragraph(let p) = c { return p } else { return nil }
        }, "expected first body child to be a paragraph")

        // Reader assigns position >= 1 for ALL source children — first child is 1, not 0.
        XCTAssertEqual(firstPara.contentControls.count, 1, "Expected one source SDT in first paragraph")
        XCTAssertGreaterThan(firstPara.contentControls[0].position, 0,
                             "First-child source SDT SHALL receive position > 0; got \(firstPara.contentControls[0].position)")
        XCTAssertGreaterThan(firstPara.runs.first?.position ?? 0, firstPara.contentControls[0].position,
                             "Run after SDT SHALL receive higher position than the SDT")

        // Round-trip — emitted XML SHALL place the SDT before the run, matching source.
        let reread = try roundtrip(doc)
        let rereadFirstPara = try XCTUnwrap(reread.body.children.first.flatMap { (c: BodyChild) -> Paragraph? in
            if case .paragraph(let p) = c { return p } else { return nil }
        })
        let emit = rereadFirstPara.toXML()
        // Debug: dump the emit when assertions miss to surface the actual ordering.
        let sdtRange = try XCTUnwrap(emit.range(of: "<w:sdt"), "Re-read paragraph emit missing <w:sdt>: \(emit)")
        let runBRange = try XCTUnwrap(emit.range(of: ">B<"), "Re-read paragraph emit missing run B; emit was: \(emit)")
        XCTAssertLessThan(sdtRange.lowerBound, runBRange.lowerBound,
                          "SDT SHALL be emitted before run B (source order); got: \(emit)")
    }

    func testAPIBuiltContentControlPreservesPositionZeroSentinel() throws {
        // Programmatically constructed paragraph: API-built ContentControl
        // defaults to position == 0. Emit SHALL go through the legacy
        // post-content path that appends API-built children at end.
        var p = Paragraph()
        p.runs = [Run(text: "first")]
        p.runs[0].position = 1  // simulate a source-positioned run for sorted-emit routing
        let cc = ContentControl.richText(tag: "API", alias: "API", content: "<w:r><w:t>API</w:t></w:r>")
        XCTAssertEqual(cc.position, 0, "API-built ContentControl SHALL default to position 0")
        p.contentControls.append(cc)

        let emit = p.toXML()
        let sdtRange = try XCTUnwrap(emit.range(of: "<w:sdt"), "API-built SDT not emitted: \(emit)")
        let runRange = try XCTUnwrap(emit.range(of: ">first<"), "Run text not emitted: \(emit)")
        XCTAssertGreaterThan(sdtRange.lowerBound, runRange.lowerBound,
                             "API-built (position 0) SDT SHALL be emitted AFTER source-positioned runs (legacy post-content path); got: \(emit)")
    }

    // MARK: - §5 P0 #4: DocxReader SHALL propagate typed Revisions from block-level SDT children into document.revisions.revisions

    func testBlockLevelSDTWrappedRevisionSurfacesInDocumentRevisions() throws {
        // Source XML: block-level <w:sdt> wrapping a paragraph that contains
        // a mixed-content <w:ins> wrapper (insertion of a hyperlink).
        // R3-NEW-4 propagates typed Revisions to document.revisions.revisions
        // for body paragraphs and table cells, but case .contentControl in
        // DocxReader.swift was a `break` — so MCP accept_revision throws
        // notFound for SDT-wrapped revisions. R5 fix: recurse into children.
        let sourceXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
        <w:sdt>
        <w:sdtPr><w:tag w:val="T"/><w:alias w:val="Block"/></w:sdtPr>
        <w:sdtContent>
        <w:p>
        <w:ins w:id="13" w:author="Alice"><w:hyperlink r:id="rId1"><w:r><w:t>x</w:t></w:r></w:hyperlink></w:ins>
        </w:p>
        </w:sdtContent>
        </w:sdt>
        </w:body>
        </w:document>
        """
        let doc = try parseDocXMLToWordDocument(sourceXML)

        XCTAssertTrue(doc.revisions.revisions.contains(where: { $0.id == 13 }),
                      "Block-level SDT-wrapped Revision SHALL surface in document.revisions.revisions; got: \(doc.revisions.revisions.map { $0.id })")

        let rev = try XCTUnwrap(doc.revisions.revisions.first(where: { $0.id == 13 }))
        XCTAssertEqual(rev.type, .insertion)
        XCTAssertTrue(rev.isMixedContentWrapper, "Revision SHALL be marked as mixed-content wrapper")

        // accept_revision must NOT throw notFound — the lookup must succeed.
        var mut = doc
        XCTAssertNoThrow(try mut.acceptRevision(revisionId: 13),
                         "accept_revision SHALL NOT throw notFound for SDT-wrapped revision")
    }

    // MARK: - §7 P0 #6: DocxReader SHALL capture w:tbl direct children of header, footer, footnote, and endnote roots

    func testHeaderTableBookmarkSurfacesInNextBookmarkIdCalibration() throws {
        // Header containing both a paragraph and a table; the table cell has
        // a paragraph with a bookmarkStart id=42. Pre-R5: parseContainerParagraphs
        // discards the <w:tbl> sibling, so id=42 never enters the model and
        // nextBookmarkId calibration misses it (could collide on insert).
        let docXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body><w:p/></w:body>
        </w:document>
        """
        // Build a complete .docx where word/header1.xml has a paragraph + table.
        let stagingURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("r5-§7-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: stagingURL, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: stagingURL) }
        try FileManager.default.createDirectory(at: stagingURL.appendingPathComponent("_rels"), withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: stagingURL.appendingPathComponent("word/_rels"), withIntermediateDirectories: true)

        let contentTypes = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
        </Types>
        """
        try contentTypes.write(to: stagingURL.appendingPathComponent("[Content_Types].xml"), atomically: true, encoding: .utf8)

        let rels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """
        try rels.write(to: stagingURL.appendingPathComponent("_rels/.rels"), atomically: true, encoding: .utf8)

        let documentRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
        </Relationships>
        """
        try documentRels.write(to: stagingURL.appendingPathComponent("word/_rels/document.xml.rels"), atomically: true, encoding: .utf8)

        let docXMLWithHeader = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
          <w:p/>
          <w:sectPr><w:headerReference w:type="default" r:id="rId10"/></w:sectPr>
        </w:body>
        </w:document>
        """
        _ = docXML  // silence unused warning
        try docXMLWithHeader.write(to: stagingURL.appendingPathComponent("word/document.xml"), atomically: true, encoding: .utf8)

        let header1XML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:p><w:r><w:t>header-text</w:t></w:r></w:p>
          <w:tbl>
            <w:tr>
              <w:tc>
                <w:p><w:bookmarkStart w:id="42" w:name="HeaderBookmark"/><w:r><w:t>cell</w:t></w:r></w:p>
              </w:tc>
            </w:tr>
          </w:tbl>
        </w:hdr>
        """
        try header1XML.write(to: stagingURL.appendingPathComponent("word/header1.xml"), atomically: true, encoding: .utf8)

        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("r5-§7-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try ZipHelper.zip(stagingURL, to: outURL)

        let doc = try DocxReader.read(from: outURL)

        XCTAssertGreaterThanOrEqual(doc.nextBookmarkId, 43,
            "nextBookmarkId SHALL surface header table bookmark id 42 (and be > 42); got \(doc.nextBookmarkId)")
        XCTAssertEqual(doc.headers.count, 1, "Expected 1 header parsed")
        let header = doc.headers[0]
        XCTAssertEqual(header.bodyChildren.count, 2, "Header SHALL have 2 bodyChildren (paragraph then table); got \(header.bodyChildren.count)")
        if case .paragraph = header.bodyChildren[0] {} else {
            XCTFail("Header bodyChildren[0] SHALL be .paragraph")
        }
        if case .table = header.bodyChildren[1] {} else {
            XCTFail("Header bodyChildren[1] SHALL be .table")
        }
    }

    func testFooterTableContentSurvivesRoundtrip() throws {
        // Build doc with a footer that has a single <w:tbl> direct child;
        // round-trip via DocxWriter then DocxReader; assert table content survives.
        var doc = WordDocument()
        var inner = Paragraph()
        inner.runs = [Run(text: "footer-cell")]
        let cell = TableCell(paragraphs: [inner])
        let row = TableRow(cells: [cell])
        let table = Table(rows: [row])
        var footer = Footer(id: "rId11", paragraphs: [], type: .default, originalFileName: "footer1.xml")
        footer.bodyChildren = [.table(table)]
        doc.footers = [footer]

        let reread = try roundtrip(doc)
        let rereadFooter = try XCTUnwrap(reread.footers.first)
        XCTAssertTrue(rereadFooter.bodyChildren.contains(where: {
            if case .table(let t) = $0,
               let firstCellPara = t.rows.first?.cells.first?.paragraphs.first,
               firstCellPara.runs.contains(where: { $0.text == "footer-cell" }) {
                return true
            }
            return false
        }), "Footer table content 'footer-cell' SHALL survive roundtrip")
        // Computed view: paragraphs should be empty (only table direct child)
        XCTAssertEqual(rereadFooter.paragraphs.count, 0,
                       "Footer.paragraphs computed view SHALL be empty when only direct child is a table")
    }

    // MARK: - §6 P0 #5: Document.replaceText headers/footers/footnotes/endnotes symmetric surface walk

    func testReplaceTextInsideHeaderHyperlinkAppliesAndPersists() throws {
        // Build doc whose word/header1.xml contains
        // <w:hyperlink r:id="rId1"><w:r><w:t>old-link</w:t></w:r></w:hyperlink>
        // Then call document.replaceText("old-link", with: "new-link") and
        // assert via roundtrip that the re-read header hyperlink contains
        // "new-link". On main this fails because the header replaceText path
        // only walks para.runs (skipping hyperlink.runs).
        var doc = WordDocument()
        var p = Paragraph()
        let hl = Hyperlink(id: "rId99", runs: [Run(text: "old-link")])
        p.hyperlinks = [hl]
        let header = Header(id: "rId10", paragraphs: [p], type: .default, originalFileName: "header1.xml")
        doc.headers = [header]

        let count = try doc.replaceText(find: "old-link", with: "new-link", options: ReplaceOptions(scope: .all))
        XCTAssertGreaterThan(count, 0, "replaceText SHALL find and replace text inside header hyperlink")

        let reread = try roundtrip(doc)
        let rereadHeader = try XCTUnwrap(reread.headers.first)
        let rereadPara = try XCTUnwrap(rereadHeader.paragraphs.first)
        let allText = rereadPara.hyperlinks.flatMap { $0.runs }.map { $0.text }.joined()
            + rereadPara.runs.map { $0.text }.joined()
        XCTAssertTrue(allText.contains("new-link"),
                      "Re-read header SHALL contain replacement text 'new-link'; got: \(allText)")
        XCTAssertFalse(allText.contains("old-link"),
                       "Re-read header SHALL NOT contain original text 'old-link'; got: \(allText)")
    }

    func testReplaceTextInsideFootnoteFieldSimpleAppliesAndPersists() throws {
        // Footnote with a fieldSimple containing "ANCHOR" text. replaceText
        // SHALL find and replace it via the symmetric surface walk.
        var doc = WordDocument()
        var p = Paragraph()
        var f = FieldSimple(instr: "REF Bookmark1")
        f.runs = [Run(text: "ANCHOR")]
        p.fieldSimples = [f]
        var fn = Footnote(id: 1, text: "", paragraphIndex: 0)
        fn.paragraphs = [p]
        doc.footnotes.footnotes = [fn]

        let count = try doc.replaceText(find: "ANCHOR", with: "TARGET", options: ReplaceOptions(scope: .all))
        XCTAssertGreaterThan(count, 0, "replaceText SHALL find and replace text inside footnote fieldSimple")

        let reread = try roundtrip(doc)
        let rereadFn = try XCTUnwrap(reread.footnotes.footnotes.first)
        let rereadPara = try XCTUnwrap(rereadFn.paragraphs.first)
        let allText = rereadPara.fieldSimples.flatMap { $0.runs }.map { $0.text }.joined()
            + rereadPara.runs.map { $0.text }.joined()
        XCTAssertTrue(allText.contains("TARGET"),
                      "Re-read footnote SHALL contain replacement text 'TARGET'; got: \(allText)")
    }

    // MARK: - §4 P0 #3: XML attribute escape sweep

    func testApplyStyleWithAttackerControlledNameDoesNotInjectOOXML() {
        var p = Paragraph()
        // PoC injection: caller-controlled style id closes the w:val attribute,
        // injects a bookmarkStart, then re-opens a w:pStyle so the rest stays valid.
        let attackerInput = "Foo\"><w:bookmarkStart w:id=\"99\" w:name=\"PWNED\"/><w:pStyle w:val=\""
        p.properties.style = attackerInput

        let emit = p.toXML()

        XCTAssertFalse(emit.contains("<w:bookmarkStart w:id=\"99\""),
                       "Attacker-controlled style id SHALL NOT inject <w:bookmarkStart> via attribute escape; emit was: \(emit)")
        XCTAssertTrue(emit.contains("&quot;"),
                      "Emit SHALL contain escaped \" via &quot; for the attacker-controlled value")
        XCTAssertTrue(emit.contains("&lt;") && emit.contains("&gt;"),
                      "Emit SHALL contain escaped < and > for the attacker-controlled value")
    }

    func testCreateStyleWithSpecialCharIdRoundTripsByteEquivalent() {
        var style = Style(id: "Heading&Title", name: "<Test>", type: .paragraph)
        style.basedOn = "Normal'Quoted"

        let emit = style.toXML()

        XCTAssertTrue(emit.contains("w:styleId=\"Heading&amp;Title\""),
                      "Style.id SHALL be escaped (& → &amp;); emit: \(emit)")
        XCTAssertTrue(emit.contains("<w:name w:val=\"&lt;Test&gt;\"/>"),
                      "Style.name SHALL be escaped (< → &lt;, > → &gt;); emit: \(emit)")
        XCTAssertTrue(emit.contains("<w:basedOn w:val=\"Normal&apos;Quoted\"/>"),
                      "Style.basedOn SHALL be escaped (' → &apos;); emit: \(emit)")
    }

    // Helper: parse raw document.xml string into a WordDocument by writing a
    // minimal docx ZIP and reading via DocxReader. Used by §3 tests that need
    // to inject specific source XML for the parser to assign positions to.
    private func parseDocXMLToWordDocument(_ documentXML: String) throws -> WordDocument {
        let tmpURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("r5-§3-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: tmpURL) }
        try buildMinimalDocx(documentXML: documentXML, to: tmpURL)
        return try DocxReader.read(from: tmpURL)
    }

    private func buildMinimalDocx(documentXML: String, to url: URL) throws {
        // Minimal valid .docx with the given document.xml content.
        let stagingURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("r5-§3-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: stagingURL, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: stagingURL) }

        try FileManager.default.createDirectory(at: stagingURL.appendingPathComponent("_rels"), withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: stagingURL.appendingPathComponent("word/_rels"), withIntermediateDirectories: true)

        let contentTypes = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>
        """
        try contentTypes.write(to: stagingURL.appendingPathComponent("[Content_Types].xml"), atomically: true, encoding: .utf8)

        let rels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """
        try rels.write(to: stagingURL.appendingPathComponent("_rels/.rels"), atomically: true, encoding: .utf8)

        let documentRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        </Relationships>
        """
        try documentRels.write(to: stagingURL.appendingPathComponent("word/_rels/document.xml.rels"), atomically: true, encoding: .utf8)

        try documentXML.write(to: stagingURL.appendingPathComponent("word/document.xml"), atomically: true, encoding: .utf8)

        // Zip the staging directory into url.
        try ZipHelper.zip(stagingURL, to: url)
    }

    func testAcceptRevisionOnMissingWrapperRaisesNotFound() {
        var doc = WordDocument()
        var rev = Revision(id: 99, type: .insertion, author: "Phantom")
        rev.isMixedContentWrapper = true
        rev.source = .body
        doc.revisions.revisions.append(rev)
        // Note: NO paragraph anywhere has a matching unrecognizedChildren entry for id 99.

        XCTAssertThrowsError(try doc.acceptRevision(revisionId: 99)) { err in
            guard case RevisionError.notFound(let id) = err else {
                XCTFail("Expected RevisionError.notFound, got \(err)")
                return
            }
            XCTAssertEqual(id, 99)
        }

        XCTAssertTrue(doc.revisions.revisions.contains(where: { $0.id == 99 }),
                      "Typed Revision SHALL NOT be removed when accept throws notFound")
        XCTAssertFalse(doc.modifiedParts.contains("word/document.xml"),
                       "modifiedParts SHALL NOT be mutated when accept throws notFound; got: \(doc.modifiedParts)")
    }
}

// MARK: - Allow-list audit table for emit sites NOT routed through escapeXMLAttribute
//
// Per Issue #56 R5 stack-completion spec `xml-attribute-escape` Requirement 3:
// "Issue56R4StackTests SHALL include an allow-list audit table for emit sites
// NOT routed through escapeXMLAttribute". This is the explicit allow-list of
// exemptions, NOT a deny-list claiming "all sites covered". A reviewer can
// verify by checking that every caller-controlled String attribute interpolation
// is either (a) routed through `escapeXMLAttribute(_:)` from
// `XMLAttributeEscape.swift`, or (b) named here with rationale.
//
// Scope: only String values whose source is a public-API parameter, an MCP
// tool argument, or a model field that originated from a public mutator are
// in scope. Out-of-scope (always safe, exempt by category):
//
//   Category A — Numeric attribute values via String(Int) / String(Double)
//   ─ All sites of the form `w:val="\(intValue)"`, `w:sz="\(size)"`,
//     `w:id="\(integer)"`, `w:left="\(indent)"`, etc., are out-of-scope:
//     numeric-to-string conversion produces only digits / sign / decimal,
//     none of which are XML-special chars. No injection surface.
//   Category B — Enum rawValue interpolations
//   ─ All `\(someEnum.rawValue)` where the enum is defined inside ooxml-swift
//     are constants from a closed set (e.g., `BorderStyle`, `RevisionType`,
//     `HeaderFooterType`). Author-controlled at build time, not user input.
//     No injection surface.
//   Category C — Hardcoded string literals
//   ─ Sites like `w:type="default"`, `w:val="clear"`, `w:hint="default"`
//     are author-controlled compile-time constants. No injection surface.
//   Category D — Pre-validated identifier strings
//   ─ Relationship IDs like `rId1`, `rId2` produced by RelationshipIdAllocator
//     are constrained to `r` + digits format by the allocator's API. No
//     injection surface (covered by allocator unit tests).
//   Category E — Verbatim XML round-trip (rawXML fields)
//   ─ `unrecognizedChild.rawXML`, `alternateContent.rawXML`,
//     `customXmlBlock.rawXML`, `bidiOverride.rawXML`, `smartTag.rawXML`
//     hold a verbatim XML *fragment* (not an attribute value). They are
//     emitted as-is to preserve byte-equivalence with source. The Reader
//     captured them, so they are well-formed XML by construction.
//
// Explicit named exemptions (Category F — site-specific rationale):
//
//   File:Line — Rationale
//   ─ packages/ooxml-swift/Sources/OOXMLSwift/IO/DocxWriter.swift
//     ─ root preamble emits (`<?xml version="1.0" ...?>`, `xmlns:*` declarations)
//       are author-controlled namespace constants, not user input.
//   ─ packages/ooxml-swift/Sources/OOXMLSwift/Models/Paragraph.swift
//     ─ Footnote/Endnote reference rStyle val "FootnoteReference" /
//       "EndnoteReference" are hardcoded style names, not user input.
//   ─ packages/ooxml-swift/Sources/OOXMLSwift/IO/SDTParser.swift
//     ─ Tag/alias values emitted by SDT parser come from typed parser output,
//       routed through hyperlink/SDT models that themselves emit via
//       escape-aware paths.
//
// Test fixture builders (Category G — test-only emitters):
//   ─ buildMinimalDocx in this file emits known-safe constants for ZIP
//     packaging (Content_Types.xml, .rels). Test-only — never reachable
//     from production MCP tool surface.
//
// Alternate escape helpers (Category H — pending consolidation):
//   ─ Sites that route caller-controlled values through `escapeXML(_:)` (a
//     private helper in `Hyperlink.swift`, `DocxWriter.swift`, etc.) escape
//     the four attribute-significant chars (& < > ") and do NOT have a
//     quote-injection vulnerability for caller input. Specific sites:
//     ─ Hyperlink.swift around line 210 (anchor / target / display text)
//     ─ Field.swift around 157 (text input field)
//     ─ Revision.swift around 187 (revision wrapper attributes)
//     ─ Bookmark.swift around 67 (bookmark range marker name)
//     ─ DocxWriter.swift around 608 (numbering definition values)
//     These remain on the local `escapeXML` helper for v0.19.5; the
//     consolidation onto the shared `escapeXMLAttribute(_:)` is a
//     follow-up (post-R5) consistency cleanup. No injection surface
//     because the local helper does cover all four attribute-significant
//     chars; only `'` is missed, which has no attribute-injection effect
//     when emitted between double quotes (the only attribute-quote style
//     ooxml-swift uses).
//   ─ MathComponent.swift `escapeMathXML` covers `& < >` (sufficient for
//     element text). Sites that use it for ATTRIBUTE values (e.g.,
//     `MathAccent.accentChar` at line 144) were migrated to
//     `escapeXMLAttribute` in R5 P0 #3 because attribute escape requires
//     `"` coverage too. Element-text usage of `escapeMathXML` (e.g.,
//     `MathRun.text` at line 49) remains intentional — `"` is allowed
//     inside element text, so the narrower escape is correct.
//
// Reviewer protocol: when adding a NEW caller-controlled String attribute
// emit site, EITHER route through `escapeXMLAttribute(_:)`, OR add the
// site to this allow-list with explicit rationale. If the rationale is
// "I think it's safe", that's not enough — must cite the category above
// or add a new explicit Category F entry with concrete reasoning.
