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

    /// §11.2 helper: minimal .docx that also drops a header part + document →
    /// header relationship + header content-type override.
    private func buildMinimalDocxWithHeader(documentXML: String, headerXML: String, headerRId: String, to url: URL) throws {
        let stagingURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("r5cont-p0-2-staging-\(UUID().uuidString)")
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
        <Relationship Id="\(headerRId)" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
        </Relationships>
        """
        try documentRels.write(to: stagingURL.appendingPathComponent("word/_rels/document.xml.rels"), atomically: true, encoding: .utf8)

        try documentXML.write(to: stagingURL.appendingPathComponent("word/document.xml"), atomically: true, encoding: .utf8)
        try headerXML.write(to: stagingURL.appendingPathComponent("word/header1.xml"), atomically: true, encoding: .utf8)

        try ZipHelper.zip(stagingURL, to: url)
    }

    /// §8.2 helper: minimal .docx that also drops a footnotes part + the
    /// document → footnotes relationship + footnotes content-type override.
    private func buildMinimalDocxWithFootnotes(documentXML: String, footnotesXML: String, to url: URL) throws {
        let stagingURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("r5-p1-2-staging-\(UUID().uuidString)")
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
        <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
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
        <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
        </Relationships>
        """
        try documentRels.write(to: stagingURL.appendingPathComponent("word/_rels/document.xml.rels"), atomically: true, encoding: .utf8)

        try documentXML.write(to: stagingURL.appendingPathComponent("word/document.xml"), atomically: true, encoding: .utf8)
        try footnotesXML.write(to: stagingURL.appendingPathComponent("word/footnotes.xml"), atomically: true, encoding: .utf8)

        try ZipHelper.zip(stagingURL, to: url)
    }

    // MARK: - §8.1 P1: Hyperlink mutation detection SHALL use deep Run equality

    /// v0.19.5+ (#56 R5 P1 #1): formatting-only mutation (same text, different
    /// `runs[0].properties.bold`) SHALL force the writer to emit from `runs`
    /// instead of replaying source `children`. Pre-fix the detection compared
    /// joined text strings only — a property-only mutation was silently
    /// dropped. Closes R4 P1 L1+L2+DA-R2.
    func testHyperlinkRunPropertyMutationDetectedByDeepEquality() throws {
        // Source XML: hyperlink whose single child run has bold=false (no rPr).
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:body>
            <w:p>
              <w:hyperlink r:id="rId1">
                <w:r><w:t>plain</w:t></w:r>
              </w:hyperlink>
            </w:p>
          </w:body>
        </w:document>
        """
        var document = try parseDocXMLToWordDocument(documentXML)

        // Sanity: parser populated children (source-loaded) and runs.
        let body = document.body
        guard case .paragraph(var para) = body.children[0] else {
            return XCTFail("Expected first body child to be a paragraph")
        }
        XCTAssertEqual(para.hyperlinks.count, 1)
        XCTAssertEqual(para.hyperlinks[0].text, "plain")
        XCTAssertFalse(para.hyperlinks[0].children.isEmpty,
                       "Reader SHALL populate children for round-trip ordering")

        // Mutation: same text, different formatting (bold=true).
        para.hyperlinks[0].runs[0].properties.bold = true
        document.body.children[0] = .paragraph(para)
        // Mark body dirty — direct value-type mutation through `body.children[i]`
        // does not propagate to `modifiedParts`; real MCP mutation paths
        // (`replaceText`, `format_text`, etc.) mark dirty themselves. This
        // test exercises the writer-side deep-equality detection only.
        document.modifiedParts.insert("word/document.xml")

        // Roundtrip and assert the bold mutation survives.
        let reread = try roundtrip(document)
        guard case .paragraph(let rPara) = reread.body.children[0] else {
            return XCTFail("Expected first body child to be a paragraph after roundtrip")
        }
        XCTAssertEqual(rPara.hyperlinks.count, 1)
        XCTAssertEqual(rPara.hyperlinks[0].text, "plain",
                       "text SHALL be preserved across roundtrip")
        XCTAssertTrue(rPara.hyperlinks[0].runs.first?.properties.bold ?? false,
                      "Property-only hyperlink mutation SHALL persist via deep-equality detection")
    }

    // MARK: - §8.2 P1: Footnote.toXML SHALL emit from paragraphs (closes DA-N6)

    /// v0.19.5+ (#56 R5 P1 #2): Footnote with multiple source paragraphs and a
    /// mutation to the second paragraph SHALL persist the mutation through
    /// roundtrip. Pre-fix `Footnote.toXML` emitted a hardcoded single-text-run
    /// template ignoring `paragraphs`/`bodyChildren`, so any per-paragraph
    /// mutation was silently dropped on save. (Code-level fix landed in
    /// §6 commit; this test pins the contract.)
    func testFootnoteMultiParagraphMutationSurvivesRoundtrip() throws {
        // Build a docx whose footnotes.xml carries a footnote with two paragraphs.
        let footnotesXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
          <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
          <w:footnote w:id="1">
            <w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:t>first-para-original</w:t></w:r></w:p>
            <w:p><w:r><w:t>second-para-original</w:t></w:r></w:p>
          </w:footnote>
        </w:footnotes>
        """
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:body><w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p></w:body>
        </w:document>
        """
        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("r5-p1-2-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: docxURL) }
        try buildMinimalDocxWithFootnotes(documentXML: documentXML, footnotesXML: footnotesXML, to: docxURL)

        var document = try DocxReader.read(from: docxURL)
        XCTAssertEqual(document.footnotes.footnotes.count, 1)
        let fn = document.footnotes.footnotes[0]
        XCTAssertEqual(fn.bodyChildren.count, 2,
                       "Footnote SHALL preserve both source paragraphs in bodyChildren")

        // Mutate the second paragraph's first run text.
        var mutated = fn
        guard case .paragraph(var p2) = mutated.bodyChildren[1] else {
            return XCTFail("Expected second bodyChild to be a paragraph")
        }
        XCTAssertEqual(p2.runs.first?.text, "second-para-original")
        p2.runs[0].text = "second-para-mutated"
        mutated.bodyChildren[1] = .paragraph(p2)
        document.footnotes.footnotes[0] = mutated
        document.modifiedParts.insert("word/footnotes.xml")

        let reread = try roundtrip(document)
        XCTAssertEqual(reread.footnotes.footnotes.count, 1)
        let rfn = reread.footnotes.footnotes[0]
        XCTAssertEqual(rfn.paragraphs.count, 2,
                       "Footnote SHALL still have 2 paragraphs post-roundtrip")
        XCTAssertEqual(rfn.paragraphs[0].runs.first?.text, "first-para-original",
                       "First paragraph SHALL be preserved verbatim")
        XCTAssertEqual(rfn.paragraphs[1].runs.first?.text, "second-para-mutated",
                       "Second-paragraph mutation SHALL persist; toXML must emit from bodyChildren, not the hardcoded single-run template")
    }

    // MARK: - §8.3 P1: updateHyperlink/deleteHyperlink SHALL walk all parts

    /// v0.19.5+ (#56 R5 P1 #3): `updateHyperlink` SHALL find a hyperlink that
    /// lives inside a header table cell (or inside footers / footnotes /
    /// endnotes / body tables / SDTs). Pre-fix it only walked
    /// `body.children[i].paragraph` directly and threw notFound for any
    /// nested location, silently breaking MCP `update_hyperlink` against
    /// realistic templates. Closes R4 P1 DA-N4.
    func testUpdateHyperlinkInsideHeaderTableSucceeds() throws {
        // Build a doc with a hyperlink inside a header → table → cell.
        var document = WordDocument()

        // Inner paragraph carrying the hyperlink (id = "h1", url placeholder).
        var innerPara = Paragraph()
        var hyper = Hyperlink(id: "h1", text: "old-link", url: "https://old.example",
                              relationshipId: "rId99", tooltip: nil, history: true)
        hyper.position = 0  // legacy emit path is fine for this in-memory build
        innerPara.hyperlinks.append(hyper)

        // 1×1 header table containing innerPara.
        let cell = TableCell(paragraphs: [innerPara])
        let row = TableRow(cells: [cell])
        let table = Table(rows: [row])

        var header = Header(id: "rId10", paragraphs: [], type: .default,
                            originalFileName: "header1.xml")
        header.bodyChildren.append(.table(table))
        document.headers.append(header)
        document.hyperlinkReferences.append(
            HyperlinkReference(relationshipId: "rId99", url: "https://old.example")
        )

        // Pre-fix this throws "Hyperlink 'h1' not found".
        try document.updateHyperlink(hyperlinkId: "h1",
                                     text: "new-link",
                                     url: "https://new.example")

        // Verify in-memory mutation.
        guard case .table(let updated) = document.headers[0].bodyChildren[0] else {
            return XCTFail("Expected first header child to be a table")
        }
        let mutated = updated.rows[0].cells[0].paragraphs[0].hyperlinks[0]
        XCTAssertEqual(mutated.text, "new-link",
                       "Hyperlink inside header table SHALL be updated by updateHyperlink")
        XCTAssertEqual(mutated.url, "https://new.example",
                       "Hyperlink URL inside header table SHALL be updated")

        // Verify hyperlinkReferences URL was synced.
        let ref = document.hyperlinkReferences.first { $0.relationshipId == "rId99" }
        XCTAssertEqual(ref?.url, "https://new.example")

        // Verify modifiedParts now includes the header file (not just document.xml).
        XCTAssertTrue(document.modifiedParts.contains("word/header1.xml"),
                      "modifiedParts SHALL include the header part containing the mutated hyperlink, got: \(document.modifiedParts)")
    }

    // MARK: - §8.4 P1: SDTParser.parseSDT recursive call SHALL pass position

    /// v0.19.5+ (#56 R5 P1 #4): nested SDT (an `<w:sdt>` inside another
    /// `<w:sdtContent>`) SHALL carry a positive `position` after parsing,
    /// not the API-built `position == 0` sentinel. Pre-fix the recursive
    /// `parseSDT(from: childEl, parentSdtId: sdt.id)` call dropped position,
    /// so every nested SDT was indistinguishable from an in-memory built
    /// ContentControl and the `Paragraph.toXMLSortedByPosition` emit path
    /// could collide it with sibling source children. Closes DA-N8.
    func testNestedSDTReceivesPositiveChildPosition() throws {
        // Doc with a paragraph-level SDT that contains a paragraph; that
        // inner paragraph carries a nested SDT.
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:body>
            <w:p>
              <w:sdt>
                <w:sdtPr><w:id w:val="100"/><w:tag w:val="outer"/></w:sdtPr>
                <w:sdtContent>
                  <w:sdt>
                    <w:sdtPr><w:id w:val="101"/><w:tag w:val="inner"/></w:sdtPr>
                    <w:sdtContent><w:r><w:t>nested</w:t></w:r></w:sdtContent>
                  </w:sdt>
                </w:sdtContent>
              </w:sdt>
            </w:p>
          </w:body>
        </w:document>
        """
        let document = try parseDocXMLToWordDocument(documentXML)
        guard case .paragraph(let para) = document.body.children[0] else {
            return XCTFail("Expected first body child to be a paragraph")
        }
        XCTAssertEqual(para.contentControls.count, 1, "Outer SDT SHALL be parsed")
        let outer = para.contentControls[0]
        XCTAssertGreaterThanOrEqual(outer.position, 1,
                                    "Outer SDT (source-loaded) SHALL have position >= 1, got \(outer.position)")
        XCTAssertEqual(outer.children.count, 1, "Outer SDT SHALL have one nested ContentControl child")
        let inner = outer.children[0]
        XCTAssertGreaterThanOrEqual(inner.position, 1,
                                    "Nested SDT SHALL have position >= 1 (not the API-built 0 sentinel), got \(inner.position)")
    }

    func testNestedSiblingSDTsReceiveOneBasedSiblingPositions() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:body>
            <w:p>
              <w:sdt>
                <w:sdtPr><w:id w:val="200"/><w:tag w:val="outer"/></w:sdtPr>
                <w:sdtContent>
                  <w:r><w:t>before</w:t></w:r>
                  <w:sdt>
                    <w:sdtPr><w:id w:val="201"/><w:tag w:val="inner_1"/></w:sdtPr>
                    <w:sdtContent><w:r><w:t>one</w:t></w:r></w:sdtContent>
                  </w:sdt>
                  <w:r><w:t>between</w:t></w:r>
                  <w:sdt>
                    <w:sdtPr><w:id w:val="202"/><w:tag w:val="inner_2"/></w:sdtPr>
                    <w:sdtContent><w:r><w:t>two</w:t></w:r></w:sdtContent>
                  </w:sdt>
                </w:sdtContent>
              </w:sdt>
            </w:p>
          </w:body>
        </w:document>
        """

        let document = try parseDocXMLToWordDocument(documentXML)
        guard case .paragraph(let para) = document.body.children[0] else {
            return XCTFail("Expected first body child to be a paragraph")
        }
        let outer = try XCTUnwrap(para.contentControls.first, "Outer SDT SHALL be parsed")

        XCTAssertEqual(outer.children.map(\.sdt.tag), ["inner_1", "inner_2"])
        XCTAssertEqual(outer.children.map(\.position), [1, 2],
                       "Nested sibling SDTs SHALL receive one-based sibling positions")
    }

    // MARK: - §8.5 P1: acceptAllRevisions/rejectAllRevisions SHALL surface aggregate failure

    /// v0.19.5+ (#56 R5 P1 #5): acceptAllRevisions / rejectAllRevisions
    /// SHALL surface a typed aggregate failure instead of silently swallowing
    /// per-revision errors with `try?`. Pre-fix, an orphan revision id
    /// (typed Revision exists but no wrapper / target paragraph) silently
    /// failed inside `try?` and stayed in `revisions.revisions` while the
    /// caller assumed all-clear. R4 P1 DA-N9 — silent corruption mode.
    func testAcceptAllRevisionsSurfacesNotFoundFromOneFailedHelper() throws {
        var doc = WordDocument()

        // Valid wrapper revision: typed Revision + matching unrecognizedChild.
        let okPara = makeMixedContentWrapperParagraph(revisionId: 5, author: "Alice", innerText: "kept")
        doc.body.children.append(.paragraph(okPara))
        var okRev = okPara.revisions[0]
        okRev.source = .body
        doc.revisions.revisions.append(okRev)

        // Orphan revision: typed Revision exists, NO matching wrapper anywhere.
        var orphan = Revision(id: 99, type: .insertion, author: "Phantom")
        orphan.isMixedContentWrapper = true
        orphan.source = .body
        doc.revisions.revisions.append(orphan)

        // Pre-fix: silently swallows the orphan's notFound, returns clean.
        // Post-fix: throws RevisionError.partialFailure([99]) via the new
        // `tryAcceptAllRevisions()` API (legacy non-throwing variant
        // preserved for che-word-mcp source-compat per R5 design).
        XCTAssertThrowsError(try doc.tryAcceptAllRevisions()) { err in
            guard case RevisionError.partialFailure(let ids) = err else {
                XCTFail("Expected RevisionError.partialFailure, got \(err)")
                return
            }
            XCTAssertEqual(ids, [99],
                           "partialFailure SHALL aggregate the failing revision ids only")
        }

        // The valid revision SHALL still have been accepted.
        XCTAssertFalse(doc.revisions.revisions.contains(where: { $0.id == 5 }),
                       "Valid wrapper revision (id 5) SHALL be accepted even when sibling revision throws")
        // The orphan SHALL remain — accept failed, so its typed Revision stays.
        XCTAssertTrue(doc.revisions.revisions.contains(where: { $0.id == 99 }),
                      "Orphan revision (id 99) SHALL stay in revisions list when accept throws notFound")
    }

    // MARK: - §11.1 R5-CONTINUATION P0 #1: handleMixedContentWrapperRevision SHALL walk container bodyChildren

    /// v0.19.5+ (#56 R5-CONT P0 #1): R5 verify (Logic L2 + Codex P0 + DA C1
    /// root cause) flagged that `handleMixedContentWrapperRevision`'s
    /// container loops walk `headers[hi].paragraphs` (the flat backward-compat
    /// computed view) instead of `bodyChildren`. A mixed-content `<w:ins>`
    /// wrapper inside a header table cell paragraph is therefore unreachable
    /// → `accept_revision` throws notFound even though the revision parsed
    /// correctly into `document.revisions`. This test pins the contract that
    /// container `bodyChildren` recursion (mirroring body's
    /// `transformInBodyChildren`) is mandatory.
    func testAcceptRevisionOnHeaderTableMixedContentWrapperUnwrapsInHeaderPart() throws {
        var doc = WordDocument()

        // Inner paragraph with mixed-content <w:ins> wrapper around a hyperlink.
        var innerPara = Paragraph()
        let raw = "<w:ins w:id=\"77\" w:author=\"Bob\"><w:hyperlink r:id=\"rId99\"><w:r><w:t>head-table-link</w:t></w:r></w:hyperlink></w:ins>"
        innerPara.unrecognizedChildren.append(UnrecognizedChild(name: "ins", rawXML: raw, position: 1))
        var rev = Revision(id: 77, type: .insertion, author: "Bob")
        rev.newText = "head-table-link"
        rev.isMixedContentWrapper = true
        innerPara.revisions.append(rev)

        // Header containing a 1×1 table whose single cell carries innerPara.
        let cell = TableCell(paragraphs: [innerPara])
        let row = TableRow(cells: [cell])
        let table = Table(rows: [row])
        var header = Header(id: "rId10", paragraphs: [], type: .default,
                            originalFileName: "header1.xml")
        header.bodyChildren.append(.table(table))
        doc.headers = [header]

        // Mirror what DocxReader does: surface the typed Revision to document.revisions.
        var docRev = rev
        docRev.source = .header(id: "rId10")
        doc.revisions.revisions.append(docRev)

        // Pre-fix: throws RevisionError.notFound(77) because the container loop
        // iterates header.paragraphs (empty — no top-level paragraphs, only the
        // table) and never reaches the inner paragraph.
        try doc.acceptRevision(revisionId: 77)

        // Post-fix: revision accepted in header part.
        XCTAssertTrue(doc.modifiedParts.contains("word/header1.xml"),
                      "Accept on header table cell wrapper SHALL mark word/header1.xml dirty; got: \(doc.modifiedParts)")
        XCTAssertFalse(doc.modifiedParts.contains("word/document.xml"),
                       "Accept on header wrapper SHALL NOT mark word/document.xml dirty; got: \(doc.modifiedParts)")
        XCTAssertFalse(doc.revisions.revisions.contains(where: { $0.id == 77 }),
                       "Typed Revision id=77 SHALL be removed after accept")

        // The inner paragraph in the table cell SHALL no longer carry the wrapper.
        guard case .table(let updatedTable) = doc.headers[0].bodyChildren[0] else {
            return XCTFail("Expected first header child to remain a table")
        }
        let updatedPara = updatedTable.rows[0].cells[0].paragraphs[0]
        let rawConcat = updatedPara.unrecognizedChildren.map { $0.rawXML }.joined()
        XCTAssertFalse(rawConcat.contains("<w:ins"),
                       "Header table cell paragraph SHALL NOT contain <w:ins> wrapper after accept; got: \(rawConcat)")
    }

    // MARK: - §11.2 R5-CONTINUATION P0 #2: DocxReader propagates revisions from container bodyChildren

    /// v0.19.5+ (#56 R5-CONT P0 #2): R5 verify (Logic L3) flagged the four
    /// per-container revision propagation loops in `DocxReader.read`
    /// (header/footer/footnote/endnote, lines 305-347) walk
    /// `container.paragraphs` (flat view), missing typed Revisions on
    /// paragraphs inside container tables / content controls.
    /// `document.revisions.revisions` therefore can't see them → MCP
    /// `get_revisions` / `accept_revision` / `reject_revision` invisibility
    /// for any tracked change inside a header table.
    ///
    /// Test fixture: a header containing `<w:tbl>` whose cell paragraph
    /// carries a mixed-content `<w:ins>` wrapper (parser surfaces the typed
    /// Revision into the inner paragraph). Pre-fix `document.revisions`
    /// stays empty for that revision; post-fix it contains the revision
    /// with `source = .header(id: <rId>)`.
    func testRevisionInsideHeaderTableSurfacesInDocumentRevisions() throws {
        // Build a minimal docx with a header containing <w:tbl> whose
        // cell paragraph has a <w:ins> wrapper.
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:body>
            <w:p><w:r><w:t>body</w:t></w:r></w:p>
            <w:sectPr><w:headerReference w:type="default" r:id="rId10"/></w:sectPr>
          </w:body>
        </w:document>
        """
        let headerXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:tbl>
            <w:tr>
              <w:tc>
                <w:p>
                  <w:ins w:id="88" w:author="Carol">
                    <w:r><w:t>tracked-in-header-table</w:t></w:r>
                  </w:ins>
                </w:p>
              </w:tc>
            </w:tr>
          </w:tbl>
        </w:hdr>
        """
        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("r5cont-p0-2-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: docxURL) }
        try buildMinimalDocxWithHeader(documentXML: documentXML, headerXML: headerXML, headerRId: "rId10", to: docxURL)

        let document = try DocxReader.read(from: docxURL)

        // Post-fix: revision id=88 surfaces with source = .header(id: "rId10").
        let rev = document.revisions.revisions.first { $0.id == 88 }
        XCTAssertNotNil(rev,
                        "Typed Revision id=88 inside header table cell SHALL surface in document.revisions.revisions; got: \(document.revisions.revisions.map { ($0.id, $0.source) })")
        if let r = rev {
            if case .header(let id) = r.source {
                XCTAssertEqual(id, "rId10",
                               "Revision source SHALL carry the header rId, not hardcoded .body")
            } else {
                XCTFail("Revision source SHALL be .header(id: \"rId10\"), got: \(r.source)")
            }
        }
    }

    // MARK: - §11.3 R5-CONTINUATION P0 #3: replaceText(.all) containers walk bodyChildren

    /// v0.19.5+ (#56 R5-CONT P0 #3): R5 verify (Codex P1 + Regression F1 +
    /// DA C1) flagged `Document.replaceText(scope: .all)` container loops
    /// iterate `headers[i].paragraphs[j]` (flat backward-compat view).
    /// After R5 P0 #6 elevated container `<w:tbl>` direct children into
    /// `bodyChildren`, header table cell text became visible to readers /
    /// writers / hyperlink CRUD but `replaceText` could not reach it →
    /// silent edit-failure recurs in the new surface.
    func testReplaceTextInsideHeaderTableCellAppliesAndPersists() throws {
        var doc = WordDocument()

        // Header containing a 1×1 table whose cell paragraph carries
        // "old-cell-text".
        var innerPara = Paragraph()
        innerPara.runs.append(Run(text: "old-cell-text"))
        let cell = TableCell(paragraphs: [innerPara])
        let row = TableRow(cells: [cell])
        let table = Table(rows: [row])
        var header = Header(id: "rId10", paragraphs: [], type: .default,
                            originalFileName: "header1.xml")
        header.bodyChildren.append(.table(table))
        doc.headers = [header]

        let count = try doc.replaceText(
            find: "old-cell-text",
            with: "new-cell-text",
            options: ReplaceOptions(scope: .all)
        )

        XCTAssertEqual(count, 1,
                       "replaceText SHALL substitute exactly 1 occurrence inside header table cell")
        XCTAssertTrue(doc.modifiedParts.contains("word/header1.xml"),
                      "modifiedParts SHALL include the header part containing the mutated table cell, got: \(doc.modifiedParts)")

        // In-memory verification: the mutated cell paragraph carries the new text.
        guard case .table(let updatedTable) = doc.headers[0].bodyChildren[0] else {
            return XCTFail("Expected first header child to remain a table")
        }
        let updatedPara = updatedTable.rows[0].cells[0].paragraphs[0]
        XCTAssertEqual(updatedPara.runs.first?.text, "new-cell-text",
                       "Header table cell paragraph SHALL carry the replaced text")

        // Roundtrip the mutation through write→reread.
        let reread = try roundtrip(doc)
        guard let rHeader = reread.headers.first(where: { $0.originalFileName == "header1.xml" }),
              case .table(let rTable) = rHeader.bodyChildren.first ?? .paragraph(Paragraph()),
              let rCellPara = rTable.rows.first?.cells.first?.paragraphs.first
        else {
            return XCTFail("Re-read header missing or empty")
        }
        XCTAssertEqual(rCellPara.runs.first?.text, "new-cell-text",
                       "Roundtrip-persisted header table cell SHALL carry replaced text")
    }

    // MARK: - §11.4 R5-CONTINUATION P0 #4: partKey alignment between DocumentWalker and Header.fileName

    /// v0.19.5+ (#56 R5-CONT P0 #4): R5 verify (Logic L1) flagged that
    /// `DocumentWalker.headerPartKey` (`header2.xml` for `.even`) and
    /// `Header.fileName` (`headerEven.xml` for `.even`) produce different
    /// strings for API-built containers (no `originalFileName`). All
    /// existing R5 tests masked the asymmetry by passing `originalFileName`
    /// explicitly. For an API-built `Header(id:..., type: .even)`,
    /// `handleMixedContentWrapperRevision` writes `"word/header2.xml"` to
    /// `modifiedParts`, but `DocxWriter.writeHeader` dirty-gates on
    /// `"word/headerEven.xml"` → mismatch → header part not re-emitted →
    /// silent loss-on-save. This test pins that the two namespaces converge
    /// for every (type, originalFileName) combination.
    func testAPIBuiltHeaderEvenAcceptRevisionMarksWriterCheckedPartDirty() throws {
        var doc = WordDocument()

        // API-built .even header WITHOUT originalFileName.
        var headerPara = makeMixedContentWrapperParagraph(revisionId: 91, author: "Eve", innerText: "even-link")
        let _ = headerPara
        var header = Header(id: "rId11", paragraphs: [headerPara], type: .even)
        XCTAssertNil(header.originalFileName,
                     "Sanity: API-built header SHALL NOT carry originalFileName")
        // Sanity: the two namespaces disagree pre-fix for .even type.
        // (Header.fileName = "headerEven.xml"; DocumentWalker default was "header2.xml")
        let walkerKey = DocumentWalker.headerPartKey(for: header)
        let writerKey = "word/\(header.fileName)"
        XCTAssertEqual(walkerKey, writerKey,
                       "DocumentWalker.headerPartKey SHALL agree with writer-checked Header.fileName for every (type, originalFileName) combination — got walker=\(walkerKey) writer=\(writerKey)")

        // End-to-end: revision accept on this API-built header should mark
        // exactly the part the writer will check.
        doc.headers.append(header)
        var docRev = header.paragraphs[0].revisions[0]
        docRev.source = .header(id: "rId11")
        doc.revisions.revisions.append(docRev)
        try doc.acceptRevision(revisionId: 91)
        XCTAssertTrue(doc.modifiedParts.contains(writerKey),
                      "Accept on API-built .even header SHALL mark \(writerKey) (the writer-checked path); got: \(doc.modifiedParts)")
    }

    /// Footer mirror of the above — pins the same partKey-alignment contract
    /// for footers across every type.
    func testAPIBuiltFooterFirstPartKeyAlignsWithWriter() {
        let footer = Footer(id: "rId12", paragraphs: [], type: .first)
        XCTAssertNil(footer.originalFileName)
        let walkerKey = DocumentWalker.footerPartKey(for: footer)
        let writerKey = "word/\(footer.fileName)"
        XCTAssertEqual(walkerKey, writerKey,
                       "DocumentWalker.footerPartKey SHALL agree with writer-checked Footer.fileName — got walker=\(walkerKey) writer=\(writerKey)")
    }

    // MARK: - §11.5 R5-CONTINUATION P0 #5: acceptRevision typed .deletion routes by revision.source

    /// v0.19.5+ (#56 R5-CONT P0 #5): R5 verify (DA C2 + H2) flagged that
    /// `acceptRevision`'s typed `.deletion` branch (Document.swift:2757-2775)
    /// only walks `body.children`. For a typed `.deletion` Revision with
    /// `source = .footnote(id:N)`, the branch indexes `body.children`
    /// (case-insensitive to source) and either silently no-ops OR deletes
    /// the wrong paragraph in body. Then `revisions.revisions.remove(at:)`
    /// + `modifiedParts.insert("word/document.xml")` fire regardless →
    /// ghost text in footnote, vanished revision marker, wrong part dirty.
    /// Worse than R4's notFound (which at least reported failure).
    func testAcceptRevisionTypedDeletionInFootnoteRemovesText() throws {
        var doc = WordDocument()

        // Footnote with one paragraph carrying "to-keep" + "to-delete-text".
        var fnPara = Paragraph()
        fnPara.runs.append(Run(text: "to-keep "))
        fnPara.runs.append(Run(text: "to-delete-text"))
        var fn = Footnote(id: 1, text: "", paragraphIndex: 0)
        fn.bodyChildren = [.paragraph(fnPara)]
        doc.footnotes.footnotes = [fn]

        // Body has an unrelated paragraph at index 0 that SHOULD NOT be touched.
        var bodyPara = Paragraph()
        bodyPara.runs.append(Run(text: "body-untouched"))
        doc.body.children = [.paragraph(bodyPara)]

        // Typed .deletion Revision pointing at the footnote.
        var rev = Revision(id: 55, type: .deletion, author: "Dan")
        rev.source = .footnote(id: 1)
        rev.paragraphIndex = 0
        rev.originalText = "to-delete-text"
        doc.revisions.revisions.append(rev)

        try doc.acceptRevision(revisionId: 55)

        // Post-fix: footnote text actually gets deleted.
        let updatedFnPara = doc.footnotes.footnotes[0].paragraphs[0]
        let fnConcat = updatedFnPara.runs.map { $0.text }.joined()
        XCTAssertFalse(fnConcat.contains("to-delete-text"),
                       "Accept .deletion(source: .footnote) SHALL remove footnote text; got: \(fnConcat)")
        XCTAssertTrue(fnConcat.contains("to-keep"),
                      "Accept .deletion SHALL preserve sibling text; got: \(fnConcat)")

        // Body paragraph SHALL NOT have been touched (pre-fix would index body
        // by revision.paragraphIndex == 0 and try to mutate body.children[0]).
        guard case .paragraph(let bodyAfter) = doc.body.children[0] else {
            return XCTFail("Body[0] SHALL remain a paragraph")
        }
        XCTAssertEqual(bodyAfter.runs.first?.text, "body-untouched",
                       "Body paragraph SHALL be untouched by footnote-source deletion")

        // modifiedParts SHALL mark the footnotes part, NOT word/document.xml.
        XCTAssertTrue(doc.modifiedParts.contains("word/footnotes.xml"),
                      "modifiedParts SHALL include word/footnotes.xml; got: \(doc.modifiedParts)")
        XCTAssertFalse(doc.modifiedParts.contains("word/document.xml"),
                       "modifiedParts SHALL NOT include word/document.xml for footnote-source deletion; got: \(doc.modifiedParts)")

        // Typed Revision SHALL be removed.
        XCTAssertFalse(doc.revisions.revisions.contains(where: { $0.id == 55 }),
                       "Typed Revision id=55 SHALL be removed after accept")
    }

    // MARK: - §11.6 R5-CONTINUATION P0 #6: toXMLSortedByPosition filters API-built runs/hyperlinks

    /// v0.19.5+ (#56 R5-CONT P0 #6): R5 verify (DA C3) flagged that
    /// `Paragraph.toXMLSortedByPosition` includes ALL runs/hyperlinks/fieldSimples
    /// /alternateContents in the positioned-emit list without filtering on
    /// `position > 0`. Only `contentControls` had the filter (per R5 P0 #2).
    /// API-built runs/hyperlinks default to `position == 0`, so when an MCP
    /// caller appends a Run to a source-loaded paragraph (e.g.,
    /// `insertTextAsRevision` at end-of-paragraph), the new Run sorts BEFORE
    /// every source-loaded child (`position >= 1`) and lands at the head of
    /// the rendered text, not the intended append position.
    func testInsertRunIntoSourceLoadedParagraphPersistsAtAppendPosition() throws {
        // Build a paragraph with two source-loaded runs at positions 1 and 2.
        var para = Paragraph()
        var sourceRun1 = Run(text: "[source-1]")
        sourceRun1.position = 1
        var sourceRun2 = Run(text: "[source-2]")
        sourceRun2.position = 2
        para.runs = [sourceRun1, sourceRun2]

        // Sanity: paragraph routes to the sorted emit path.
        XCTAssertTrue(para.hasSourcePositionedChildren,
                      "Sanity: source-positioned runs SHALL route to toXMLSortedByPosition")

        // Append a new API-built Run with default position (0).
        let appendedRun = Run(text: "[appended]")
        XCTAssertEqual(appendedRun.position, 0,
                       "Sanity: API-built Run defaults to position == 0")
        para.runs.append(appendedRun)

        let xml = para.toXML()
        // Re-parse and walk runs in source-document order.
        let wrapped = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>\(xml)</w:body></w:document>"
        let parsed = try XMLDocument(xmlString: wrapped, options: [])
        let runNodes = try parsed.nodes(forXPath: "//*[local-name()='r']/*[local-name()='t']")
        let texts = runNodes.compactMap { $0.stringValue }
        XCTAssertEqual(texts, ["[source-1]", "[source-2]", "[appended]"],
                       "Appended API-built Run SHALL emit AFTER source-loaded runs in document order; got: \(texts)")
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
