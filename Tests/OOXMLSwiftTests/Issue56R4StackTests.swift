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
// verify by checking that every emit site is either (a) routed through
// `escapeXMLAttribute(_:)` from `XMLAttributeEscape.swift`, or (b) named here
// with rationale.
//
// Format: `<file>:<line(s)>` — <rationale>
//
// Initial allow-list (populated as §4 sweep proceeds):
// - (none — every emit site SHALL be routed through escapeXMLAttribute or
//   listed here with rationale before §4.13 is marked done)
