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
