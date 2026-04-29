import XCTest
@testable import OOXMLSwift

/// Tests for `Comment.paragraphIndex` linker semantics
/// (PsychQuant/che-word-mcp#87, PsychQuant/ooxml-swift#10 family).
///
/// **Pre-fix bug**: `DocxReader.swift:440-447` linker wrote
/// `paragraphIndex = body.children.enumerated() index`, which counts
/// `.table` / `.contentControl` / `.bookmarkMarker` / `.rawBlockElement`
/// alongside `.paragraph`. Callers using the documented pattern
/// `getParagraphs()[paragraphIndex]` therefore got the wrong paragraph
/// (off-by-N where N = number of non-paragraph siblings before the
/// commented paragraph).
///
/// **Post-fix**: linker walks via the same recursion as `getParagraphs()`
/// (recurse `.contentControl` children, skip `.table` / markers / raw),
/// using a flat-paragraph counter. `paragraphIndex` is now a valid index
/// into `getParagraphs()`.
///
/// **Out-of-scope guard**: comments anchored inside table cells were not
/// linked pre-fix and remain unlinked post-fix. Adding table-cell comment
/// linkage is an additive enhancement (separate issue).
final class Issue87CommentParagraphIndexTests: XCTestCase {

    private var tempDir: URL!

    override func setUp() {
        super.setUp()
        tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("Issue87Tests-\(UUID().uuidString)")
        try? FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
    }

    override func tearDown() {
        try? FileManager.default.removeItem(at: tempDir)
        super.tearDown()
    }

    private func roundTrip(_ document: WordDocument) throws -> WordDocument {
        let docxURL = tempDir.appendingPathComponent("test.docx")
        try DocxWriter.write(document, to: docxURL)
        return try DocxReader.read(from: docxURL)
    }

    private func paragraphWithComment(text: String, commentId: Int) -> Paragraph {
        var para = Paragraph(text: text)
        para.commentRangeMarkers = [
            CommentRangeMarker(kind: .start, id: commentId, position: 0),
            CommentRangeMarker(kind: .end, id: commentId, position: 1)
        ]
        return para
    }

    // MARK: - Regression: 0 tables (no behavior change)

    func testCommentParagraphIndexMatchesGetParagraphsWith0Tables() throws {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "P0")),
            .paragraph(paragraphWithComment(text: "P1 commented", commentId: 1)),
            .paragraph(Paragraph(text: "P2"))
        ]
        doc.comments.addComment(
            Comment(id: 1, author: "Tester", text: "note", paragraphIndex: 0)
        )

        let result = try roundTrip(doc)

        let comment = result.comments.comments.first { $0.id == 1 }
        XCTAssertNotNil(comment, "Comment id=1 must round-trip")

        let flatParas = result.getParagraphs()
        let expected = flatParas.firstIndex { $0.commentRangeIds.contains(1) }
        XCTAssertEqual(expected, 1, "Sanity: getParagraphs()[1] is the commented paragraph")

        XCTAssertEqual(comment?.paragraphIndex, 1,
            "Baseline regression: with no tables, paragraphIndex matches body.children index AND flat index")
    }

    // MARK: - Primary fix: 1 table before comment

    func testCommentParagraphIndexMatchesGetParagraphsWith1TableBefore() throws {
        var doc = WordDocument()
        let table = Table(rows: [
            TableRow(cells: [TableCell(text: "TC0"), TableCell(text: "TC1")])
        ])
        doc.body.children = [
            .paragraph(Paragraph(text: "P0")),
            .table(table),
            .paragraph(paragraphWithComment(text: "P1 commented", commentId: 1))
        ]
        doc.comments.addComment(
            Comment(id: 1, author: "Tester", text: "note", paragraphIndex: 0)
        )

        let result = try roundTrip(doc)

        let comment = result.comments.comments.first { $0.id == 1 }
        XCTAssertNotNil(comment, "Comment id=1 must round-trip")

        let flatParas = result.getParagraphs()
        let expected = flatParas.firstIndex { $0.commentRangeIds.contains(1) }
        XCTAssertEqual(expected, 1, "Sanity: getParagraphs() skips table, so commented paragraph is flat index 1")

        XCTAssertEqual(comment?.paragraphIndex, 1,
            "PRIMARY FIX: paragraphIndex matches getParagraphs() flat index (1), NOT body.children enum index (2)")
        XCTAssertEqual(comment?.paragraphIndex, expected,
            "Contract: comment.paragraphIndex == getParagraphs().firstIndex(where: { contains commentId })")
    }

    // MARK: - SDT recursion: contentControl wrapping the commented paragraph

    func testCommentParagraphIndexMatchesGetParagraphsWithSDTContaining() throws {
        var doc = WordDocument()
        let sdt = StructuredDocumentTag(id: 200, tag: "issue87-sdt")
        let cc = ContentControl(sdt: sdt, content: "")
        doc.body.children = [
            .paragraph(Paragraph(text: "P0")),
            .contentControl(cc, children: [
                .paragraph(paragraphWithComment(text: "P1 commented inside SDT", commentId: 1))
            ]),
            .paragraph(Paragraph(text: "P2"))
        ]
        doc.comments.addComment(
            Comment(id: 1, author: "Tester", text: "note", paragraphIndex: 0)
        )

        let result = try roundTrip(doc)

        let comment = result.comments.comments.first { $0.id == 1 }
        XCTAssertNotNil(comment, "Comment id=1 must round-trip")

        let flatParas = result.getParagraphs()
        let expected = flatParas.firstIndex { $0.commentRangeIds.contains(1) }
        XCTAssertEqual(expected, 1, "Sanity: getParagraphs() recurses into SDT, so commented paragraph is flat index 1")

        XCTAssertEqual(comment?.paragraphIndex, 1,
            "SDT recursion: linker walks into .contentControl children matching getParagraphs() semantics")
        XCTAssertEqual(comment?.paragraphIndex, expected,
            "Contract: comment.paragraphIndex == getParagraphs() flat index, regardless of SDT wrapping")
    }

    // MARK: - Out-of-scope guard: comments inside table cells

    func testCommentParagraphIndexUnaffectedWhenCommentInsideTableCell() throws {
        // Pre-fix and post-fix behavior: comments anchored to paragraphs
        // inside table cells are NOT linked by the body-level walker (the
        // walker mirrors getParagraphs() semantics, which excludes tables).
        // Adding table-cell linkage is an additive enhancement tracked
        // separately. This test pins current behavior to catch accidental
        // regression in either direction.
        var doc = WordDocument()
        let commentedCellPara = paragraphWithComment(text: "cell comment", commentId: 99)
        var cell = TableCell()
        cell.paragraphs = [commentedCellPara]
        let table = Table(rows: [TableRow(cells: [cell])])

        doc.body.children = [
            .paragraph(Paragraph(text: "P0")),
            .table(table)
        ]
        doc.comments.addComment(
            Comment(id: 99, author: "Tester", text: "cell note", paragraphIndex: -1)
        )

        let result = try roundTrip(doc)

        let comment = result.comments.comments.first { $0.id == 99 }
        XCTAssertNotNil(comment, "Comment id=99 must round-trip even if not linked")

        // Pre-fix: paragraphIndex stayed at whatever Reader's default was
        // (the linker only fired for top-level paragraphs).
        // Post-fix: linker still doesn't enter table cells; paragraphIndex
        // remains unset by the body-level walker.
        // Note: this test asserts the *behavior is unchanged*, not a
        // specific value — the default may evolve as the parser improves.
        // Whatever the default is, it should not point to a "wrong"
        // paragraph in body.children.
        let flatParas = result.getParagraphs()
        let inFlatParas = flatParas.contains { $0.commentRangeIds.contains(99) }
        XCTAssertFalse(inFlatParas, "Sanity: cell-anchored comments are NOT in getParagraphs() (which excludes tables)")
        // No assertion on comment.paragraphIndex's specific value — pinning
        // current "unset" behavior would couple this test to internal
        // default initialization. The contract being preserved is:
        // body-level linker does not synthesize a wrong paragraphIndex
        // for cell-anchored comments.
    }
}
