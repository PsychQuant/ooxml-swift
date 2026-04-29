import XCTest
@testable import OOXMLSwift

/// Public API surface tests for `WordDocument.findBodyChildContainingText` +
/// static helpers `bodyChildContainsText` / `tableContainsText`, exposed as
/// `public` in v0.21.7 per PsychQuant/che-word-mcp#86.
///
/// External Swift SPM consumers (rescue scripts, dxedit CLI, third-party tooling)
/// previously had to reimplement this logic with diverging semantics — esp. the
/// `.contentControl(_, children:)` recursion and `.table` cell traversal added
/// in #68. These tests pin the canonical public behavior so consumers can rely
/// on it across releases.
final class Issue86PublicAnchorLookupTests: XCTestCase {

    // MARK: - findBodyChildContainingText (instance method)

    func testFindBodyChildContainingText_topLevelParagraphMatch() {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "alpha")),
            .paragraph(Paragraph(text: "beta")),
            .paragraph(Paragraph(text: "gamma")),
        ]

        XCTAssertEqual(doc.findBodyChildContainingText("alpha"), 0)
        XCTAssertEqual(doc.findBodyChildContainingText("beta"), 1)
        XCTAssertEqual(doc.findBodyChildContainingText("gamma"), 2)
        XCTAssertNil(doc.findBodyChildContainingText("delta"))
    }

    func testFindBodyChildContainingText_nthInstanceDisambiguation() {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "圖 4-1：apple")),
            .paragraph(Paragraph(text: "intermediate text")),
            .paragraph(Paragraph(text: "圖 4-2：banana")),
            .paragraph(Paragraph(text: "圖 4-3：cherry")),
        ]

        // Default instance = 1 → first match
        XCTAssertEqual(doc.findBodyChildContainingText("圖 4-"), 0)
        // Explicit instance disambiguates among multiple matches
        XCTAssertEqual(doc.findBodyChildContainingText("圖 4-", nthInstance: 1), 0)
        XCTAssertEqual(doc.findBodyChildContainingText("圖 4-", nthInstance: 2), 2)
        XCTAssertEqual(doc.findBodyChildContainingText("圖 4-", nthInstance: 3), 3)
        XCTAssertNil(doc.findBodyChildContainingText("圖 4-", nthInstance: 4))
    }

    func testFindBodyChildContainingText_invalidInputsReturnNil() {
        var doc = WordDocument()
        doc.body.children = [.paragraph(Paragraph(text: "hello"))]

        // Empty needle → nil (defensive contract)
        XCTAssertNil(doc.findBodyChildContainingText(""))
        // nthInstance < 1 → nil (defensive contract)
        XCTAssertNil(doc.findBodyChildContainingText("hello", nthInstance: 0))
        XCTAssertNil(doc.findBodyChildContainingText("hello", nthInstance: -1))
    }

    func testFindBodyChildContainingText_traversesContentControlChildren() {
        // SDT wrapping a paragraph — pre-#68 callers might miss this
        var doc = WordDocument()
        let sdt = StructuredDocumentTag(id: 100, tag: "test-sdt")
        let cc = ContentControl(sdt: sdt, content: "")
        doc.body.children = [
            .paragraph(Paragraph(text: "before")),
            .contentControl(cc, children: [
                .paragraph(Paragraph(text: "inside the SDT"))
            ]),
            .paragraph(Paragraph(text: "after")),
        ]

        // Returns the BodyChild index (1, the SDT), NOT the inner paragraph
        XCTAssertEqual(doc.findBodyChildContainingText("inside the SDT"), 1)
    }

    func testFindBodyChildContainingText_traversesTableCells() {
        // Per #68: table cells are now in scope by default
        var cell = TableCell()
        cell.paragraphs = [Paragraph(text: "cell content xyz")]
        let row = TableRow(cells: [cell])
        let table = Table(rows: [row])

        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "before table")),
            .table(table),
            .paragraph(Paragraph(text: "after table")),
        ]

        // Returns the table's BodyChild index (1), not the cell paragraph
        XCTAssertEqual(doc.findBodyChildContainingText("cell content xyz"), 1)
    }

    func testFindBodyChildContainingText_skipsBookmarkAndRawElements() {
        // BookmarkMarker and rawBlockElement carry no flattened text → return false
        var doc = WordDocument()
        let bm = BookmarkRangeMarker(kind: .start, id: 1, name: "bm1")
        let raw = RawElement(name: "w:customXml", xml: "<w:customXml>nothing</w:customXml>")
        doc.body.children = [
            .bookmarkMarker(bm),
            .rawBlockElement(raw),
            .paragraph(Paragraph(text: "real content")),
        ]

        XCTAssertEqual(doc.findBodyChildContainingText("real content"), 2)
        // Confirm the bookmark/raw don't match (even if their internal repr would)
        XCTAssertNil(doc.findBodyChildContainingText("bm1"))
        XCTAssertNil(doc.findBodyChildContainingText("nothing"))
    }

    // MARK: - bodyChildContainsText (static primitive)

    func testBodyChildContainsText_paragraphMatch() {
        let child: BodyChild = .paragraph(Paragraph(text: "needle in haystack"))
        XCTAssertTrue(WordDocument.bodyChildContainsText(child, needle: "needle"))
        XCTAssertFalse(WordDocument.bodyChildContainsText(child, needle: "missing"))
    }

    func testBodyChildContainsText_skipsNonTextChildren() {
        let bm = BookmarkRangeMarker(kind: .start, id: 1, name: "bm1")
        let raw = RawElement(name: "w:customXml", xml: "<w:customXml/>")
        let bookmark: BodyChild = .bookmarkMarker(bm)
        let rawChild: BodyChild = .rawBlockElement(raw)

        XCTAssertFalse(WordDocument.bodyChildContainsText(bookmark, needle: "anything"))
        XCTAssertFalse(WordDocument.bodyChildContainsText(rawChild, needle: "anything"))
    }

    // MARK: - tableContainsText (static primitive)

    func testTableContainsText_walksAllCellsAndNestedTables() {
        // Outer cell with nested table; nested table's cell has the needle
        var nestedCell = TableCell()
        nestedCell.paragraphs = [Paragraph(text: "deep needle here")]
        let nestedTable = Table(rows: [TableRow(cells: [nestedCell])])

        var outerCell = TableCell()
        outerCell.paragraphs = [Paragraph(text: "outer noise")]
        outerCell.nestedTables = [nestedTable]
        let outerTable = Table(rows: [TableRow(cells: [outerCell])])

        XCTAssertTrue(WordDocument.tableContainsText(outerTable, needle: "deep needle here"))
        XCTAssertTrue(WordDocument.tableContainsText(outerTable, needle: "outer noise"))
        XCTAssertFalse(WordDocument.tableContainsText(outerTable, needle: "absent"))
    }

    // MARK: - Round-trip with insertParagraph (smoke test the public API matches internal behavior)

    func testPublicLookup_matchesInternalAfterTextResolution() throws {
        // Confirm: result of public findBodyChildContainingText(needle, nthInstance: 2)
        // matches the BodyChild that .afterText(needle, instance: 2) would resolve to.
        // This is the contract that lets external consumers compute insertion sites
        // without throwing, then use .paragraphIndex for the actual insert.
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "段落 1: 前言")),
            .paragraph(Paragraph(text: "段落 2: 內文")),
            .paragraph(Paragraph(text: "段落 3: 內文 補充")),
            .paragraph(Paragraph(text: "段落 4: 結論")),
        ]

        let publicIdx = doc.findBodyChildContainingText("內文", nthInstance: 2)
        XCTAssertEqual(publicIdx, 2)

        // Internal afterText path lands the new paragraph at publicIdx + 1
        try doc.insertParagraph(Paragraph(text: "新段落"), at: .afterText("內文", instance: 2))
        guard case .paragraph(let inserted) = doc.body.children[3] else {
            XCTFail("Expected new paragraph at index 3"); return
        }
        XCTAssertEqual(inserted.runs.first?.text, "新段落")
    }
}
