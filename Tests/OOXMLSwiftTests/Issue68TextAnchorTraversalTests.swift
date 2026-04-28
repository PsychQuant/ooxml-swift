import XCTest
@testable import OOXMLSwift

/// PsychQuant/che-word-mcp#68 — text anchor lookup must traverse table-cell
/// paragraphs and block-level SDT children.
///
/// Pre-fix `findBodyChildContainingText` only iterated `.paragraph` BodyChild
/// cases (post-#63 it correctly handled inline SDT inside paragraphs via
/// `flattenedDisplayText`, but `.table` and block-level `.contentControl`
/// BodyChild were silently skipped — anchor text inside a table cell or block
/// SDT child was not findable, causing `textNotFound` even though the text
/// was present in the document).
///
/// All tests drive the public API via `insertParagraph(at: .afterText)` so
/// the private traversal helper is exercised end-to-end. Returned positions
/// must always be the **top-level** `body.children` index of the BodyChild
/// containing the match (not a nested index inside table/SDT).
final class Issue68TextAnchorTraversalTests: XCTestCase {

    // MARK: - .table recursion

    /// Anchor text in a 1-level table cell paragraph: must find and return
    /// the table's body-level index (so insert places new paragraph AFTER
    /// the table at body level, not nested inside the cell).
    func testAfterTextFindsAnchorInTableCellParagraph() throws {
        var doc = WordDocument()
        let cell = TableCell(paragraphs: [Paragraph(text: "anchor inside cell")])
        let row = TableRow(cells: [cell])
        let table = Table(rows: [row])
        doc.body.children = [
            .paragraph(Paragraph(text: "before-table")),
            .table(table),
            .paragraph(Paragraph(text: "after-table")),
        ]

        try doc.insertParagraph(
            Paragraph(text: "inserted"),
            at: .afterText("anchor inside cell", instance: 1)
        )

        // Insert lands at index 2 (right after the .table at index 1).
        XCTAssertEqual(doc.body.children.count, 4)
        guard case .paragraph(let p) = doc.body.children[2] else {
            XCTFail("Expected inserted paragraph after table; got \(doc.body.children[2])")
            return
        }
        XCTAssertEqual(p.runs.first?.text, "inserted")
    }

    /// Anchor in nested table cell (table > row > cell > nestedTable > row > cell > para):
    /// must still return the TOP-LEVEL table body idx, not the nested table.
    func testAfterTextFindsAnchorInNestedTableCellParagraph() throws {
        var doc = WordDocument()

        var nestedCell = TableCell(paragraphs: [Paragraph(text: "deep anchor")])
        nestedCell.properties = TableCellProperties()
        let nestedRow = TableRow(cells: [nestedCell])
        let nestedTable = Table(rows: [nestedRow])

        var outerCell = TableCell(paragraphs: [Paragraph(text: "outer cell text")])
        outerCell.nestedTables = [nestedTable]
        let outerRow = TableRow(cells: [outerCell])
        let outerTable = Table(rows: [outerRow])

        doc.body.children = [
            .paragraph(Paragraph(text: "before")),
            .table(outerTable),
        ]

        try doc.insertParagraph(
            Paragraph(text: "inserted-deep"),
            at: .afterText("deep anchor", instance: 1)
        )

        // outerTable is at body idx 1; insert should land at idx 2.
        XCTAssertEqual(doc.body.children.count, 3)
        guard case .paragraph(let p) = doc.body.children[2] else {
            XCTFail("Expected paragraph at body idx 2; got \(doc.body.children[2])")
            return
        }
        XCTAssertEqual(p.runs.first?.text, "inserted-deep")
    }

    // MARK: - .contentControl (block-level SDT) recursion

    /// Anchor text in a block-level SDT's child paragraph: must find and
    /// return the SDT's body idx so insert lands after the SDT at body level.
    func testAfterTextFindsAnchorInBlockSDTChildParagraph() throws {
        var doc = WordDocument()

        let sdt101 = StructuredDocumentTag(id: 101, tag: "t101", alias: "T101", type: .plainText)

        let cc = ContentControl(sdt: sdt101, content: "")
        let sdtChildren: [BodyChild] = [
            .paragraph(Paragraph(text: "sdt-child anchor"))
        ]
        doc.body.children = [
            .paragraph(Paragraph(text: "before-sdt")),
            .contentControl(cc, children: sdtChildren),
            .paragraph(Paragraph(text: "after-sdt")),
        ]

        try doc.insertParagraph(
            Paragraph(text: "inserted-after-sdt"),
            at: .afterText("sdt-child anchor", instance: 1)
        )

        XCTAssertEqual(doc.body.children.count, 4)
        guard case .paragraph(let p) = doc.body.children[2] else {
            XCTFail("Expected paragraph at body idx 2; got \(doc.body.children[2])")
            return
        }
        XCTAssertEqual(p.runs.first?.text, "inserted-after-sdt")
    }

    /// Block SDT containing nested block SDT containing paragraph — recursion
    /// must walk the inner children list.
    func testAfterTextFindsAnchorInNestedBlockSDT() throws {
        var doc = WordDocument()

        let sdt102 = StructuredDocumentTag(id: 102, tag: "t102", alias: "T102", type: .plainText)

        let inner = ContentControl(sdt: sdt102, content: "")
        let sdt103 = StructuredDocumentTag(id: 103, tag: "t103", alias: "T103", type: .plainText)
        let outer = ContentControl(sdt: sdt103, content: "")
        let innerChildren: [BodyChild] = [
            .paragraph(Paragraph(text: "deeply nested sdt anchor"))
        ]
        let outerChildren: [BodyChild] = [
            .contentControl(inner, children: innerChildren)
        ]
        doc.body.children = [
            .contentControl(outer, children: outerChildren),
        ]

        try doc.insertParagraph(
            Paragraph(text: "found-deeply"),
            at: .afterText("deeply nested sdt anchor", instance: 1)
        )

        // outer SDT at body idx 0; insert should land at idx 1.
        XCTAssertEqual(doc.body.children.count, 2)
        guard case .paragraph(let p) = doc.body.children[1] else {
            XCTFail("Expected paragraph at body idx 1; got \(doc.body.children[1])")
            return
        }
        XCTAssertEqual(p.runs.first?.text, "found-deeply")
    }

    // MARK: - Mixed nesting (.contentControl > .table > cell > paragraph)

    func testAfterTextFindsAnchorInBlockSDTContainingTable() throws {
        var doc = WordDocument()
        let cell = TableCell(paragraphs: [Paragraph(text: "sdt > table > cell anchor")])
        let row = TableRow(cells: [cell])
        let table = Table(rows: [row])
        let sdt104 = StructuredDocumentTag(id: 104, tag: "t104", alias: "T104", type: .plainText)
        let cc = ContentControl(sdt: sdt104, content: "")
        let sdtChildren: [BodyChild] = [.table(table)]
        doc.body.children = [
            .paragraph(Paragraph(text: "lead")),
            .contentControl(cc, children: sdtChildren),
        ]

        try doc.insertParagraph(
            Paragraph(text: "after-sdt-with-table"),
            at: .afterText("sdt > table > cell anchor", instance: 1)
        )

        XCTAssertEqual(doc.body.children.count, 3)
        guard case .paragraph(let p) = doc.body.children[2] else {
            XCTFail("Expected paragraph at body idx 2; got \(doc.body.children[2])")
            return
        }
        XCTAssertEqual(p.runs.first?.text, "after-sdt-with-table")
    }

    // MARK: - nthInstance ordering across mixed locations

    /// Multiple matches across paragraph + table cell + SDT: nthInstance
    /// must count in document order (depth-first per top-level BodyChild).
    func testAfterTextNthInstanceAcrossMixedLocations() throws {
        var doc = WordDocument()
        let cellPara = TableCell(paragraphs: [Paragraph(text: "needle in cell")])
        let table = Table(rows: [TableRow(cells: [cellPara])])

        let sdt105 = StructuredDocumentTag(id: 105, tag: "t105", alias: "T105", type: .plainText)

        let cc = ContentControl(sdt: sdt105, content: "")
        let sdtChildren: [BodyChild] = [.paragraph(Paragraph(text: "needle in sdt"))]

        doc.body.children = [
            .paragraph(Paragraph(text: "needle at top")),     // instance 1 (idx 0)
            .table(table),                                    // instance 2 (idx 1)
            .contentControl(cc, children: sdtChildren),       // instance 3 (idx 2)
            .paragraph(Paragraph(text: "trailer")),
        ]

        try doc.insertParagraph(
            Paragraph(text: "after-2nd"),
            at: .afterText("needle", instance: 2)
        )

        // 2nd "needle" is in the table (idx 1); insert at idx 2.
        XCTAssertEqual(doc.body.children.count, 5)
        guard case .paragraph(let p) = doc.body.children[2] else {
            XCTFail("Expected paragraph after 2nd-instance table; got \(doc.body.children[2])")
            return
        }
        XCTAssertEqual(p.runs.first?.text, "after-2nd")
    }

    // MARK: - Regression pin: pre-existing behavior

    /// #63 inline-SDT anchor (text inside paragraph.contentControls) still works.
    func testInlineSDTAnchorStillFindsViaFlattenedDisplayText() throws {
        var doc = WordDocument()
        var para = Paragraph()
        let sdt106 = StructuredDocumentTag(id: 106, tag: "t106", alias: "T106", type: .plainText)
        let cc = ContentControl(sdt: sdt106, content: "")
        // Note: inline SDT lookup currently uses Paragraph.flattenedDisplayText
        // which walks paragraph.contentControls. We construct a minimal scenario
        // where the anchor text is in a top-level paragraph.
        para.runs = [Run(text: "inline anchor text")]
        para.contentControls = [cc]

        doc.body.children = [
            .paragraph(para),
            .paragraph(Paragraph(text: "trailer")),
        ]

        try doc.insertParagraph(
            Paragraph(text: "inserted"),
            at: .afterText("inline anchor text", instance: 1)
        )

        XCTAssertEqual(doc.body.children.count, 3)
        guard case .paragraph(let p) = doc.body.children[1] else {
            XCTFail("Expected inserted at idx 1; got \(doc.body.children[1])")
            return
        }
        XCTAssertEqual(p.runs.first?.text, "inserted")
    }

    // MARK: - Verify-68 P3 fixes

    /// Counting rule pin (Verify-68 Logic P2): a single table containing the
    /// needle in MULTIPLE cells still counts as ONE `nthInstance`. This locks
    /// in the design choice (1 BodyChild = 1 instance) so future refactors
    /// can't silently switch to "1 cell paragraph = 1 instance".
    func testAfterTextOneTableWithMultipleCellsContainingNeedleIsOneInstance() throws {
        var doc = WordDocument()
        let cellA = TableCell(paragraphs: [Paragraph(text: "needle in cell A")])
        let cellB = TableCell(paragraphs: [Paragraph(text: "needle in cell B")])
        let cellC = TableCell(paragraphs: [Paragraph(text: "needle in cell C")])
        let row = TableRow(cells: [cellA, cellB, cellC])
        let table = Table(rows: [row])

        doc.body.children = [
            .table(table),                                    // instance 1 (single table, 3 cells)
            .paragraph(Paragraph(text: "needle outside table")), // instance 2
        ]

        try doc.insertParagraph(
            Paragraph(text: "after-instance-2"),
            at: .afterText("needle", instance: 2)
        )

        // 2nd instance is the trailing paragraph at original idx 1 → insert at idx 2.
        XCTAssertEqual(doc.body.children.count, 3)
        guard case .paragraph(let p) = doc.body.children[2] else {
            XCTFail("Expected paragraph at idx 2; got \(doc.body.children[2])")
            return
        }
        XCTAssertEqual(p.runs.first?.text, "after-instance-2")

        // Sanity: instance 3 must NOT exist (the 3 cells of the table count as 1).
        XCTAssertThrowsError(try doc.insertParagraph(
            Paragraph(text: "x"),
            at: .afterText("needle", instance: 3)
        ))
    }

    /// Empty needle MUST NOT match anything (Verify-68 Logic P3 / DA P2).
    /// Pre-fix `String.contains("")` returns true, silently inserting at idx 1
    /// (after the first BodyChild). Post-fix: explicit guard returns nil →
    /// `textNotFound` thrown.
    func testAfterTextEmptyNeedleThrowsNotFound() {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "hello")),
            .paragraph(Paragraph(text: "world")),
        ]

        XCTAssertThrowsError(try doc.insertParagraph(
            Paragraph(text: "x"),
            at: .afterText("", instance: 1)
        )) { error in
            if let we = error as? InsertLocationError,
               case let .textNotFound(searchText: s, instance: _) = we {
                XCTAssertEqual(s, "")
            } else {
                XCTFail("Expected InsertLocationError.textNotFound; got \(error)")
            }
        }
    }

    /// Pre-existing: anchor not found anywhere → throws textNotFound.
    func testAfterTextNotFoundStillThrows() {
        var doc = WordDocument()
        doc.body.children = [.paragraph(Paragraph(text: "haystack"))]

        XCTAssertThrowsError(try doc.insertParagraph(
            Paragraph(text: "x"),
            at: .afterText("needle-nope", instance: 1)
        )) { error in
            if let we = error as? InsertLocationError,
               case let .textNotFound(searchText: s, instance: _) = we {
                XCTAssertEqual(s, "needle-nope")
            } else {
                XCTFail("Expected InsertLocationError.textNotFound; got \(error)")
            }
        }
    }
}
