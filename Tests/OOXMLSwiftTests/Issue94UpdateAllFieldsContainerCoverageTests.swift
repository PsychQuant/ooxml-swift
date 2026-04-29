import XCTest
@testable import OOXMLSwift

/// Tests for `WordDocument.updateAllFields(...)` container traversal coverage
/// (PsychQuant/che-word-mcp#94, post-#62 follow-up).
///
/// **Pre-fix bug**: body loop only processed top-level `.paragraph` BodyChild
/// cases — silently skipped `.table` and `.contentControl`. SEQ fields inside
/// table cells or block-level SDTs were never updated, surfacing as "no SEQ
/// fields found" / stale cachedResults for callers.
///
/// **Post-fix**: body loop recurses into `.table` (rows × cells × paragraphs +
/// nestedTables) and `.contentControl(_, children:)` matching the recursion
/// pattern established in #68's `findBodyChildContainingText` (v0.20.6).
///
/// **Heading-count semantics decision**: only top-level direct `.paragraph`
/// body children count toward `currentHeadingCount` (chapter-reset). Headings
/// nested inside tables / SDTs do NOT increment. Rationale: thesis workflows
/// put chapter headings at body top level; SDT/table-internal headings would
/// create false resets. Test pinned below.
final class Issue94UpdateAllFieldsContainerCoverageTests: XCTestCase {

    // MARK: - Test fixture builders (mirror UpdateAllFieldsTests pattern)

    private func captionParagraph(identifier: String, resetLevel: Int? = nil, initialCached: String = "1") -> Paragraph {
        let field = SequenceField(identifier: identifier, resetLevel: resetLevel, cachedResult: initialCached)
        var run = Run(text: "")
        run.rawXML = field.toFieldXML()
        var para = Paragraph()
        para.runs = [Run(text: "\(identifier) "), run]
        para.properties.style = "Caption"
        return para
    }

    private func headingParagraph(level: Int, text: String) -> Paragraph {
        var para = Paragraph(text: text)
        para.properties.style = "Heading \(level)"
        return para
    }

    private func extractCachedResult(_ rawXML: String) -> String? {
        guard let sepRange = rawXML.range(of: "fldCharType=\"separate\"") else { return nil }
        let after = rawXML[sepRange.upperBound...]
        guard let openRange = after.range(of: "<w:t") else { return nil }
        let afterOpen = after[openRange.upperBound...]
        guard let gtRange = afterOpen.range(of: ">") else { return nil }
        let body = afterOpen[gtRange.upperBound...]
        guard let closeRange = body.range(of: "</w:t>") else { return nil }
        return String(body[..<closeRange.lowerBound])
    }

    private func cachedResultOfFirstSEQ(in para: Paragraph) -> String? {
        for run in para.runs {
            if let raw = run.rawXML, raw.contains("SEQ "),
               let cached = extractCachedResult(raw) {
                return cached
            }
        }
        return nil
    }

    // MARK: - 94.1 Primary reproducer: SEQ inside table cell

    func testUpdateAllFieldsRecursesIntoTableCellParagraphs() {
        var doc = WordDocument()

        let cellPara1 = captionParagraph(identifier: "Figure", initialCached: "1")
        let cellPara2 = captionParagraph(identifier: "Figure", initialCached: "1")

        let table = Table(rows: [
            TableRow(cells: [
                {
                    var cell = TableCell()
                    cell.paragraphs = [cellPara1]
                    return cell
                }(),
                {
                    var cell = TableCell()
                    cell.paragraphs = [cellPara2]
                    return cell
                }()
            ])
        ])

        doc.body.children = [.table(table)]

        let result = doc.updateAllFields()
        XCTAssertEqual(result, ["Figure": 2],
            "Updater must traverse table cells and increment Figure counter for each cell-anchored SEQ")

        // Verify cached results were rewritten in cell paragraphs
        guard case .table(let updatedTable) = doc.body.children[0] else {
            return XCTFail("expected .table")
        }
        let cell1Cached = cachedResultOfFirstSEQ(in: updatedTable.rows[0].cells[0].paragraphs[0])
        let cell2Cached = cachedResultOfFirstSEQ(in: updatedTable.rows[0].cells[1].paragraphs[0])
        XCTAssertEqual(cell1Cached, "1", "cell 0 SEQ cachedResult should be 1")
        XCTAssertEqual(cell2Cached, "2", "cell 1 SEQ cachedResult should be 2")
    }

    // MARK: - 94.2 SEQ inside block-level SDT (.contentControl)

    func testUpdateAllFieldsRecursesIntoSDTChildParagraphs() {
        var doc = WordDocument()

        let innerPara1 = captionParagraph(identifier: "Figure", initialCached: "1")
        let innerPara2 = captionParagraph(identifier: "Figure", initialCached: "1")

        let sdt = StructuredDocumentTag(id: 100, tag: "issue94-sdt")
        let cc = ContentControl(sdt: sdt, content: "")

        doc.body.children = [
            .contentControl(cc, children: [
                .paragraph(innerPara1),
                .paragraph(innerPara2)
            ])
        ]

        let result = doc.updateAllFields()
        XCTAssertEqual(result, ["Figure": 2],
            "Updater must recurse into block-level SDT children and increment Figure counter")

        guard case .contentControl(_, let updatedChildren) = doc.body.children[0] else {
            return XCTFail("expected .contentControl")
        }
        guard case .paragraph(let p1) = updatedChildren[0],
              case .paragraph(let p2) = updatedChildren[1] else {
            return XCTFail("expected paragraphs inside SDT")
        }
        XCTAssertEqual(cachedResultOfFirstSEQ(in: p1), "1", "SDT child[0] SEQ cachedResult should be 1")
        XCTAssertEqual(cachedResultOfFirstSEQ(in: p2), "2", "SDT child[1] SEQ cachedResult should be 2")
    }

    // MARK: - 94.3 Regression: top-level SEQ behavior unchanged

    func testUpdateAllFieldsTopLevelUnaffectedByTableContents() {
        var doc = WordDocument()

        let topPara1 = captionParagraph(identifier: "Figure", initialCached: "1")
        let topPara2 = captionParagraph(identifier: "Figure", initialCached: "1")
        let cellPara = captionParagraph(identifier: "Figure", initialCached: "1")

        let table = Table(rows: [
            TableRow(cells: [
                {
                    var cell = TableCell()
                    cell.paragraphs = [cellPara]
                    return cell
                }()
            ])
        ])

        doc.body.children = [
            .paragraph(topPara1),     // Figure 1
            .table(table),            // contains Figure 2 inside cell
            .paragraph(topPara2)      // Figure 3
        ]

        let result = doc.updateAllFields()
        XCTAssertEqual(result, ["Figure": 3],
            "Counter walks document order: top-paragraph + cell-paragraph + top-paragraph = 3 increments")

        // Verify document-order counter assignment
        guard case .paragraph(let updatedTop1) = doc.body.children[0] else { return XCTFail() }
        guard case .table(let updatedTable) = doc.body.children[1] else { return XCTFail() }
        guard case .paragraph(let updatedTop2) = doc.body.children[2] else { return XCTFail() }

        XCTAssertEqual(cachedResultOfFirstSEQ(in: updatedTop1), "1")
        XCTAssertEqual(cachedResultOfFirstSEQ(in: updatedTable.rows[0].cells[0].paragraphs[0]), "2")
        XCTAssertEqual(cachedResultOfFirstSEQ(in: updatedTop2), "3")
    }

    // MARK: - 94.4 Heading-reset semantics: container-nested headings don't reset

    func testUpdateAllFieldsHeadingResetIgnoresContainerNestedHeadings() {
        var doc = WordDocument()

        // Layout:
        //   [Heading 1] Chapter 1                                  ← top-level, level1Count = 1
        //   [Caption]   Figure 1   (resetLevel=1, count=1)         ← Figure 1-1
        //   [Caption]   Figure 1   (resetLevel=1, count=1)         ← Figure 1-2
        //   [Table containing [Heading 1] inside cell]              ← container-nested heading: should NOT count
        //   [Caption]   Figure 1   (resetLevel=1, count=1)         ← Figure 1-3 (NOT reset)
        //
        // Pre-fix: table is silent → no heading counted, no caption updated inside
        // Post-fix-conservative: table heading does NOT count toward level1Count;
        //   trailing top-level Figure stays at counter 3 (no reset)

        let h1Top = headingParagraph(level: 1, text: "Chapter 1")
        let figResetLevel1A = captionParagraph(identifier: "Figure", resetLevel: 1, initialCached: "1")
        let figResetLevel1B = captionParagraph(identifier: "Figure", resetLevel: 1, initialCached: "1")

        let h1InsideTable = headingParagraph(level: 1, text: "Section in table cell")
        let table = Table(rows: [
            TableRow(cells: [
                {
                    var cell = TableCell()
                    cell.paragraphs = [h1InsideTable]
                    return cell
                }()
            ])
        ])

        let figAfterTable = captionParagraph(identifier: "Figure", resetLevel: 1, initialCached: "1")

        doc.body.children = [
            .paragraph(h1Top),
            .paragraph(figResetLevel1A),
            .paragraph(figResetLevel1B),
            .table(table),
            .paragraph(figAfterTable)
        ]

        let result = doc.updateAllFields()

        // Counter should be 3, NOT reset by the table-cell heading.
        XCTAssertEqual(result, ["Figure": 3],
            "Heading inside table cell should NOT trigger SEQ chapter-reset; counter stays continuous")

        guard case .paragraph(let f1) = doc.body.children[1],
              case .paragraph(let f2) = doc.body.children[2],
              case .paragraph(let f3) = doc.body.children[4] else {
            return XCTFail("expected paragraphs at indices 1, 2, 4")
        }
        XCTAssertEqual(cachedResultOfFirstSEQ(in: f1), "1")
        XCTAssertEqual(cachedResultOfFirstSEQ(in: f2), "2")
        XCTAssertEqual(cachedResultOfFirstSEQ(in: f3), "3",
            "Trailing top-level Figure must continue counter (3), not reset to 1 by container-nested heading")
    }
}
