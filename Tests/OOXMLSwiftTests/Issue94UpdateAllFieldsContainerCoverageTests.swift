import XCTest
@testable import OOXMLSwift

/// Tests for `WordDocument.updateAllFields(...)` container traversal coverage.
///
/// **#94 (v0.21.9)** — body-side fix: pre-fix, the body loop only processed
/// top-level `.paragraph` BodyChild cases — silently skipped `.table` and
/// `.contentControl`. SEQ fields inside table cells or block-level SDTs were
/// never updated, surfacing as "no SEQ fields found" / stale cachedResults
/// for callers. Post-fix: body loop recurses into `.table` (rows × cells ×
/// paragraphs + nestedTables) and `.contentControl(_, children:)` matching
/// the recursion pattern established in #68's `findBodyChildContainingText`
/// (v0.20.6). Tests at MARK 94.1-94.4.
///
/// **#25 (v0.22.x)** — extends the same recursion pattern to header / footer /
/// footnote / endnote `bodyChildren` (previously those container loops
/// iterated the legacy `.paragraphs` flat-view computed property, which
/// silently dropped `.table` and `.contentControl` siblings). Tests at MARK
/// 25.1-25.5 (25.5 pins the P1 regression for resetLevel false-reset under
/// shared-counter mode).
///
/// **#51** — extends #25's coverage to `.contentControl` wrapping `.table`
/// (not just `.contentControl` wrapping bare paragraphs, or `.table` at the
/// container top level) inside header / footer / footnote / endnote. Tests
/// at MARK 51.1-51.4.
///
/// **#53** — extends #25's coverage to multi-instance containers (multiple
/// headers / multiple footnotes / multiple endnotes), pinning both counter
/// chaining across instances and (for footnotes/endnotes, which share a
/// single dirty bit across all instances in the collection) that a later
/// instance whose own SEQ rewrite is a no-op must not clear a dirty bit an
/// earlier instance already set. Tests at MARK 53.1-53.3.
///
/// **Heading-count semantics decision**: only top-level direct `.paragraph`
/// body children count toward `currentHeadingCount` (chapter-reset). Headings
/// nested inside tables / SDTs / containers do NOT increment. Rationale:
/// thesis workflows put chapter headings at body top level; SDT/table
/// -internal headings would create false resets. Test pinned below.
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

    // MARK: - 25.1 SEQ inside header table cell

    /// PsychQuant/ooxml-swift#25 — header SEQ scan must recurse through bodyChildren.
    /// Pre-#25 fix: header loop iterated `.paragraphs` (computed flat view, drops
    /// `.table` siblings); SEQ inside header table cell was silently skipped.
    /// Post-fix: header loop iterates `bodyChildren` through walkAndProcessBodyChildForFields.
    func testUpdateAllFieldsRecursesIntoHeaderTableCellParagraphs() {
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

        var header = Header(id: "rId-h1")
        header.bodyChildren = [.table(table)]
        doc.headers = [header]

        let result = doc.updateAllFields()
        XCTAssertEqual(result, ["Figure": 2],
            "Header SEQ scan must traverse table cells inside Header.bodyChildren")

        // Verify cached results were rewritten in the header's table cells
        guard case .table(let updatedTable) = doc.headers[0].bodyChildren[0] else {
            return XCTFail("expected .table inside header bodyChildren")
        }
        let cell1Cached = cachedResultOfFirstSEQ(in: updatedTable.rows[0].cells[0].paragraphs[0])
        let cell2Cached = cachedResultOfFirstSEQ(in: updatedTable.rows[0].cells[1].paragraphs[0])
        XCTAssertEqual(cell1Cached, "1", "header table cell 0 SEQ cachedResult should be 1")
        XCTAssertEqual(cell2Cached, "2", "header table cell 1 SEQ cachedResult should be 2")

        // Dirty-bit propagation: header file marked modified
        XCTAssertTrue(doc.modifiedParts.contains("word/\(doc.headers[0].fileName)"),
            "modifiedParts must include the header file when its SEQ fields were updated")
    }

    // MARK: - 25.2 SEQ inside footer block-level SDT (.contentControl)

    /// PsychQuant/ooxml-swift#25 — footer SEQ scan must recurse into block-level
    /// SDT children. Pre-fix: footer loop iterated `.paragraphs` (drops
    /// `.contentControl` siblings); SEQ inside footer SDT was silently skipped.
    func testUpdateAllFieldsRecursesIntoFooterSDTChildParagraphs() {
        var doc = WordDocument()

        let innerPara1 = captionParagraph(identifier: "Figure", initialCached: "1")
        let innerPara2 = captionParagraph(identifier: "Figure", initialCached: "1")

        let sdt = StructuredDocumentTag(id: 200, tag: "issue25-footer-sdt")
        let cc = ContentControl(sdt: sdt, content: "")

        var footer = Footer(id: "rId-f1")
        footer.bodyChildren = [
            .contentControl(cc, children: [
                .paragraph(innerPara1),
                .paragraph(innerPara2)
            ])
        ]
        doc.footers = [footer]

        let result = doc.updateAllFields()
        XCTAssertEqual(result, ["Figure": 2],
            "Footer SEQ scan must recurse into block-level SDT children")

        guard case .contentControl(_, let updatedChildren) = doc.footers[0].bodyChildren[0] else {
            return XCTFail("expected .contentControl inside footer bodyChildren")
        }
        guard case .paragraph(let p1) = updatedChildren[0],
              case .paragraph(let p2) = updatedChildren[1] else {
            return XCTFail("expected paragraphs inside footer SDT")
        }
        XCTAssertEqual(cachedResultOfFirstSEQ(in: p1), "1",
            "footer SDT child[0] SEQ cachedResult should be 1")
        XCTAssertEqual(cachedResultOfFirstSEQ(in: p2), "2",
            "footer SDT child[1] SEQ cachedResult should be 2")

        // Dirty-bit propagation: footer file marked modified
        XCTAssertTrue(doc.modifiedParts.contains("word/\(doc.footers[0].fileName)"),
            "modifiedParts must include the footer file when its SEQ fields were updated")
    }

    // MARK: - 25.3 SEQ inside footnote table cell

    /// PsychQuant/ooxml-swift#25 — footnote SEQ scan must recurse through
    /// bodyChildren so SEQ inside footnote-internal tables is visible.
    func testUpdateAllFieldsRecursesIntoFootnoteTableCellParagraphs() {
        var doc = WordDocument()

        let cellPara1 = captionParagraph(identifier: "Source", initialCached: "1")
        let cellPara2 = captionParagraph(identifier: "Source", initialCached: "1")

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

        var footnote = Footnote(id: 1, text: "", paragraphIndex: 0)
        footnote.bodyChildren = [.table(table)]
        doc.footnotes = FootnotesCollection(footnotes: [footnote])

        let result = doc.updateAllFields()
        XCTAssertEqual(result, ["Source": 2],
            "Footnote SEQ scan must traverse table cells inside Footnote.bodyChildren")

        guard case .table(let updatedTable) = doc.footnotes.footnotes[0].bodyChildren[0] else {
            return XCTFail("expected .table inside footnote bodyChildren")
        }
        let cell1Cached = cachedResultOfFirstSEQ(in: updatedTable.rows[0].cells[0].paragraphs[0])
        let cell2Cached = cachedResultOfFirstSEQ(in: updatedTable.rows[0].cells[1].paragraphs[0])
        XCTAssertEqual(cell1Cached, "1", "footnote table cell 0 SEQ cachedResult should be 1")
        XCTAssertEqual(cell2Cached, "2", "footnote table cell 1 SEQ cachedResult should be 2")

        // Dirty-bit propagation: footnotes.xml marked modified
        XCTAssertTrue(doc.modifiedParts.contains("word/footnotes.xml"),
            "modifiedParts must include word/footnotes.xml when footnote SEQ fields were updated")
    }

    // MARK: - 25.4 SEQ inside endnote container

    /// PsychQuant/ooxml-swift#25 — endnote SEQ scan must recurse through
    /// bodyChildren so SEQ inside endnote-internal tables / SDTs is visible.
    ///
    /// #52: the cell's cache starts stale (`"0"`, not the correctly-computed
    /// `"1"`) specifically so the rewrite is a REAL change — a single-cell
    /// fixture whose cache already matched would never flip `endnotesDirty`,
    /// making the `modifiedParts` assertion below vacuously true for the
    /// wrong reason.
    func testUpdateAllFieldsRecursesIntoEndnoteContainerSEQ() {
        var doc = WordDocument()

        let cellPara = captionParagraph(identifier: "Source", initialCached: "0")

        let table = Table(rows: [
            TableRow(cells: [
                {
                    var cell = TableCell()
                    cell.paragraphs = [cellPara]
                    return cell
                }()
            ])
        ])

        var endnote = Endnote(id: 1, text: "", paragraphIndex: 0)
        endnote.bodyChildren = [.table(table)]
        doc.endnotes = EndnotesCollection(endnotes: [endnote])

        let result = doc.updateAllFields()
        XCTAssertEqual(result, ["Source": 1],
            "Endnote SEQ scan must traverse table cells inside Endnote.bodyChildren")

        // #52 — asymmetry vs 25.1/25.2/25.3, which all assert modifiedParts.
        XCTAssertTrue(doc.modifiedParts.contains("word/endnotes.xml"),
            "modifiedParts must include word/endnotes.xml when endnote SEQ fields were updated")

        guard case .table(let updatedTable) = doc.endnotes.endnotes[0].bodyChildren[0] else {
            return XCTFail("expected .table inside endnote bodyChildren")
        }
        let cellCached = cachedResultOfFirstSEQ(in: updatedTable.rows[0].cells[0].paragraphs[0])
        XCTAssertEqual(cellCached, "1", "endnote table cell SEQ cachedResult should be 1")
    }

    // MARK: - 51.1 SEQ inside header .contentControl wrapping .table

    /// PsychQuant/ooxml-swift#51 — walker supports `.contentControl` recursing
    /// into `.table` (both cases are independently handled by
    /// `walkAndProcessBodyChildForFields`), but no #25-family fixture combined
    /// them nested inside header/footnote/endnote — a thesis-style "SDT
    /// wrapping a caption table" inside a container was untested.
    func testUpdateAllFieldsHeaderSDTWrappingTableSEQ() {
        var doc = WordDocument()

        let cellPara = captionParagraph(identifier: "Figure", initialCached: "0")
        let table = Table(rows: [
            TableRow(cells: [
                {
                    var cell = TableCell()
                    cell.paragraphs = [cellPara]
                    return cell
                }()
            ])
        ])

        let sdt = StructuredDocumentTag(id: 211, tag: "issue51-header-sdt-table")
        let cc = ContentControl(sdt: sdt, content: "")

        var header = Header(id: "rId-h1")
        header.bodyChildren = [.contentControl(cc, children: [.table(table)])]
        doc.headers = [header]

        let result = doc.updateAllFields()
        XCTAssertEqual(result, ["Figure": 1],
            "Header SEQ scan must recurse through .contentControl wrapping .table")

        guard case .contentControl(_, let updatedChildren) = doc.headers[0].bodyChildren[0],
              case .table(let updatedTable) = updatedChildren[0] else {
            return XCTFail("expected .contentControl wrapping .table inside header bodyChildren")
        }
        let cellCached = cachedResultOfFirstSEQ(in: updatedTable.rows[0].cells[0].paragraphs[0])
        XCTAssertEqual(cellCached, "1", "header SDT-wrapped table cell SEQ cachedResult should be 1")

        XCTAssertTrue(doc.modifiedParts.contains("word/\(doc.headers[0].fileName)"),
            "modifiedParts must include the header file when its SEQ fields were updated")
    }

    // MARK: - 51.2 SEQ inside footer .contentControl wrapping .table

    /// PsychQuant/ooxml-swift#51 — footer 25.2 already covers `.contentControl`
    /// wrapping bare paragraphs; this pins the nested `.table` variant.
    func testUpdateAllFieldsFooterSDTWrappingTableSEQ() {
        var doc = WordDocument()

        let cellPara = captionParagraph(identifier: "Figure", initialCached: "0")
        let table = Table(rows: [
            TableRow(cells: [
                {
                    var cell = TableCell()
                    cell.paragraphs = [cellPara]
                    return cell
                }()
            ])
        ])

        let sdt = StructuredDocumentTag(id: 212, tag: "issue51-footer-sdt-table")
        let cc = ContentControl(sdt: sdt, content: "")

        var footer = Footer(id: "rId-f1")
        footer.bodyChildren = [.contentControl(cc, children: [.table(table)])]
        doc.footers = [footer]

        let result = doc.updateAllFields()
        XCTAssertEqual(result, ["Figure": 1],
            "Footer SEQ scan must recurse through .contentControl wrapping .table")

        guard case .contentControl(_, let updatedChildren) = doc.footers[0].bodyChildren[0],
              case .table(let updatedTable) = updatedChildren[0] else {
            return XCTFail("expected .contentControl wrapping .table inside footer bodyChildren")
        }
        let cellCached = cachedResultOfFirstSEQ(in: updatedTable.rows[0].cells[0].paragraphs[0])
        XCTAssertEqual(cellCached, "1", "footer SDT-wrapped table cell SEQ cachedResult should be 1")

        XCTAssertTrue(doc.modifiedParts.contains("word/\(doc.footers[0].fileName)"),
            "modifiedParts must include the footer file when its SEQ fields were updated")
    }

    // MARK: - 51.3 SEQ inside footnote .contentControl wrapping .table

    /// PsychQuant/ooxml-swift#51 — footnote 25.3 only covers a bare `.table`
    /// at the footnote's bodyChildren top level; this pins the SDT-wrapped
    /// variant.
    func testUpdateAllFieldsFootnoteSDTWrappingTableSEQ() {
        var doc = WordDocument()

        let cellPara = captionParagraph(identifier: "Source", initialCached: "0")
        let table = Table(rows: [
            TableRow(cells: [
                {
                    var cell = TableCell()
                    cell.paragraphs = [cellPara]
                    return cell
                }()
            ])
        ])

        let sdt = StructuredDocumentTag(id: 213, tag: "issue51-footnote-sdt-table")
        let cc = ContentControl(sdt: sdt, content: "")

        var footnote = Footnote(id: 1, text: "", paragraphIndex: 0)
        footnote.bodyChildren = [.contentControl(cc, children: [.table(table)])]
        doc.footnotes = FootnotesCollection(footnotes: [footnote])

        let result = doc.updateAllFields()
        XCTAssertEqual(result, ["Source": 1],
            "Footnote SEQ scan must recurse through .contentControl wrapping .table")

        guard case .contentControl(_, let updatedChildren) = doc.footnotes.footnotes[0].bodyChildren[0],
              case .table(let updatedTable) = updatedChildren[0] else {
            return XCTFail("expected .contentControl wrapping .table inside footnote bodyChildren")
        }
        let cellCached = cachedResultOfFirstSEQ(in: updatedTable.rows[0].cells[0].paragraphs[0])
        XCTAssertEqual(cellCached, "1", "footnote SDT-wrapped table cell SEQ cachedResult should be 1")

        XCTAssertTrue(doc.modifiedParts.contains("word/footnotes.xml"),
            "modifiedParts must include word/footnotes.xml when footnote SEQ fields were updated")
    }

    // MARK: - 51.4 SEQ inside endnote .contentControl wrapping .table

    /// PsychQuant/ooxml-swift#51 — endnote 25.4 only covers a bare `.table`
    /// at the endnote's bodyChildren top level; this pins the SDT-wrapped
    /// variant.
    func testUpdateAllFieldsEndnoteSDTWrappingTableSEQ() {
        var doc = WordDocument()

        let cellPara = captionParagraph(identifier: "Source", initialCached: "0")
        let table = Table(rows: [
            TableRow(cells: [
                {
                    var cell = TableCell()
                    cell.paragraphs = [cellPara]
                    return cell
                }()
            ])
        ])

        let sdt = StructuredDocumentTag(id: 214, tag: "issue51-endnote-sdt-table")
        let cc = ContentControl(sdt: sdt, content: "")

        var endnote = Endnote(id: 1, text: "", paragraphIndex: 0)
        endnote.bodyChildren = [.contentControl(cc, children: [.table(table)])]
        doc.endnotes = EndnotesCollection(endnotes: [endnote])

        let result = doc.updateAllFields()
        XCTAssertEqual(result, ["Source": 1],
            "Endnote SEQ scan must recurse through .contentControl wrapping .table")

        guard case .contentControl(_, let updatedChildren) = doc.endnotes.endnotes[0].bodyChildren[0],
              case .table(let updatedTable) = updatedChildren[0] else {
            return XCTFail("expected .contentControl wrapping .table inside endnote bodyChildren")
        }
        let cellCached = cachedResultOfFirstSEQ(in: updatedTable.rows[0].cells[0].paragraphs[0])
        XCTAssertEqual(cellCached, "1", "endnote SDT-wrapped table cell SEQ cachedResult should be 1")

        XCTAssertTrue(doc.modifiedParts.contains("word/endnotes.xml"),
            "modifiedParts must include word/endnotes.xml when endnote SEQ fields were updated")
    }

    // MARK: - 25.5 Container SEQ with resetLevel must not false-reset (P1 regression)

    /// PsychQuant/ooxml-swift#25 verify P1 finding: pre-fix-commit, containers got
    /// `containerHeadingCount = [:]` (always 0) while `lastResetHeadingCount`
    /// inherited body's accumulated state. processParagraph's resetLevel logic
    /// compares currentCount[level] vs lastReset[level] — `0 != N` triggered a
    /// false reset, breaking SEQ chain into containers under shared-counter mode.
    ///
    /// Repro from Codex's blind verify: body Heading 1 + 2x Figure(\s 1), then
    /// header Figure(\s 1) — pre-fix returned cached "1" (false-reset to start);
    /// post-fix returns "3" (continues body counter). The 4 prior tests all use
    /// `resetLevel: nil` so the resetLevel codepath was never exercised.
    func testUpdateAllFieldsHeaderSEQWithResetLevelChainsBodyCounter() {
        var doc = WordDocument()

        // Body: Heading 1 + 2x Figure with resetLevel=1
        let h1 = headingParagraph(level: 1, text: "Chapter 1")
        let bodyFig1 = captionParagraph(identifier: "Figure", resetLevel: 1, initialCached: "1")
        let bodyFig2 = captionParagraph(identifier: "Figure", resetLevel: 1, initialCached: "1")
        doc.body.children = [
            .paragraph(h1),
            .paragraph(bodyFig1),
            .paragraph(bodyFig2)
        ]

        // Header: 1x Figure with resetLevel=1 (same identifier as body)
        let headerFig = captionParagraph(identifier: "Figure", resetLevel: 1, initialCached: "1")
        var header = Header(id: "rId-h1")
        header.bodyChildren = [.paragraph(headerFig)]
        doc.headers = [header]

        // Default isolatePerContainer=false — header should chain body's counter.
        let result = doc.updateAllFields()

        // Pre-fix: result == ["Figure": 1] (header false-reset to start)
        // Post-fix: result == ["Figure": 3] (header continues body's counter)
        XCTAssertEqual(result, ["Figure": 3],
            "Header SEQ with resetLevel=1 must chain body's accumulated counter when " +
            "isolatePerContainer=false. Pre-fix bug: containerHeadingCount=[:] " +
            "+ inherited lastResetHeadingCount caused false reset.")

        guard case .paragraph(let updatedHeaderPara) = doc.headers[0].bodyChildren[0] else {
            return XCTFail("expected .paragraph in header bodyChildren")
        }
        XCTAssertEqual(cachedResultOfFirstSEQ(in: updatedHeaderPara), "3",
            "Header Figure SEQ cachedResult must be 3 (continues body's counter), not 1 (false-reset)")
    }

    // MARK: - 53.1 Multi-header SEQ counter chains

    /// PsychQuant/ooxml-swift#53 — #25-family fixtures only ever used ONE
    /// header/footer/footnote/endnote instance. Thesis workflows commonly
    /// have three header instances (default + first + even); this pins that
    /// the shared SEQ counter chains across ALL of them in registration
    /// order, and that each header's own file is independently marked dirty
    /// (headers already use a per-file `Set<String>`, not a single shared
    /// bit — unlike footnotes/endnotes below).
    func testUpdateAllFieldsMultiHeaderSEQCounterChains() {
        var doc = WordDocument()

        // #53 — h1's cache starts stale ("0") so its own rewrite is a REAL
        // change too; if it started already-correct ("1", matching the first
        // occurrence's computed value) the rewrite would be a no-op and
        // `word/header1.xml` would never be marked dirty regardless of the
        // per-header dirty-bit logic being correct or not.
        var h1 = Header(id: "rId-h1", type: .default)
        h1.bodyChildren = [.paragraph(captionParagraph(identifier: "Figure", initialCached: "0"))]
        var h2 = Header(id: "rId-h2", type: .first)
        h2.bodyChildren = [.paragraph(captionParagraph(identifier: "Figure", initialCached: "1"))]
        var h3 = Header(id: "rId-h3", type: .even)
        h3.bodyChildren = [.paragraph(captionParagraph(identifier: "Figure", initialCached: "1"))]
        doc.headers = [h1, h2, h3]

        let result = doc.updateAllFields()
        XCTAssertEqual(result, ["Figure": 3],
            "multi-header SEQ counter must chain across header instances (default + first + " +
            "even) under shared-counter mode (isolatePerContainer=false)")

        guard case .paragraph(let p1) = doc.headers[0].bodyChildren[0],
              case .paragraph(let p2) = doc.headers[1].bodyChildren[0],
              case .paragraph(let p3) = doc.headers[2].bodyChildren[0] else {
            return XCTFail("expected paragraphs in header bodyChildren")
        }
        XCTAssertEqual(cachedResultOfFirstSEQ(in: p1), "1", "first header instance should chain to 1")
        XCTAssertEqual(cachedResultOfFirstSEQ(in: p2), "2", "second header instance should chain to 2")
        XCTAssertEqual(cachedResultOfFirstSEQ(in: p3), "3", "third header instance should chain to 3")

        XCTAssertTrue(doc.modifiedParts.contains("word/header1.xml"))
        XCTAssertTrue(doc.modifiedParts.contains("word/headerFirst.xml"))
        XCTAssertTrue(doc.modifiedParts.contains("word/headerEven.xml"))
    }

    // MARK: - 53.2 Multi-footnote: shared dirty bit must not be reset by a later no-op

    /// PsychQuant/ooxml-swift#53 — footnotes share ONE `footnotesDirty` bit
    /// across every footnote instance (`updateAllFields()` only sets it via
    /// `if walkAndProcess... { footnotesDirty = true }`, never assigns the
    /// walker's per-instance return value directly). fn1's cache is stale
    /// (needs rewriting → dirty); fn2's cache is already correct for its
    /// counter value (rewrite is a no-op → NOT dirty on its own). A
    /// regression to `footnotesDirty = <last instance's return value>`
    /// would silently clear fn1's dirty bit here because fn2 runs last.
    func testUpdateAllFieldsMultiFootnoteSingleDirtyBit() {
        var doc = WordDocument()

        var fn1 = Footnote(id: 1, text: "", paragraphIndex: 0)
        fn1.bodyChildren = [.paragraph(captionParagraph(identifier: "Figure", initialCached: "0"))]
        var fn2 = Footnote(id: 2, text: "", paragraphIndex: 1)
        fn2.bodyChildren = [.paragraph(captionParagraph(identifier: "Figure", initialCached: "2"))]
        doc.footnotes = FootnotesCollection(footnotes: [fn1, fn2])

        let result = doc.updateAllFields()
        XCTAssertEqual(result, ["Figure": 2],
            "multi-footnote SEQ counter must chain: fn1 -> 1, fn2 -> 2")

        guard case .paragraph(let p1) = doc.footnotes.footnotes[0].bodyChildren[0],
              case .paragraph(let p2) = doc.footnotes.footnotes[1].bodyChildren[0] else {
            return XCTFail("expected paragraphs in footnote bodyChildren")
        }
        XCTAssertEqual(cachedResultOfFirstSEQ(in: p1), "1", "fn1's stale cache must be rewritten to 1")
        XCTAssertEqual(cachedResultOfFirstSEQ(in: p2), "2", "fn2's cache was already correct (no-op rewrite)")

        XCTAssertTrue(doc.modifiedParts.contains("word/footnotes.xml"),
            "footnotesDirty must stay set from fn1's real change — a later footnote whose own " +
            "SEQ rewrite is a no-op (fn2) must NOT reset the shared dirty bit")
    }

    // MARK: - 53.3 Multi-endnote SEQ counter chains + shared dirty bit

    /// PsychQuant/ooxml-swift#53 — same shape as 53.2 for endnotes.
    func testUpdateAllFieldsMultiEndnoteSEQ() {
        var doc = WordDocument()

        var en1 = Endnote(id: 1, text: "", paragraphIndex: 0)
        en1.bodyChildren = [.paragraph(captionParagraph(identifier: "Source", initialCached: "0"))]
        var en2 = Endnote(id: 2, text: "", paragraphIndex: 1)
        en2.bodyChildren = [.paragraph(captionParagraph(identifier: "Source", initialCached: "2"))]
        doc.endnotes = EndnotesCollection(endnotes: [en1, en2])

        let result = doc.updateAllFields()
        XCTAssertEqual(result, ["Source": 2],
            "multi-endnote SEQ counter must chain: en1 -> 1, en2 -> 2")

        guard case .paragraph(let p1) = doc.endnotes.endnotes[0].bodyChildren[0],
              case .paragraph(let p2) = doc.endnotes.endnotes[1].bodyChildren[0] else {
            return XCTFail("expected paragraphs in endnote bodyChildren")
        }
        XCTAssertEqual(cachedResultOfFirstSEQ(in: p1), "1", "en1's stale cache must be rewritten to 1")
        XCTAssertEqual(cachedResultOfFirstSEQ(in: p2), "2", "en2's cache was already correct (no-op rewrite)")

        XCTAssertTrue(doc.modifiedParts.contains("word/endnotes.xml"),
            "endnotesDirty must stay set from en1's real change — a later endnote whose own " +
            "SEQ rewrite is a no-op (en2) must NOT reset the shared dirty bit")
    }
}
