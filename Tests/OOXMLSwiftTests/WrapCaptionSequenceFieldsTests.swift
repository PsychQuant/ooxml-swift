import XCTest
@testable import OOXMLSwift

/// PsychQuant/che-word-mcp#62 — `Document.wrapCaptionSequenceFields` lib API.
///
/// Spec: openspec/specs/ooxml-content-insertion-primitives/spec.md (Requirement
/// `Document.wrapCaptionSequenceFields converts plain-text caption number
/// portions to SEQ-field runs`).
final class WrapCaptionSequenceFieldsTests: XCTestCase {

    // MARK: - 1.9.1 Body scope wraps three plain-text figure captions

    func testBodyScopeWrapsThreePlainTextFigureCaptions() throws {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "圖 4-1：架構圖")),
            .paragraph(Paragraph(text: "圖 4-2：流程圖")),
            .paragraph(Paragraph(text: "圖 4-3：時序圖")),
            .paragraph(Paragraph(text: "lorem ipsum")),
            .paragraph(Paragraph(text: "dolor sit amet")),
            .paragraph(Paragraph(text: "consectetur adipiscing")),
            .paragraph(Paragraph(text: "elit sed do")),
            .paragraph(Paragraph(text: "eiusmod tempor")),
        ]
        let pattern = try NSRegularExpression(pattern: "圖 4-(\\d+)：")
        let result = try doc.wrapCaptionSequenceFields(
            pattern: pattern,
            sequenceName: "Figure"
        )

        XCTAssertEqual(result.matchedParagraphs, 3)
        XCTAssertEqual(result.fieldsInserted, 3)
        XCTAssertEqual(result.paragraphsModified, [0, 1, 2])
        XCTAssertTrue(result.skipped.isEmpty, "no idempotency skips on first run")

        // Each modified paragraph still displays the original digits AND now
        // contains a SEQ Figure field somewhere in its run rawXML.
        for idx in [0, 1, 2] {
            guard case .paragraph(let p) = doc.body.children[idx] else {
                XCTFail("expected paragraph at idx \(idx); got \(doc.body.children[idx])")
                return
            }
            let containsSEQ = p.runs.contains { ($0.rawXML ?? "").contains("SEQ Figure") }
            XCTAssertTrue(containsSEQ, "paragraph \(idx) must carry SEQ Figure field rawXML")
        }
    }

    // MARK: - 1.9.2 Idempotent re-run skips already-wrapped paragraphs

    func testIdempotentSecondCallSkipsAlreadyWrapped() throws {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "圖 4-1：架構圖")),
            .paragraph(Paragraph(text: "圖 4-2：流程圖")),
            .paragraph(Paragraph(text: "圖 4-3：時序圖")),
        ]
        let pattern = try NSRegularExpression(pattern: "圖 4-(\\d+)：")
        _ = try doc.wrapCaptionSequenceFields(pattern: pattern, sequenceName: "Figure")

        // Second run — all 3 should report as skipped, none modified.
        let result = try doc.wrapCaptionSequenceFields(pattern: pattern, sequenceName: "Figure")

        XCTAssertEqual(result.matchedParagraphs, 3)
        XCTAssertEqual(result.fieldsInserted, 0)
        XCTAssertTrue(result.paragraphsModified.isEmpty)
        XCTAssertEqual(result.skipped.count, 3)
        for s in result.skipped {
            XCTAssertEqual(s.reason, "already wraps SEQ Figure")
            XCTAssertNil(s.container)
        }
    }

    // MARK: - 1.9.3 Pattern with zero capture groups throws

    func testPatternWithZeroCaptureGroupsThrows() throws {
        var doc = WordDocument()
        doc.body.children = [.paragraph(Paragraph(text: "圖 4-1：title"))]
        let pattern = try NSRegularExpression(pattern: "圖 4-\\d+：")  // no group

        XCTAssertThrowsError(try doc.wrapCaptionSequenceFields(
            pattern: pattern,
            sequenceName: "Figure"
        )) { error in
            guard case WrapCaptionError.patternMissingCaptureGroup(let actual) = error else {
                XCTFail("expected patternMissingCaptureGroup; got \(error)")
                return
            }
            XCTAssertEqual(actual, 0)
        }
        // Document should not be mutated by a rejected call.
        guard case .paragraph(let p) = doc.body.children[0] else { return XCTFail() }
        XCTAssertEqual(p.runs.first?.text, "圖 4-1：title")
    }

    // MARK: - 1.9.4 Pattern with two capture groups throws

    func testPatternWithTwoCaptureGroupsThrows() throws {
        var doc = WordDocument()
        doc.body.children = [.paragraph(Paragraph(text: "Figure 1.2"))]
        let pattern = try NSRegularExpression(pattern: "Figure (\\d+)\\.(\\d+)")  // two groups

        XCTAssertThrowsError(try doc.wrapCaptionSequenceFields(
            pattern: pattern,
            sequenceName: "Figure"
        )) { error in
            guard case WrapCaptionError.patternMissingCaptureGroup(let actual) = error else {
                XCTFail("expected patternMissingCaptureGroup; got \(error)")
                return
            }
            XCTAssertEqual(actual, 2)
        }
    }

    // MARK: - 1.9.5 Bookmark wrapping with template substitution

    func testBookmarkWrappingWithTemplateSubstitution() throws {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "Figure 7. Distribution of MoCA scores")),
        ]
        let pattern = try NSRegularExpression(pattern: "Figure (\\d+)\\.")
        let result = try doc.wrapCaptionSequenceFields(
            pattern: pattern,
            sequenceName: "Figure",
            insertBookmark: true,
            bookmarkTemplate: "fig${number}"
        )

        XCTAssertEqual(result.fieldsInserted, 1)
        XCTAssertEqual(result.paragraphsModified, [0])

        guard case .paragraph(let p) = doc.body.children[0] else {
            XCTFail("expected paragraph")
            return
        }
        // bookmarkStart should appear before SEQ field, bookmarkEnd after.
        // Concatenating raw XML in run order should expose both markers around the SEQ run.
        let allRawXML = p.runs.map { $0.rawXML ?? "" }.joined()
        XCTAssertTrue(allRawXML.contains("<w:bookmarkStart"))
        XCTAssertTrue(allRawXML.contains("w:name=\"fig7\""))
        XCTAssertTrue(allRawXML.contains("<w:bookmarkEnd"))
        XCTAssertTrue(allRawXML.contains("SEQ Figure"))

        // bookmarkStart MUST appear before SEQ rawXML, bookmarkEnd after.
        let startRange = allRawXML.range(of: "<w:bookmarkStart")!
        let seqRange = allRawXML.range(of: "SEQ Figure")!
        let endRange = allRawXML.range(of: "<w:bookmarkEnd")!
        XCTAssertLessThan(startRange.lowerBound, seqRange.lowerBound)
        XCTAssertLessThan(seqRange.lowerBound, endRange.lowerBound)
    }

    // MARK: - 1.9.6 insert_bookmark = true without bookmark_template throws

    func testInsertBookmarkTrueWithoutTemplateThrows() throws {
        var doc = WordDocument()
        doc.body.children = [.paragraph(Paragraph(text: "Figure 1. caption"))]
        let pattern = try NSRegularExpression(pattern: "Figure (\\d+)")

        XCTAssertThrowsError(try doc.wrapCaptionSequenceFields(
            pattern: pattern,
            sequenceName: "Figure",
            insertBookmark: true,
            bookmarkTemplate: nil
        )) { error in
            guard case WrapCaptionError.bookmarkTemplateMissing = error else {
                XCTFail("expected bookmarkTemplateMissing; got \(error)")
                return
            }
        }

        // Template missing ${number} placeholder also throws.
        XCTAssertThrowsError(try doc.wrapCaptionSequenceFields(
            pattern: pattern,
            sequenceName: "Figure",
            insertBookmark: true,
            bookmarkTemplate: "fig"  // no ${number}
        )) { error in
            guard case WrapCaptionError.bookmarkTemplateMissing = error else {
                XCTFail("expected bookmarkTemplateMissing; got \(error)")
                return
            }
        }
    }

    // MARK: - 1.9.7 Idempotency covers both fldSimple AND rawXML emissions

    func testIdempotencyCoversBothFldSimpleAndRawXMLEmissions() throws {
        var doc = WordDocument()

        // Paragraph 0 — already has FieldSimple-style SEQ.
        var paraFS = Paragraph(text: "Figure 1. existing fldSimple caption")
        let fs = FieldSimple(instr: " SEQ Figure \\* ARABIC ", runs: [Run(text: "1")])
        paraFS.fieldSimples = [fs]

        // Paragraph 1 — already has rawXML-embedded SEQ (insertCaption emission style).
        var paraRaw = Paragraph(text: "Figure 2. existing rawXML caption")
        var seqRun = Run(text: "")
        seqRun.rawXML = SequenceField(
            identifier: "Figure",
            format: .arabic,
            cachedResult: "2"
        ).toFieldXML()
        paraRaw.runs.append(seqRun)

        doc.body.children = [.paragraph(paraFS), .paragraph(paraRaw)]
        let pattern = try NSRegularExpression(pattern: "Figure (\\d+)\\.")
        let result = try doc.wrapCaptionSequenceFields(pattern: pattern, sequenceName: "Figure")

        XCTAssertEqual(result.matchedParagraphs, 2)
        XCTAssertEqual(result.fieldsInserted, 0)
        XCTAssertEqual(result.skipped.count, 2)
        for s in result.skipped {
            XCTAssertEqual(s.reason, "already wraps SEQ Figure")
        }
    }

    // MARK: - 1.9.8 Cached result preserves user numerals before F9

    func testCachedResultPreservesUserNumeralsBeforeF9() throws {
        var doc = WordDocument()
        doc.body.children = [.paragraph(Paragraph(text: "圖 4-7：架構圖"))]
        let pattern = try NSRegularExpression(pattern: "圖 4-(\\d+)：")
        _ = try doc.wrapCaptionSequenceFields(pattern: pattern, sequenceName: "Figure")

        guard case .paragraph(let p) = doc.body.children[0] else {
            XCTFail("expected paragraph")
            return
        }
        // The SEQ run rawXML should encode cachedResult = "7" so Word's
        // first-open render shows the original numbering before F9.
        let seqRun = p.runs.first { ($0.rawXML ?? "").contains("SEQ Figure") }
        XCTAssertNotNil(seqRun)
        XCTAssertTrue(seqRun!.rawXML!.contains("<w:t>7</w:t>"),
                      "cachedResult must preserve captured digit '7' for first-open render")
    }

    // MARK: - 1.9.9 Table cell anchor paragraph wraps

    func testTableCellAnchorParagraphWraps() throws {
        var doc = WordDocument()
        let cell = TableCell(paragraphs: [Paragraph(text: "Table 1. Demographics")])
        let row = TableRow(cells: [cell])
        let table = Table(rows: [row])
        doc.body.children = [
            .paragraph(Paragraph(text: "intro paragraph")),
            .table(table),
            .paragraph(Paragraph(text: "trailer")),
        ]

        let pattern = try NSRegularExpression(pattern: "Table (\\d+)\\.")
        let result = try doc.wrapCaptionSequenceFields(pattern: pattern, sequenceName: "Table")

        XCTAssertEqual(result.matchedParagraphs, 1)
        XCTAssertEqual(result.fieldsInserted, 1)
        XCTAssertEqual(result.paragraphsModified, [1], "matched paragraph reports table's body idx (#68 semantic)")

        guard case .table(let t) = doc.body.children[1] else {
            XCTFail("expected table at idx 1")
            return
        }
        let cellPara = t.rows[0].cells[0].paragraphs[0]
        let containsSEQ = cellPara.runs.contains { ($0.rawXML ?? "").contains("SEQ Table") }
        XCTAssertTrue(containsSEQ, "cell paragraph must carry SEQ Table field rawXML")
    }

    // MARK: - 1.9.10 Scope .all throws scopeNotImplemented in Phase 1

    func testScopeAllThrowsScopeNotImplementedInPhase1() throws {
        var doc = WordDocument()
        doc.body.children = [.paragraph(Paragraph(text: "Figure 1. caption"))]
        let pattern = try NSRegularExpression(pattern: "Figure (\\d+)\\.")

        XCTAssertThrowsError(try doc.wrapCaptionSequenceFields(
            pattern: pattern,
            sequenceName: "Figure",
            scope: .all
        )) { error in
            guard case WrapCaptionError.scopeNotImplemented(let scope) = error else {
                XCTFail("expected scopeNotImplemented; got \(error)")
                return
            }
            XCTAssertEqual(scope, .all)
        }
    }
}
