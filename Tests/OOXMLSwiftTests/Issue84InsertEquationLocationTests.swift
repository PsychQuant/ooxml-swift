import XCTest
@testable import OOXMLSwift

/// `insertEquation(at: InsertLocation)` overload tests for
/// [PsychQuant/che-word-mcp#84](https://github.com/PsychQuant/che-word-mcp/issues/84).
///
/// Pre-fix `WordDocument.insertEquation(at: Int?, latex:, displayMode:)` was
/// the only insert API on the lib that did NOT accept `InsertLocation`,
/// forcing every external Swift SPM consumer (rescue CLI, planned dxedit
/// CLI) to manually reimplement text → bodyChild Int conversion. This test
/// suite pins the new `insertEquation(at: InsertLocation, ...)` overload's
/// behaviour: matches `insertImage` / `insertParagraph` for display-mode
/// equations; rejects non-`paragraphIndex` anchors in inline mode (per
/// che-word-mcp#67 F2 inline-mode guarantee).
final class Issue84InsertEquationLocationTests: XCTestCase {

    /// Spec: `.afterText("anchor")` inserts the new equation paragraph
    /// AFTER the matching paragraph (display mode = new paragraph).
    func testAfterTextResolvesToParagraphAfterMatch() throws {
        var doc = try buildDocWithThreeParagraphs()
        try doc.insertEquation(
            at: .afterText("middle", instance: 1),
            latex: "x = 1",
            displayMode: true
        )
        let texts = paragraphTexts(in: doc)
        // Original 3 paragraphs ("first" / "middle" / "last") become 4
        // with the new (empty-text, equation-bearing) paragraph between
        // "middle" and "last".
        XCTAssertEqual(texts.count, 4, "expected 4 paragraphs after insert: \(texts)")
        XCTAssertEqual(texts[0], "first")
        XCTAssertEqual(texts[1], "middle")
        XCTAssertEqual(texts[3], "last",
            "new equation paragraph should land at index 2, not displace 'last' from index 3")
    }

    /// Spec: `.beforeText("anchor")` inserts the new equation paragraph
    /// BEFORE the matching paragraph.
    func testBeforeTextResolvesToParagraphBeforeMatch() throws {
        var doc = try buildDocWithThreeParagraphs()
        try doc.insertEquation(
            at: .beforeText("middle", instance: 1),
            latex: "y = 2",
            displayMode: true
        )
        let texts = paragraphTexts(in: doc)
        XCTAssertEqual(texts.count, 4)
        XCTAssertEqual(texts[0], "first")
        XCTAssertEqual(texts[2], "middle",
            "new equation paragraph should land at index 1, pushing 'middle' to 2")
        XCTAssertEqual(texts[3], "last")
    }

    /// Spec: `.paragraphIndex(N)` works the same as the legacy `Int?` overload.
    /// Regression guard: the new overload's `.paragraphIndex` path delegates
    /// correctly.
    func testParagraphIndexInsertsAtGivenIndex() throws {
        var doc = try buildDocWithThreeParagraphs()
        try doc.insertEquation(
            at: .paragraphIndex(1),
            latex: "z = 3",
            displayMode: true
        )
        let texts = paragraphTexts(in: doc)
        XCTAssertEqual(texts.count, 4)
        XCTAssertEqual(texts[0], "first")
        // index 1 insertion → original "middle" pushes to 2
        XCTAssertEqual(texts[2], "middle")
        XCTAssertEqual(texts[3], "last")
    }

    /// Spec / che-word-mcp#67 F2 + che-word-mcp#91: inline equation explicitly
    /// rejects non-`paragraphIndex` anchors. `.afterText` in inline mode throws
    /// the dedicated `InsertLocationError.inlineModeRequiresParagraphIndexForAnchor`
    /// case (post-#23; pre-#91 used the misleading `.invalidParagraphIndex(-1)`
    /// sentinel — see `Issue91InlineModeRejectionTests` for full coverage of
    /// all 5 non-paragraphIndex anchor cases).
    func testInlineModeRejectsAfterTextAnchor() throws {
        var doc = try buildDocWithThreeParagraphs()
        XCTAssertThrowsError(try doc.insertEquation(
            at: .afterText("middle", instance: 1),
            latex: "t",
            displayMode: false
        )) { error in
            XCTAssertEqual(error as? InsertLocationError,
                .inlineModeRequiresParagraphIndexForAnchor("afterText"),
                "inline mode must reject non-paragraphIndex anchors with the dedicated error case (#91)")
        }
    }

    /// Verify finding (DA §4.2 + §2.1, BLOCKING in batched-verify of e53fa00):
    /// API-inserted display equations must be visible to `flattenedDisplayText()`
    /// IMMEDIATELY (before any save → reload). Without the in-scope fix that
    /// also sets `run.rawXML` alongside `run.properties.rawXML`, the canonical
    /// batch-CLI workflow (rescue script Phase 5: insert equation → next
    /// anchor lookup) silently resolves anchors against a doc whose prior
    /// inserts are invisible.
    func testInsertEquationThenFlattenSeesMathText() throws {
        var doc = try buildDocWithThreeParagraphs()
        // Insert display equation at index 1 (between "first" and "middle").
        try doc.insertEquation(
            at: .paragraphIndex(1),
            latex: "x = 1",
            displayMode: true
        )
        // The newly-inserted equation paragraph is at body.children[1].
        guard case .paragraph(let eqPara) = doc.body.children[1] else {
            XCTFail("Expected equation paragraph at index 1")
            return
        }
        let flat = eqPara.flattenedDisplayText()
        // The OMML emitted by MathEquation contains the latex tokens; we
        // only assert that flatten is non-empty (the math text appears) —
        // exact content depends on MathEquation's internal escape pipeline.
        XCTAssertFalse(flat.isEmpty,
            "API-inserted equation must flatten with non-empty math text immediately, " +
            "without requiring save → reload round-trip. Got empty flatten — \(flat)")
    }

    /// Spec: text-not-found surface returns `InsertLocationError.textNotFound`.
    /// Mirrors `insertParagraph(at: .afterText)` error semantics.
    func testAfterTextThrowsWhenAnchorNotFound() throws {
        var doc = try buildDocWithThreeParagraphs()
        XCTAssertThrowsError(try doc.insertEquation(
            at: .afterText("does-not-exist", instance: 1),
            latex: "x",
            displayMode: true
        )) { error in
            guard case let InsertLocationError.textNotFound(needle, instance) = error else {
                XCTFail("Expected InsertLocationError.textNotFound, got \(error)")
                return
            }
            XCTAssertEqual(needle, "does-not-exist")
            XCTAssertEqual(instance, 1)
        }
    }

    // MARK: - Helpers

    private func buildDocWithThreeParagraphs() throws -> WordDocument {
        var doc = WordDocument()
        for text in ["first", "middle", "last"] {
            let para = Paragraph(runs: [Run(text: text)])
            doc.body.children.append(.paragraph(para))
        }
        return doc
    }

    private func paragraphTexts(in doc: WordDocument) -> [String] {
        return doc.body.children.compactMap { child in
            guard case .paragraph(let p) = child else { return nil }
            return p.runs.map { $0.text }.joined()
        }
    }
}
