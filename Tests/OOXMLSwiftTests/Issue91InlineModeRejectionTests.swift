import XCTest
@testable import OOXMLSwift

/// Inline-mode rejection contract tests for `WordDocument.insertEquation`
/// (`Sources/OOXMLSwift/Models/Document.swift` `insertEquation(at: InsertLocation, ...)`).
///
/// Per [PsychQuant/che-word-mcp#91](https://github.com/PsychQuant/che-word-mcp/issues/91)
/// — verify findings from #84's 6-AI ensemble (Logic §2.5 + Devil's Advocate
/// §2.2 + §2.5):
///
/// 1. **Sentinel misuse**: pre-fix, inline-mode rejection threw
///    `InsertLocationError.invalidParagraphIndex(-1)` — a structural lie because
///    the case is documented for "out-of-range paragraph index", not
///    "non-paragraphIndex anchor". Caller patterns like
///    `catch let .invalidParagraphIndex(idx) { print("idx \(idx) out of range") }`
///    mis-reported inline rejection as `"-1 is out of range"`.
/// 2. **Silent no-op on bad inline index**: `.paragraphIndex(9999)` in inline
///    mode silently delegated to the deprecated `Int?` overload, which silently
///    does nothing on out-of-range index. Display mode threw on the same input
///    — asymmetric error semantics for one public API.
///
/// Post-fix contract: BOTH cases throw, with the new dedicated
/// `InsertLocationError.inlineModeRequiresParagraphIndex` for non-paragraphIndex
/// anchors and the existing `InsertLocationError.invalidParagraphIndex(idx)` for
/// bad index in inline mode.
final class Issue91InlineModeRejectionTests: XCTestCase {

    // MARK: - Defect 2: silent no-op on bad inline index

    /// Spec: inline-mode `insertEquation` with `.paragraphIndex(9999)` (or any
    /// out-of-range index) MUST throw, mirroring display-mode's bounds-check.
    /// Pre-fix: silent no-op (delegated to deprecated `Int?` overload).
    func testInlineModeThrowsOnBadParagraphIndex() throws {
        var doc = try buildDocWithThreeParagraphs()
        XCTAssertThrowsError(try doc.insertEquation(
            at: .paragraphIndex(9999),
            latex: "x",
            displayMode: false
        )) { error in
            guard case let InsertLocationError.invalidParagraphIndex(idx) = error else {
                XCTFail("Expected InsertLocationError.invalidParagraphIndex, got \(error)")
                return
            }
            XCTAssertEqual(idx, 9999,
                "error case should carry the actual bad index, not the -1 sentinel")
        }
    }

    /// Spec: inline-mode `insertEquation` with negative index MUST throw.
    func testInlineModeThrowsOnNegativeParagraphIndex() throws {
        var doc = try buildDocWithThreeParagraphs()
        XCTAssertThrowsError(try doc.insertEquation(
            at: .paragraphIndex(-1),
            latex: "x",
            displayMode: false
        )) { error in
            guard case let InsertLocationError.invalidParagraphIndex(idx) = error else {
                XCTFail("Expected InsertLocationError.invalidParagraphIndex, got \(error)")
                return
            }
            XCTAssertEqual(idx, -1)
        }
    }

    // MARK: - Defect 1: sentinel misuse for non-paragraphIndex anchors

    /// Spec: inline-mode `insertEquation` with `.afterImageId` MUST throw the
    /// dedicated `inlineModeRequiresParagraphIndex` error, NOT the
    /// `invalidParagraphIndex(-1)` sentinel.
    func testInlineModeThrowsOnAfterImageId() throws {
        var doc = try buildDocWithThreeParagraphs()
        XCTAssertThrowsError(try doc.insertEquation(
            at: .afterImageId("rId-nonexistent"),
            latex: "x",
            displayMode: false
        )) { error in
            XCTAssertEqual(error as? InsertLocationError,
                .inlineModeRequiresParagraphIndex,
                "inline mode rejection must use the dedicated error case, not invalidParagraphIndex(-1)")
        }
    }

    /// Spec: inline-mode `insertEquation` with `.intoTableCell` MUST throw the
    /// dedicated `inlineModeRequiresParagraphIndex` error.
    func testInlineModeThrowsOnIntoTableCell() throws {
        var doc = try buildDocWithThreeParagraphs()
        XCTAssertThrowsError(try doc.insertEquation(
            at: .intoTableCell(tableIndex: 0, row: 0, col: 0),
            latex: "x",
            displayMode: false
        )) { error in
            XCTAssertEqual(error as? InsertLocationError,
                .inlineModeRequiresParagraphIndex)
        }
    }

    /// Spec: inline-mode `insertEquation` with `.afterTableIndex` MUST throw
    /// the dedicated `inlineModeRequiresParagraphIndex` error.
    func testInlineModeThrowsOnAfterTableIndex() throws {
        var doc = try buildDocWithThreeParagraphs()
        XCTAssertThrowsError(try doc.insertEquation(
            at: .afterTableIndex(0),
            latex: "x",
            displayMode: false
        )) { error in
            XCTAssertEqual(error as? InsertLocationError,
                .inlineModeRequiresParagraphIndex)
        }
    }

    /// Spec: inline-mode `insertEquation` with `.beforeText` MUST throw the
    /// dedicated `inlineModeRequiresParagraphIndex` error (the symmetric
    /// counterpart to the existing `.afterText` test in
    /// `Issue84InsertEquationLocationTests.testInlineModeRejectsAfterTextAnchor`).
    func testInlineModeThrowsOnBeforeText() throws {
        var doc = try buildDocWithThreeParagraphs()
        XCTAssertThrowsError(try doc.insertEquation(
            at: .beforeText("middle", instance: 1),
            latex: "x",
            displayMode: false
        )) { error in
            XCTAssertEqual(error as? InsertLocationError,
                .inlineModeRequiresParagraphIndex)
        }
    }

    // MARK: - F1 corrective: SDT-nested + empty-doc + boundary

    /// Verify F1 corrective (convergent finding: Codex P1 + Devil's Advocate
    /// DA-1): pre-corrective bounds-check used `getParagraphs().count` which
    /// recurses into block-level `.contentControl`, but the legacy delegate at
    /// `Document.swift:4044` only walks top-level `.paragraph` body children
    /// (no SDT descent). For a doc shape `[.contentControl(_, [.paragraph])]`,
    /// the SDT-nested paragraph counted toward `getParagraphs().count == 1`,
    /// so `.paragraphIndex(0)` passed the old guard but the delegate found
    /// zero top-level paragraph matches and silently no-op'd — exactly the
    /// failure class Defect 2 was meant to eliminate, just shifted to
    /// SDT-nested docs.
    ///
    /// Post-corrective: bounds-check uses top-level `.paragraph` count only,
    /// so `.paragraphIndex(0)` against an SDT-nested-only doc throws
    /// `invalidParagraphIndex(0)` instead of silent no-op.
    func testInlineModeThrowsOnSdtNestedOnlyDoc() throws {
        var doc = WordDocument()
        // Doc shape: only an SDT containing a paragraph; ZERO top-level paragraphs
        let sdt = StructuredDocumentTag(id: 100, tag: "test-sdt")
        let cc = ContentControl(sdt: sdt, content: "")
        let nestedPara = Paragraph(runs: [Run(text: "inside-sdt")])
        doc.body.children = [.contentControl(cc, children: [.paragraph(nestedPara)])]

        XCTAssertThrowsError(try doc.insertEquation(
            at: .paragraphIndex(0),
            latex: "x",
            displayMode: false
        )) { error in
            guard case let InsertLocationError.invalidParagraphIndex(idx) = error else {
                XCTFail("Expected InsertLocationError.invalidParagraphIndex(0) for SDT-nested-only doc, got \(error)")
                return
            }
            XCTAssertEqual(idx, 0,
                "F1 corrective: bounds-check must count only top-level paragraphs, not recurse into SDTs")
        }
    }

    /// Verify F1 corrective (mixed top-level + SDT). Doc shape:
    /// `[.paragraph(p0), .contentControl(_, [.paragraph(sdt-p)])]` has
    /// `getParagraphs().count == 2` (recurses) but only ONE top-level paragraph.
    /// Pre-corrective: `.paragraphIndex(1)` passed `idx < 2` guard then silent
    /// no-op'd in delegate. Post-corrective: throws `invalidParagraphIndex(1)`.
    func testInlineModeThrowsOnIdxBeyondTopLevelButWithinGetParagraphs() throws {
        var doc = WordDocument()
        let p0 = Paragraph(runs: [Run(text: "top-level-0")])
        let sdt = StructuredDocumentTag(id: 100, tag: "test-sdt")
        let cc = ContentControl(sdt: sdt, content: "")
        let nestedPara = Paragraph(runs: [Run(text: "inside-sdt")])
        doc.body.children = [
            .paragraph(p0),
            .contentControl(cc, children: [.paragraph(nestedPara)]),
        ]

        // getParagraphs() returns 2 (recurses); body.children top-level .paragraph count = 1
        XCTAssertEqual(doc.getParagraphs().count, 2,
            "sanity: getParagraphs recurses into SDT")

        // Pre-corrective: passes idx < 2 guard then silent no-op
        // Post-corrective: throws invalidParagraphIndex(1)
        XCTAssertThrowsError(try doc.insertEquation(
            at: .paragraphIndex(1),
            latex: "x",
            displayMode: false
        )) { error in
            guard case let InsertLocationError.invalidParagraphIndex(idx) = error else {
                XCTFail("Expected invalidParagraphIndex(1) — bounds-check should narrow to top-level only, got \(error)")
                return
            }
            XCTAssertEqual(idx, 1)
        }

        // Sanity: idx 0 (top-level p0) still works correctly (does NOT throw)
        XCTAssertNoThrow(try doc.insertEquation(
            at: .paragraphIndex(0),
            latex: "y",
            displayMode: false
        ), "valid top-level idx must still succeed post-corrective")
    }

    /// F6 verify follow-up: empty-doc edge case (`paragraphCount == 0`) +
    /// `.paragraphIndex(0)` inline mode throws `invalidParagraphIndex(0)`.
    func testInlineModeThrowsOnEmptyDocument() throws {
        var doc = WordDocument()
        // Empty body — no children at all
        XCTAssertEqual(doc.body.children.count, 0)

        XCTAssertThrowsError(try doc.insertEquation(
            at: .paragraphIndex(0),
            latex: "x",
            displayMode: false
        )) { error in
            guard case let InsertLocationError.invalidParagraphIndex(idx) = error else {
                XCTFail("Expected invalidParagraphIndex(0), got \(error)")
                return
            }
            XCTAssertEqual(idx, 0)
        }
    }

    /// F7 verify follow-up: boundary case (`.paragraphIndex(paragraphCount)`)
    /// — the off-by-one most likely to hit callers by mistake. Inline mode
    /// must throw (no append-at-end semantics for inline; that's display-mode's
    /// job via `insertParagraph`).
    func testInlineModeThrowsOnBoundaryAtParagraphCount() throws {
        var doc = try buildDocWithThreeParagraphs()
        // 3 top-level paragraphs; .paragraphIndex(3) is exactly at boundary
        XCTAssertThrowsError(try doc.insertEquation(
            at: .paragraphIndex(3),
            latex: "x",
            displayMode: false
        )) { error in
            guard case let InsertLocationError.invalidParagraphIndex(idx) = error else {
                XCTFail("Expected invalidParagraphIndex(3), got \(error)")
                return
            }
            XCTAssertEqual(idx, 3,
                "boundary case: idx == paragraphCount must throw inline (no append semantics)")
        }
    }

    // MARK: - Display-mode regression guard

    /// Spec: display-mode behaviour is UNCHANGED. `.afterImageId` non-existent
    /// id still throws `InsertLocationError.imageIdNotFound`, NOT the new
    /// inline-mode error case.
    func testDisplayModeAfterImageIdStillThrowsImageIdNotFound() throws {
        var doc = try buildDocWithThreeParagraphs()
        XCTAssertThrowsError(try doc.insertEquation(
            at: .afterImageId("rId-nonexistent"),
            latex: "x",
            displayMode: true
        )) { error in
            guard case let InsertLocationError.imageIdNotFound(rId) = error else {
                XCTFail("Expected InsertLocationError.imageIdNotFound, got \(error)")
                return
            }
            XCTAssertEqual(rId, "rId-nonexistent")
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
}
