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
