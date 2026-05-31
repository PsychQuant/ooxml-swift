// InsertParagraphTests.swift
// EditAlgebra — Phase 2 of ooxml-edit-isomorphism-foundation (PsychQuant/macdoc#99)
//
// §3 of #105 tasks — insertParagraph + insertParagraphBefore emission tests.
//
// SCOPE: Validates OOXMLEdit → Operation translation (the `.operations()`
// method). Pure-data assertions; does NOT invoke OperationReducer.
//
// END-TO-END (apply actually mutates xmlTrees) is DEFERRED — see design.md
// Decision 6 (OpLog Phase 2c dependency). The Reducer currently throws
// `malformedOp("Phase 2c implements this op")` for `insertParagraphAfter`
// and `insertParagraphBefore`. Once OpLog Phase 2c lands, add e2e tests
// that build a synthesized WordDocument, apply the Edit, and assert
// xmlTrees["word/document.xml"] contains the new paragraph at the
// expected position.

import XCTest
@testable import OOXMLSwift

final class InsertParagraphTests: XCTestCase {

    // MARK: - insertParagraph (after) emission

    func testInsertParagraphEmitsInsertParagraphAfter() throws {
        let target = ElementID(libraryUUID: UUID())
        let edit = OOXMLEdit.insertParagraph(
            after: target,
            content: "Hello world",
            styleId: nil
        )

        let ops = try edit.operations()
        XCTAssertEqual(ops.count, 1, "insertParagraph lowers to exactly one Operation")

        guard case .insertParagraphAfter(let after, let payload) = ops[0] else {
            XCTFail("Expected Operation.insertParagraphAfter, got \(ops[0])")
            return
        }
        XCTAssertEqual(after, target, "after: ElementID round-trips")
        XCTAssertEqual(payload.text, "Hello world", "content lowers to ParagraphPayload.text")
        XCTAssertNil(payload.styleId, "nil styleId stays nil")
    }

    func testInsertParagraphPreservesStyleId() throws {
        let target = ElementID(libraryUUID: UUID())
        let edit = OOXMLEdit.insertParagraph(
            after: target,
            content: "Section heading",
            styleId: "Heading1"
        )

        let ops = try edit.operations()
        guard case .insertParagraphAfter(_, let payload) = ops[0] else {
            XCTFail("Expected Operation.insertParagraphAfter")
            return
        }
        XCTAssertEqual(payload.styleId, "Heading1", "styleId lowers verbatim")
    }

    func testInsertParagraphEmptyContent() throws {
        let target = ElementID(libraryUUID: UUID())
        let edit = OOXMLEdit.insertParagraph(after: target, content: "", styleId: nil)

        let ops = try edit.operations()
        guard case .insertParagraphAfter(_, let payload) = ops[0] else {
            XCTFail("Expected Operation.insertParagraphAfter")
            return
        }
        XCTAssertEqual(payload.text, "", "Empty content lowers to empty ParagraphPayload.text")
    }

    // MARK: - insertParagraphBefore emission

    func testInsertParagraphBeforeEmitsInsertParagraphBefore() throws {
        let target = ElementID(libraryUUID: UUID())
        let edit = OOXMLEdit.insertParagraphBefore(
            before: target,
            content: "Prologue",
            styleId: "Quote"
        )

        let ops = try edit.operations()
        XCTAssertEqual(ops.count, 1, "insertParagraphBefore lowers to exactly one Operation")

        guard case .insertParagraphBefore(let before, let payload) = ops[0] else {
            XCTFail("Expected Operation.insertParagraphBefore, got \(ops[0])")
            return
        }
        XCTAssertEqual(before, target, "before: ElementID round-trips")
        XCTAssertEqual(payload.text, "Prologue")
        XCTAssertEqual(payload.styleId, "Quote")
    }

    // MARK: - Stub still throws for unimplemented cases

    func testSetBoldStillThrowsNotImplemented() {
        let edit = OOXMLEdit.setBold(target: ElementID(libraryUUID: UUID()), value: true)

        XCTAssertThrowsError(try edit.operations()) { error in
            guard case EditError.notImplemented(let msg) = error else {
                XCTFail("Expected .notImplemented, got \(error)")
                return
            }
            XCTAssertTrue(msg.contains("§4-§6"),
                          "Error message references task batch: \(msg)")
        }
    }

    func testRemoveParagraphStillThrowsNotImplemented() {
        let edit = OOXMLEdit.removeParagraph(target: ElementID(libraryUUID: UUID()))
        XCTAssertThrowsError(try edit.operations()) { error in
            guard case EditError.notImplemented = error else {
                XCTFail("Expected .notImplemented, got \(error)")
                return
            }
        }
    }

    func testInsertHyperlinkStillThrowsNotImplemented() {
        let edit = OOXMLEdit.insertHyperlink(
            target: ElementID(libraryUUID: UUID()),
            href: URL(string: "https://example.com")!,
            displayText: "click"
        )
        XCTAssertThrowsError(try edit.operations()) { error in
            guard case EditError.notImplemented = error else {
                XCTFail("Expected .notImplemented, got \(error)")
                return
            }
        }
    }
}
