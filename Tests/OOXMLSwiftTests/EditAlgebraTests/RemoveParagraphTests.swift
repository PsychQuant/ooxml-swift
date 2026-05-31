// RemoveParagraphTests.swift
// EditAlgebra — Phase 2 of ooxml-edit-isomorphism-foundation (PsychQuant/macdoc#99)
//
// §6 of #105 tasks — removeParagraph emission tests.
//
// SCOPE: Validates OOXMLEdit.removeParagraph → Operation.removeParagraph
// translation. End-to-end deferred behind ooxml-swift#71.
//
// Note: Operation.removeParagraph uses `id:` label, OOXMLEdit.removeParagraph
// uses `target:`. The lowering must translate label names — this test pins
// the label translation to prevent silent breakage if either side renames.

import XCTest
@testable import OOXMLSwift

final class RemoveParagraphTests: XCTestCase {

    func testRemoveParagraphEmitsOperationRemoveParagraph() throws {
        let target = ElementID(libraryUUID: UUID())
        let edit = OOXMLEdit.removeParagraph(target: target)

        let ops = try edit.operations()
        XCTAssertEqual(ops.count, 1, "removeParagraph lowers to exactly one Operation")

        guard case .removeParagraph(let opId) = ops[0] else {
            XCTFail("Expected Operation.removeParagraph, got \(ops[0])")
            return
        }
        XCTAssertEqual(opId, target,
                       "OOXMLEdit.removeParagraph(target:) → Operation.removeParagraph(id:) — ElementID round-trips across label rename")
    }

    func testRemoveParagraphLowerReturnsSelf() {
        let target = ElementID(libraryUUID: UUID())
        let edit = OOXMLEdit.removeParagraph(target: target)

        let lowered = edit.lower()
        XCTAssertEqual(lowered.count, 1)
        XCTAssertEqual(lowered[0], edit, "OOXMLEdit.lower() is identity")
    }

    // MARK: - Insertions remain stubbed (insertHyperlink only — §5 pending)

    func testInsertHyperlinkStillThrowsAfterRemoveParagraphImplemented() {
        let edit = OOXMLEdit.insertHyperlink(
            target: ElementID(libraryUUID: UUID()),
            href: URL(string: "https://example.com")!,
            displayText: "click"
        )
        XCTAssertThrowsError(try edit.operations()) { error in
            guard case EditError.notImplemented(let msg) = error else {
                XCTFail("Expected .notImplemented, got \(error)")
                return
            }
            XCTAssertTrue(msg.contains("§5"),
                          "Error references composite design pending: \(msg)")
        }
    }
}
