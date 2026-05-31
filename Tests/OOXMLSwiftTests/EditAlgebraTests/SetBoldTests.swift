// SetBoldTests.swift
// EditAlgebra — Phase 2 of ooxml-edit-isomorphism-foundation (PsychQuant/macdoc#99)
//
// §4 of #105 tasks — setBold emission tests.
//
// SCOPE: Validates OOXMLEdit.setBold → Operation.setRunFormat translation.
// End-to-end deferred behind ooxml-swift#71 (OpLog Phase 2c).

import XCTest
@testable import OOXMLSwift

final class SetBoldTests: XCTestCase {

    // MARK: - Emission: setBold(value: true) → RunFormatPayload(bold: true)

    func testSetBoldTrueEmitsSetRunFormatWithBoldTrue() throws {
        let target = ElementID(libraryUUID: UUID())
        let edit = OOXMLEdit.setBold(target: target, value: true)

        let ops = try edit.operations()
        XCTAssertEqual(ops.count, 1, "setBold lowers to exactly one Operation")

        guard case .setRunFormat(let opTarget, let payload) = ops[0] else {
            XCTFail("Expected Operation.setRunFormat, got \(ops[0])")
            return
        }
        XCTAssertEqual(opTarget, target, "target: ElementID round-trips")
        XCTAssertEqual(payload.bold, true, "value: true lowers to RunFormatPayload.bold = true")
    }

    // MARK: - Emission: setBold(value: false) → RunFormatPayload(bold: false)
    //
    // Important contract: `false` is EXPLICIT — payload.bold == false, NOT nil.
    // nil would mean "leave bold unchanged", but the user said "set to false"
    // (remove bold). See design.md Decision 1 + RunFormatPayload field-semantics
    // (nil = leave unchanged, Bool = set to that value).

    func testSetBoldFalseEmitsSetRunFormatWithBoldFalse() throws {
        let target = ElementID(libraryUUID: UUID())
        let edit = OOXMLEdit.setBold(target: target, value: false)

        let ops = try edit.operations()
        guard case .setRunFormat(_, let payload) = ops[0] else {
            XCTFail("Expected Operation.setRunFormat")
            return
        }
        XCTAssertEqual(payload.bold, false,
                       "value: false → explicit RunFormatPayload.bold = false (NOT nil)")
        XCTAssertNotNil(payload.bold,
                        "bold MUST be non-nil; nil means 'leave unchanged'")
    }

    // MARK: - Emission: other format fields stay nil (single-purpose Edit)

    func testSetBoldDoesNotTouchOtherFormatFields() throws {
        let target = ElementID(libraryUUID: UUID())
        let edit = OOXMLEdit.setBold(target: target, value: true)

        let ops = try edit.operations()
        guard case .setRunFormat(_, let payload) = ops[0] else {
            XCTFail("Expected Operation.setRunFormat")
            return
        }
        XCTAssertNil(payload.italic, "setBold leaves italic unchanged (nil)")
        XCTAssertNil(payload.underline, "setBold leaves underline unchanged (nil)")
        XCTAssertNil(payload.fontSizeHalfPoints, "setBold leaves fontSize unchanged (nil)")
        XCTAssertNil(payload.color, "setBold leaves color unchanged (nil)")
    }

    // MARK: - lower() identity (OOXMLEdit is its own lowering)

    func testSetBoldLowerReturnsSelf() {
        let target = ElementID(libraryUUID: UUID())
        let edit = OOXMLEdit.setBold(target: target, value: true)

        let lowered = edit.lower()
        XCTAssertEqual(lowered.count, 1, "OOXMLEdit.lower() is identity (returns [self])")
        XCTAssertEqual(lowered[0], edit, "OOXMLEdit.lower() returns self verbatim")
    }

    // MARK: - Remaining stubs still throw
    //
    // removeParagraph landed in §6 — coverage in RemoveParagraphTests.
    // insertHyperlink is the only remaining stub (§5).

    func testInsertHyperlinkStillThrowsAfterSetBoldImplemented() {
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
