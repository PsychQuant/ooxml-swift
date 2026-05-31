// DocumentApplyTests.swift
// EditAlgebra — Phase 2 of ooxml-edit-isomorphism-foundation (PsychQuant/macdoc#99)
//
// §2 of #105 tasks — Document.apply public API smoke tests.
//
// §2 ships the apply pipeline wiring (lower → operations → log append →
// materialize → return new WordDocument). Tests here validate the pipeline
// surfaces errors correctly through each layer:
//   - notImplemented from OOXMLEdit.operations() (§1 stub) propagates
//   - Empty edit list → no-op (identity)
//   - operationLog field starts empty + is excluded from Equatable
//   - Sequence-folding apply<S> chains correctly
//
// End-to-end behavior (apply actually modifies xmlTrees) lands in §3+ when
// per-OOXMLEdit-case operations() implementations exist.

import XCTest
@testable import OOXMLSwift

final class DocumentApplyTests: XCTestCase {

    // MARK: - operationLog field tests

    func testOperationLogStartsEmpty() {
        let doc = WordDocument()
        XCTAssertEqual(doc.operationLog.entries.count, 0,
                       "Fresh WordDocument has empty operationLog")
    }

    func testOperationLogExcludedFromEquatable() {
        // Two WordDocuments with identical content but different logs should
        // still be Equatable-equal (per design.md Decision 3).
        var doc1 = WordDocument()
        let doc2 = WordDocument()

        let elementID = ElementID(libraryUUID: UUID())
        doc1.operationLog.append(
            .setText(target: elementID, text: "hello"),
            source: .swift
        )

        XCTAssertNotEqual(doc1.operationLog, doc2.operationLog,
                          "Logs differ — explicit log comparison shows it")
        XCTAssertEqual(doc1, doc2,
                       "Content equal — Equatable excludes log per Decision 3")
    }

    // MARK: - apply pipeline wiring tests

    func testApplyReturnsNewDocument() throws {
        // Use empty edit sequence — should be identity (no ops emitted)
        let doc = WordDocument()
        let result = try doc.apply([] as [any Edit])

        XCTAssertEqual(doc, result, "Empty apply is identity (content equal)")
        XCTAssertEqual(doc.operationLog.entries.count,
                       result.operationLog.entries.count,
                       "Empty apply doesn't append to log")
    }

    func testApplyPropagatesNotImplementedFromOperationsStub() {
        // §1 scaffold has OOXMLEdit.operations() throwing notImplemented.
        // §2 apply pipeline should surface that error through lower → operations.
        let doc = WordDocument()
        let edit = OOXMLEdit.removeParagraph(target: ElementID(libraryUUID: UUID()))

        XCTAssertThrowsError(try doc.apply(edit)) { error in
            guard case EditError.notImplemented(let message) = error else {
                XCTFail("Expected .notImplemented (from §1 stub), got \(error)")
                return
            }
            XCTAssertTrue(message.contains("§5-§6"),
                          "Stub error message references remaining task batch: \(message)")
        }
    }

    func testApplyPropagatesNotImplementedFromWordEditLowerStub() {
        // §1 WordEdit.lower() returns [] — empty list → no ops → no error.
        // This documents the current stub behavior (will change in §7 when
        // WordEdit.lower() returns real translations).
        let doc = WordDocument()
        let range = WordRange(
            startRun: ElementID(libraryUUID: UUID()),
            startOffset: 0,
            endRun: ElementID(libraryUUID: UUID()),
            endOffset: 5
        )
        let edit = WordEdit.applyBold(range: range)

        // Currently stub: WordEdit.lower() returns [] → no operations() called.
        // §7 will make this throw notImplemented for the actual OOXMLEdit.
        let result = try? doc.apply(edit)
        XCTAssertNotNil(result, "Stub WordEdit.lower() empty list → no-op apply (will change in §7)")
        XCTAssertEqual(doc, result, "Stub WordEdit apply is identity")
    }

    // MARK: - Sequence-folding apply

    func testApplyEmptySequence() throws {
        let doc = WordDocument()
        let result = try doc.apply([] as [any Edit])
        XCTAssertEqual(doc, result, "Empty sequence apply is identity")
    }

    func testApplySingleEditViaSequence() {
        // Single-element sequence should match single-edit apply behavior.
        // Both throw the same notImplemented stub error (asserted separately
        // to avoid nested XCTAssertThrowsError — Swift's throwing-closure
        // type contract doesn't allow a throwing `try` inside the non-throwing
        // ErrorHandler closure of the outer XCTAssertThrowsError).
        let doc = WordDocument()
        let edit = OOXMLEdit.removeParagraph(target: ElementID(libraryUUID: UUID()))

        // Sequence apply path
        XCTAssertThrowsError(try doc.apply([edit] as [any Edit])) { seqError in
            guard case EditError.notImplemented = seqError else {
                XCTFail("Expected .notImplemented from sequence apply, got \(seqError)")
                return
            }
        }

        // Single apply path (separately)
        XCTAssertThrowsError(try doc.apply(edit)) { singleError in
            guard case EditError.notImplemented = singleError else {
                XCTFail("Expected .notImplemented from single apply, got \(singleError)")
                return
            }
        }
    }

    // MARK: - Immutability tests

    func testApplyDoesNotMutateInput() throws {
        // Even on error, input doc should be unchanged
        let doc = WordDocument()
        let initialLogCount = doc.operationLog.entries.count

        let edit = OOXMLEdit.removeParagraph(target: ElementID(libraryUUID: UUID()))
        _ = try? doc.apply(edit)  // expected to throw notImplemented

        XCTAssertEqual(doc.operationLog.entries.count, initialLogCount,
                       "Input doc's log is unchanged after failed apply")
    }

    func testApplyMutationIsolation() throws {
        // If apply DID succeed and modify newDoc, original doc's xmlTrees
        // should NOT be affected (value semantics)
        let doc = WordDocument()
        let snapshotTrees = doc.xmlTrees

        // Try a stub apply (will throw, but isolation should hold)
        _ = try? doc.apply([] as [any Edit])

        XCTAssertEqual(doc.xmlTrees.keys.sorted(), snapshotTrees.keys.sorted(),
                       "Input doc's xmlTrees keys unchanged")
    }
}
