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

    // §5 (insertHyperlink + wrapWithHyperlink emission) shipped — no OOXMLEdit
    // case stubs remain at operations(). The stub-mechanism test
    // (testApplyPropagatesNotImplementedFromOperationsStub) has been deleted
    // per its original "DELETE once §5 ships" instruction. Error-pipeline
    // coverage continues via testApplyWrapsElementNotFoundAsOperationLogFailure
    // (Reducer-level error wrap, lines below) and the WordEdit empty-lower
    // guard test (next method).

    func testApplyThrowsOnStubWordEditEmptyLower() {
        // §1 WordEdit.lower() returns [] — defensive check in WordDocument.apply
        // detects non-OOXMLEdit returning empty list and throws notImplemented.
        // This prevents the silent no-op trap that would let callers think
        // their applyBold succeeded when actually nothing happened.
        // §7 of macdoc#105 ships per-case WordEdit.lower() implementations;
        // this guard becomes dormant for real cases.
        let doc = WordDocument()
        let range = WordRange(
            startRun: ElementID(libraryUUID: UUID()),
            startOffset: 0,
            endRun: ElementID(libraryUUID: UUID()),
            endOffset: 5
        )
        let edit = WordEdit.applyBold(range: range)

        XCTAssertThrowsError(try doc.apply(edit)) { error in
            guard case EditError.notImplemented(let message) = error else {
                XCTFail("Expected .notImplemented on stub WordEdit, got \(error)")
                return
            }
            XCTAssertTrue(message.contains("WordEdit") || message.contains("stub"),
                          "Error references WordEdit/stub: \(message)")
            XCTAssertTrue(message.contains("§7") || message.contains("macdoc#105"),
                          "Error references task batch: \(message)")
        }
    }

    func testApplyWrapsElementNotFoundAsOperationLogFailure() throws {
        // Per spec.md "WordDocument.apply Public Method" item #4 (PHASED):
        // target-not-found surfaces via Reducer-wrapping as
        // EditError.operationLogFailure (until Phase 2c follow-up ships
        // upfront EditError.pathNotFound validation).
        //
        // Build a WordDocument with a known paragraph, then issue an
        // insertParagraph against a NON-EXISTENT ElementID. The Reducer
        // throws elementNotFound, which WordDocument.apply wraps.
        let paraUUID = UUID()
        let textNode = XmlNode.text("present")
        let wt = XmlNode.element(prefix: "w", localName: "t", children: [textNode])
        let wr = XmlNode.element(prefix: "w", localName: "r", children: [wt])
        let wp = XmlNode.element(prefix: "w", localName: "p", children: [wr])
        wp.libraryUUID = paraUUID
        let body = XmlNode.element(prefix: "w", localName: "body", children: [wp])
        let root = XmlNode.element(prefix: "w", localName: "document", children: [body])
        var doc = WordDocument()
        doc.xmlTrees["word/document.xml"] = XmlTree.synthesized(root: root)

        // Edit references an ElementID NOT in the doc
        let bogusID = ElementID(libraryUUID: UUID())
        let edit = OOXMLEdit.insertParagraph(after: bogusID, content: "new", styleId: nil)

        XCTAssertThrowsError(try doc.apply(edit)) { error in
            // Pin the CASE — spec.md PHASED behavior #4 says target-not-found
            // surfaces as operationLogFailure (NOT as pathNotFound until Phase
            // 2c upfront-validation lands). The wrapped `underlying` string
            // format is intentionally not pinned (Swift's default
            // localizedDescription is locale-dependent and the structured
            // ElementID is lost in the wrap — a documented limitation of the
            // PHASED contract; Phase 2c follow-up restores structured info).
            guard case EditError.operationLogFailure = error else {
                XCTFail("Expected .operationLogFailure (PHASED behavior per spec.md #4), got \(error)")
                return
            }
        }
    }

    // MARK: - Sequence-folding apply

    func testApplyEmptySequence() throws {
        let doc = WordDocument()
        let result = try doc.apply([] as [any Edit])
        XCTAssertEqual(doc, result, "Empty sequence apply is identity")
    }

    func testApplySingleEditViaSequence() throws {
        // Single-element sequence should match single-edit apply behavior.
        // Both succeed on insertParagraph through a single-part doc.
        let paraUUID = UUID()
        let wt = XmlNode.element(prefix: "w", localName: "t",
                                  children: [XmlNode.text("existing")])
        let wr = XmlNode.element(prefix: "w", localName: "r", children: [wt])
        let wp = XmlNode.element(prefix: "w", localName: "p", children: [wr])
        wp.libraryUUID = paraUUID
        let body = XmlNode.element(prefix: "w", localName: "body", children: [wp])
        let root = XmlNode.element(prefix: "w", localName: "document",
                                    children: [body])

        var doc = WordDocument()
        doc.xmlTrees["word/document.xml"] = XmlTree.synthesized(root: root)
        let paraID = ElementID(libraryUUID: paraUUID)
        let edit = OOXMLEdit.insertParagraph(after: paraID, content: "new", styleId: nil)

        let viaSequence = try doc.apply([edit] as [any Edit])
        let viaSingle = try doc.apply(edit)
        XCTAssertEqual(viaSequence, viaSingle,
                       "Sequence apply with one element ≡ single apply (content-equal)")
    }

    // MARK: - Immutability tests

    func testApplyDoesNotMutateInput() throws {
        // Even on error, input doc should be unchanged. We use a target
        // ElementID that doesn't resolve in the (empty) doc → apply throws
        // operationLogFailure (PHASED #4 — Reducer wraps elementNotFound).
        let doc = WordDocument()
        let initialLogCount = doc.operationLog.entries.count

        let edit = OOXMLEdit.removeParagraph(target: ElementID(libraryUUID: UUID()))
        _ = try? doc.apply(edit)  // expected to throw (no tree to address)

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
