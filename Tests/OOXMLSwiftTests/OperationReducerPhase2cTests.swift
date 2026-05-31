// OperationReducerPhase2cTests.swift
// OpLog Phase 2c — tree-mutating Operation cases in OperationReducer.
// PsychQuant/ooxml-swift#71
//
// Implements the cases macdoc#105 (Edit algebra) needs end-to-end:
//   1. insertParagraphAfter  (this file — pioneer case)
//   2. insertParagraphBefore (symmetric — added once 1 is green)
//   3. removeParagraph
//   4. setRunFormat
//   5. insertNode + updateAttribute (composite primitives)
//
// Each case validates:
//   (a) tree mutated correctly per Operation semantics
//   (b) ID stability — existing nodes' ElementIDs unchanged
//   (c) New node ID is deterministic from entry.opID (so log replay is stable)

import XCTest
@testable import OOXMLSwift

final class OperationReducerPhase2cTests: XCTestCase {

    // MARK: - Test helpers

    /// Builds a synthesized WordDocument body with N paragraphs, each
    /// containing one text run with the given text. Returns the tree +
    /// each paragraph's ElementID for addressability in tests.
    private func makeBodyWithParagraphs(_ texts: [String]) -> (XmlTree, [ElementID]) {
        var paraIDs: [ElementID] = []
        var paragraphs: [XmlNode] = []

        for text in texts {
            let paraUUID = UUID()
            paraIDs.append(ElementID(libraryUUID: paraUUID))

            let textNode = XmlNode.text(text)
            let wt = XmlNode.element(prefix: "w", localName: "t", children: [textNode])
            let wr = XmlNode.element(prefix: "w", localName: "r", children: [wt])
            let wp = XmlNode.element(prefix: "w", localName: "p", children: [wr])
            wp.libraryUUID = paraUUID
            paragraphs.append(wp)
        }

        let body = XmlNode.element(prefix: "w", localName: "body", children: paragraphs)
        let doc = XmlNode.element(prefix: "w", localName: "document", children: [body])
        return (XmlTree.synthesized(root: doc), paraIDs)
    }

    /// Walks tree, returns text content of all <w:t> descendants in document order.
    private func extractAllText(_ tree: XmlTree) -> [String] {
        var result: [String] = []
        func walk(_ node: XmlNode) {
            if node.kind == .element && node.localName == "t" {
                for child in node.children where child.kind == .text {
                    result.append(child.textContent)
                }
            }
            for child in node.children {
                walk(child)
            }
        }
        walk(tree.root)
        return result
    }

    /// Returns the libraryUUIDs of all <w:p> descendants in document order.
    private func extractParagraphIDs(_ tree: XmlTree) -> [UUID?] {
        var result: [UUID?] = []
        func walk(_ node: XmlNode) {
            if node.kind == .element && node.localName == "p" {
                result.append(node.libraryUUID)
            }
            for child in node.children {
                walk(child)
            }
        }
        walk(tree.root)
        return result
    }

    // MARK: - insertParagraphAfter

    func testInsertParagraphAfterAppendsToBody() throws {
        // GIVEN body with one paragraph "first"
        let (base, paraIDs) = makeBodyWithParagraphs(["first"])
        XCTAssertEqual(extractAllText(base), ["first"])

        // WHEN insertParagraphAfter with content "second"
        var log = OperationLog()
        log.append(
            .insertParagraphAfter(
                after: paraIDs[0],
                paragraph: ParagraphPayload(text: "second", styleId: nil)
            ),
            source: .swift
        )

        let result = try OperationReducer.materialize(log: log, base: base)

        // THEN body contains 2 paragraphs in order: first, second
        XCTAssertEqual(extractAllText(result), ["first", "second"],
                       "New paragraph appears AFTER target in document order")
    }

    func testInsertParagraphAfterPreservesExistingParagraphID() throws {
        let (base, paraIDs) = makeBodyWithParagraphs(["first"])
        let originalFirstID = paraIDs[0].libraryUUID

        var log = OperationLog()
        log.append(
            .insertParagraphAfter(
                after: paraIDs[0],
                paragraph: ParagraphPayload(text: "second", styleId: nil)
            ),
            source: .swift
        )

        let result = try OperationReducer.materialize(log: log, base: base)
        let resultIDs = extractParagraphIDs(result)

        XCTAssertEqual(resultIDs.count, 2, "Two paragraphs in result")
        XCTAssertEqual(resultIDs[0], originalFirstID,
                       "Existing paragraph's libraryUUID unchanged")
        XCTAssertNotNil(resultIDs[1], "New paragraph has a libraryUUID")
    }

    func testInsertParagraphAfterNewParagraphIDDerivesFromOpID() throws {
        // Determinism check: new paragraph's libraryUUID == entry.opID.
        // This makes log replay produce the same tree every time.
        let (base, paraIDs) = makeBodyWithParagraphs(["first"])

        let determinedOpID = UUID()
        var log = OperationLog()
        log.append(
            .insertParagraphAfter(
                after: paraIDs[0],
                paragraph: ParagraphPayload(text: "second", styleId: nil)
            ),
            source: .swift,
            opID: determinedOpID
        )

        let result = try OperationReducer.materialize(log: log, base: base)
        let resultIDs = extractParagraphIDs(result)

        XCTAssertEqual(resultIDs[1], determinedOpID,
                       "New paragraph's libraryUUID == entry.opID (deterministic replay)")
    }

    func testInsertParagraphAfterReplayIsIdempotent() throws {
        // Same log applied twice → same tree shape (modulo deep-clone).
        let (base, paraIDs) = makeBodyWithParagraphs(["first"])

        var log = OperationLog()
        log.append(
            .insertParagraphAfter(
                after: paraIDs[0],
                paragraph: ParagraphPayload(text: "second", styleId: nil)
            ),
            source: .swift,
            opID: UUID()  // fixed opID
        )

        let first = try OperationReducer.materialize(log: log, base: base)
        let second = try OperationReducer.materialize(log: log, base: base)

        XCTAssertEqual(extractAllText(first), extractAllText(second),
                       "Replaying same log produces same text content")
        XCTAssertEqual(extractParagraphIDs(first), extractParagraphIDs(second),
                       "Replaying same log produces same paragraph IDs")
    }

    func testInsertParagraphAfterTargetNotFoundThrows() {
        let (base, _) = makeBodyWithParagraphs(["first"])

        // Reference a non-existent ElementID
        var log = OperationLog()
        log.append(
            .insertParagraphAfter(
                after: ElementID(libraryUUID: UUID()),  // never inserted into base
                paragraph: ParagraphPayload(text: "second", styleId: nil)
            ),
            source: .swift
        )

        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: base)) { error in
            guard case ReducerError.elementNotFound = error else {
                XCTFail("Expected .elementNotFound, got \(error)")
                return
            }
        }
    }

    func testInsertParagraphAfterMultipleCases() throws {
        // Sequence: start with [a], insert "b" after a → [a, b]; insert "c" after b → [a, b, c].
        let (base, paraIDs) = makeBodyWithParagraphs(["a"])
        let aID = paraIDs[0]

        let opIDB = UUID()
        let opIDC = UUID()

        var log = OperationLog()
        log.append(
            .insertParagraphAfter(after: aID, paragraph: ParagraphPayload(text: "b", styleId: nil)),
            source: .swift,
            opID: opIDB
        )
        // After the first op, new paragraph "b"'s libraryUUID == opIDB.
        log.append(
            .insertParagraphAfter(
                after: ElementID(libraryUUID: opIDB),
                paragraph: ParagraphPayload(text: "c", styleId: nil)
            ),
            source: .swift,
            opID: opIDC
        )

        let result = try OperationReducer.materialize(log: log, base: base)
        XCTAssertEqual(extractAllText(result), ["a", "b", "c"],
                       "Sequential insertParagraphAfter chains via deterministic new IDs")
    }
}
