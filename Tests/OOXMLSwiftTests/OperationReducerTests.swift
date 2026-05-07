import XCTest
@testable import OOXMLSwift

/// Unit tests for the operation log reducer (Phase 2b).
///
/// Spectra change: `operation-reducer-impl`, target ooxml-swift v0.31.4.
/// Capability: `ooxml-operation-reducer`
///
/// Pins 13 contracts from the spec scenarios:
/// 1. `testMaterialize_pureFunction` — same input twice produces same output.
/// 2. `testMaterialize_doesNotMutateBase` — caller's base tree is untouched.
/// 3. `testMaterialize_appliesSetText` — `setText` op produces expected text.
/// 4. `testState_indexZeroReturnsBaseUnchanged`
/// 5. `testState_indexEqualToCountIsLatest`
/// 6. `testState_timestampFilters`
/// 7. `testState_outOfRangeIndexThrows`
/// 8. `testUndo_setTextReverts`
/// 9. `testUndo_unsupportedOpThrows`
/// 10. `testRedo_restoresOriginalOpEffect`
/// 11. `testBlame_returnsMostRecentTouchingOp`
/// 12. `testCache_tailReplayOnHit`
/// 13. `testReducerError_elementNotFoundOnMissingTarget`
final class OperationReducerTests: XCTestCase {

    // MARK: - Fixture helpers

    /// Build a synthesized `<w:document><w:body>` tree containing one or more
    /// `<w:r>` runs, each with a `libraryUUID` for ElementID addressability and
    /// an initial `<w:t>` text payload. Returns the tree plus the array of
    /// per-run ElementIDs in source order.
    private func makeDocumentWithRuns(_ initialTexts: [String]) -> (XmlTree, [ElementID]) {
        var ids: [ElementID] = []
        var runs: [XmlNode] = []
        for text in initialTexts {
            let runUUID = UUID()
            let t = XmlNode.element(prefix: "w", localName: "t", children: [XmlNode.text(text)])
            let r = XmlNode.element(prefix: "w", localName: "r", children: [t])
            r.libraryUUID = runUUID
            runs.append(r)
            ids.append(ElementID(libraryUUID: runUUID))
        }
        let paras = runs.map { run -> XmlNode in
            XmlNode.element(prefix: "w", localName: "p", children: [run])
        }
        let body = XmlNode.element(prefix: "w", localName: "body", children: paras)
        let doc = XmlNode.element(prefix: "w", localName: "document", children: [body])
        return (XmlTree.synthesized(root: doc), ids)
    }

    /// Walks `tree` looking for the run with the given `ElementID` and returns
    /// the concatenated text of all descendant `<w:t>` element children.
    private func extractRunText(tree: XmlTree, runID: ElementID) -> String? {
        guard let run = OperationReducer.findNode(elementID: runID, in: tree) else {
            return nil
        }
        // Concatenate text from direct <w:t> children (the post-setText shape).
        var pieces: [String] = []
        for child in run.children where child.kind == .element && child.localName == "t" {
            for grand in child.children where grand.kind == .text {
                pieces.append(grand.textContent)
            }
        }
        return pieces.joined()
    }

    // MARK: - 1. Materialize: pure function

    /// Spec: same `(log, base)` input always produces the same output.
    func testMaterialize_pureFunction() throws {
        let (base, ids) = makeDocumentWithRuns(["Old"])
        var log = OperationLog()
        log.append(.setText(target: ids[0], text: "First"), source: .swift)
        log.append(.setText(target: ids[0], text: "Second"), source: .swift)

        let r1 = try OperationReducer.materialize(log: log, base: base)
        let r2 = try OperationReducer.materialize(log: log, base: base)

        XCTAssertEqual(r1.root.normalizedFingerprint(), r2.root.normalizedFingerprint(),
                       "Two materializations of the same input must produce fingerprint-equal trees")
    }

    // MARK: - 2. Materialize: caller's base tree not mutated

    /// Spec: returned tree contains new text; original `base` is unchanged.
    func testMaterialize_doesNotMutateBase() throws {
        let (base, ids) = makeDocumentWithRuns(["Original"])
        let preFingerprint = base.root.normalizedFingerprint()

        var log = OperationLog()
        log.append(.setText(target: ids[0], text: "Mutated"), source: .swift)

        let result = try OperationReducer.materialize(log: log, base: base)

        XCTAssertEqual(extractRunText(tree: result, runID: ids[0]), "Mutated",
                       "Returned tree must reflect the setText op")
        XCTAssertEqual(extractRunText(tree: base, runID: ids[0]), "Original",
                       "Caller's base tree must be unchanged")
        XCTAssertEqual(base.root.normalizedFingerprint(), preFingerprint,
                       "Caller's base tree fingerprint must be unchanged")
    }

    // MARK: - 3. Materialize: applies setText

    func testMaterialize_appliesSetText() throws {
        let (base, ids) = makeDocumentWithRuns(["Old"])
        var log = OperationLog()
        log.append(.setText(target: ids[0], text: "New"), source: .swift)

        let result = try OperationReducer.materialize(log: log, base: base)

        XCTAssertEqual(extractRunText(tree: result, runID: ids[0]), "New")
    }

    // MARK: - 4. State at index(0) returns base unchanged (fingerprint-equal)

    func testState_indexZeroReturnsBaseUnchanged() throws {
        let (base, ids) = makeDocumentWithRuns(["A", "B"])
        var log = OperationLog()
        log.append(.setText(target: ids[0], text: "X"), source: .swift)
        log.append(.setText(target: ids[1], text: "Y"), source: .swift)
        log.append(.setText(target: ids[0], text: "Z"), source: .swift)

        let snapshot = try OperationReducer.state(log: log, base: base, at: .index(0))

        XCTAssertEqual(snapshot.root.normalizedFingerprint(), base.root.normalizedFingerprint(),
                       "state(at: .index(0)) must fingerprint-equal base")
    }

    // MARK: - 5. State at index(N) where N == count is identical to .latest

    func testState_indexEqualToCountIsLatest() throws {
        let (base, ids) = makeDocumentWithRuns(["A", "B"])
        var log = OperationLog()
        log.append(.setText(target: ids[0], text: "X"), source: .swift)
        log.append(.setText(target: ids[1], text: "Y"), source: .swift)
        log.append(.setText(target: ids[0], text: "Z"), source: .swift)

        let viaIndex = try OperationReducer.state(log: log, base: base, at: .index(3))
        let viaLatest = try OperationReducer.state(log: log, base: base, at: .latest)

        XCTAssertEqual(viaIndex.root.normalizedFingerprint(), viaLatest.root.normalizedFingerprint())
    }

    // MARK: - 6. State at timestamp filters by cutoff

    func testState_timestampFilters() throws {
        let (base, ids) = makeDocumentWithRuns(["A"])
        let t0 = Date(timeIntervalSince1970: 1_700_000_000)
        let t1 = Date(timeIntervalSince1970: 1_700_000_100)
        let t2 = Date(timeIntervalSince1970: 1_700_000_200)

        var log = OperationLog()
        log.append(.setText(target: ids[0], text: "First"), source: .swift, at: t0)
        log.append(.setText(target: ids[0], text: "Second"), source: .swift, at: t1)
        log.append(.setText(target: ids[0], text: "Third"), source: .swift, at: t2)

        let snapshot = try OperationReducer.state(log: log, base: base, at: .timestamp(t1))

        XCTAssertEqual(extractRunText(tree: snapshot, runID: ids[0]), "Second",
                       "timestamp cutoff t1 must include entries[0] and entries[1] only; entry[2] (t2 > t1) is excluded")
    }

    // MARK: - 7. State out-of-range index throws

    func testState_outOfRangeIndexThrows() throws {
        let (base, ids) = makeDocumentWithRuns(["A"])
        var log = OperationLog()
        log.append(.setText(target: ids[0], text: "X"), source: .swift)
        log.append(.setText(target: ids[0], text: "Y"), source: .swift)
        log.append(.setText(target: ids[0], text: "Z"), source: .swift)

        XCTAssertThrowsError(try OperationReducer.state(log: log, base: base, at: .index(5))) { err in
            guard case ReducerError.malformedOp(_, let reason) = err else {
                XCTFail("Expected ReducerError.malformedOp, got \(err)")
                return
            }
            XCTAssertEqual(reason, "index out of range")
        }
    }

    // MARK: - 8. Undo of setText reverts to prior text

    func testUndo_setTextReverts() throws {
        let (base, ids) = makeDocumentWithRuns(["Initial"])
        let opA = UUID(uuidString: "AAAAAAAA-AAAA-4AAA-8AAA-AAAAAAAAAAAA")!
        let opB = UUID(uuidString: "BBBBBBBB-BBBB-4BBB-8BBB-BBBBBBBBBBBB")!

        var log = OperationLog()
        log.append(.setText(target: ids[0], text: "Old"), source: .swift, opID: opA)
        log.append(.setText(target: ids[0], text: "New"), source: .swift, opID: opB)

        let result = try OperationReducer.undo(opB, log: log, base: base)

        XCTAssertEqual(extractRunText(tree: result, runID: ids[0]), "Old",
                       "Undoing opB must revert to opA's text")
    }

    // MARK: - 9. Undo of unsupported op throws

    func testUndo_unsupportedOpThrows() throws {
        let (base, ids) = makeDocumentWithRuns(["A"])
        let opID = UUID()

        var log = OperationLog()
        log.append(
            .insertTable(at: ids[0], table: TablePayload(rows: 2, columns: 3)),
            source: .swift,
            opID: opID
        )

        XCTAssertThrowsError(try OperationReducer.undo(opID, log: log, base: base)) { err in
            guard case ReducerError.cannotUndo(let target) = err else {
                XCTFail("Expected ReducerError.cannotUndo, got \(err)")
                return
            }
            XCTAssertEqual(target, opID)
        }
    }

    // MARK: - 10. Redo restores the original op's effect

    func testRedo_restoresOriginalOpEffect() throws {
        let (base, ids) = makeDocumentWithRuns(["X-original", "Y-original"])
        let opA = UUID(uuidString: "11111111-1111-4111-8111-111111111111")!
        let opB = UUID(uuidString: "22222222-2222-4222-8222-222222222222")!
        let opC = UUID(uuidString: "33333333-3333-4333-8333-333333333333")!

        var log = OperationLog()
        log.append(.setText(target: ids[0], text: "Original"), source: .swift, opID: opA)
        log.append(.undo(targetOpID: opA), source: .swift, opID: opB)
        log.append(.setText(target: ids[1], text: "Other"), source: .swift, opID: opC)

        let result = try OperationReducer.redo(opA, log: log, base: base)

        XCTAssertEqual(extractRunText(tree: result, runID: ids[0]), "Original",
                       "redo(opA) must skip the .undo entry so opA stays in effect")
        XCTAssertEqual(extractRunText(tree: result, runID: ids[1]), "Other",
                       "Other ops (opC) must continue to apply normally")
    }

    // MARK: - 11. Blame returns the most recent touching op

    func testBlame_returnsMostRecentTouchingOp() throws {
        let (_, ids) = makeDocumentWithRuns(["X", "Y"])
        let opA = UUID(uuidString: "AAAA1111-1111-4111-8111-111111111111")!
        let opB = UUID(uuidString: "BBBB2222-2222-4222-8222-222222222222")!
        let opC = UUID(uuidString: "CCCC3333-3333-4333-8333-333333333333")!

        var log = OperationLog()
        log.append(.setText(target: ids[0], text: "A"), source: .swift, opID: opA)
        log.append(.setText(target: ids[1], text: "B"), source: .swift, opID: opB)
        log.append(.setText(target: ids[0], text: "C"), source: .swift, opID: opC)

        let blamed = OperationReducer.blame(elementID: ids[0], log: log)

        XCTAssertEqual(blamed?.opID, opC, "blame on element X must return opC (most recent touching op)")

        // Untouched ElementID returns nil.
        let phantom = ElementID(libraryUUID: UUID())
        XCTAssertNil(OperationReducer.blame(elementID: phantom, log: log))
    }

    // MARK: - 12. Cache: tail-replay on hit produces same result as fresh materialize

    func testCache_tailReplayOnHit() async throws {
        let (base, ids) = makeDocumentWithRuns(["A", "B"])

        var log = OperationLog()
        log.append(.setText(target: ids[0], text: "step1"), source: .swift)
        log.append(.setText(target: ids[1], text: "step2"), source: .swift)
        log.append(.setText(target: ids[0], text: "step3"), source: .swift)

        let cache = OperationReducerCache()
        let firstResult = try await cache.materialize(log: log, base: base)

        // Sanity: cache returns the same materialization as the pure reducer.
        let pureFirst = try OperationReducer.materialize(log: log, base: base)
        XCTAssertEqual(firstResult.root.normalizedFingerprint(), pureFirst.root.normalizedFingerprint())

        // Append two more entries and call again — tail-replay path.
        log.append(.setText(target: ids[1], text: "step4"), source: .swift)
        log.append(.setText(target: ids[0], text: "step5"), source: .swift)

        let secondResult = try await cache.materialize(log: log, base: base)
        let pureSecond = try OperationReducer.materialize(log: log, base: base)

        XCTAssertEqual(secondResult.root.normalizedFingerprint(), pureSecond.root.normalizedFingerprint(),
                       "Cache tail-replay result must fingerprint-equal a fresh materialize call")
        XCTAssertEqual(extractRunText(tree: secondResult, runID: ids[0]), "step5")
        XCTAssertEqual(extractRunText(tree: secondResult, runID: ids[1]), "step4")
    }

    // MARK: - 13. ReducerError.elementNotFound on missing target

    func testReducerError_elementNotFoundOnMissingTarget() throws {
        let (base, _) = makeDocumentWithRuns(["A"])
        let phantomID = ElementID(libraryUUID: UUID())
        let opID = UUID()

        var log = OperationLog()
        log.append(.setText(target: phantomID, text: "x"), source: .swift, opID: opID)

        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: base)) { err in
            guard case ReducerError.elementNotFound(let errOpID, let errElementID) = err else {
                XCTFail("Expected ReducerError.elementNotFound, got \(err)")
                return
            }
            XCTAssertEqual(errOpID, opID)
            XCTAssertEqual(errElementID, phantomID)
        }
    }
}
