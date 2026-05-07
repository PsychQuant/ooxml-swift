import XCTest
@testable import OOXMLSwift

/// Unit tests for the operation log scaffold (Phase 2a).
///
/// Spectra change: `operation-log-scaffold-impl`, target ooxml-swift v0.31.3.
/// Capability: `ooxml-operation-log`
///
/// Pins eight contracts from the spec scenarios:
/// 1. `Operation` enum has all 21 cases; each constructs and pattern-matches.
/// 2. `ElementID` derives from `w14:paraId` via priority chain.
/// 3. `ElementID` falls back to `libraryUUID` when no native stable ID.
/// 4. `ElementID(node:)` returns nil for bare elements with no identity.
/// 5. `OperationLog.append` increases `entries.count` and preserves source.
/// 6. `OperationLog.batch` wraps body ops in `batchBegin` / `batchEnd`.
/// 7. JSONL round-trip on known ops is byte-equal.
/// 8. JSONL forward-compat: unknown `op_type` round-trips byte-equal via
///    `.unknown(opType:payload:)` fallback (sorted-key input required).
final class OperationLogTests: XCTestCase {

    // MARK: - 1. Operation enum exhaustive case coverage

    /// **Decision pinned**: `Operation` has 21 cases (16 element-level + 4
    /// tree-node-level + 1 unknown). Each constructs and pattern-matches.
    func testOperationEnumEachCaseConstructsAndMatches() {
        let id = ElementID(rawString: "w14:paraId=A")
        let id2 = ElementID(rawString: "w14:paraId=B")
        let uuid = UUID(uuidString: "11111111-1111-4111-8111-111111111111")!

        let cases: [OOXMLSwift.Operation] = [
            .insertParagraphAfter(after: id, paragraph: ParagraphPayload(text: "p")),
            .insertParagraphBefore(before: id, paragraph: ParagraphPayload(text: "p")),
            .removeParagraph(id: id),
            .setText(target: id, text: "Hello"),
            .setParagraphStyle(target: id, styleId: "Heading1"),
            .insertTable(at: id, table: TablePayload(rows: 2, columns: 3)),
            .removeTable(id: id),
            .setCellText(table: id, row: 0, column: 1, text: "cell"),
            .insertRun(in: id, position: 0, run: RunPayload(text: "r")),
            .setRunFormat(target: id, format: RunFormatPayload(bold: true)),
            .insertBookmark(at: id, bookmarkId: 7, name: "anchor"),
            .insertComment(anchor: id, commentId: 3, text: "ct", author: "auth"),
            .undo(targetOpID: uuid),
            .redo(targetOpID: uuid),
            .batchBegin(label: "rename"),
            .batchEnd,
            .insertNode(parent: id, position: 0, nodeXML: "<w:p/>"),
            .removeNode(target: id),
            .updateAttribute(target: id, prefix: "w", localName: "id", value: "5"),
            .moveNode(source: id, destinationParent: id2, destinationIndex: 0),
            .unknown(opType: "future", payload: JSONValue.object(["k": JSONValue.int(1)]))
        ]

        XCTAssertEqual(cases.count, 21, "Operation MUST have exactly 21 cases enumerated in the test")

        // Pattern-match: each case maps to its expected discriminator.
        for op in cases {
            switch op {
            case .insertParagraphAfter, .insertParagraphBefore, .removeParagraph,
                 .setText, .setParagraphStyle,
                 .insertTable, .removeTable, .setCellText,
                 .insertRun, .setRunFormat,
                 .insertBookmark, .insertComment,
                 .undo, .redo,
                 .batchBegin, .batchEnd,
                 .insertNode, .removeNode, .updateAttribute, .moveNode,
                 .unknown:
                break // matched
            }
        }
    }

    // MARK: - 2-4. ElementID derivation

    /// Spec scenario: ElementID derives from w14:paraId
    func testElementIDDerivesFromW14ParaId() {
        let node = XmlNode.element(prefix: "w", localName: "p")
        node.setAttribute(prefix: "w14", localName: "paraId", value: "0ABC1234")

        let id = ElementID(node: node)
        XCTAssertNotNil(id, "ElementID(node:) MUST succeed when w14:paraId is present")
        XCTAssertEqual(id?.raw, "w14:paraId=0ABC1234",
                       "raw MUST byte-align with XmlNode.stableID format")
    }

    /// Spec scenario: ElementID falls back to libraryUUID when no native stable ID
    func testElementIDFallsBackToLibraryUUID() {
        let node = XmlNode.element(prefix: "w", localName: "p")
        node.libraryUUID = UUID(uuidString: "550E8400-E29B-41D4-A716-446655440000")

        let id = ElementID(node: node)
        XCTAssertNotNil(id, "ElementID(node:) MUST fall back to libraryUUID")
        XCTAssertEqual(id?.raw, "lib:550E8400-E29B-41D4-A716-446655440000",
                       "libraryUUID format SHALL be 'lib:<UUID>'")
    }

    /// Spec scenario: ElementID returns nil when no stable identity exists
    func testElementIDReturnsNilForBareElement() {
        let node = XmlNode.element(prefix: "w", localName: "p")
        // No attributes, no libraryUUID.
        let id = ElementID(node: node)
        XCTAssertNil(id, "ElementID(node:) MUST return nil for bare elements")
    }

    // MARK: - 5-6. OperationLog append + batch

    /// Spec scenario: append increases entries count and preserves source
    func testOperationLogAppendIncreasesCount() {
        var log = OperationLog()
        let id = ElementID(rawString: "w14:paraId=X")
        log.append(.setText(target: id, text: "Hello"), source: .swift)

        XCTAssertEqual(log.entries.count, 1)
        XCTAssertEqual(log.entries[0].source, .swift)
        if case .setText(let target, let text) = log.entries[0].op {
            XCTAssertEqual(target, id)
            XCTAssertEqual(text, "Hello")
        } else {
            XCTFail("entry op SHALL be the setText we appended")
        }

        // opIDs are unique across appends.
        log.append(.removeParagraph(id: id), source: .swift)
        XCTAssertNotEqual(log.entries[0].opID, log.entries[1].opID,
                          "default-opID appends MUST produce distinct UUIDs")
    }

    /// Spec scenario: batch wraps body ops with begin/end markers
    func testOperationLogBatchWrapsBodyOps() {
        var log = OperationLog()
        let id = ElementID(rawString: "w14:paraId=X")

        log.batch(.swift, label: "rename") { lb in
            lb.append(.setText(target: id, text: "X"), source: .swift)
            lb.append(.setParagraphStyle(target: id, styleId: "Heading1"), source: .swift)
        }

        XCTAssertEqual(log.entries.count, 4, "batch SHALL emit 4 entries: begin + 2 inner + end")
        if case .batchBegin(let label) = log.entries[0].op {
            XCTAssertEqual(label, "rename")
        } else {
            XCTFail("entries[0] SHALL be batchBegin(label: 'rename')")
        }
        if case .setText = log.entries[1].op {} else { XCTFail("entries[1] SHALL be setText") }
        if case .setParagraphStyle = log.entries[2].op {} else { XCTFail("entries[2] SHALL be setParagraphStyle") }
        if case .batchEnd = log.entries[3].op {} else { XCTFail("entries[3] SHALL be batchEnd") }
    }

    // MARK: - 7-8. JSONL round-trip

    /// Spec scenario: known-ops JSONL round-trip is byte-equal
    func testJSONLKnownOpsRoundTripByteEqual() throws {
        var log = OperationLog()
        log.append(
            .setText(
                target: ElementID(rawString: "w14:paraId=0ABC1234"),
                text: "Hello"
            ),
            source: .swift,
            opID: UUID(uuidString: "11111111-1111-4111-8111-111111111111")!,
            at: Date(timeIntervalSince1970: 1_747_500_000)
        )

        let bytes1 = log.encodeJSONL()
        let decoded = try OperationLog.decodeJSONL(bytes1)
        let bytes2 = decoded.encodeJSONL()

        XCTAssertEqual(bytes1, bytes2, "encode → decode → encode SHALL be byte-equal for known ops")
        XCTAssertEqual(decoded, log, "round-trip Equatable SHALL hold")

        // Sanity: the encoded line has the four required fields in the
        // documented order.
        let line = String(decoding: bytes1, as: UTF8.self).trimmingCharacters(in: .newlines)
        XCTAssertTrue(line.hasPrefix("{\"op_id\":\"11111111-1111-4111-8111-111111111111\""))
        XCTAssertTrue(line.contains("\"source\":\"swift\""))
        XCTAssertTrue(line.contains("\"op_type\":\"setText\""))
        XCTAssertTrue(line.contains("\"target\":\"w14:paraId=0ABC1234\""))
        XCTAssertTrue(line.contains("\"text\":\"Hello\""))
    }

    /// Spec scenario: unknown op_type round-trips byte-equal via the .unknown fallback.
    ///
    /// The spec's Decision 4 + the byte-equal guarantee require payload object
    /// keys to be lexicographically sorted on output. Per ASCII order,
    /// `strike` (0x73) sorts before `target` (0x74). The input bytes therefore
    /// MUST be sorted at write time for round-trip byte-equality to hold —
    /// the test fixture provides keys in sorted order to match the contract.
    func testJSONLForwardCompatRoundTripByteEqual() throws {
        let inputLine = #"{"op_id":"22222222-2222-4222-8222-222222222222","ts":"2026-05-07T02:00:00Z","source":"swift","op_type":"setRunStrikethrough","strike":true,"target":"w14:paraId=Z"}"#
        let inputBytes = Data((inputLine + "\n").utf8)

        let decoded = try OperationLog.decodeJSONL(inputBytes)
        XCTAssertEqual(decoded.entries.count, 1)
        if case .unknown(let opType, let payload) = decoded.entries[0].op {
            XCTAssertEqual(opType, "setRunStrikethrough")
            // Payload SHALL contain target + strike (the only non-required fields).
            if case .object(let dict) = payload {
                XCTAssertEqual(dict["strike"], .bool(true))
                XCTAssertEqual(dict["target"], .string("w14:paraId=Z"))
            } else {
                XCTFail("payload SHALL be a JSONValue.object")
            }
        } else {
            XCTFail("decoded op SHALL be .unknown for unrecognized op_type")
        }

        let reencoded = decoded.encodeJSONL()
        XCTAssertEqual(reencoded, inputBytes,
                       "unknown op_type SHALL round-trip byte-equal when input keys are sorted")
    }

    /// Bonus coverage: malformed line throws .malformedLine
    func testJSONLMalformedLineThrows() {
        // Missing op_type field — should throw.
        let badStr = #"{"op_id":"11111111-1111-4111-8111-111111111111","ts":"2026-05-07T01:00:00Z","source":"swift"}"# + "\n"
        let bad = Data(badStr.utf8)
        XCTAssertThrowsError(try OperationLog.decodeJSONL(bad)) { error in
            guard case OperationLogJSONLError.malformedLine(let lineIndex) = error else {
                XCTFail("expected OperationLogJSONLError.malformedLine; got \(error)")
                return
            }
            XCTAssertEqual(lineIndex, 0)
        }
    }
}
