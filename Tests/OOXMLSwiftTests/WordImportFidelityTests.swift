import XCTest
@testable import OOXMLSwift

/// 7.4 verify panel findings (v33-correctness) — Word-import fidelity:
/// P0: Word-inferred paragraph inserts must carry the REAL w14:paraId
///     (WordImport payload + reducer stamp), or flush() erases Word's
///     identity from disk.
/// P2: DocxChangeDetector.poll() must not commit the mtime baseline
///     before the content-hash read succeeds.
/// P3: setRuns on a non-paragraph target throws malformedOp instead of
///     silently corrupting the node's children.
final class WordImportFidelityTests: XCTestCase {

    private func para(_ text: String, paraId: String?) -> XmlNode {
        let ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        let t = XmlNode.text(text)
        let wt = XmlNode.element(prefix: "w", localName: "t", namespaceURI: ns, children: [t])
        let r = XmlNode.element(prefix: "w", localName: "r", namespaceURI: ns, children: [wt])
        let p = XmlNode.element(prefix: "w", localName: "p", namespaceURI: ns, children: [r])
        if let paraId { p.setAttribute(prefix: "w14", localName: "paraId", value: paraId) }
        return p
    }

    private func bodyTree(_ paras: [XmlNode]) -> XmlTree {
        let ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        let body = XmlNode.element(prefix: "w", localName: "body", namespaceURI: ns, children: paras)
        let doc = XmlNode.element(prefix: "w", localName: "document", namespaceURI: ns, children: [body])
        return XmlTree.synthesized(root: doc)
    }

    // MARK: P0 — inferred inserts carry the real Word paraId end-to-end

    func testWordInsertedParagraphKeepsItsParaIdInPayload() throws {
        let snapshot = bodyTree([para("first", paraId: "AAAA1111")])
        let current = bodyTree([para("first", paraId: "AAAA1111"),
                                para("added in Word", paraId: "BBBB2222")])

        let ops = WordImport.diff(snapshot: snapshot, current: current)

        guard case .insertParagraphAfter(_, let payload)? = ops.operations.first(where: {
            if case .insertParagraphAfter = $0 { return true } else { return false }
        }) else {
            return XCTFail("expected an inferred insertParagraphAfter, got \(ops.operations)")
        }
        XCTAssertEqual(payload.paraId, "BBBB2222",
                       "the Word-assigned w14:paraId must ride the payload (P0: identity loss on flush)")
    }

    func testReducerStampsParaIdOnInsertAfterAndBefore() throws {
        let tree = bodyTree([para("anchor", paraId: "ANCH0001")])
        var log = OperationLog()
        log.append(.insertParagraphAfter(
            after: ElementID(rawString: "w14:paraId=ANCH0001"),
            paragraph: ParagraphPayload(text: "after", styleId: nil, paraId: "AFT00001")), source: .word)
        log.append(.insertParagraphBefore(
            before: ElementID(rawString: "w14:paraId=ANCH0001"),
            paragraph: ParagraphPayload(text: "before", styleId: nil, paraId: "BEF00001")), source: .word)

        let out = try OperationReducer.materialize(log: log, base: tree)
        let body = out.root.children.first { $0.localName == "body" }!
        let ids = body.children.compactMap { node in
            node.attributes.first { $0.prefix == "w14" && $0.localName == "paraId" }?.value
        }
        XCTAssertEqual(Set(ids), ["ANCH0001", "AFT00001", "BEF00001"],
                       "insertParagraphAfter/Before must stamp payload.paraId like appendParagraph does; got \(ids)")
    }

    // MARK: P2 — poll() baseline not committed before hash read succeeds

    func testPollFailureDoesNotAdvanceMtimeBaseline() throws {
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("poll-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try Data("v1".utf8).write(to: url)
        var detector = try DocxChangeDetector(url: url)

        // Change content (mtime + bytes), then make the file unreadable so
        // the hash read throws AFTER the mtime fast-path would have differed.
        try Data("v2 — changed".utf8).write(to: url)
        try FileManager.default.setAttributes([.posixPermissions: 0o000], ofItemAtPath: url.path)
        XCTAssertThrowsError(try detector.poll(), "unreadable file must throw")

        // Restore readability WITHOUT touching mtime: the retry must still
        // report the change (the old code had already swallowed the mtime).
        try FileManager.default.setAttributes([.posixPermissions: 0o644], ofItemAtPath: url.path)
        XCTAssertTrue(try detector.poll(),
                      "a change whose first poll failed mid-read must still be reported on retry")
    }

    // MARK: P3 — setRuns target-kind validation

    func testSetRunsOnNonParagraphThrowsMalformedOp() throws {
        let ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        let tbl = XmlNode.element(prefix: "w", localName: "tbl", namespaceURI: ns)
        tbl.setAttribute(prefix: "w14", localName: "paraId", value: "NOTAPARA")
        let tree = bodyTree([tbl])
        var log = OperationLog()
        log.append(.setRuns(target: ElementID(rawString: "w14:paraId=NOTAPARA"),
                            runs: [RunPayload(text: "x")]), source: .swift)

        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: tree)) { error in
            guard case ReducerError.malformedOp = error else {
                return XCTFail("expected malformedOp on non-<w:p> setRuns target, got \(error)")
            }
        }
    }
}
