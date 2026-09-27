import XCTest
@testable import OOXMLSwift

/// R2 (RED→GREEN) regression coverage for the three findings from the
/// independent adversarial review of `Operation.setRunFormatInRange`
/// (PsychQuant/macdoc#232, review report `review-ox-232.md`):
///
/// - CRITICAL-1: split fragments shared the pre-split run's ElementID.
/// - CRITICAL-2: a run with multiple `<w:t>` (e.g. interleaved `<w:tab/>`)
///   silently produced duplicated text instead of being refused.
/// - HIGH-3: a match could silently cross a tracked-insertion/deletion,
///   hyperlink, or SDT boundary and treat the text on either side as
///   adjacent, when it visibly is not.
///
/// See `OperationReducer.applySetRunFormatInRange`'s doc comment for the
/// normative "Flatten/adjacency rule" and "ElementID allocation rule" this
/// file pins.
final class OperationReducerSetRunFormatInRangeR2Tests: XCTestCase {

    // MARK: - Shared helpers

    private func makeParagraph(runs texts: [String]) -> (XmlTree, paragraphID: ElementID, runIDs: [ElementID]) {
        let paraUUID = UUID()
        var runNodes: [XmlNode] = []
        var runIDs: [ElementID] = []
        for text in texts {
            let wt = XmlNode.element(prefix: "w", localName: "t", children: [XmlNode.text(text)])
            let wr = XmlNode.element(prefix: "w", localName: "r", children: [wt])
            let uuid = UUID()
            wr.libraryUUID = uuid
            runNodes.append(wr)
            runIDs.append(ElementID(libraryUUID: uuid))
        }
        let wp = XmlNode.element(prefix: "w", localName: "p", children: runNodes)
        wp.libraryUUID = paraUUID
        let body = XmlNode.element(prefix: "w", localName: "body", children: [wp])
        let doc = XmlNode.element(prefix: "w", localName: "document", children: [body])
        return (XmlTree.synthesized(root: doc), ElementID(libraryUUID: paraUUID), runIDs)
    }

    private func runsOf(_ tree: XmlTree, paragraphID: ElementID) -> [XmlNode] {
        guard let paragraph = OperationReducer.findNode(elementID: paragraphID, in: tree) else {
            return []
        }
        return paragraph.children.filter { $0.kind == .element && $0.localName == "r" }
    }

    private func text(of run: XmlNode) -> String {
        run.children
            .filter { $0.kind == .element && $0.localName == "t" }
            .flatMap { $0.children.filter { $0.kind == .text }.map(\.textContent) }
            .joined()
    }

    private func isBold(_ run: XmlNode) -> Bool {
        guard let rPr = run.children.first(where: { $0.kind == .element && $0.localName == "rPr" })
        else { return false }
        return rPr.children.contains { $0.kind == .element && $0.localName == "b" }
    }

    private func isItalic(_ run: XmlNode) -> Bool {
        guard let rPr = run.children.first(where: { $0.kind == .element && $0.localName == "rPr" })
        else { return false }
        return rPr.children.contains { $0.kind == .element && $0.localName == "i" }
    }

    private func reason(of error: Error) -> String {
        guard case ReducerError.malformedOp(_, let reason) = error else {
            XCTFail("Expected .malformedOp, got \(error)")
            return ""
        }
        return reason
    }

    // MARK: - CRITICAL-1: distinct ElementIDs after split

    /// Single-run split into 3 fragments: prefix keeps the ORIGINAL run's
    /// id; middle and suffix get NEW ids derived from the op's opID
    /// (index 1 and 2 respectively — see the "ElementID allocation rule"
    /// doc comment). All three must be pairwise distinct.
    func testSplitFragmentsHaveDistinctElementIDsAndFollowTheAllocationRule() throws {
        let (base, paragraphID, runIDs) = makeParagraph(runs: ["Hello world today"])
        let originalRunID = runIDs[0]
        let opID = UUID()

        var log = OperationLog()
        log.append(
            .setRunFormatInRange(
                target: paragraphID, substring: "world", instance: 1,
                format: RunFormatPayload(bold: true)),
            source: .swift, opID: opID)

        let result = try OperationReducer.materialize(log: log, base: base)
        let runs = runsOf(result, paragraphID: paragraphID)
        XCTAssertEqual(runs.map(text(of:)), ["Hello ", "world", " today"])

        let ids = runs.map { ElementID(node: $0) }
        XCTAssertEqual(Set(ids.compactMap { $0 }).count, 3, "all three fragments must have distinct ElementIDs")

        XCTAssertEqual(ids[0], originalRunID, "prefix (first fragment in document order) keeps the original run's id")
        XCTAssertEqual(ids[1], ElementID.derivedFragmentID(fromOpID: opID, index: 1),
                       "matched middle gets the index-1 derived id (split needed at the match's START)")
        XCTAssertEqual(ids[2], ElementID.derivedFragmentID(fromOpID: opID, index: 2),
                       "suffix gets the index-2 derived id (split needed at the match's END)")
    }

    /// A follow-up op in the SAME op-log addressing the PRE-SPLIT run's
    /// ElementID must land on the fragment that kept that id (the prefix,
    /// "Hello ") — not silently hit an unrelated fragment.
    func testFollowUpOpOnOriginalRunIDHitsThePrefixFragment() throws {
        let (base, paragraphID, runIDs) = makeParagraph(runs: ["Hello world today"])
        let originalRunID = runIDs[0]

        var log = OperationLog()
        log.append(
            .setRunFormatInRange(
                target: paragraphID, substring: "world", instance: 1,
                format: RunFormatPayload(bold: true)),
            source: .swift)
        log.append(
            .setRunFormat(target: originalRunID, format: RunFormatPayload(italic: true)),
            source: .swift)

        let result = try OperationReducer.materialize(log: log, base: base)
        let runs = runsOf(result, paragraphID: paragraphID)

        let italicized = runs.filter(isItalic)
        XCTAssertEqual(italicized.count, 1, "exactly one fragment should be italicized")
        XCTAssertEqual(italicized.first.map(text(of:)), "Hello ",
                       "a follow-up op on the pre-split run's id lands on the fragment that kept it")
    }

    /// A follow-up op addressing the NEW fragment id (obtained via
    /// `ElementID.derivedFragmentID(fromOpID:index:)`) lands precisely on
    /// the matched "world" fragment.
    func testFollowUpOpOnDerivedFragmentIDHitsTheMatchedFragment() throws {
        let (base, paragraphID, _) = makeParagraph(runs: ["Hello world today"])
        let opID = UUID()

        var log = OperationLog()
        log.append(
            .setRunFormatInRange(
                target: paragraphID, substring: "world", instance: 1,
                format: RunFormatPayload(bold: true)),
            source: .swift, opID: opID)
        log.append(
            .setRunFormat(
                target: ElementID.derivedFragmentID(fromOpID: opID, index: 1),
                format: RunFormatPayload(italic: true)),
            source: .swift)

        let result = try OperationReducer.materialize(log: log, base: base)
        let runs = runsOf(result, paragraphID: paragraphID)

        let italicized = runs.filter(isItalic)
        XCTAssertEqual(italicized.count, 1)
        XCTAssertEqual(italicized.first.map(text(of:)), "world",
                       "a follow-up op on the derived fragment id lands on the matched fragment")
    }

    /// Cross-run split (start run's tail + end run's head both split off):
    /// walking every run in the resulting document, no two runs may share
    /// an ElementID.
    func testNoDuplicateElementIDsAfterCrossRunSplit() throws {
        let (base, paragraphID, _) = makeParagraph(runs: ["foo", "bar", "baz", "qux"])

        var log = OperationLog()
        log.append(
            .setRunFormatInRange(
                target: paragraphID, substring: "obarbazq", instance: 1,
                format: RunFormatPayload(bold: true)),
            source: .swift)

        let result = try OperationReducer.materialize(log: log, base: base)
        let runs = runsOf(result, paragraphID: paragraphID)
        XCTAssertEqual(runs.map(text(of:)), ["fo", "o", "bar", "baz", "q", "ux"])

        let ids = runs.compactMap { ElementID(node: $0) }
        XCTAssertEqual(ids.count, runs.count, "every fragment must have a resolvable ElementID")
        XCTAssertEqual(Set(ids).count, ids.count, "no two fragments may share an ElementID")
    }

    /// Replaying the same log twice (fresh materialize from `base` each
    /// time) must reproduce byte-identical derived fragment ids — the
    /// reducer's pure-function/replay-determinism contract.
    func testDerivedFragmentIDsAreReplayDeterministic() throws {
        let (base, paragraphID, _) = makeParagraph(runs: ["Hello world today"])
        var log = OperationLog()
        log.append(
            .setRunFormatInRange(
                target: paragraphID, substring: "world", instance: 1,
                format: RunFormatPayload(bold: true)),
            source: .swift, opID: UUID(uuidString: "11111111-1111-4111-8111-111111111111")!)

        let result1 = try OperationReducer.materialize(log: log, base: base)
        let result2 = try OperationReducer.materialize(log: log, base: base)

        let ids1 = runsOf(result1, paragraphID: paragraphID).map { ElementID(node: $0)?.raw }
        let ids2 = runsOf(result2, paragraphID: paragraphID).map { ElementID(node: $0)?.raw }
        XCTAssertEqual(ids1, ids2, "replaying the same log must reproduce identical fragment ids")
    }

    // MARK: - CRITICAL-2: runs with multiple <w:t> / interleaved tab / br

    private func makeMixedContentRunParagraph(
        _ children: [XmlNode]
    ) -> (XmlTree, paragraphID: ElementID) {
        let run = XmlNode.element(prefix: "w", localName: "r", children: children)
        run.libraryUUID = UUID()
        let paraUUID = UUID()
        let wp = XmlNode.element(prefix: "w", localName: "p", children: [run])
        wp.libraryUUID = paraUUID
        let body = XmlNode.element(prefix: "w", localName: "body", children: [wp])
        let doc = XmlNode.element(prefix: "w", localName: "document", children: [body])
        return (XmlTree.synthesized(root: doc), ElementID(libraryUUID: paraUUID))
    }

    private func abTabCdRun() -> [XmlNode] {
        [
            XmlNode.element(prefix: "w", localName: "t", children: [XmlNode.text("AB")]),
            XmlNode.element(prefix: "w", localName: "tab"),
            XmlNode.element(prefix: "w", localName: "t", children: [XmlNode.text("CD")]),
        ]
    }

    /// The reviewer's exact repro: a run with "AB"<tab/>"CD" must not let
    /// "ABCD" match across the tab (which would previously produce
    /// duplicated "CD" text). Refused, not silently corrupted.
    func testMixedContentRunSpanningTabIsRefusedNotCorrupted() throws {
        let (base, paragraphID) = makeMixedContentRunParagraph(abTabCdRun())
        var log = OperationLog()
        log.append(
            .setRunFormatInRange(
                target: paragraphID, substring: "ABCD", instance: 1,
                format: RunFormatPayload(bold: true)),
            source: .swift)

        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: base)) { error in
            XCTAssertTrue(self.reason(of: error).contains("mixed content"), self.reason(of: error))
        }
    }

    func testMixedContentRunSubstringEntirelyBeforeTabIsRefused() throws {
        let (base, paragraphID) = makeMixedContentRunParagraph(abTabCdRun())
        var log = OperationLog()
        log.append(
            .setRunFormatInRange(target: paragraphID, substring: "AB", instance: 1,
                                  format: RunFormatPayload(bold: true)),
            source: .swift)
        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: base)) { error in
            XCTAssertTrue(self.reason(of: error).contains("mixed content"), self.reason(of: error))
        }
    }

    func testMixedContentRunSubstringEntirelyAfterTabIsRefused() throws {
        let (base, paragraphID) = makeMixedContentRunParagraph(abTabCdRun())
        var log = OperationLog()
        log.append(
            .setRunFormatInRange(target: paragraphID, substring: "CD", instance: 1,
                                  format: RunFormatPayload(bold: true)),
            source: .swift)
        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: base)) { error in
            XCTAssertTrue(self.reason(of: error).contains("mixed content"), self.reason(of: error))
        }
    }

    func testMixedContentRunWithBreakElementIsAlsoRefused() throws {
        let children = [
            XmlNode.element(prefix: "w", localName: "t", children: [XmlNode.text("AB")]),
            XmlNode.element(prefix: "w", localName: "br"),
            XmlNode.element(prefix: "w", localName: "t", children: [XmlNode.text("CD")]),
        ]
        let (base, paragraphID) = makeMixedContentRunParagraph(children)
        var log = OperationLog()
        log.append(
            .setRunFormatInRange(target: paragraphID, substring: "ABCD", instance: 1,
                                  format: RunFormatPayload(bold: true)),
            source: .swift)
        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: base)) { error in
            XCTAssertTrue(self.reason(of: error).contains("mixed content"), self.reason(of: error))
        }
    }

    /// A mixed-content run elsewhere in the paragraph must NOT block a
    /// match that falls entirely inside an ordinary, simple run.
    func testMatchInPlainRunSucceedsDespiteMixedContentRunElsewhere() throws {
        let beforeRun = XmlNode.element(
            prefix: "w", localName: "r",
            children: [XmlNode.element(prefix: "w", localName: "t", children: [XmlNode.text("before ")])])
        beforeRun.libraryUUID = UUID()
        let mixedRun = XmlNode.element(prefix: "w", localName: "r", children: abTabCdRun())
        mixedRun.libraryUUID = UUID()
        let afterRun = XmlNode.element(
            prefix: "w", localName: "r",
            children: [XmlNode.element(prefix: "w", localName: "t", children: [XmlNode.text(" after")])])
        afterRun.libraryUUID = UUID()

        let paraUUID = UUID()
        let wp = XmlNode.element(prefix: "w", localName: "p", children: [beforeRun, mixedRun, afterRun])
        wp.libraryUUID = paraUUID
        let body = XmlNode.element(prefix: "w", localName: "body", children: [wp])
        let doc = XmlNode.element(prefix: "w", localName: "document", children: [body])
        let base = XmlTree.synthesized(root: doc)
        let paragraphID = ElementID(libraryUUID: paraUUID)

        var log = OperationLog()
        log.append(
            .setRunFormatInRange(target: paragraphID, substring: "before", instance: 1,
                                  format: RunFormatPayload(bold: true)),
            source: .swift)
        let result = try OperationReducer.materialize(log: log, base: base)
        let runs = runsOf(result, paragraphID: paragraphID)
        XCTAssertTrue(runs.contains { self.text(of: $0) == "before" && self.isBold($0) },
                      "the 'before' group should still be matchable and bolded")

        var log2 = OperationLog()
        log2.append(
            .setRunFormatInRange(target: paragraphID, substring: "after", instance: 1,
                                  format: RunFormatPayload(bold: true)),
            source: .swift)
        let result2 = try OperationReducer.materialize(log: log2, base: base)
        let runs2 = runsOf(result2, paragraphID: paragraphID)
        XCTAssertTrue(runs2.contains { self.text(of: $0) == "after" && self.isBold($0) },
                      "the 'after' group (on the far side of the mixed-content run) should still be matchable")
    }

    // MARK: - HIGH-3: <w:ins> / <w:del> / <w:hyperlink> / <w:sdt> boundaries

    private func makeParagraph(children: [XmlNode]) -> (XmlTree, paragraphID: ElementID) {
        let paraUUID = UUID()
        let wp = XmlNode.element(prefix: "w", localName: "p", children: children)
        wp.libraryUUID = paraUUID
        let body = XmlNode.element(prefix: "w", localName: "body", children: [wp])
        let doc = XmlNode.element(prefix: "w", localName: "document", children: [body])
        return (XmlTree.synthesized(root: doc), ElementID(libraryUUID: paraUUID))
    }

    private func plainRun(_ text: String) -> XmlNode {
        let wt = XmlNode.element(prefix: "w", localName: "t", children: [XmlNode.text(text)])
        let r = XmlNode.element(prefix: "w", localName: "r", children: [wt])
        r.libraryUUID = UUID()
        return r
    }

    private func trackedInsertion(_ text: String) -> XmlNode {
        let ins = XmlNode.element(prefix: "w", localName: "ins", children: [plainRun(text)])
        ins.libraryUUID = UUID()
        return ins
    }

    private func trackedDeletion(_ text: String) -> XmlNode {
        let delText = XmlNode.element(prefix: "w", localName: "delText", children: [XmlNode.text(text)])
        let r = XmlNode.element(prefix: "w", localName: "r", children: [delText])
        let del = XmlNode.element(prefix: "w", localName: "del", children: [r])
        del.libraryUUID = UUID()
        return del
    }

    private func hyperlink(_ text: String) -> XmlNode {
        let hyperlink = XmlNode.element(
            prefix: "w", localName: "hyperlink",
            attributes: [XmlAttribute(prefix: "r", localName: "id", value: "rId1")],
            children: [plainRun(text)])
        hyperlink.libraryUUID = UUID()
        return hyperlink
    }

    private func sdt(_ text: String) -> XmlNode {
        let content = XmlNode.element(prefix: "w", localName: "sdtContent", children: [plainRun(text)])
        let sdt = XmlNode.element(prefix: "w", localName: "sdt", children: [content])
        sdt.libraryUUID = UUID()
        return sdt
    }

    func testMatchCrossingTrackedInsertionIsRefusedNamingIns() throws {
        let (base, paragraphID) = makeParagraph(children: [plainRun("foo"), trackedInsertion("XXX"), plainRun("bar")])
        var log = OperationLog()
        // "ooX" spans the tail of "foo" into the tracked insertion "XXX".
        log.append(
            .setRunFormatInRange(target: paragraphID, substring: "ooX", instance: 1,
                                  format: RunFormatPayload(bold: true)),
            source: .swift)
        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: base)) { error in
            XCTAssertTrue(self.reason(of: error).contains("<w:ins>"), self.reason(of: error))
        }
    }

    func testMatchEntirelyInsideTrackedInsertionIsRefusedNamingIns() throws {
        let (base, paragraphID) = makeParagraph(children: [plainRun("foo"), trackedInsertion("XXX"), plainRun("bar")])
        var log = OperationLog()
        log.append(
            .setRunFormatInRange(target: paragraphID, substring: "XXX", instance: 1,
                                  format: RunFormatPayload(bold: true)),
            source: .swift)
        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: base)) { error in
            XCTAssertTrue(self.reason(of: error).contains("<w:ins>"), self.reason(of: error))
        }
    }

    /// The false-adjacency probe from the review report: "foobar" is not
    /// contiguous visible text ("foo" + tracked-insertion "XXX" + "bar" is
    /// actually "fooXXXbar"), so it must be refused, not silently matched
    /// by gluing "foo" and "bar" together across the <w:ins>.
    func testFalseAdjacencyAcrossTrackedInsertionIsRefused() throws {
        let (base, paragraphID) = makeParagraph(children: [plainRun("foo"), trackedInsertion("XXX"), plainRun("bar")])
        var log = OperationLog()
        log.append(
            .setRunFormatInRange(target: paragraphID, substring: "foobar", instance: 1,
                                  format: RunFormatPayload(bold: true)),
            source: .swift)
        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: base))
    }

    func testMatchCrossingHyperlinkIsRefusedNamingHyperlink() throws {
        let (base, paragraphID) = makeParagraph(
            children: [plainRun("pre "), hyperlink("click here"), plainRun(" post")])
        var log = OperationLog()
        log.append(
            .setRunFormatInRange(target: paragraphID, substring: "pre click", instance: 1,
                                  format: RunFormatPayload(bold: true)),
            source: .swift)
        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: base)) { error in
            XCTAssertTrue(self.reason(of: error).contains("<w:hyperlink>"), self.reason(of: error))
        }
    }

    func testMatchEntirelyInsideHyperlinkIsRefusedNamingHyperlink() throws {
        let (base, paragraphID) = makeParagraph(children: [hyperlink("clickhere")])
        var log = OperationLog()
        log.append(
            .setRunFormatInRange(target: paragraphID, substring: "click", instance: 1,
                                  format: RunFormatPayload(bold: true)),
            source: .swift)
        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: base)) { error in
            XCTAssertTrue(self.reason(of: error).contains("<w:hyperlink>"), self.reason(of: error))
        }
    }

    func testMatchCrossingSDTIsRefusedNamingSDT() throws {
        let (base, paragraphID) = makeParagraph(children: [plainRun("foo"), sdt("XXX"), plainRun("bar")])
        var log = OperationLog()
        log.append(
            .setRunFormatInRange(target: paragraphID, substring: "fooX", instance: 1,
                                  format: RunFormatPayload(bold: true)),
            source: .swift)
        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: base)) { error in
            XCTAssertTrue(self.reason(of: error).contains("<w:sdt>"), self.reason(of: error))
        }
    }

    /// Deleted text must never appear in the flattened text used for
    /// matching, NOR in the diagnostic text used to build a specific
    /// "crosses <container>" reason — it is not visible content at all.
    func testDeletedTextNeverAppearsInFlattenedOrDiagnosticText() throws {
        let (base, paragraphID) = makeParagraph(children: [plainRun("foo"), trackedDeletion("XXX"), plainRun("bar")])

        var log = OperationLog()
        log.append(
            .setRunFormatInRange(target: paragraphID, substring: "XXX", instance: 1,
                                  format: RunFormatPayload(bold: true)),
            source: .swift)
        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: base)) { error in
            let r = self.reason(of: error)
            XCTAssertFalse(r.contains("<w:del>"), "deleted text must never be cited as a reachable container: \(r)")
        }

        // "foobar" must ALSO be refused — <w:del> is a boundary like <w:ins>,
        // not a silent no-op splice that makes "foo"/"bar" adjacent.
        var log2 = OperationLog()
        log2.append(
            .setRunFormatInRange(target: paragraphID, substring: "foobar", instance: 1,
                                  format: RunFormatPayload(bold: true)),
            source: .swift)
        XCTAssertThrowsError(try OperationReducer.materialize(log: log2, base: base))
    }

    /// A tracked-insertion elsewhere in the paragraph must not block a
    /// match entirely contained in an ordinary run before it.
    func testMatchInPlainRunSucceedsDespiteTrackedInsertionElsewhere() throws {
        let (base, paragraphID) = makeParagraph(children: [plainRun("foo"), trackedInsertion("XXX"), plainRun("bar")])
        var log = OperationLog()
        log.append(
            .setRunFormatInRange(target: paragraphID, substring: "foo", instance: 1,
                                  format: RunFormatPayload(bold: true)),
            source: .swift)
        let result = try OperationReducer.materialize(log: log, base: base)
        let runs = runsOf(result, paragraphID: paragraphID)
        XCTAssertTrue(runs.contains { self.text(of: $0) == "foo" && self.isBold($0) })
    }
}
