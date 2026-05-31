// NaturalityTests.swift
// EditAlgebra — addresses macdoc#110 item #2 (§9 of macdoc#105 tasks.md).
//
// Naturality property tests for the WordEdit → OOXMLEdit lowering functor.
//
// The naturality invariant (per macdoc#99 foundation ADR-002 + macdoc#105
// spec.md §123-132): for any pair of composable WordEdits (a, b),
//
//   (WordEdit.a ∘ WordEdit.b).lower() == WordEdit.a.lower() ∘ WordEdit.b.lower()
//
// Since WordEdit doesn't have a built-in ∘ operator, the practical
// formulation is:
//
//   doc.apply(a).apply(b)                                 // Path A
//     ≡ (modulo Equatable)
//   doc.apply([a.lower() + b.lower()] as [any Edit])      // Path B
//
// Equality is via WordDocument's custom Equatable, which excludes
// operationLog from content equality (per design.md Decision 3). So:
// - Path A appends 2 log entries (one per WordEdit.apply call)
// - Path B appends N log entries (one per OOXMLEdit from a.lower() + b.lower())
// - Log counts differ; content (xmlTrees) should match.
//
// Composable pair-types per spec.md §130:
//   1. applyBold       ∘ applyLink            — applyLink → throws at OOXMLEdit layer (§5 stub)
//   2. applyBold       ∘ applyInsertParagraph — both paths functional
//   3. applyLink       ∘ applyInsertParagraph — applyLink → throws
//
// For pairs involving applyLink, naturality holds in the failure mode:
// both Path A and Path B should throw `EditError.notImplemented` (the
// throw originates from OOXMLEdit.insertHyperlink.operations() — same
// site for both paths, so error parity is the naturality assertion).
//
// **Implementation deviation from spec.md Decision 5**: uses XCTest loops
// instead of swift-testing `@Test(arguments:)`. Same rationale as
// FullyFaithfulFunctorTests — package is swift-tools-version 5.9, no
// swift-testing dependency.

import XCTest
@testable import OOXMLSwift

final class NaturalityTests: XCTestCase {

    // MARK: - Configuration

    /// Samples per naturality pair, per spec.md §130 minimum (50).
    private static let samplesPerPair = 50

    /// Number of paragraphs in the synthesized fixture.
    private static let fixtureParagraphCount = 10

    /// Number of runs per paragraph in the synthesized fixture.
    private static let runsPerParagraph = 3

    // MARK: - Fixture substrate (shared with FullyFaithfulFunctorTests)

    /// Builds a deterministic multi-part WordDocument with N paragraphs ×
    /// M runs. Returns the doc + (paragraphID, [runID]) arrays for random
    /// sample selection.
    private func makeFixture(
        paragraphs: Int = fixtureParagraphCount,
        runsPerPara: Int = runsPerParagraph
    ) -> (WordDocument, [ElementID], [[ElementID]]) {
        var paraIDs: [ElementID] = []
        var runIDsByPara: [[ElementID]] = []
        var paragraphNodes: [XmlNode] = []

        for paraIdx in 0..<paragraphs {
            let paraUUID = UUID()
            paraIDs.append(ElementID(libraryUUID: paraUUID))

            var runIDsForThisPara: [ElementID] = []
            var runNodes: [XmlNode] = []
            for runIdx in 0..<runsPerPara {
                let runUUID = UUID()
                runIDsForThisPara.append(ElementID(libraryUUID: runUUID))

                let textNode = XmlNode.text("p\(paraIdx)r\(runIdx)")
                let wt = XmlNode.element(prefix: "w", localName: "t", children: [textNode])
                let wr = XmlNode.element(prefix: "w", localName: "r", children: [wt])
                wr.libraryUUID = runUUID
                runNodes.append(wr)
            }
            runIDsByPara.append(runIDsForThisPara)

            let wp = XmlNode.element(prefix: "w", localName: "p", children: runNodes)
            wp.libraryUUID = paraUUID
            paragraphNodes.append(wp)
        }

        let body = XmlNode.element(prefix: "w", localName: "body", children: paragraphNodes)
        let docRoot = XmlNode.element(prefix: "w", localName: "document", children: [body])
        let style = XmlNode.element(prefix: "w", localName: "style")
        style.setAttribute(prefix: "w", localName: "styleId", value: "Normal")
        let styles = XmlNode.element(prefix: "w", localName: "styles", children: [style])

        var doc = WordDocument()
        doc.xmlTrees["word/document.xml"] = XmlTree.synthesized(root: docRoot)
        doc.xmlTrees["word/styles.xml"] = XmlTree.synthesized(root: styles)
        return (doc, paraIDs, runIDsByPara)
    }

    // MARK: - Naturality assertion helper

    /// Applies a pair (a, b) of WordEdits via both paths and asserts naturality.
    /// Returns nothing; uses XCTAssert internally. `sampleContext` is logged
    /// on failure for reproducibility.
    private func assertNaturality(
        a: WordEdit,
        b: WordEdit,
        doc: WordDocument,
        sampleContext: String
    ) {
        // Path A: apply WordEdits in sequence
        var resultA: WordDocument?
        var errorA: Error?
        do {
            resultA = try doc.apply(a).apply(b)
        } catch {
            errorA = error
        }

        // Path B: lower both, concatenate, apply at OOXMLEdit level
        let lowered: [any Edit] = a.lower().map { $0 as any Edit }
            + b.lower().map { $0 as any Edit }
        var resultB: WordDocument?
        var errorB: Error?
        do {
            resultB = try doc.apply(lowered)
        } catch {
            errorB = error
        }

        // Naturality of throw: both succeed or both fail
        let aSucceeded = errorA == nil
        let bSucceeded = errorB == nil
        XCTAssertEqual(aSucceeded, bSucceeded,
                       "\(sampleContext): naturality of throw — Path A errorA=\(String(describing: errorA)), Path B errorB=\(String(describing: errorB))")

        // If both succeeded, content equality (WordDocument Equatable excludes log)
        if let rA = resultA, let rB = resultB {
            XCTAssertEqual(rA, rB,
                           "\(sampleContext): naturality of content — Path A xmlTrees ≠ Path B xmlTrees")
        }

        // If both failed, both should be EditError.notImplemented (the
        // observed failure mode for applyLink-involved pairs in the
        // current §5-stubbed state).
        if let eA = errorA, let eB = errorB {
            if case EditError.notImplemented = eA, case EditError.notImplemented = eB {
                // Naturality of error case holds — both throw same error type
            } else {
                XCTFail("\(sampleContext): naturality of error type — Path A=\(eA), Path B=\(eB)")
            }
        }
    }

    // MARK: - Pair 1: applyBold ∘ applyLink (BOTH PATHS THROW — naturality of error)

    func testNaturality_applyBold_applyLink() {
        for sampleIndex in 0..<Self.samplesPerPair {
            let (doc, _, runIDsByPara) = makeFixture()

            // Pick two random Runs (single-Run case for both edits)
            let paraIdx1 = Int.random(in: 0..<runIDsByPara.count)
            let runIdx1 = Int.random(in: 0..<runIDsByPara[paraIdx1].count)
            let runID1 = runIDsByPara[paraIdx1][runIdx1]

            let paraIdx2 = Int.random(in: 0..<runIDsByPara.count)
            let runIdx2 = Int.random(in: 0..<runIDsByPara[paraIdx2].count)
            let runID2 = runIDsByPara[paraIdx2][runIdx2]

            let range1 = WordRange(startRun: runID1, startOffset: 0, endRun: runID1, endOffset: 5)
            let range2 = WordRange(startRun: runID2, startOffset: 0, endRun: runID2, endOffset: 5)
            let url = URL(string: "https://example.com/test-\(sampleIndex)")!

            let editA = WordEdit.applyBold(range: range1)
            let editB = WordEdit.applyLink(range: range2, url: url)

            assertNaturality(
                a: editA, b: editB, doc: doc,
                sampleContext: "applyBold(p\(paraIdx1)r\(runIdx1)) ∘ applyLink(p\(paraIdx2)r\(runIdx2), sample \(sampleIndex))"
            )
        }
    }

    // MARK: - Pair 2: applyBold ∘ applyInsertParagraph (BOTH PATHS FUNCTIONAL)

    func testNaturality_applyBold_applyInsertParagraph() {
        for sampleIndex in 0..<Self.samplesPerPair {
            let (doc, paraIDs, runIDsByPara) = makeFixture()

            // Random Run for applyBold (single-Run case)
            let boldParaIdx = Int.random(in: 0..<runIDsByPara.count)
            let boldRunIdx = Int.random(in: 0..<runIDsByPara[boldParaIdx].count)
            let runID = runIDsByPara[boldParaIdx][boldRunIdx]
            let range = WordRange(startRun: runID, startOffset: 0, endRun: runID, endOffset: 5)

            // Random paragraph for applyInsertParagraph
            let targetParaIdx = Int.random(in: 0..<paraIDs.count)
            let paraRef = ParagraphRef(paraIDs[targetParaIdx])

            let editA = WordEdit.applyBold(range: range)
            let editB = WordEdit.applyInsertParagraph(after: paraRef, content: "naturality-\(sampleIndex)")

            assertNaturality(
                a: editA, b: editB, doc: doc,
                sampleContext: "applyBold(p\(boldParaIdx)r\(boldRunIdx)) ∘ applyInsertParagraph(after p\(targetParaIdx), sample \(sampleIndex))"
            )
        }
    }

    // MARK: - Pair 3: applyLink ∘ applyInsertParagraph (Path A throws on applyLink)

    func testNaturality_applyLink_applyInsertParagraph() {
        for sampleIndex in 0..<Self.samplesPerPair {
            let (doc, paraIDs, runIDsByPara) = makeFixture()

            // Random Run for applyLink
            let linkParaIdx = Int.random(in: 0..<runIDsByPara.count)
            let linkRunIdx = Int.random(in: 0..<runIDsByPara[linkParaIdx].count)
            let runID = runIDsByPara[linkParaIdx][linkRunIdx]
            let range = WordRange(startRun: runID, startOffset: 0, endRun: runID, endOffset: 5)
            let url = URL(string: "https://example.com/test-\(sampleIndex)")!

            // Random paragraph for applyInsertParagraph
            let targetParaIdx = Int.random(in: 0..<paraIDs.count)
            let paraRef = ParagraphRef(paraIDs[targetParaIdx])

            let editA = WordEdit.applyLink(range: range, url: url)
            let editB = WordEdit.applyInsertParagraph(after: paraRef, content: "naturality-\(sampleIndex)")

            assertNaturality(
                a: editA, b: editB, doc: doc,
                sampleContext: "applyLink(p\(linkParaIdx)r\(linkRunIdx)) ∘ applyInsertParagraph(after p\(targetParaIdx), sample \(sampleIndex))"
            )
        }
    }

    // MARK: - Reversed order checks (a∘b vs b∘a) — sanity for order-dependence

    func testNaturality_applyInsertParagraph_applyBold_reverseOrder() {
        // Same as Pair 2 but with arguments swapped. Naturality still holds:
        // path A and path B should both succeed and produce equal docs (even
        // though the result is DIFFERENT from applyBold ∘ applyInsertParagraph
        // — order matters for the semantic outcome, but naturality only
        // claims path A == path B for a GIVEN order).
        for sampleIndex in 0..<Self.samplesPerPair {
            let (doc, paraIDs, runIDsByPara) = makeFixture()

            let boldParaIdx = Int.random(in: 0..<runIDsByPara.count)
            let boldRunIdx = Int.random(in: 0..<runIDsByPara[boldParaIdx].count)
            let runID = runIDsByPara[boldParaIdx][boldRunIdx]
            let range = WordRange(startRun: runID, startOffset: 0, endRun: runID, endOffset: 5)

            let targetParaIdx = Int.random(in: 0..<paraIDs.count)
            let paraRef = ParagraphRef(paraIDs[targetParaIdx])

            let editA = WordEdit.applyInsertParagraph(after: paraRef, content: "rev-\(sampleIndex)")
            let editB = WordEdit.applyBold(range: range)

            assertNaturality(
                a: editA, b: editB, doc: doc,
                sampleContext: "applyInsertParagraph(after p\(targetParaIdx)) ∘ applyBold(p\(boldParaIdx)r\(boldRunIdx)), sample \(sampleIndex)"
            )
        }
    }
}
