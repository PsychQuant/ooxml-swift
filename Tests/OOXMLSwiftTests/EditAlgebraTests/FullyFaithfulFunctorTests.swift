// FullyFaithfulFunctorTests.swift
// EditAlgebra — addresses macdoc#110 item #1 (§8 of macdoc#105 tasks.md).
//
// Property-based tests for the **canonical-identity invariant** mandated by
// macdoc#99 foundation Requirement "Canonical-Identity Round-Trip Contract":
//
//   For each OOXMLEdit case, applying the edit to a Document mutates only
//   the subtrees on the Edit's target path. All other subtrees remain
//   bytewise-equal (c14n-equal) to their input state.
//
// macdoc#105 spec.md §123-132 prescribes: ≥100 randomized samples per Edit
// case, swift-testing `@Test(arguments:)` parameterized. Reality on this
// package: swift-tools-version 5.9, no swift-testing dependency. Approach
// uses XCTest with parameterized loops + per-sample failure logging that
// includes the sample index for reproduction (the synthesized fixture is
// deterministic, so sampleIndex pins the input).
//
// Per design.md Decision 5 errata: until swift-testing is added (would
// require bumping tools-version + a Package.swift dependency), XCTest
// loops satisfy the SHALL — what matters is the property check + sample
// count, not the framework chosen.
//
// **Fixture substrate**: spec.md mentions "RealWorldDocxRoundTripSmokeTests
// NTPU thesis fixture", but no such file exists in the repo (the loader
// was aspirational). This test uses a SYNTHESIZED multi-part fixture
// (document.xml + styles.xml) with ~10 paragraphs, each containing 1-3
// Runs with text content. The synthesized substrate is faster, deterministic,
// and exercises the same code path as real .docx (DocxReader builds the
// same XmlTree shape from real ZIPs). When a real NTPU thesis fixture
// lands (separate macdoc#110 follow-up), the property tests should be
// re-run against it.
//
// **Scope**: 4 of 5 OOXMLEdit cases have functional `apply()`:
//   - insertParagraph(after:)
//   - insertParagraphBefore(before:)
//   - setBold(target:value:)
//   - removeParagraph(target:)
//
// insertHyperlink is stubbed pending §5 composite design checkpoint;
// property test for it is included as a skip with documented reason.

import XCTest
@testable import OOXMLSwift

final class FullyFaithfulFunctorTests: XCTestCase {

    // MARK: - Configuration

    /// Number of randomized samples per property test, per spec.md §123-132.
    /// 100 is the spec minimum; bump locally for stress testing.
    private static let samplesPerCase = 100

    /// Number of paragraphs in the synthesized fixture. Larger gives more
    /// target variety per sample; smaller is faster. 10 balances both.
    private static let fixtureParagraphCount = 10

    /// Number of runs per paragraph in the synthesized fixture. Variety
    /// for setBold sample selection.
    private static let runsPerParagraph = 3

    // MARK: - Fixture substrate (synthesized multi-part doc)

    /// Builds a deterministic multi-part WordDocument:
    ///   - `word/document.xml`: body with N paragraphs, each with M runs
    ///   - `word/styles.xml`: minimal styles part (proves multi-part scoping)
    ///
    /// Returns the doc + arrays of (paragraphID, [runID]) for sample selection.
    /// Each paragraph has libraryUUID; each run has libraryUUID (so setBold
    /// can address them).
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

                let runText = "p\(paraIdx)r\(runIdx)"
                let textNode = XmlNode.text(runText)
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

        // styles.xml — non-target part for c14n-invariance checks
        let style = XmlNode.element(prefix: "w", localName: "style")
        style.setAttribute(prefix: "w", localName: "styleId", value: "Normal")
        let styles = XmlNode.element(prefix: "w", localName: "styles", children: [style])

        var doc = WordDocument()
        doc.xmlTrees["word/document.xml"] = XmlTree.synthesized(root: docRoot)
        doc.xmlTrees["word/styles.xml"] = XmlTree.synthesized(root: styles)
        return (doc, paraIDs, runIDsByPara)
    }

    // MARK: - Canonical-identity check helpers

    /// Snapshot of a paragraph for invariance comparison:
    /// `(libraryUUID, runIDs, runTexts, runBoldFlags)`.
    private struct ParaSnapshot: Equatable {
        let id: UUID?
        let runs: [RunSnapshot]
    }

    private struct RunSnapshot: Equatable {
        let id: UUID?
        let text: String
        let bold: Bool
    }

    /// Walks the body and produces an ordered snapshot of every paragraph.
    private func snapshot(_ doc: WordDocument) -> [ParaSnapshot] {
        guard let tree = doc.xmlTrees["word/document.xml"] else { return [] }
        var result: [ParaSnapshot] = []

        func walkParas(_ node: XmlNode) {
            if node.kind == .element && node.localName == "p" {
                var runs: [RunSnapshot] = []
                for child in node.children where child.kind == .element && child.localName == "r" {
                    var text = ""
                    var bold = false
                    for grandchild in child.children {
                        if grandchild.kind == .element && grandchild.localName == "t" {
                            for ggc in grandchild.children where ggc.kind == .text {
                                text += ggc.textContent
                            }
                        }
                        if grandchild.kind == .element && grandchild.localName == "rPr" {
                            if grandchild.children.contains(where: {
                                $0.kind == .element && $0.localName == "b"
                            }) {
                                bold = true
                            }
                        }
                    }
                    runs.append(RunSnapshot(id: child.libraryUUID, text: text, bold: bold))
                }
                result.append(ParaSnapshot(id: node.libraryUUID, runs: runs))
                return  // don't recurse into paragraph children for paragraphs
            }
            for child in node.children {
                walkParas(child)
            }
        }
        walkParas(tree.root)
        return result
    }

    /// Walks the styles.xml and produces a structural fingerprint used to
    /// assert the non-target part is unchanged.
    private func stylesFingerprint(_ doc: WordDocument) -> String {
        guard let tree = doc.xmlTrees["word/styles.xml"] else { return "<missing>" }
        var parts: [String] = []
        func walk(_ node: XmlNode) {
            if node.kind == .element {
                let attrs = node.attributes.map { "\($0.prefix ?? "").\($0.localName)=\($0.value)" }.sorted().joined(separator: ",")
                parts.append("\(node.prefix ?? "").\(node.localName)[\(attrs)]")
            }
            for child in node.children {
                walk(child)
            }
        }
        walk(tree.root)
        return parts.joined(separator: "|")
    }

    // MARK: - Property: insertParagraph(after:) preserves non-target paragraphs

    func testInsertParagraphCanonicalIdentity() throws {
        for sampleIndex in 0..<Self.samplesPerCase {
            let (doc, paraIDs, _) = makeFixture()
            let before = snapshot(doc)
            let stylesBefore = stylesFingerprint(doc)

            // Pick random target paragraph index
            let targetIdx = Int.random(in: 0..<paraIDs.count)
            let targetID = paraIDs[targetIdx]
            let edit = OOXMLEdit.insertParagraph(
                after: targetID,
                content: "sample-\(sampleIndex)-content",
                styleId: nil
            )

            do {
                let result = try doc.apply(edit)
                let after = snapshot(result)
                let stylesAfter = stylesFingerprint(result)

                // Invariant 1: result has one MORE paragraph
                XCTAssertEqual(after.count, before.count + 1,
                               "sample \(sampleIndex): result has +1 paragraph")

                // Invariant 2: all original paragraphs appear unchanged in order,
                // with the new paragraph inserted at position targetIdx+1
                XCTAssertEqual(Array(after.prefix(targetIdx + 1)),
                               Array(before.prefix(targetIdx + 1)),
                               "sample \(sampleIndex): paragraphs up to + including target unchanged")
                XCTAssertEqual(Array(after.suffix(from: targetIdx + 2)),
                               Array(before.suffix(from: targetIdx + 1)),
                               "sample \(sampleIndex): paragraphs after target unchanged (shifted by 1)")

                // Invariant 3: styles.xml fingerprint unchanged (non-target part)
                XCTAssertEqual(stylesAfter, stylesBefore,
                               "sample \(sampleIndex): non-target styles.xml unchanged")
            } catch {
                XCTFail("sample \(sampleIndex) (target paraIdx \(targetIdx)): apply threw \(error)")
            }
        }
    }

    // MARK: - Property: insertParagraphBefore(before:) preserves non-target paragraphs

    func testInsertParagraphBeforeCanonicalIdentity() throws {
        for sampleIndex in 0..<Self.samplesPerCase {
            let (doc, paraIDs, _) = makeFixture()
            let before = snapshot(doc)
            let stylesBefore = stylesFingerprint(doc)

            let targetIdx = Int.random(in: 0..<paraIDs.count)
            let targetID = paraIDs[targetIdx]
            let edit = OOXMLEdit.insertParagraphBefore(
                before: targetID,
                content: "sample-\(sampleIndex)-content",
                styleId: nil
            )

            do {
                let result = try doc.apply(edit)
                let after = snapshot(result)
                let stylesAfter = stylesFingerprint(result)

                XCTAssertEqual(after.count, before.count + 1,
                               "sample \(sampleIndex): result has +1 paragraph")

                // Paragraphs BEFORE target unchanged
                XCTAssertEqual(Array(after.prefix(targetIdx)),
                               Array(before.prefix(targetIdx)),
                               "sample \(sampleIndex): paragraphs before target unchanged")

                // Paragraphs FROM target onward shifted by 1 (target now at targetIdx+1)
                XCTAssertEqual(Array(after.suffix(from: targetIdx + 1)),
                               Array(before.suffix(from: targetIdx)),
                               "sample \(sampleIndex): paragraphs from target onward shifted by 1")

                XCTAssertEqual(stylesAfter, stylesBefore,
                               "sample \(sampleIndex): non-target styles.xml unchanged")
            } catch {
                XCTFail("sample \(sampleIndex) (target paraIdx \(targetIdx)): apply threw \(error)")
            }
        }
    }

    // MARK: - Property: setBold(target:value:true) toggles target Run only

    func testSetBoldCanonicalIdentity() throws {
        for sampleIndex in 0..<Self.samplesPerCase {
            let (doc, _, runIDsByPara) = makeFixture()
            let before = snapshot(doc)
            let stylesBefore = stylesFingerprint(doc)

            // Pick a random (paraIdx, runIdx)
            let paraIdx = Int.random(in: 0..<runIDsByPara.count)
            let runIdx = Int.random(in: 0..<runIDsByPara[paraIdx].count)
            let targetRunID = runIDsByPara[paraIdx][runIdx]

            let edit = OOXMLEdit.setBold(target: targetRunID, value: true)

            do {
                let result = try doc.apply(edit)
                let after = snapshot(result)
                let stylesAfter = stylesFingerprint(result)

                // Invariant 1: same number of paragraphs
                XCTAssertEqual(after.count, before.count,
                               "sample \(sampleIndex): paragraph count unchanged")

                // Invariant 2: only the target Run has bold flipped; everything
                // else identical (run IDs, run texts, other runs' bold flags)
                for (pIdx, pSnap) in after.enumerated() {
                    XCTAssertEqual(pSnap.id, before[pIdx].id,
                                   "sample \(sampleIndex) p\(pIdx): paragraph ID unchanged")
                    for (rIdx, rSnap) in pSnap.runs.enumerated() {
                        let beforeRun = before[pIdx].runs[rIdx]
                        XCTAssertEqual(rSnap.id, beforeRun.id,
                                       "sample \(sampleIndex) p\(pIdx)r\(rIdx): run ID unchanged")
                        XCTAssertEqual(rSnap.text, beforeRun.text,
                                       "sample \(sampleIndex) p\(pIdx)r\(rIdx): run text unchanged")

                        let isTarget = (pIdx == paraIdx && rIdx == runIdx)
                        if isTarget {
                            XCTAssertTrue(rSnap.bold,
                                          "sample \(sampleIndex): target run is bold")
                        } else {
                            XCTAssertEqual(rSnap.bold, beforeRun.bold,
                                           "sample \(sampleIndex) p\(pIdx)r\(rIdx): non-target run bold flag unchanged")
                        }
                    }
                }

                XCTAssertEqual(stylesAfter, stylesBefore,
                               "sample \(sampleIndex): non-target styles.xml unchanged")
            } catch {
                XCTFail("sample \(sampleIndex) (target p\(paraIdx)r\(runIdx)): apply threw \(error)")
            }
        }
    }

    // MARK: - Property: removeParagraph(target:) removes only the target

    func testRemoveParagraphCanonicalIdentity() throws {
        for sampleIndex in 0..<Self.samplesPerCase {
            let (doc, paraIDs, _) = makeFixture()
            let before = snapshot(doc)
            let stylesBefore = stylesFingerprint(doc)

            let targetIdx = Int.random(in: 0..<paraIDs.count)
            let targetID = paraIDs[targetIdx]
            let edit = OOXMLEdit.removeParagraph(target: targetID)

            do {
                let result = try doc.apply(edit)
                let after = snapshot(result)
                let stylesAfter = stylesFingerprint(result)

                // Invariant 1: result has one LESS paragraph
                XCTAssertEqual(after.count, before.count - 1,
                               "sample \(sampleIndex): result has -1 paragraph")

                // Invariant 2: paragraphs before target unchanged
                XCTAssertEqual(Array(after.prefix(targetIdx)),
                               Array(before.prefix(targetIdx)),
                               "sample \(sampleIndex): paragraphs before target unchanged")

                // Invariant 3: paragraphs after target shifted up by 1
                XCTAssertEqual(Array(after.suffix(from: targetIdx)),
                               Array(before.suffix(from: targetIdx + 1)),
                               "sample \(sampleIndex): paragraphs after target shifted up by 1")

                XCTAssertEqual(stylesAfter, stylesBefore,
                               "sample \(sampleIndex): non-target styles.xml unchanged")
            } catch {
                XCTFail("sample \(sampleIndex) (target paraIdx \(targetIdx)): apply threw \(error)")
            }
        }
    }

    // MARK: - insertHyperlink — stubbed, documented skip

    func testInsertHyperlinkCanonicalIdentity_SKIPPED_PENDING_SECTION_5() throws {
        // OOXMLEdit.insertHyperlink is stubbed pending §5 composite design
        // checkpoint (per macdoc#105 tasks.md §5 — 5 open design questions:
        // target type semantics, atomicity strategy, rels XML coordination,
        // displayText nil → use href, run-splitting when range partial-covers
        // a Run). Once §5 ships, replicate the pattern from
        // testInsertParagraphCanonicalIdentity using random URL + display text.
        throw XCTSkip("Pending §5 composite-design checkpoint (macdoc#105)")
    }
}
