// WordEditLowerTests.swift
// EditAlgebra — addresses macdoc#110 item #6 (§7 of macdoc#105 tasks.md).
//
// Per-case `WordEdit.lower() -> [OOXMLEdit]` translation tests, satisfying
// the spec.md Requirement "WordEdit Enum with 3 Canonical Cases" lowering
// contract and design.md Decision 2.
//
// Implemented:
//   - applyBold(range:) — single-Run case
//   - applyLink(range:url:) — single-Run case (lowered OOXMLEdit is itself
//     stubbed pending §5 composite design)
//   - applyInsertParagraph(after:content:) — trivial 1:1 mapping
//
// Cases that lower() can't resolve without document context return [] and
// surface as `EditError.notImplemented` via the silent-noop guard in
// `WordDocument.apply`. See WordEdit.swift for the bounded-context
// constraint rationale.

import XCTest
@testable import OOXMLSwift

final class WordEditLowerTests: XCTestCase {

    // MARK: - applyBold(range:) single-Run case

    func testApplyBoldSingleRunLowersToSetBold() {
        // startRun == endRun → 1:1 mapping to OOXMLEdit.setBold(target: startRun)
        let runID = ElementID(libraryUUID: UUID())
        let range = WordRange(
            startRun: runID,
            startOffset: 0,
            endRun: runID,
            endOffset: 5
        )
        let edit = WordEdit.applyBold(range: range)

        let lowered = edit.lower()
        XCTAssertEqual(lowered.count, 1, "single-Run applyBold lowers to exactly 1 OOXMLEdit")

        guard case .setBold(let target, let value) = lowered[0] else {
            XCTFail("Expected OOXMLEdit.setBold, got \(lowered[0])")
            return
        }
        XCTAssertEqual(target, runID,
                       "lowered setBold target == range.startRun (single-Run case)")
        XCTAssertTrue(value, "lowered setBold value == true (applyBold always enables)")
    }

    func testApplyBoldSingleRunIgnoresOffsets() {
        // Offsets are ignored at this layer — setBold applies to entire Run.
        // Partial-Run bold (substring) would require run-splitting (separate
        // OOXMLEdit case, tracked for future design).
        let runID = ElementID(libraryUUID: UUID())
        let range1 = WordRange(startRun: runID, startOffset: 0, endRun: runID, endOffset: 100)
        let range2 = WordRange(startRun: runID, startOffset: 3, endRun: runID, endOffset: 7)

        let lowered1 = WordEdit.applyBold(range: range1).lower()
        let lowered2 = WordEdit.applyBold(range: range2).lower()

        XCTAssertEqual(lowered1, lowered2,
                       "Single-Run applyBold lowering ignores offsets (whole-Run mutation)")
    }

    // MARK: - applyBold(range:) multi-Run case (returns []) — surfaces via apply

    func testApplyBoldMultiRunReturnsEmptyLower() {
        // startRun != endRun → cross-Run case. lower() can't resolve
        // intermediate Runs without doc context, so returns []. The empty
        // list triggers the silent-noop guard in WordDocument.apply.
        let startRunID = ElementID(libraryUUID: UUID())
        let endRunID = ElementID(libraryUUID: UUID())
        let range = WordRange(startRun: startRunID, startOffset: 0, endRun: endRunID, endOffset: 5)
        let edit = WordEdit.applyBold(range: range)

        XCTAssertEqual(edit.lower(), [],
                       "Multi-Run applyBold.lower() returns [] (unsupported input combination)")
    }

    func testApplyBoldMultiRunSurfacesViaApply() {
        // doc.apply(multiRunApplyBold) should throw notImplemented (the
        // silent-noop guard in WordDocument.apply catches the empty lower()
        // and throws with a clear message).
        let doc = WordDocument()
        let startRunID = ElementID(libraryUUID: UUID())
        let endRunID = ElementID(libraryUUID: UUID())
        let range = WordRange(startRun: startRunID, startOffset: 0, endRun: endRunID, endOffset: 5)
        let edit = WordEdit.applyBold(range: range)

        XCTAssertThrowsError(try doc.apply(edit)) { error in
            guard case EditError.notImplemented(let message) = error else {
                XCTFail("Expected .notImplemented, got \(error)")
                return
            }
            XCTAssertTrue(message.contains("WordEdit"),
                          "Error message references WordEdit type: \(message)")
        }
    }

    // MARK: - applyLink(range:url:) single-Run case

    func testApplyLinkSingleRunLowersToInsertHyperlink() {
        let runID = ElementID(libraryUUID: UUID())
        let url = URL(string: "https://example.com/test")!
        let range = WordRange(startRun: runID, startOffset: 0, endRun: runID, endOffset: 5)
        let edit = WordEdit.applyLink(range: range, url: url)

        let lowered = edit.lower()
        XCTAssertEqual(lowered.count, 1, "single-Run applyLink lowers to exactly 1 OOXMLEdit")

        guard case .insertHyperlink(let target, let href, let displayText) = lowered[0] else {
            XCTFail("Expected OOXMLEdit.insertHyperlink, got \(lowered[0])")
            return
        }
        XCTAssertEqual(target, runID, "lowered insertHyperlink target == range.startRun")
        XCTAssertEqual(href, url, "lowered insertHyperlink href == applyLink url")
        XCTAssertNil(displayText,
                     "displayText == nil (lower() can't extract substring without doc context; §5 design will resolve nil → use href)")
    }

    func testApplyLinkLoweredOOXMLEditStillStubbed() {
        // The lowered OOXMLEdit.insertHyperlink is itself pending §5 design.
        // doc.apply(applyLink) should throw notImplemented at the
        // OOXMLEdit.operations() step, not at the WordEdit.lower() step.
        // This test pins the layering: WordEdit.lower() works; OOXMLEdit
        // path throws.
        let doc = WordDocument()
        let runID = ElementID(libraryUUID: UUID())
        let url = URL(string: "https://example.com")!
        let range = WordRange(startRun: runID, startOffset: 0, endRun: runID, endOffset: 5)
        let edit = WordEdit.applyLink(range: range, url: url)

        XCTAssertThrowsError(try doc.apply(edit)) { error in
            guard case EditError.notImplemented(let message) = error else {
                XCTFail("Expected .notImplemented from OOXMLEdit.insertHyperlink, got \(error)")
                return
            }
            // The error originates from OOXMLEdit.operations() for insertHyperlink,
            // not from WordEdit.lower(). Message should reference §5 (the
            // OOXMLEdit stub) not §7 (the WordEdit lower stub).
            XCTAssertTrue(message.contains("insertHyperlink") || message.contains("§5"),
                          "Error from OOXMLEdit layer (§5), not WordEdit layer (§7): \(message)")
        }
    }

    // MARK: - applyInsertParagraph(after:content:) — trivial 1:1

    func testApplyInsertParagraphLowersToInsertParagraph() {
        let paraID = ElementID(libraryUUID: UUID())
        let paraRef = ParagraphRef(paraID)
        let edit = WordEdit.applyInsertParagraph(after: paraRef, content: "Hello world")

        let lowered = edit.lower()
        XCTAssertEqual(lowered.count, 1)

        guard case .insertParagraph(let after, let content, let styleId) = lowered[0] else {
            XCTFail("Expected OOXMLEdit.insertParagraph, got \(lowered[0])")
            return
        }
        XCTAssertEqual(after, paraID, "after: ParagraphRef → ElementID round-trip")
        XCTAssertEqual(content, "Hello world")
        XCTAssertNil(styleId, "WordEdit.applyInsertParagraph doesn't take styleId; lowered styleId == nil")
    }

    func testApplyInsertParagraphEmptyContent() {
        // Pin: empty content is legal (creates empty paragraph)
        let paraRef = ParagraphRef(ElementID(libraryUUID: UUID()))
        let edit = WordEdit.applyInsertParagraph(after: paraRef, content: "")

        let lowered = edit.lower()
        guard case .insertParagraph(_, let content, _) = lowered[0] else {
            XCTFail("Expected OOXMLEdit.insertParagraph")
            return
        }
        XCTAssertEqual(content, "", "Empty content propagates through lower()")
    }

    // MARK: - End-to-end: WordEdit chains through to xmlTrees mutation

    func testApplyInsertParagraphEndToEndMutatesDocPart() throws {
        // The full path: WordEdit.applyInsertParagraph → lower() →
        // OOXMLEdit.insertParagraph → operations() → log append → materialize
        // → xmlTrees mutation. This is the first proof that WordEdit is
        // production-ready for at least one case.
        let paraUUID = UUID()
        let textNode = XmlNode.text("first")
        let wt = XmlNode.element(prefix: "w", localName: "t", children: [textNode])
        let wr = XmlNode.element(prefix: "w", localName: "r", children: [wt])
        let wp = XmlNode.element(prefix: "w", localName: "p", children: [wr])
        wp.libraryUUID = paraUUID
        let body = XmlNode.element(prefix: "w", localName: "body", children: [wp])
        let root = XmlNode.element(prefix: "w", localName: "document", children: [body])

        var doc = WordDocument()
        doc.xmlTrees["word/document.xml"] = XmlTree.synthesized(root: root)

        let paraRef = ParagraphRef(ElementID(libraryUUID: paraUUID))
        let edit = WordEdit.applyInsertParagraph(after: paraRef, content: "second")
        let result = try doc.apply(edit)

        // Walk result's body and verify second paragraph appears
        var texts: [String] = []
        func walk(_ node: XmlNode) {
            if node.kind == .element && node.localName == "t" {
                for child in node.children where child.kind == .text {
                    texts.append(child.textContent)
                }
            }
            for child in node.children { walk(child) }
        }
        walk(result.xmlTrees["word/document.xml"]!.root)
        XCTAssertEqual(texts, ["first", "second"],
                       "WordEdit.applyInsertParagraph end-to-end produces new paragraph in doc")
    }

    func testApplyBoldSingleRunEndToEndMutatesRPr() throws {
        // applyBold single-Run case lowers to setBold which has a working
        // reducer. End-to-end proves WordEdit ergonomics work for this case.
        let runUUID = UUID()
        let textNode = XmlNode.text("hello")
        let wt = XmlNode.element(prefix: "w", localName: "t", children: [textNode])
        let wr = XmlNode.element(prefix: "w", localName: "r", children: [wt])
        wr.libraryUUID = runUUID
        let wp = XmlNode.element(prefix: "w", localName: "p", children: [wr])
        let body = XmlNode.element(prefix: "w", localName: "body", children: [wp])
        let root = XmlNode.element(prefix: "w", localName: "document", children: [body])

        var doc = WordDocument()
        doc.xmlTrees["word/document.xml"] = XmlTree.synthesized(root: root)

        let runID = ElementID(libraryUUID: runUUID)
        let range = WordRange(startRun: runID, startOffset: 0, endRun: runID, endOffset: 5)
        let edit = WordEdit.applyBold(range: range)
        let result = try doc.apply(edit)

        // Verify <w:b/> appeared in target Run's rPr
        guard let tree = result.xmlTrees["word/document.xml"],
              let run = OperationReducer.findNode(elementID: runID, in: tree),
              let rPr = run.children.first(where: { $0.kind == .element && $0.localName == "rPr" }) else {
            XCTFail("Expected run with rPr after applyBold")
            return
        }
        let hasBold = rPr.children.contains { $0.kind == .element && $0.localName == "b" }
        XCTAssertTrue(hasBold, "applyBold end-to-end produces <w:b/> in target Run's rPr")
    }
}
