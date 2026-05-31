// InsertParagraphE2ETests.swift
// EditAlgebra — Phase 2 of ooxml-edit-isomorphism-foundation (PsychQuant/macdoc#99)
//
// §3.3 of #105 tasks — END-TO-END test for OOXMLEdit.insertParagraph.
//
// Now unblocked by ooxml-swift#71 (Phase 2c reducer): the full chain
//   Edit → lower() → operations() → log append → materialize → new WordDocument
// can mutate xmlTrees and we can inspect the result.
//
// Scope: synthesized single-part WordDocument (xmlTrees["word/document.xml"]
// only — no styles.xml, comments.xml, etc.). Real-.docx fixture tests come
// after the multi-part scoping fix is in place — see follow-up note in
// WordDocument+Apply.swift.

import XCTest
@testable import OOXMLSwift

final class InsertParagraphE2ETests: XCTestCase {

    // MARK: - Helpers

    /// Builds a WordDocument whose `xmlTrees["word/document.xml"]` contains
    /// the given paragraphs. Returns the doc + each paragraph's ElementID.
    private func makeSinglePartDoc(paragraphs texts: [String]) -> (WordDocument, [ElementID]) {
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

        var wordDoc = WordDocument()
        wordDoc.xmlTrees["word/document.xml"] = XmlTree.synthesized(root: doc)
        return (wordDoc, paraIDs)
    }

    /// Extracts text of all <w:t> in the document.xml part, in order.
    private func extractTextFromDocPart(_ doc: WordDocument) -> [String] {
        guard let tree = doc.xmlTrees["word/document.xml"] else { return [] }
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

    // MARK: - End-to-end: OOXMLEdit.insertParagraph mutates xmlTrees

    func testInsertParagraphAfterMutatesDocumentPart() throws {
        let (doc, paraIDs) = makeSinglePartDoc(paragraphs: ["first"])
        XCTAssertEqual(extractTextFromDocPart(doc), ["first"], "baseline")

        let edit = OOXMLEdit.insertParagraph(after: paraIDs[0], content: "second", styleId: nil)
        let result = try doc.apply(edit)

        XCTAssertEqual(extractTextFromDocPart(result), ["first", "second"],
                       "After apply, document.xml has new paragraph after target")
    }

    func testInsertParagraphBeforeMutatesDocumentPart() throws {
        let (doc, paraIDs) = makeSinglePartDoc(paragraphs: ["second"])

        let edit = OOXMLEdit.insertParagraphBefore(before: paraIDs[0], content: "first", styleId: nil)
        let result = try doc.apply(edit)

        XCTAssertEqual(extractTextFromDocPart(result), ["first", "second"],
                       "After apply, new paragraph appears BEFORE target")
    }

    func testRemoveParagraphMutatesDocumentPart() throws {
        let (doc, paraIDs) = makeSinglePartDoc(paragraphs: ["a", "b", "c"])

        let edit = OOXMLEdit.removeParagraph(target: paraIDs[1])
        let result = try doc.apply(edit)

        XCTAssertEqual(extractTextFromDocPart(result), ["a", "c"],
                       "After apply, target paragraph removed; siblings preserved")
    }

    // MARK: - End-to-end: WordDocument immutability + operationLog accumulation

    func testApplyDoesNotMutateSelfOnSuccess() throws {
        let (doc, paraIDs) = makeSinglePartDoc(paragraphs: ["first"])

        let edit = OOXMLEdit.insertParagraph(after: paraIDs[0], content: "second", styleId: nil)
        _ = try doc.apply(edit)

        XCTAssertEqual(extractTextFromDocPart(doc), ["first"],
                       "Original doc unchanged after apply (value semantics)")
    }

    func testApplyAppendsToOperationLog() throws {
        let (doc, paraIDs) = makeSinglePartDoc(paragraphs: ["first"])
        let logCountBefore = doc.operationLog.entries.count

        let edit = OOXMLEdit.insertParagraph(after: paraIDs[0], content: "second", styleId: nil)
        let result = try doc.apply(edit)

        XCTAssertEqual(result.operationLog.entries.count, logCountBefore + 1,
                       "After single-op apply, log has exactly one more entry")
        XCTAssertEqual(doc.operationLog.entries.count, logCountBefore,
                       "Original doc's log unchanged (immutable apply)")

        // The new entry's op should be the expected insertParagraphAfter.
        let lastEntry = result.operationLog.entries.last!
        guard case .insertParagraphAfter(let after, let payload) = lastEntry.op else {
            XCTFail("Expected insertParagraphAfter in log, got \(lastEntry.op)")
            return
        }
        XCTAssertEqual(after, paraIDs[0])
        XCTAssertEqual(payload.text, "second")
    }

    // MARK: - End-to-end: setBold via WordDocument.apply

    func testSetBoldMutatesRunRPr() throws {
        // Synthesize doc with a paragraph containing one Run with libraryUUID.
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
        let edit = OOXMLEdit.setBold(target: runID, value: true)
        let result = try doc.apply(edit)

        // Inspect result: target run should have <w:rPr><w:b/></w:rPr>
        guard let tree = result.xmlTrees["word/document.xml"],
              let run = OperationReducer.findNode(elementID: runID, in: tree),
              let rPr = run.children.first(where: { $0.kind == .element && $0.localName == "rPr" }) else {
            XCTFail("Expected run with rPr in result")
            return
        }
        let hasBold = rPr.children.contains { $0.kind == .element && $0.localName == "b" }
        XCTAssertTrue(hasBold, "After setBold(value: true), run has <w:b/> in rPr")
    }

    // MARK: - End-to-end: edit chain through apply<Sequence>

    func testApplySequenceChainsEditsViaDeterministicIDs() throws {
        let (doc, paraIDs) = makeSinglePartDoc(paragraphs: ["a"])

        // First insert "b" after "a"; new paragraph's libraryUUID == opID-of-first-edit.
        // But OOXMLEdit doesn't expose opID — apply() generates fresh UUID for each op.
        // So we can't chain via deterministic IDs from OOXMLEdit alone.
        //
        // What we CAN test: applying two independent edits in sequence yields
        // the expected text order, even if the second edit can't reference
        // the first's new paragraph by ID.
        let editB = OOXMLEdit.insertParagraph(after: paraIDs[0], content: "b", styleId: nil)
        let editC = OOXMLEdit.insertParagraph(after: paraIDs[0], content: "c", styleId: nil)
        // Both inserts target the SAME "a" — second insert goes AFTER "a",
        // before the previously-inserted "b". Result: [a, c, b].
        let result = try doc.apply([editB, editC] as [any Edit])

        XCTAssertEqual(extractTextFromDocPart(result), ["a", "c", "b"],
                       "Two inserts after same target: second goes between (LIFO order)")
    }
}
