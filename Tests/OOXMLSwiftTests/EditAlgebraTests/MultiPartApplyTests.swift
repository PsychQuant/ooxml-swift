// MultiPartApplyTests.swift
// EditAlgebra — addresses macdoc#110 multi-part scoping fix.
//
// Two related bugs in WordDocument+Apply.swift §2 scaffold:
//
// 1. **Multi-part scoping**: the scaffold iterates ALL parts of the document
//    and applies the full op log to each. For multi-part documents (real
//    .docx has document.xml + styles.xml + comments.xml + etc.), an op
//    targeting document.xml fails on styles.xml because the target ElementID
//    isn't in that tree → Reducer throws elementNotFound → apply throws
//    operationLogFailure even though the op SHOULD have succeeded.
//
// 2. **opID-determinism**: the scaffold appends each op to newLog AND tempLog
//    via separate `.append(op, source:)` calls, each generating a fresh
//    UUID. New nodes' libraryUUIDs derive from tempLog's opID (used during
//    materialize); the PERSISTED log carries newLog's opID. Re-materializing
//    the persisted log later produces nodes with DIFFERENT IDs than the
//    freshly-applied doc — replay-determinism violation.
//
// These tests should FAIL against the current §2 scaffold and PASS after the
// fix lands.

import XCTest
@testable import OOXMLSwift

final class MultiPartApplyTests: XCTestCase {

    // MARK: - Helpers (multi-part synthesized doc)

    /// Builds a WordDocument with TWO parts:
    ///   - "word/document.xml" — contains one paragraph
    ///   - "word/styles.xml" — contains one style element
    /// Returns the doc + the document.xml paragraph's ElementID.
    private func makeMultiPartDoc() -> (WordDocument, ElementID) {
        // document.xml part
        let paraUUID = UUID()
        let textNode = XmlNode.text("hello")
        let wt = XmlNode.element(prefix: "w", localName: "t", children: [textNode])
        let wr = XmlNode.element(prefix: "w", localName: "r", children: [wt])
        let wp = XmlNode.element(prefix: "w", localName: "p", children: [wr])
        wp.libraryUUID = paraUUID
        let body = XmlNode.element(prefix: "w", localName: "body", children: [wp])
        let docRoot = XmlNode.element(prefix: "w", localName: "document", children: [body])

        // styles.xml part — different tree shape, NO paragraph addressable from document.xml
        let style = XmlNode.element(prefix: "w", localName: "style")
        style.setAttribute(prefix: "w", localName: "styleId", value: "Normal")
        let styles = XmlNode.element(prefix: "w", localName: "styles", children: [style])

        var doc = WordDocument()
        doc.xmlTrees["word/document.xml"] = XmlTree.synthesized(root: docRoot)
        doc.xmlTrees["word/styles.xml"] = XmlTree.synthesized(root: styles)
        return (doc, ElementID(libraryUUID: paraUUID))
    }

    /// Walks a tree finding all <w:p> nodes and returns their text content in
    /// document order.
    private func paragraphTextsIn(_ tree: XmlTree?) -> [String] {
        guard let tree = tree else { return [] }
        var result: [String] = []
        func walk(_ node: XmlNode) {
            if node.kind == .element && node.localName == "p" {
                var text = ""
                func collectText(_ n: XmlNode) {
                    if n.kind == .text { text += n.textContent }
                    for c in n.children { collectText(c) }
                }
                collectText(node)
                result.append(text)
            }
            for child in node.children {
                walk(child)
            }
        }
        walk(tree.root)
        return result
    }

    // MARK: - Bug 1: multi-part scoping

    func testApplyMultiPartDocSucceeds() throws {
        // Bug 1 reproduction: apply insertParagraph to a doc with
        // document.xml + styles.xml. With the scaffold bug, this throws
        // operationLogFailure because the same op is applied to styles.xml
        // (where the target doesn't exist) → elementNotFound → wrap.
        let (doc, paraID) = makeMultiPartDoc()

        let edit = OOXMLEdit.insertParagraph(after: paraID, content: "second", styleId: nil)
        let result = try doc.apply(edit)

        // document.xml should have the new paragraph
        XCTAssertEqual(paragraphTextsIn(result.xmlTrees["word/document.xml"]),
                       ["hello", "second"],
                       "Multi-part apply succeeds: target part mutated correctly")

        // styles.xml should be UNCHANGED — no <w:p> in styles.xml regardless
        XCTAssertEqual(paragraphTextsIn(result.xmlTrees["word/styles.xml"]), [],
                       "Non-target part untouched (no paragraphs)")

        // styles.xml tree should be equivalent to input (no spurious mutations)
        let beforeStylesXML = doc.xmlTrees["word/styles.xml"]
        let afterStylesXML = result.xmlTrees["word/styles.xml"]
        XCTAssertNotNil(beforeStylesXML)
        XCTAssertNotNil(afterStylesXML)
        // Identity check: both trees should describe a single <w:styles><w:style w:styleId="Normal"/></w:styles>
        // We assert via structural walk (avoid pointer identity since deep copy)
        XCTAssertEqual(afterStylesXML?.root.localName, "styles",
                       "styles.xml root unchanged")
        XCTAssertEqual(afterStylesXML?.root.children.count, 1,
                       "styles.xml has exactly 1 child element")
        XCTAssertEqual(afterStylesXML?.root.children.first?.attributeValue(prefix: "w", localName: "styleId"),
                       "Normal",
                       "styles.xml style@styleId preserved")
    }

    // MARK: - Bug 2: opID-determinism (same applied doc vs replayed log)

    func testApplyOpIDDeterminism() throws {
        // Bug 2 reproduction: after apply, the new paragraph's libraryUUID
        // should equal the opID stored in the persisted log. Re-materializing
        // the persisted log against the original tree should produce a tree
        // with the SAME new-paragraph ID.
        let (doc, paraID) = makeMultiPartDoc()

        let edit = OOXMLEdit.insertParagraph(after: paraID, content: "second", styleId: nil)
        let result = try doc.apply(edit)

        // Extract the persisted log entry's opID
        XCTAssertEqual(result.operationLog.entries.count, 1)
        let persistedOpID = result.operationLog.entries[0].opID

        // Extract the new paragraph's libraryUUID from the freshly-applied tree
        let docXMLTree = result.xmlTrees["word/document.xml"]!
        var newParaUUID: UUID?
        func findNewPara(_ node: XmlNode) {
            if node.kind == .element && node.localName == "p" {
                // The "new" paragraph is the one whose libraryUUID isn't paraID's UUID
                if let uuid = node.libraryUUID, uuid != paraID.libraryUUID {
                    newParaUUID = uuid
                }
            }
            for child in node.children {
                findNewPara(child)
            }
        }
        findNewPara(docXMLTree.root)

        XCTAssertNotNil(newParaUUID, "New paragraph has a libraryUUID")
        XCTAssertEqual(newParaUUID, persistedOpID,
                       "Replay determinism: new paragraph's libraryUUID == persisted log entry's opID. " +
                       "Without this, re-materializing the persisted log later would produce different IDs.")
    }

    func testApplyPersistedLogReplaysToSameTree() throws {
        // Stronger replay-determinism check: re-materialize the persisted log
        // against the ORIGINAL document and assert the resulting tree is
        // structurally identical (same paragraph IDs, same content) to the
        // freshly-applied result.
        let (doc, paraID) = makeMultiPartDoc()

        let edit = OOXMLEdit.insertParagraph(after: paraID, content: "second", styleId: nil)
        let freshResult = try doc.apply(edit)

        // Re-materialize freshResult's persisted log against doc's original tree
        let originalDocTree = doc.xmlTrees["word/document.xml"]!
        let replayedTree = try OperationReducer.materialize(
            log: freshResult.operationLog,
            base: originalDocTree
        )

        // Compare: both should have same paragraph IDs in same order
        func extractParaIDs(_ tree: XmlTree) -> [UUID?] {
            var ids: [UUID?] = []
            func walk(_ node: XmlNode) {
                if node.kind == .element && node.localName == "p" {
                    ids.append(node.libraryUUID)
                }
                for child in node.children { walk(child) }
            }
            walk(tree.root)
            return ids
        }

        XCTAssertEqual(extractParaIDs(freshResult.xmlTrees["word/document.xml"]!),
                       extractParaIDs(replayedTree),
                       "Replayed log produces same paragraph IDs as freshly-applied doc")
    }

    // MARK: - Sequential apply through a chain (in-flight created IDs)

    func testApplySequenceChainsThroughCreatedIDs() throws {
        // After op1 inserts paragraph "B" with libraryUUID == op1.opID, op2
        // can reference "B" via ElementID(libraryUUID: op1.opID). The fix
        // needs to handle in-flight created IDs (op2's target was created by
        // op1, exists only in the partial state).
        let (doc, aID) = makeMultiPartDoc()

        // First insert B after A. We don't control opID, so we have to
        // chain via doc.apply().operationLog.entries.last.opID afterwards.
        let editB = OOXMLEdit.insertParagraph(after: aID, content: "B", styleId: nil)
        let afterB = try doc.apply(editB)

        let bOpID = afterB.operationLog.entries.last!.opID
        let bID = ElementID(libraryUUID: bOpID)

        // Now insert C after B
        let editC = OOXMLEdit.insertParagraph(after: bID, content: "C", styleId: nil)
        let afterC = try afterB.apply(editC)

        XCTAssertEqual(paragraphTextsIn(afterC.xmlTrees["word/document.xml"]),
                       ["hello", "B", "C"],
                       "Chain through in-flight created ID works")
    }
}
