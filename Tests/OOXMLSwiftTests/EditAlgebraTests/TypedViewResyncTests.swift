// TypedViewResyncTests.swift
// EditAlgebra — addresses macdoc#110 item #8.
//
// After WordDocument.apply(_:), the typed views (body, styles, etc.) were
// stale relative to the new xmlTrees. Callers reading result.body.children
// would see PRE-edit content even though xmlTrees was correctly mutated.
//
// This PR fixes body.children specifically: it's rebuilt from the new
// document.xml tree after apply, with each <w:p> and <w:tbl> becoming a
// tree-backed Paragraph(xmlNode:) / Table(xmlNode:).
//
// Limitations (documented in WordDocument+Apply.swift):
//   - Only <w:p> and <w:tbl> become typed body children. Other body-level
//     elements (<w:sdt>, <w:bookmarkStart>/End, etc.) are NOT re-typed.
//   - Only document.xml's body is resynced. styles/headers/footers/etc.
//     remain stale after apply.
//
// Tests cover the working scope: body.children paragraph resync for the
// 4 functional OOXMLEdit cases.

import XCTest
@testable import OOXMLSwift

final class TypedViewResyncTests: XCTestCase {

    // MARK: - Helpers (single-part doc with N paragraphs)

    private func makeFixture(paragraphs texts: [String]) -> (WordDocument, [ElementID]) {
        var paraIDs: [ElementID] = []
        var paragraphNodes: [XmlNode] = []

        for text in texts {
            let paraUUID = UUID()
            paraIDs.append(ElementID(libraryUUID: paraUUID))

            let textNode = XmlNode.text(text)
            let wt = XmlNode.element(prefix: "w", localName: "t", children: [textNode])
            let wr = XmlNode.element(prefix: "w", localName: "r", children: [wt])
            let wp = XmlNode.element(prefix: "w", localName: "p", children: [wr])
            wp.libraryUUID = paraUUID
            paragraphNodes.append(wp)
        }

        let body = XmlNode.element(prefix: "w", localName: "body", children: paragraphNodes)
        let root = XmlNode.element(prefix: "w", localName: "document", children: [body])

        var doc = WordDocument()
        doc.xmlTrees["word/document.xml"] = XmlTree.synthesized(root: root)
        // Note: body.children is EMPTY before resync. The fixture doesn't
        // populate it from the synthesized tree — that's what resync does.
        return (doc, paraIDs)
    }

    /// Extracts body.children text contents in order. Returns nil for non-Paragraph entries.
    private func bodyParagraphTexts(_ doc: WordDocument) -> [String] {
        return doc.body.children.compactMap { child -> String? in
            guard case .paragraph(let p) = child else { return nil }
            return p.text
        }
    }

    // MARK: - insertParagraph resyncs body.children

    func testInsertParagraphResyncsBody() throws {
        let (doc, paraIDs) = makeFixture(paragraphs: ["first"])

        // Sanity: pre-apply, body.children is empty (fixture didn't populate it)
        XCTAssertEqual(doc.body.children.count, 0,
                       "Synthesized fixture has empty body.children pre-apply (unset)")

        let edit = OOXMLEdit.insertParagraph(after: paraIDs[0], content: "second", styleId: nil)
        var result = try doc.apply(edit)
        result.resyncBodyFromDocumentTree()

        // Post-apply + resync: body.children has 2 paragraphs (existing + new)
        XCTAssertEqual(bodyParagraphTexts(result), ["first", "second"],
                       "After insertParagraph apply + resync, body.children reflects new paragraph")
    }

    func testInsertParagraphBeforeResyncsBody() throws {
        let (doc, paraIDs) = makeFixture(paragraphs: ["second"])

        let edit = OOXMLEdit.insertParagraphBefore(before: paraIDs[0], content: "first", styleId: nil)
        var result = try doc.apply(edit)
        result.resyncBodyFromDocumentTree()

        XCTAssertEqual(bodyParagraphTexts(result), ["first", "second"],
                       "After insertParagraphBefore apply + resync, body.children reflects new paragraph at start")
    }

    // MARK: - removeParagraph resyncs body.children

    func testRemoveParagraphResyncsBody() throws {
        let (doc, paraIDs) = makeFixture(paragraphs: ["a", "b", "c"])

        let edit = OOXMLEdit.removeParagraph(target: paraIDs[1])  // remove "b"
        var result = try doc.apply(edit)
        result.resyncBodyFromDocumentTree()

        XCTAssertEqual(bodyParagraphTexts(result), ["a", "c"],
                       "After removeParagraph apply + resync, body.children reflects removal")
    }

    // MARK: - setBold doesn't change paragraph count but reflects new run state

    func testSetBoldResyncsBodyContent() throws {
        let (doc, paraIDs) = makeFixture(paragraphs: ["hello"])

        // Get the run ID inside the first paragraph's <w:r>
        let tree = doc.xmlTrees["word/document.xml"]!
        let bodyNode = tree.root.children.first { $0.kind == .element && $0.localName == "body" }!
        let firstP = bodyNode.children.first { $0.kind == .element && $0.localName == "p" }!
        let firstR = firstP.children.first { $0.kind == .element && $0.localName == "r" }!
        // The fixture run doesn't have a libraryUUID, so set one for addressing
        let runUUID = UUID()
        firstR.libraryUUID = runUUID
        let runID = ElementID(libraryUUID: runUUID)

        let edit = OOXMLEdit.setBold(target: runID, value: true)
        var result = try doc.apply(edit)
        result.resyncBodyFromDocumentTree()

        // body still has 1 paragraph
        XCTAssertEqual(result.body.children.count, 1, "Paragraph count unchanged after setBold + resync")

        // The paragraph IS tree-backed (xmlNode wired) so reads reflect new state
        guard case .paragraph(let p) = result.body.children[0] else {
            XCTFail("Expected paragraph in body.children[0]")
            return
        }
        XCTAssertNotNil(p.xmlNode, "Resync produces tree-backed paragraph (xmlNode != nil)")
        XCTAssertEqual(p.text, "hello", "Paragraph text unchanged by setBold")

        // The run inside should now have <w:b/> in rPr (verify via xmlNode walk)
        guard let pNode = p.xmlNode,
              let rNode = pNode.children.first(where: { $0.kind == .element && $0.localName == "r" }),
              let rPr = rNode.children.first(where: { $0.kind == .element && $0.localName == "rPr" }) else {
            XCTFail("Expected rPr in paragraph's run after setBold")
            return
        }
        let hasBold = rPr.children.contains { $0.kind == .element && $0.localName == "b" }
        XCTAssertTrue(hasBold, "Resynced paragraph's run has <w:b/> in rPr (post-edit state)")
    }

    // MARK: - resync produces tree-backed paragraphs (xmlNode != nil)

    func testResyncProducesTreeBackedParagraphs() throws {
        let (doc, paraIDs) = makeFixture(paragraphs: ["one", "two"])

        let edit = OOXMLEdit.insertParagraph(after: paraIDs[1], content: "three", styleId: nil)
        var result = try doc.apply(edit)
        result.resyncBodyFromDocumentTree()

        XCTAssertEqual(result.body.children.count, 3)
        for (i, child) in result.body.children.enumerated() {
            guard case .paragraph(let p) = child else {
                XCTFail("body.children[\(i)] should be a paragraph")
                continue
            }
            XCTAssertNotNil(p.xmlNode,
                            "body.children[\(i)] should be tree-backed (xmlNode != nil) after apply resync")
        }
    }

    // MARK: - Sequence apply preserves resync at each step

    func testSequenceApplyResyncsBodyAtEachStep() throws {
        let (doc, paraIDs) = makeFixture(paragraphs: ["x"])

        let edit1 = OOXMLEdit.insertParagraph(after: paraIDs[0], content: "y", styleId: nil)
        var intermediate = try doc.apply(edit1)
        intermediate.resyncBodyFromDocumentTree()
        XCTAssertEqual(bodyParagraphTexts(intermediate), ["x", "y"],
                       "After first apply + resync, body has 2 paragraphs")

        // Use the inserted "y"'s opID to chain a subsequent insert after it
        let yOpID = intermediate.operationLog.entries.last!.opID
        let yID = ElementID(libraryUUID: yOpID)
        let edit2 = OOXMLEdit.insertParagraph(after: yID, content: "z", styleId: nil)
        var final = try intermediate.apply(edit2)
        final.resyncBodyFromDocumentTree()

        XCTAssertEqual(bodyParagraphTexts(final), ["x", "y", "z"],
                       "After sequence apply + resync, body reflects chained insertions")
    }
}
