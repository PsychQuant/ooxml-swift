// InsertHyperlinkE2ETests.swift
// EditAlgebra — addresses ooxml-swift#71 Phase 2c reducer for
// OOXMLEdit.insertHyperlink end-to-end mutation.
//
// After this PR ships Reducer support for:
//   - Operation.insertSiblingAfter (new typed primitive)
//   - Operation.addRelationship
//
// OOXMLEdit.insertHyperlink becomes functional end-to-end: doc.apply
// produces a WordDocument whose xmlTrees["word/document.xml"] has the
// <w:hyperlink> wrapper inserted after the target Run, and
// xmlTrees["word/_rels/document.xml.rels"] has the new <Relationship>
// entry.
//
// Wrap semantics (OOXMLEdit.wrapWithHyperlink) is NOT covered here —
// the Reducer for wrapWithHyperlink (placeholder substitution or typed
// wrap primitive) ships in a follow-up PR.

import XCTest
@testable import OOXMLSwift

final class InsertHyperlinkE2ETests: XCTestCase {

    // MARK: - Fixture

    /// Builds a multi-part WordDocument:
    ///   - word/document.xml: paragraph with one <w:r><w:t>before</w:t></w:r>
    ///   - (no rels part initially — apply creates it on-demand)
    /// Returns the doc + the target run's ElementID.
    private func makeFixture() -> (WordDocument, ElementID) {
        let runUUID = UUID()
        let textNode = XmlNode.text("before")
        let wt = XmlNode.element(prefix: "w", localName: "t", children: [textNode])
        let wr = XmlNode.element(prefix: "w", localName: "r", children: [wt])
        wr.libraryUUID = runUUID
        let wp = XmlNode.element(prefix: "w", localName: "p", children: [wr])
        let body = XmlNode.element(prefix: "w", localName: "body", children: [wp])
        let root = XmlNode.element(prefix: "w", localName: "document", children: [body])

        var doc = WordDocument()
        doc.xmlTrees["word/document.xml"] = XmlTree.synthesized(root: root)
        return (doc, ElementID(libraryUUID: runUUID))
    }

    // MARK: - End-to-end: insertHyperlink mutates both parts

    func testInsertHyperlinkAddsWrapperAfterTargetRun() throws {
        let (doc, runID) = makeFixture()
        let url = URL(string: "https://example.com")!

        let edit = OOXMLEdit.insertHyperlink(
            target: runID,
            href: url,
            displayText: "click here"
        )
        let result = try doc.apply(edit)

        // document.xml side: paragraph now has [run, hyperlink] in order
        let docTree = result.xmlTrees["word/document.xml"]!
        let body = docTree.root.children.first { $0.kind == .element && $0.localName == "body" }!
        let para = body.children.first { $0.kind == .element && $0.localName == "p" }!

        let paraChildren = para.children.filter { $0.kind == .element }
        XCTAssertEqual(paraChildren.count, 2,
                       "Paragraph has 2 element children: original run + hyperlink wrapper")
        XCTAssertEqual(paraChildren[0].localName, "r",
                       "First child is the original run")
        XCTAssertEqual(paraChildren[1].localName, "hyperlink",
                       "Second child is the new hyperlink wrapper")

        // hyperlink wrapper contains a run with "click here"
        let hyperlink = paraChildren[1]
        let hyperlinkRId = hyperlink.attributeValue(prefix: "r", localName: "id")
        XCTAssertNotNil(hyperlinkRId, "Hyperlink has r:id attribute")
        XCTAssertTrue(hyperlinkRId?.hasPrefix("rIdEdit") ?? false,
                      "rId follows freshRelationshipId convention: \(hyperlinkRId ?? "nil")")

        let hyperlinkRun = hyperlink.children.first {
            $0.kind == .element && $0.localName == "r"
        }
        XCTAssertNotNil(hyperlinkRun, "Hyperlink wraps one <w:r>")
        let hyperlinkText = hyperlinkRun?.children.first {
            $0.kind == .element && $0.localName == "t"
        }?.children.first { $0.kind == .text }?.textContent
        XCTAssertEqual(hyperlinkText, "click here", "Display text is 'click here'")
    }

    func testInsertHyperlinkCreatesRelsPartOnDemand() throws {
        let (doc, runID) = makeFixture()
        let url = URL(string: "https://example.com/test")!

        // Sanity: rels part doesn't exist before apply
        XCTAssertNil(doc.xmlTrees["word/_rels/document.xml.rels"],
                     "Fixture has no rels part initially")

        let edit = OOXMLEdit.insertHyperlink(target: runID, href: url, displayText: nil)
        let result = try doc.apply(edit)

        // rels part is created
        let relsTree = result.xmlTrees["word/_rels/document.xml.rels"]
        XCTAssertNotNil(relsTree, "Apply created rels part on-demand")
        XCTAssertEqual(relsTree?.root.localName, "Relationships",
                       "Rels part root is <Relationships>")

        // Single <Relationship> child added
        let relationships = relsTree?.root.children.filter {
            $0.kind == .element && $0.localName == "Relationship"
        } ?? []
        XCTAssertEqual(relationships.count, 1, "One <Relationship> entry added")

        // Attributes pin the contract
        let rel = relationships[0]
        XCTAssertEqual(rel.attributeValue(prefix: nil, localName: "Type"),
                       "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink")
        XCTAssertEqual(rel.attributeValue(prefix: nil, localName: "Target"),
                       "https://example.com/test")
        XCTAssertEqual(rel.attributeValue(prefix: nil, localName: "TargetMode"), "External")

        // The Id attribute matches the hyperlink wrapper's r:id (referential integrity)
        let docTree = result.xmlTrees["word/document.xml"]!
        let para = docTree.root
            .children.first { $0.kind == .element && $0.localName == "body" }!
            .children.first { $0.kind == .element && $0.localName == "p" }!
        let hyperlink = para.children.first { $0.kind == .element && $0.localName == "hyperlink" }!
        let docRId = hyperlink.attributeValue(prefix: "r", localName: "id")
        let relsId = rel.attributeValue(prefix: nil, localName: "Id")
        XCTAssertEqual(docRId, relsId,
                       "Hyperlink's r:id matches Relationship's Id (referential integrity)")
    }

    func testInsertHyperlinkNilDisplayTextFallsBackToHref() throws {
        let (doc, runID) = makeFixture()
        let url = URL(string: "https://example.com/x")!

        let edit = OOXMLEdit.insertHyperlink(target: runID, href: url, displayText: nil)
        let result = try doc.apply(edit)

        // displayText nil → uses href as displayed text
        let docTree = result.xmlTrees["word/document.xml"]!
        let para = docTree.root
            .children.first { $0.kind == .element && $0.localName == "body" }!
            .children.first { $0.kind == .element && $0.localName == "p" }!
        let hyperlink = para.children.first { $0.kind == .element && $0.localName == "hyperlink" }!
        let displayText = hyperlink.children
            .first { $0.kind == .element && $0.localName == "r" }?
            .children.first { $0.kind == .element && $0.localName == "t" }?
            .children.first { $0.kind == .text }?.textContent
        XCTAssertEqual(displayText, "https://example.com/x",
                       "Nil displayText falls back to href.absoluteString per §5 Q4 verdict")
    }

    // MARK: - wrapWithHyperlink end-to-end (via OOXMLEdit + via WordEdit.applyLink)

    func testWrapWithHyperlinkReplacesTargetWithWrappedRun() throws {
        let (doc, runID) = makeFixture()
        let url = URL(string: "https://example.com/wrap")!

        let edit = OOXMLEdit.wrapWithHyperlink(target: runID, href: url)
        let result = try doc.apply(edit)

        // The paragraph now contains the hyperlink wrapper directly
        // (not as a sibling — Wrap REPLACES target's position).
        let docTree = result.xmlTrees["word/document.xml"]!
        let para = docTree.root
            .children.first { $0.kind == .element && $0.localName == "body" }!
            .children.first { $0.kind == .element && $0.localName == "p" }!

        let paraChildren = para.children.filter { $0.kind == .element }
        XCTAssertEqual(paraChildren.count, 1,
                       "After wrap: paragraph has exactly 1 element child (the hyperlink wrapper)")
        XCTAssertEqual(paraChildren[0].localName, "hyperlink",
                       "That child IS the hyperlink wrapper, not a sibling Run")

        // Wrapper's r:id matches the rels entry's Id (referential integrity)
        let wrapper = paraChildren[0]
        let wrapperRId = wrapper.attributeValue(prefix: "r", localName: "id")
        XCTAssertNotNil(wrapperRId)

        let relsTree = result.xmlTrees["word/_rels/document.xml.rels"]
        XCTAssertNotNil(relsTree, "Rels part auto-created")
        let relationship = relsTree?.root.children.first {
            $0.kind == .element && $0.localName == "Relationship"
        }
        let relsId = relationship?.attributeValue(prefix: nil, localName: "Id")
        XCTAssertEqual(wrapperRId, relsId,
                       "Wrapper's r:id matches rels Relationship's Id (referential integrity)")

        // The wrapped run is INSIDE the hyperlink and preserves the original text
        let innerRun = wrapper.children.first { $0.kind == .element && $0.localName == "r" }
        XCTAssertNotNil(innerRun, "Hyperlink wraps a <w:r>")
        let innerText = innerRun?.children.first { $0.kind == .element && $0.localName == "t" }?
            .children.first { $0.kind == .text }?.textContent
        XCTAssertEqual(innerText, "before",
                       "Wrapped run preserves the original run's text content")
    }

    func testWordEditApplyLinkSingleRunEndToEnd() throws {
        // WordEdit.applyLink lowers to OOXMLEdit.wrapWithHyperlink (Design Y).
        // With wrapWithHyperlink Reducer now functional, applyLink works
        // end-to-end. This proves the full chain: WordEdit semantic-layer →
        // lower → OOXMLEdit syntactic-layer → operations → log → materialize.
        let (doc, runID) = makeFixture()
        let url = URL(string: "https://example.com/applylink")!
        let range = WordRange(startRun: runID, startOffset: 0, endRun: runID, endOffset: 6)
        let edit = WordEdit.applyLink(range: range, url: url)

        let result = try doc.apply(edit)

        // Verify the wrap happened
        let docTree = result.xmlTrees["word/document.xml"]!
        let para = docTree.root
            .children.first { $0.kind == .element && $0.localName == "body" }!
            .children.first { $0.kind == .element && $0.localName == "p" }!
        let hyperlink = para.children.first { $0.kind == .element && $0.localName == "hyperlink" }
        XCTAssertNotNil(hyperlink, "applyLink produced <w:hyperlink> wrapping the run")

        // Rels entry exists with the right URL
        let relsTree = result.xmlTrees["word/_rels/document.xml.rels"]
        let rel = relsTree?.root.children.first {
            $0.kind == .element && $0.localName == "Relationship"
        }
        XCTAssertEqual(rel?.attributeValue(prefix: nil, localName: "Target"),
                       "https://example.com/applylink")
    }
}
