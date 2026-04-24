import XCTest
@testable import OOXMLSwift

/// Model-layer tests for the ContentControl and RepeatingSection changes
/// in `che-word-mcp-content-controls-read-write` (Phase 2 tasks 2.1–2.3).
///
/// Specs covered (from `specs/ooxml-content-insertion-primitives/spec.md`):
/// - ContentControl model supports nested children
/// - RepeatingSection model supports item-level update
/// - RepeatingSection emits allowInsertDeleteSections attribute
/// - WordDocument.allocateSdtId uses max-plus-one strategy
final class ContentControlModelTests: XCTestCase {

    // MARK: - Task 2.1: ContentControl nested children

    /// A Group ContentControl holds nested ContentControls via the new
    /// `children` field; each nested entry carries a `parentSdtId` pointing
    /// back to the Group.
    func testContentControlChildrenCarryParentSdtId() {
        let childSdt = StructuredDocumentTag(
            id: 101,
            tag: "city",
            alias: "City",
            type: .plainText
        )
        let child = ContentControl(sdt: childSdt, content: "Taipei", parentSdtId: 100)

        let groupSdt = StructuredDocumentTag(
            id: 100,
            tag: "address",
            alias: "Address",
            type: .group
        )
        let group = ContentControl(
            sdt: groupSdt,
            content: "",
            children: [child]
        )

        XCTAssertEqual(group.children.count, 1)
        XCTAssertEqual(group.children[0].sdt.tag, "city")
        XCTAssertEqual(group.children[0].parentSdtId, 100)
        XCTAssertNil(group.parentSdtId, "top-level SDT has no parent")
    }

    /// Default initializer should still work with only `sdt` + `content`
    /// (backward compatibility for existing callers).
    func testContentControlBackwardCompatibleInit() {
        let sdt = StructuredDocumentTag(
            id: 200,
            tag: "legacy",
            alias: "Legacy",
            type: .plainText
        )
        let control = ContentControl(sdt: sdt, content: "hello")

        XCTAssertEqual(control.children.count, 0)
        XCTAssertNil(control.parentSdtId)
    }

    /// `toXML()` must emit child SDTs inside the parent's `<w:sdtContent>`
    /// in insertion order, after any text content.
    func testContentControlToXMLEmitsChildrenInOrder() {
        let childSdt = StructuredDocumentTag(
            id: 101,
            tag: "city",
            alias: "City",
            type: .plainText
        )
        let child = ContentControl(sdt: childSdt, content: "Taipei", parentSdtId: 100)

        let groupSdt = StructuredDocumentTag(
            id: 100,
            tag: "address",
            alias: "Address",
            type: .group
        )
        let group = ContentControl(
            sdt: groupSdt,
            content: "",
            children: [child]
        )

        let xml = group.toXML()
        XCTAssertTrue(xml.contains("<w:tag w:val=\"address\"/>"))
        XCTAssertTrue(xml.contains("<w:tag w:val=\"city\"/>"))
        XCTAssertTrue(xml.contains("<w:group/>"))
        XCTAssertTrue(xml.contains("<w:text/>"))

        // Nested SDT must appear inside parent's sdtContent.
        guard let parentContentRange = xml.range(of: "<w:sdtContent>"),
              let parentContentEndRange = xml.range(of: "</w:sdtContent>", range: parentContentRange.upperBound..<xml.endIndex) else {
            XCTFail("parent sdtContent region not found")
            return
        }
        let innerRegion = String(xml[parentContentRange.upperBound..<parentContentEndRange.lowerBound])
        XCTAssertTrue(innerRegion.contains("<w:tag w:val=\"city\"/>"),
                      "nested child SDT should be inside parent's sdtContent")
    }

    // MARK: - Task 2.2: RepeatingSection updateItem + allowInsertDeleteSections

    /// `RepeatingSection.updateItem(atIndex:newText:)` replaces the item's
    /// content without disturbing other items.
    func testRepeatingSectionUpdateItem() throws {
        var section = RepeatingSection(
            tag: "items",
            alias: "Items",
            items: [
                RepeatingSectionItem(content: "A"),
                RepeatingSectionItem(content: "B"),
                RepeatingSectionItem(content: "C"),
            ]
        )

        try section.updateItem(atIndex: 1, newText: "B-updated")

        XCTAssertEqual(section.items.count, 3)
        XCTAssertEqual(section.items[0].content, "A")
        XCTAssertEqual(section.items[1].content, "B-updated")
        XCTAssertEqual(section.items[2].content, "C")
    }

    /// Out-of-range index throws `repeatingSectionItemOutOfBounds`.
    func testRepeatingSectionUpdateItemOutOfRange() {
        var section = RepeatingSection(
            tag: "items",
            alias: "Items",
            items: [RepeatingSectionItem(content: "only")]
        )

        XCTAssertThrowsError(try section.updateItem(atIndex: 5, newText: "x")) { error in
            if case WordError.repeatingSectionItemOutOfBounds(let index, let count) = error {
                XCTAssertEqual(index, 5)
                XCTAssertEqual(count, 1)
            } else {
                XCTFail("expected repeatingSectionItemOutOfBounds, got \(error)")
            }
        }
    }

    /// When `allowInsertDeleteSections=false`, the emit must produce the
    /// explicit `w15:val="0"` attribute (not omit the element).
    func testRepeatingSectionEmitsFalseAllowInsertDelete() {
        let section = RepeatingSection(
            tag: "items",
            alias: "Items",
            items: [RepeatingSectionItem(content: "A")],
            allowInsertDeleteSections: false
        )

        let xml = section.toXML()
        XCTAssertTrue(
            xml.contains("w15:allowInsertDeleteSections=\"0\"")
            || xml.contains("w15:allowInsertDeleteSection w15:val=\"false\"")
            || xml.contains("w15:allowInsertDeleteSection w15:val=\"0\""),
            "xml should record the disabled flag explicitly; got: \(xml)"
        )
    }

    /// When `allowInsertDeleteSections=true` (the default), the emit records
    /// the permission explicitly so parsers/readers see the intended value.
    func testRepeatingSectionEmitsTrueAllowInsertDelete() {
        let section = RepeatingSection(
            tag: "items",
            alias: "Items",
            items: [RepeatingSectionItem(content: "A")],
            allowInsertDeleteSections: true
        )

        let xml = section.toXML()
        XCTAssertTrue(
            xml.contains("w15:allowInsertDeleteSections=\"1\"")
            || xml.contains("w15:allowInsertDeleteSection w15:val=\"true\"")
            || xml.contains("w15:allowInsertDeleteSection w15:val=\"1\""),
            "xml should record the enabled flag explicitly; got: \(xml)"
        )
    }

    // MARK: - Task 2.3: WordDocument.allocateSdtId

    /// An empty WordDocument allocates id 1.
    func testAllocateSdtIdOnEmptyDocument() {
        var doc = WordDocument()
        XCTAssertEqual(doc.allocateSdtId(), 1)
    }

    /// Existing sequential ids cause max+1 allocation.
    func testAllocateSdtIdSequential() throws {
        var doc = WordDocument()
        let ccA = ContentControl(
            sdt: StructuredDocumentTag(id: 1, tag: "a", alias: "A", type: .plainText),
            content: ""
        )
        let ccB = ContentControl(
            sdt: StructuredDocumentTag(id: 2, tag: "b", alias: "B", type: .plainText),
            content: ""
        )
        let ccC = ContentControl(
            sdt: StructuredDocumentTag(id: 3, tag: "c", alias: "C", type: .plainText),
            content: ""
        )
        try doc.insertContentControl(ccA, at: 0)
        try doc.insertContentControl(ccB, at: 1)
        try doc.insertContentControl(ccC, at: 2)

        XCTAssertEqual(doc.allocateSdtId(), 4)
    }

    /// Existing non-sequential / large ids still produce max+1.
    func testAllocateSdtIdLargeExisting() throws {
        var doc = WordDocument()
        let ccBig = ContentControl(
            sdt: StructuredDocumentTag(id: 789012, tag: "big", alias: "Big", type: .plainText),
            content: ""
        )
        try doc.insertContentControl(ccBig, at: 0)

        XCTAssertEqual(doc.allocateSdtId(), 789013)
    }
}
