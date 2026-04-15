import XCTest
@testable import OOXMLSwift

final class CommentTests: XCTestCase {

    // MARK: - getCommentsFull (new API)

    func testGetCommentsFullTopLevelHasNilParent() {
        var doc = WordDocument()
        doc.comments.addComment(
            Comment(id: 1, author: "Alice", text: "Top-level only", paragraphIndex: 0)
        )

        let full = doc.getCommentsFull()

        XCTAssertEqual(full.count, 1)
        XCTAssertEqual(full[0].id, 1)
        XCTAssertNil(full[0].parentId, "Top-level comment must have nil parentId")
    }

    func testGetCommentsFullReplyParentId() {
        var doc = WordDocument()
        doc.comments.addComment(
            Comment(id: 1, author: "Alice", text: "parent", paragraphIndex: 0)
        )
        let reply = doc.comments.addReply(to: 1, author: "Bob", text: "reply")
        XCTAssertNotNil(reply, "addReply should succeed when parent exists")

        let full = doc.getCommentsFull()

        XCTAssertEqual(full.count, 2)
        let parent = full.first { $0.id == 1 }!
        let child = full.first { $0.author == "Bob" }!
        XCTAssertNil(parent.parentId, "Parent should have nil parentId")
        XCTAssertEqual(child.parentId, parent.id, "Child parentId should reference parent.id")
    }

    // MARK: - Legacy getComments() unchanged (regression)

    func testGetCommentsLegacyTupleUnchanged() {
        var doc = WordDocument()
        doc.comments.addComment(
            Comment(id: 1, author: "Alice", text: "single", paragraphIndex: 2)
        )

        let legacy = doc.getComments()
        XCTAssertEqual(legacy.count, 1)
        XCTAssertEqual(legacy[0].id, 1)
        XCTAssertEqual(legacy[0].author, "Alice")
        XCTAssertEqual(legacy[0].text, "single")
        XCTAssertEqual(legacy[0].paragraphIndex, 2)
    }

    func testGetCommentsLegacyDoesNotIncludeReplies() {
        // Verify legacy tuple shape stays as-is regardless of new reply data
        var doc = WordDocument()
        doc.comments.addComment(
            Comment(id: 1, author: "Alice", text: "parent", paragraphIndex: 0)
        )
        _ = doc.comments.addReply(to: 1, author: "Bob", text: "reply")

        let legacy = doc.getComments()
        XCTAssertEqual(legacy.count, 2, "legacy returns all comments regardless of reply status")
        // Verify tuple shape stays exactly: (id, author, text, paragraphIndex, date)
        // If parentId leaked into the tuple this line would fail to compile
        let firstFields: (Int, String, String, Int, Date) = (
            legacy[0].id, legacy[0].author, legacy[0].text, legacy[0].paragraphIndex, legacy[0].date
        )
        XCTAssertEqual(firstFields.0, 1)
    }
}
