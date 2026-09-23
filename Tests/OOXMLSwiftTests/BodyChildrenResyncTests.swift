import XCTest
@testable import OOXMLSwift

/// `resyncBodyFromDocumentTree` must rebuild every body-child kind (#104).
///
/// **Pre-fix bug**: #96 gave `appendParagraph` an op-log branch that fires
/// whenever the document came from disk (`xmlTrees["word/document.xml"] != nil`),
/// ending in `resyncBodyFromDocumentTree()` + `return` — bypassing the plain
/// `body.children.append(...)`. But that resync only re-typed `p` and `tbl`;
/// every other body-level element hit `default: continue` and vanished from the
/// typed view.
///
/// The two together are not an off-by-one — they are **data loss in the typed
/// view**. Measured before the fix:
///
///   after round-trip : ["paragraph", "bookmarkMarker", "paragraph"]  count 3
///   after append     : ["paragraph", "paragraph",      "paragraph"]  count 3
///
/// The append did not append, and the marker was overwritten. XML bytes stayed
/// intact in `xmlTrees` — only the typed projection lost them, which is exactly
/// the surface downstream consumers index against.
///
/// Downstream: che-word-mcp reports append position as
/// `body.children.count - 1` (its #69 decision — `getParagraphs()` skips
/// tables/SDTs, so its count mis-reports in documents with non-paragraph
/// children). Dropping those children silently corrupts that index space.
///
/// Refs #104, #96, PsychQuant/che-word-mcp#61, #69.
final class BodyChildrenResyncTests: XCTestCase {

    private var tempDir: URL!

    override func setUp() {
        super.setUp()
        tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("BodyChildrenResync-\(UUID().uuidString)")
        try? FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
    }

    override func tearDown() {
        try? FileManager.default.removeItem(at: tempDir)
        super.tearDown()
    }

    private func roundTrip(_ doc: WordDocument) throws -> WordDocument {
        let url = tempDir.appendingPathComponent("t-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: url)
        return try DocxReader.read(from: url)
    }

    private func kinds(_ doc: WordDocument) -> [String] {
        doc.body.children.map {
            switch $0 {
            case .paragraph: return "paragraph"
            case .table: return "table"
            case .bookmarkMarker: return "bookmarkMarker"
            case .rawBlockElement: return "rawBlockElement"
            case .contentControl: return "contentControl"
            default: return "other"
            }
        }
    }

    /// The regression that took down che-word-mcp: appending must grow
    /// `body.children` and must not consume a neighbouring non-paragraph child.
    func testAppendParagraphPreservesBookmarkMarker() throws {
        var doc = WordDocument()
        doc.body.children.append(.paragraph(Paragraph(runs: [Run(text: "before")])))
        doc.body.children.append(.bookmarkMarker(BookmarkRangeMarker(kind: .start, id: 1, position: 0)))
        doc.body.children.append(.paragraph(Paragraph(runs: [Run(text: "after")])))

        var reread = try roundTrip(doc)
        XCTAssertEqual(kinds(reread), ["paragraph", "bookmarkMarker", "paragraph"],
                       "round-trip itself must keep the marker")

        reread.appendParagraph(Paragraph(runs: [Run(text: "APPENDED")]))

        XCTAssertEqual(reread.body.children.count, 4,
                       "appendParagraph must grow body.children; got \(kinds(reread))")
        XCTAssertEqual(kinds(reread),
                       ["paragraph", "bookmarkMarker", "paragraph", "paragraph"],
                       "the bookmarkMarker must survive the append")
    }

    /// Same shape for an unknown body-level element: it must land as
    /// `.rawBlockElement` rather than being dropped, mirroring `DocxReader`'s
    /// post-#58 default (unknown children are captured as raw).
    func testResyncKeepsUnknownBodyElementAsRaw() throws {
        var doc = WordDocument()
        doc.body.children.append(.paragraph(Paragraph(runs: [Run(text: "p1")])))
        doc.body.children.append(.rawBlockElement(RawElement(
            name: "w:customXmlInsRangeStart",
            xml: "<w:customXmlInsRangeStart w:id=\"7\"/>")))
        doc.body.children.append(.paragraph(Paragraph(runs: [Run(text: "p2")])))

        var reread = try roundTrip(doc)
        let beforeCount = reread.body.children.count
        reread.appendParagraph(Paragraph(runs: [Run(text: "APPENDED")]))

        XCTAssertEqual(reread.body.children.count, beforeCount + 1,
                       "append must grow the typed view; got \(kinds(reread))")
        XCTAssertTrue(kinds(reread).contains("rawBlockElement"),
                      "unknown body element must not be dropped; got \(kinds(reread))")
    }

    /// A document with only paragraphs must be unaffected — guards against a
    /// fix that changes the common path.
    func testPlainParagraphDocumentAppendUnchanged() throws {
        var doc = WordDocument()
        doc.body.children.append(.paragraph(Paragraph(runs: [Run(text: "a")])))
        doc.body.children.append(.paragraph(Paragraph(runs: [Run(text: "b")])))

        var reread = try roundTrip(doc)
        let before = reread.body.children.count
        reread.appendParagraph(Paragraph(runs: [Run(text: "c")]))

        XCTAssertEqual(reread.body.children.count, before + 1)
        XCTAssertEqual(kinds(reread), Array(repeating: "paragraph", count: before + 1))
    }
}
