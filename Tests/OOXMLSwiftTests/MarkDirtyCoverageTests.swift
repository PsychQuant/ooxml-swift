import XCTest
@testable import OOXMLSwift

/// Coverage tests for the v0.13.0 dirty tracking architecture
/// (`che-word-mcp-true-byte-preservation` Spectra change).
///
/// **Audited mutators** (one test case per method below). When adding a new
/// mutating method to `WordDocument` or its substructs, also add a test case
/// here verifying the expected part path appears in `modifiedPartsView`.
///
/// Body / paragraphs / tables → `"word/document.xml"`:
/// - `appendParagraph`, `insertParagraph`, `updateParagraph`, `deleteParagraph`
/// - `replaceText`, `insertText`
/// - `insertTable`, `updateCell`, `deleteTable`, `mergeCells`
/// - row/column add/delete on tables
///
/// Style / numbering / properties:
/// - `addStyle` / `updateStyle` / `deleteStyle` → `"word/styles.xml"`
/// - properties set → `"docProps/core.xml"`
///
/// Multi-instance parts:
/// - `addHeader` → `"word/<header.fileName>"`
/// - `addFooter` → `"word/<footer.fileName>"`
final class MarkDirtyCoverageTests: XCTestCase {

    // MARK: - Set<String> dirty tracking foundation

    func testFreshlyInitializedDocumentHasEmptyModifiedParts() {
        let doc = WordDocument()
        XCTAssertTrue(doc.modifiedPartsView.isEmpty)
    }

    func testMarkPartDirtyInsertsPathExternally() {
        var doc = WordDocument()
        doc.markPartDirty("word/theme/theme1.xml")
        XCTAssertTrue(doc.modifiedPartsView.contains("word/theme/theme1.xml"))
    }

    func testMarkPartDirtyIsIdempotent() {
        var doc = WordDocument()
        doc.markPartDirty("word/theme/theme1.xml")
        doc.markPartDirty("word/theme/theme1.xml")
        XCTAssertEqual(doc.modifiedPartsView.filter { $0 == "word/theme/theme1.xml" }.count, 1)
    }

    func testMarkMultiplePartsAccumulates() {
        var doc = WordDocument()
        doc.markPartDirty("word/document.xml")
        doc.markPartDirty("word/styles.xml")
        XCTAssertEqual(doc.modifiedPartsView.count, 2)
    }

    // MARK: - Body / paragraph mutators → word/document.xml

    func testAppendParagraphMarksDocumentXMLDirty() {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "Hello"))
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))
    }

    func testInsertParagraphMarksDocumentXMLDirty() {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "first"))
        doc.modifiedParts.removeAll()
        doc.insertParagraph(Paragraph(text: "second"), at: 0)
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))
    }

    func testUpdateParagraphMarksDocumentXMLDirty() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "before"))
        doc.modifiedParts.removeAll()
        try doc.updateParagraph(at: 0, text: "after")
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))
    }

    func testDeleteParagraphMarksDocumentXMLDirty() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "a"))
        doc.appendParagraph(Paragraph(text: "b"))
        doc.modifiedParts.removeAll()
        try doc.deleteParagraph(at: 0)
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))
    }

    func testReplaceTextMarksDocumentXMLDirty() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "hello world"))
        doc.modifiedParts.removeAll()
        _ = try doc.replaceText(find: "hello", with: "goodbye")
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))
    }

    func testFormatParagraphMarksDocumentXMLDirty() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "x"))
        doc.modifiedParts.removeAll()
        try doc.formatParagraph(at: 0, with: RunProperties(bold: true))
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))
    }

    func testSetParagraphFormatMarksDocumentXMLDirty() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "x"))
        doc.modifiedParts.removeAll()
        var pp = ParagraphProperties()
        pp.alignment = .center
        try doc.setParagraphFormat(at: 0, properties: pp)
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))
    }

    func testApplyStyleMarksDocumentXMLDirty() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "x"))
        doc.modifiedParts.removeAll()
        try doc.applyStyle(at: 0, style: "Heading1")
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))
    }

    // MARK: - Table mutators → word/document.xml

    func testAppendTableMarksDocumentXMLDirty() {
        var doc = WordDocument()
        doc.appendTable(Table(rowCount: 1, columnCount: 1))
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))
    }

    func testInsertTableMarksDocumentXMLDirty() {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "x"))
        doc.modifiedParts.removeAll()
        doc.insertTable(Table(rowCount: 2, columnCount: 2), at: 0)
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))
    }

    func testUpdateCellMarksDocumentXMLDirty() throws {
        var doc = WordDocument()
        doc.appendTable(Table(rowCount: 1, columnCount: 1))
        doc.modifiedParts.removeAll()
        try doc.updateCell(tableIndex: 0, row: 0, col: 0, text: "cell text")
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))
    }

    func testDeleteTableMarksDocumentXMLDirty() throws {
        var doc = WordDocument()
        doc.appendTable(Table(rowCount: 1, columnCount: 1))
        doc.modifiedParts.removeAll()
        try doc.deleteTable(at: 0)
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))
    }

    func testMergeCellsHorizontalMarksDocumentXMLDirty() throws {
        var doc = WordDocument()
        doc.appendTable(Table(rowCount: 2, columnCount: 3))
        doc.modifiedParts.removeAll()
        try doc.mergeCellsHorizontal(tableIndex: 0, row: 0, startCol: 0, endCol: 2)
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))
    }

    func testMergeCellsVerticalMarksDocumentXMLDirty() throws {
        var doc = WordDocument()
        doc.appendTable(Table(rowCount: 3, columnCount: 2))
        doc.modifiedParts.removeAll()
        try doc.mergeCellsVertical(tableIndex: 0, col: 0, startRow: 0, endRow: 2)
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))
    }

    // MARK: - List mutators → word/document.xml + word/numbering.xml

    func testInsertBulletListMarksDocumentAndNumberingDirty() {
        var doc = WordDocument()
        _ = doc.insertBulletList(items: ["a", "b"])
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))
        XCTAssertTrue(doc.modifiedPartsView.contains("word/numbering.xml"))
    }

    func testInsertNumberedListMarksDocumentAndNumberingDirty() {
        var doc = WordDocument()
        _ = doc.insertNumberedList(items: ["one", "two"])
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))
        XCTAssertTrue(doc.modifiedPartsView.contains("word/numbering.xml"))
    }

    // MARK: - Page setup mutators → word/document.xml (sectPr lives in body)

    func testSetPageSizeMarksDocumentXMLDirty() {
        var doc = WordDocument()
        doc.setPageSize(.a4)
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))
    }

    func testSetPageMarginsMarksDocumentXMLDirty() {
        var doc = WordDocument()
        doc.setPageMargins(.normal)
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))
    }

    func testSetPageOrientationMarksDocumentXMLDirty() {
        var doc = WordDocument()
        doc.setPageOrientation(.landscape)
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))
    }

    func testInsertPageBreakMarksDocumentXMLDirty() {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "x"))
        doc.modifiedParts.removeAll()
        doc.insertPageBreak()
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))
    }

    func testInsertSectionBreakMarksDocumentXMLDirty() {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "x"))
        doc.modifiedParts.removeAll()
        doc.insertSectionBreak()
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))
    }

    // MARK: - Style mutators → word/styles.xml

    func testAddStyleMarksStylesXMLDirty() throws {
        var doc = WordDocument()
        try doc.addStyle(Style(id: "Custom1", name: "Custom 1", type: .paragraph))
        XCTAssertTrue(doc.modifiedPartsView.contains("word/styles.xml"))
    }

    func testUpdateStyleMarksStylesXMLDirty() throws {
        var doc = WordDocument()
        try doc.addStyle(Style(id: "Custom2", name: "Custom 2", type: .paragraph))
        doc.modifiedParts.removeAll()
        try doc.updateStyle(id: "Custom2", with: StyleUpdate(name: "Renamed"))
        XCTAssertTrue(doc.modifiedPartsView.contains("word/styles.xml"))
    }

    func testDeleteStyleMarksStylesXMLDirty() throws {
        var doc = WordDocument()
        try doc.addStyle(Style(id: "Custom3", name: "Custom 3", type: .paragraph))
        doc.modifiedParts.removeAll()
        try doc.deleteStyle(id: "Custom3")
        XCTAssertTrue(doc.modifiedPartsView.contains("word/styles.xml"))
    }

    // MARK: - Header / Footer mutators → word/<header.fileName>

    func testAddHeaderMarksHeaderFileDirty() {
        var doc = WordDocument()
        let header = doc.addHeader(text: "Page header", type: .first)
        XCTAssertTrue(doc.modifiedPartsView.contains("word/\(header.fileName)"))
    }

    func testUpdateHeaderMarksHeaderFileDirty() throws {
        var doc = WordDocument()
        let header = doc.addHeader(text: "before", type: .default)
        doc.modifiedParts.removeAll()
        try doc.updateHeader(id: header.id, text: "after")
        XCTAssertTrue(doc.modifiedPartsView.contains("word/\(header.fileName)"))
    }

    func testAddFooterMarksFooterFileDirty() {
        var doc = WordDocument()
        let footer = doc.addFooter(text: "Page footer", type: .first)
        XCTAssertTrue(doc.modifiedPartsView.contains("word/\(footer.fileName)"))
    }

    func testUpdateFooterMarksFooterFileDirty() throws {
        var doc = WordDocument()
        let footer = doc.addFooter(text: "before", type: .default)
        doc.modifiedParts.removeAll()
        try doc.updateFooter(id: footer.id, text: "after")
        XCTAssertTrue(doc.modifiedPartsView.contains("word/\(footer.fileName)"))
    }

    // MARK: - Comment mutators → word/comments.xml

    func testInsertCommentMarksCommentsXMLDirty() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "comment target"))
        doc.modifiedParts.removeAll()
        _ = try doc.insertComment(text: "review note", author: "tester", paragraphIndex: 0)
        XCTAssertTrue(doc.modifiedPartsView.contains("word/comments.xml"))
    }

    func testUpdateCommentMarksCommentsXMLDirty() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "comment target"))
        let id = try doc.insertComment(text: "v1", author: "tester", paragraphIndex: 0)
        doc.modifiedParts.removeAll()
        try doc.updateComment(commentId: id, text: "v2")
        XCTAssertTrue(doc.modifiedPartsView.contains("word/comments.xml"))
    }

    func testDeleteCommentMarksCommentsXMLDirty() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "comment target"))
        let id = try doc.insertComment(text: "x", author: "tester", paragraphIndex: 0)
        doc.modifiedParts.removeAll()
        try doc.deleteComment(commentId: id)
        XCTAssertTrue(doc.modifiedPartsView.contains("word/comments.xml"))
    }

    // MARK: - Footnote / Endnote mutators

    func testInsertFootnoteMarksFootnotesXMLDirty() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "fn anchor"))
        doc.modifiedParts.removeAll()
        _ = try doc.insertFootnote(text: "fn body", paragraphIndex: 0)
        XCTAssertTrue(doc.modifiedPartsView.contains("word/footnotes.xml"))
    }

    func testDeleteFootnoteMarksFootnotesXMLDirty() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "anchor"))
        let id = try doc.insertFootnote(text: "fn body", paragraphIndex: 0)
        doc.modifiedParts.removeAll()
        try doc.deleteFootnote(footnoteId: id)
        XCTAssertTrue(doc.modifiedPartsView.contains("word/footnotes.xml"))
    }

    func testInsertEndnoteMarksEndnotesXMLDirty() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "en anchor"))
        doc.modifiedParts.removeAll()
        _ = try doc.insertEndnote(text: "en body", paragraphIndex: 0)
        XCTAssertTrue(doc.modifiedPartsView.contains("word/endnotes.xml"))
    }

    func testDeleteEndnoteMarksEndnotesXMLDirty() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "anchor"))
        let id = try doc.insertEndnote(text: "en body", paragraphIndex: 0)
        doc.modifiedParts.removeAll()
        try doc.deleteEndnote(endnoteId: id)
        XCTAssertTrue(doc.modifiedPartsView.contains("word/endnotes.xml"))
    }

    // MARK: - Track changes mutators → word/settings.xml

    func testEnableTrackChangesMarksSettingsXMLDirty() {
        var doc = WordDocument()
        doc.enableTrackChanges(author: "tester")
        XCTAssertTrue(doc.modifiedPartsView.contains("word/settings.xml"))
    }

    func testDisableTrackChangesMarksSettingsXMLDirty() {
        var doc = WordDocument()
        doc.enableTrackChanges(author: "tester")
        doc.modifiedParts.removeAll()
        doc.disableTrackChanges()
        XCTAssertTrue(doc.modifiedPartsView.contains("word/settings.xml"))
    }

    // MARK: - Multi-mutation accumulation

    func testMultipleSequentialMutationsAccumulatePaths() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "body"))
        try doc.addStyle(Style(id: "Combo", name: "Combo", type: .paragraph))
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))
        XCTAssertTrue(doc.modifiedPartsView.contains("word/styles.xml"))
    }
}
