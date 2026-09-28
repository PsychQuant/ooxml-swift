import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Regression tests for [PsychQuant/ooxml-swift#196](https://github.com/PsychQuant/ooxml-swift/issues/196).
///
/// #195 gave `TableCell.toXML(depth:)` a write-side recursion depth guard
/// (`Table.maxSerializationDepth = 10`) so a pathologically deep table
/// (reachable only via the public struct API — `TableCell.nestedTables` is
/// a plain array setter with no depth validation, or via
/// `resyncBodyFromDocumentTree()`) cannot stack-overflow the writer. But the
/// guard's own failure mode was silent: past the limit, `toXML(depth:)`
/// simply `continue`s past the excess nested tables (dropping them from the
/// output) while `DocxWriter.write`/`writeData` still reports success. The
/// saved file has fewer tables than the in-memory model — a silent
/// data-loss shape indistinguishable, from the caller's perspective, from a
/// clean save.
///
/// The fix adds `Table.measuredNestingDepth()` (a read-only measurement of
/// a table's own nesting, independent of and prior to `toXML`) plus a
/// pre-write scan (`DocxWriter.validateTableNestingDepthBeforeWrite`,
/// private — exercised here only through the public `write`/`writeData`
/// entry points, matching how a real caller would hit it) covering every
/// location that can carry a `Table`: body, headers, footers, footnotes,
/// endnotes. The scan runs before any part is written and throws
/// `WordError.nestedTooDeep(depth:max:)` — the SAME case
/// `Document.insertNestedTable`'s API-level guard already uses — with zero
/// side effects, instead of writing a truncated package.
final class Issue196TableDepthPreWriteValidationTests: XCTestCase {

    /// Builds a chain of `depth` levels of table-in-single-cell nesting
    /// (innermost first), purely through the public struct API — the exact
    /// technique `Issue195TwoLevelNestedTableCrashTests
    /// .testAPIConstructedExcessiveDepthDoesNotCrashAndTruncates` uses to
    /// bypass `DocxReader.parseTable`'s separate, unrelated
    /// `MAX_TABLE_NEST_DEPTH = 5` reader-side guard. The returned table's
    /// `measuredNestingDepth()` equals `depth`.
    private func makeNestedTable(depth: Int) -> Table {
        var innermost = Table(rows: [TableRow(cells: [TableCell(text: "leaf")])])
        for i in 0..<depth {
            var cell = TableCell(paragraphs: [Paragraph(text: "L\(i)")])
            cell.nestedTables = [innermost]
            innermost = Table(rows: [TableRow(cells: [cell])])
        }
        return innermost
    }

    // MARK: - Measurement sanity

    func testMeasuredNestingDepthMatchesConstructionDepth() {
        for depth in [0, 1, 5, 10, 11, 15] {
            let table = makeNestedTable(depth: depth)
            XCTAssertEqual(table.measuredNestingDepth(), depth, "depth=\(depth)")
        }
    }

    // MARK: - Body

    /// Exactly at the limit: must write successfully, all levels intact.
    /// Reads the raw `word/document.xml` bytes directly out of the saved
    /// ZIP (NOT through `DocxReader.read`, which has its own, separate,
    /// unrelated `MAX_TABLE_NEST_DEPTH = 5` reader-side guard — reading
    /// this fixture back through the full reader would fail for a reason
    /// that has nothing to do with what this test is checking).
    func testBodyTableAtLimitWritesSuccessfullyWithAllLevelsIntact() throws {
        let table = makeNestedTable(depth: Table.maxSerializationDepth)
        var doc = WordDocument()
        doc.body.children = [.table(table)]

        let data = try DocxWriter.writeData(doc)
        let xml = try extractRawDocumentXML(from: data)
        let tblOpens = xml.components(separatedBy: "<w:tbl>").count - 1
        XCTAssertEqual(tblOpens, Table.maxSerializationDepth + 1,
                       "all \(Table.maxSerializationDepth + 1) tables (root + \(Table.maxSerializationDepth) nested) must round-trip at exactly the limit")
    }

    /// One past the limit: must refuse BEFORE writing anything, not
    /// silently truncate. This is the core #196 assertion — RED before the
    /// fix (write "succeeds" with a truncated package), GREEN after.
    func testBodyTableOneOverLimitThrowsBeforeWritingAnyBytes() throws {
        let table = makeNestedTable(depth: Table.maxSerializationDepth + 1)
        var doc = WordDocument()
        doc.body.children = [.table(table)]

        XCTAssertThrowsError(try DocxWriter.writeData(doc)) { error in
            guard case WordError.nestedTooDeep(let depth, let max) = error else {
                XCTFail("expected WordError.nestedTooDeep, got \(error)")
                return
            }
            XCTAssertEqual(depth, Table.maxSerializationDepth + 1)
            XCTAssertEqual(max, Table.maxSerializationDepth)
        }
    }

    /// Same over-limit table, reached via `DocxWriter.write(_:to:)` (the
    /// atomic-rename file path, not just the in-memory `writeData`) — the
    /// target file must not exist afterward (zero side effects), and any
    /// pre-existing file at that path must be untouched.
    func testBodyTableOneOverLimitLeavesNoFileOnDiskAndPreservesExisting() throws {
        let table = makeNestedTable(depth: Table.maxSerializationDepth + 1)
        var doc = WordDocument()
        doc.body.children = [.table(table)]

        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue196-\(UUID().uuidString).docx")
        let sentinel = Data("not a docx — must survive untouched".utf8)
        try sentinel.write(to: outURL)
        defer { try? FileManager.default.removeItem(at: outURL) }

        XCTAssertThrowsError(try DocxWriter.write(doc, to: outURL)) { error in
            guard case WordError.nestedTooDeep = error else {
                XCTFail("expected WordError.nestedTooDeep, got \(error)")
                return
            }
        }
        let survived = try Data(contentsOf: outURL)
        XCTAssertEqual(survived, sentinel, "pre-existing file at the target path must be untouched")
    }

    // MARK: - Content-control-wrapped body table

    /// A table over the limit, but wrapped in a block-level content
    /// control (`BodyChild.contentControl`) rather than sitting directly
    /// in `body.children` — the scan must recurse into content-control
    /// children, not just the top level.
    func testContentControlWrappedTableOverLimitIsCaught() throws {
        let table = makeNestedTable(depth: Table.maxSerializationDepth + 1)
        var doc = WordDocument()
        let sdt = ContentControl(sdt: StructuredDocumentTag(tag: "wrapper"), content: "")
        doc.body.children = [.contentControl(sdt, children: [.table(table)])]

        XCTAssertThrowsError(try DocxWriter.writeData(doc)) { error in
            guard case WordError.nestedTooDeep = error else {
                XCTFail("expected WordError.nestedTooDeep, got \(error)")
                return
            }
        }
    }

    // MARK: - Headers / footers / footnotes / endnotes

    func testHeaderTableOverLimitIsCaught() throws {
        let table = makeNestedTable(depth: Table.maxSerializationDepth + 1)
        var doc = WordDocument()
        doc.body.children = [.paragraph(Paragraph(text: "body"))]
        var header = Header(id: "rIdHeader1", originalFileName: "header1.xml")
        header.bodyChildren = [.table(table)]
        doc.headers = [header]

        XCTAssertThrowsError(try DocxWriter.writeData(doc)) { error in
            guard case WordError.nestedTooDeep = error else {
                XCTFail("expected WordError.nestedTooDeep, got \(error)")
                return
            }
        }
    }

    func testFooterTableOverLimitIsCaught() throws {
        let table = makeNestedTable(depth: Table.maxSerializationDepth + 1)
        var doc = WordDocument()
        doc.body.children = [.paragraph(Paragraph(text: "body"))]
        var footer = Footer(id: "rIdFooter1", originalFileName: "footer1.xml")
        footer.bodyChildren = [.table(table)]
        doc.footers = [footer]

        XCTAssertThrowsError(try DocxWriter.writeData(doc)) { error in
            guard case WordError.nestedTooDeep = error else {
                XCTFail("expected WordError.nestedTooDeep, got \(error)")
                return
            }
        }
    }

    func testFootnoteTableOverLimitIsCaught() throws {
        let table = makeNestedTable(depth: Table.maxSerializationDepth + 1)
        var doc = WordDocument()
        doc.body.children = [.paragraph(Paragraph(text: "body"))]
        var footnote = Footnote(id: 1, text: "", paragraphIndex: 0)
        footnote.bodyChildren = [.table(table)]
        doc.footnotes.footnotes = [footnote]

        XCTAssertThrowsError(try DocxWriter.writeData(doc)) { error in
            guard case WordError.nestedTooDeep = error else {
                XCTFail("expected WordError.nestedTooDeep, got \(error)")
                return
            }
        }
    }

    func testEndnoteTableOverLimitIsCaught() throws {
        let table = makeNestedTable(depth: Table.maxSerializationDepth + 1)
        var doc = WordDocument()
        doc.body.children = [.paragraph(Paragraph(text: "body"))]
        var endnote = Endnote(id: 1, text: "", paragraphIndex: 0)
        endnote.bodyChildren = [.table(table)]
        doc.endnotes.endnotes = [endnote]

        XCTAssertThrowsError(try DocxWriter.writeData(doc)) { error in
            guard case WordError.nestedTooDeep = error else {
                XCTFail("expected WordError.nestedTooDeep, got \(error)")
                return
            }
        }
    }

    // MARK: - Sanity: ordinary documents unaffected

    func testOrdinaryShallowDocumentWritesUnaffected() throws {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "hello")),
            .table(Table(rows: [TableRow(cells: [TableCell(text: "cell")])])),
        ]
        let data = try DocxWriter.writeData(doc)
        XCTAssertGreaterThan(data.count, 0)
    }

    // MARK: - Helpers

    /// Extracts `word/document.xml`'s raw bytes directly from a saved
    /// `.docx`'s ZIP archive — deliberately NOT via `DocxReader.read`
    /// (which has its own separate `MAX_TABLE_NEST_DEPTH = 5` reader-side
    /// guard, unrelated to the writer-side limit this suite tests).
    private func extractRawDocumentXML(from docxData: Data) throws -> String {
        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue196-\(UUID().uuidString).docx")
        try docxData.write(to: docxURL)
        defer { try? FileManager.default.removeItem(at: docxURL) }

        let archive = try Archive(url: docxURL, accessMode: .read)
        let entry = try XCTUnwrap(archive["word/document.xml"])
        var data = Data()
        _ = try archive.extract(entry) { data.append($0) }
        return try XCTUnwrap(String(data: data, encoding: .utf8))
    }
}
