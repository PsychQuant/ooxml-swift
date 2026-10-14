import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Regression tests for [PsychQuant/che-word-mcp#275](https://github.com/PsychQuant/che-word-mcp/issues/275).
///
/// Before this fix there was no tracked way to insert a paragraph: callers
/// reached `insertParagraph(_:at:)`, which inserts the value as given, so a
/// paragraph added to a collaborator's draft under track changes read as the
/// original author's text. Word marks an inserted paragraph twice — `<w:ins>`
/// around its runs AND `<w:ins/>` on the paragraph mark
/// (`<w:pPr><w:rPr><w:ins …/></w:rPr></w:pPr>`); the model had no field for
/// the second.
///
/// After this fix `insertParagraphAsRevision(_:at:author:date:)` records one
/// revision for both, `Paragraph.markRevisionId` carries the paragraph-mark
/// half, and accept / reject handle it.
final class Issue275TrackedParagraphInsertionTests: XCTestCase {

    // MARK: - Fixture

    /// Optional table first (to exercise the flat paragraph index across
    /// table-cell paragraphs), then "Alpha." and "Omega.".
    private func documentXML(withLeadingTable: Bool) -> String {
        let table = withLeadingTable
            ? "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
            : ""
        return """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
            \(table)\
            <w:p><w:pPr><w:ind w:firstLine="720"/></w:pPr><w:r><w:t>Alpha.</w:t></w:r></w:p>\
            <w:p><w:r><w:t>Omega.</w:t></w:r></w:p>\
            <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>\
            </w:body></w:document>
            """
    }

    private func buildFixture(withLeadingTable: Bool = false) throws -> URL {
        let staging = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue275-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: staging, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: staging) }

        func write(_ content: String, to relativePath: String) throws {
            let url = staging.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(), withIntermediateDirectories: true)
            try content.write(to: url, atomically: true, encoding: .utf8)
        }

        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="xml" ContentType="application/xml"/>
                <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>
            """, to: "[Content_Types].xml")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """, to: "_rels/.rels")
        try write(documentXML(withLeadingTable: withLeadingTable), to: "word/document.xml")

        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue275-\(UUID().uuidString).docx")
        let archive = try Archive(url: docxURL, accessMode: .create)
        let base = staging.resolvingSymlinksInPath().path
        let enumerator = FileManager.default.enumerator(
            at: staging, includingPropertiesForKeys: [.isDirectoryKey])!
        for case let fileURL as URL in enumerator {
            let isDir = (try? fileURL.resourceValues(forKeys: [.isDirectoryKey]))?.isDirectory ?? false
            if isDir { continue }
            let entry = String(fileURL.resolvingSymlinksInPath().path.dropFirst(base.count + 1))
            try archive.addEntry(with: entry, fileURL: fileURL, compressionMethod: .deflate)
        }
        return docxURL
    }

    private func savedDocumentXML(_ doc: WordDocument) throws -> String {
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue275-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(doc, to: outURL)
        let archive = try Archive(url: outURL, accessMode: .read)
        let entry = try XCTUnwrap(archive["word/document.xml"])
        var data = Data()
        _ = try archive.extract(entry) { data.append($0) }
        return String(decoding: data, as: UTF8.self)
    }

    private func paragraph(containing marker: String, in xml: String) -> String? {
        let pattern = try! NSRegularExpression(pattern: "<w:p[ >].*?</w:p>", options: [.dotMatchesLineSeparators])
        let ns = xml as NSString
        for match in pattern.matches(in: xml, range: NSRange(location: 0, length: ns.length)) {
            let p = ns.substring(with: match.range)
            if p.contains(marker) { return p }
        }
        return nil
    }

    private func openTracked(withLeadingTable: Bool = false) throws -> (WordDocument, URL) {
        let fixture = try buildFixture(withLeadingTable: withLeadingTable)
        var doc = try DocxReader.read(from: fixture)
        doc.enableTrackChanges(author: "Editor B")
        return (doc, fixture)
    }

    // MARK: - Tests

    func testTrackedInsertionWrapsRunsAndMarksTheParagraphMark() throws {
        var (doc, fixture) = try openTracked()
        defer { try? FileManager.default.removeItem(at: fixture); doc.close() }

        let id = try doc.insertParagraphAsRevision(
            Paragraph(text: "Inserted."), at: .afterText("Alpha.", instance: 1, options: .exact))

        let inserted = try XCTUnwrap(paragraph(containing: "Inserted.", in: try savedDocumentXML(doc)))
        XCTAssertTrue(inserted.contains("<w:pPr><w:rPr><w:ins w:id=\"\(id)\" w:author=\"Editor B\""),
                      "paragraph mark must be marked inserted — got: \(inserted)")
        XCTAssertTrue(inserted.range(of: "<w:ins w:id=\"\(id)\" w:author=\"Editor B\" w:date=\"[^\"]+\"><w:r>",
                                     options: .regularExpression) != nil,
                      "runs must be wrapped in <w:ins> — got: \(inserted)")
    }

    func testTrackedInsertionLandsAfterTheAnchor() throws {
        var (doc, fixture) = try openTracked()
        defer { try? FileManager.default.removeItem(at: fixture); doc.close() }

        _ = try doc.insertParagraphAsRevision(
            Paragraph(text: "Inserted."), at: .afterText("Alpha.", instance: 1, options: .exact))

        let texts = doc.getParagraphs().map { $0.getText() }
        XCTAssertEqual(texts, ["Alpha.", "Inserted.", "Omega."])
    }

    func testTrackedInsertionThrowsWhenTrackChangesIsOff() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        XCTAssertThrowsError(try doc.insertParagraphAsRevision(
            Paragraph(text: "Inserted."), at: .afterText("Alpha.", instance: 1, options: .exact)))
    }

    /// The revision must point at the paragraph it was made in. Flat
    /// paragraph indices count table-cell paragraphs, so with one table
    /// before "Alpha." the inserted paragraph is flat index 2.
    func testRevisionRecordsTheRealFlatParagraphIndex() throws {
        var (doc, fixture) = try openTracked(withLeadingTable: true)
        defer { try? FileManager.default.removeItem(at: fixture); doc.close() }

        let id = try doc.insertParagraphAsRevision(
            Paragraph(text: "Inserted."), at: .afterText("Alpha.", instance: 1, options: .exact))

        let revision = try XCTUnwrap(doc.revisions.revisions.first { $0.id == id })
        XCTAssertEqual(revision.type, .insertion)
        XCTAssertEqual(revision.author, "Editor B")
        XCTAssertEqual(revision.paragraphIndex, 2)
    }

    func testAcceptKeepsTheParagraphAndDropsBothMarkers() throws {
        var (doc, fixture) = try openTracked(withLeadingTable: true)
        defer { try? FileManager.default.removeItem(at: fixture); doc.close() }

        let id = try doc.insertParagraphAsRevision(
            Paragraph(text: "Inserted."), at: .afterText("Alpha.", instance: 1, options: .exact))
        try doc.acceptRevision(revisionId: id)

        let inserted = try XCTUnwrap(paragraph(containing: "Inserted.", in: try savedDocumentXML(doc)))
        XCTAssertFalse(inserted.contains("<w:ins"), "accepted insertion must carry no revision markup — got: \(inserted)")
    }

    func testRejectRemovesTheInsertedParagraph() throws {
        var (doc, fixture) = try openTracked(withLeadingTable: true)
        defer { try? FileManager.default.removeItem(at: fixture); doc.close() }

        let id = try doc.insertParagraphAsRevision(
            Paragraph(text: "Inserted."), at: .afterText("Alpha.", instance: 1, options: .exact))
        try doc.rejectRevision(revisionId: id)

        let xml = try savedDocumentXML(doc)
        XCTAssertNil(paragraph(containing: "Inserted.", in: xml), "rejected paragraph must be gone — got: \(xml)")
        XCTAssertEqual(doc.getParagraphs().map { $0.getText() }.filter { $0 != "cell" }, ["Alpha.", "Omega."])
        XCTAssertFalse(xml.contains("w:author=\"Editor B\""), "no markup of the rejected revision may remain — got: \(xml)")
    }

    /// Rejecting an inserted paragraph MARK while some of the paragraph's
    /// content is not part of that insertion joins the leftover onto the
    /// next paragraph (whose mark survives) — Word's behaviour.
    func testRejectedMarkJoinsLeftoverContentOntoTheNextParagraph() {
        var marked = Paragraph(text: "left ")
        marked.markRevisionId = 9
        var children: [BodyChild] = [.paragraph(marked), .paragraph(Paragraph(text: "next"))]
        var counter = 0

        WordDocument.rejectInsertedParagraphMark(in: &children, target: 0, counter: &counter, revisionId: 9)

        XCTAssertEqual(children.count, 1)
        guard case .paragraph(let joined) = children[0] else { return XCTFail("expected a paragraph") }
        XCTAssertEqual(joined.runs.map(\.text).joined(), "left next")
        XCTAssertNil(joined.markRevisionId)
    }

    /// A table cell must keep at least one paragraph: an emptied, rejected
    /// paragraph that is the cell's only one stays (marker cleared).
    func testRejectedMarkNeverEmptiesATableCell() {
        var marked = Paragraph(text: "")
        marked.markRevisionId = 9
        var table = Table(rowCount: 1, columnCount: 1)
        table.rows[0].cells[0].paragraphs = [marked]
        var children: [BodyChild] = [.table(table)]
        var counter = 0

        WordDocument.rejectInsertedParagraphMark(in: &children, target: 0, counter: &counter, revisionId: 9)

        guard case .table(let after) = children[0] else { return XCTFail("expected the table") }
        XCTAssertEqual(after.rows[0].cells[0].paragraphs.count, 1)
        XCTAssertNil(after.rows[0].cells[0].paragraphs[0].markRevisionId)
    }

    /// Untracked insertion is unchanged: no revision markup, no mark rPr.
    func testPlainInsertParagraphOutputIsUnchanged() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        try doc.insertParagraph(Paragraph(text: "Plain."), at: .afterText("Alpha.", instance: 1, options: .exact))

        let plain = try XCTUnwrap(paragraph(containing: "Plain.", in: try savedDocumentXML(doc)))
        XCTAssertFalse(plain.contains("<w:ins"), "got: \(plain)")
        XCTAssertFalse(plain.contains("<w:pPr>"), "got: \(plain)")
    }
}
