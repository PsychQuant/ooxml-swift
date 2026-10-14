import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Regression tests for [PsychQuant/che-word-mcp#280](https://github.com/PsychQuant/che-word-mcp/issues/280)
/// (folded into che-word-mcp#275 round 2).
///
/// Before this fix: `acceptRevision` / `rejectRevision` found their target
/// through `Revision.paragraphIndex` — recorded once, never updated, and
/// left at 0 by `insertTextAsRevision` / `deleteTextAsRevision` — and removed
/// text by matching the revision's text snapshot against every run of that
/// paragraph. Any paragraph inserted or deleted above a revision, or any edit
/// after it was made, sent accept / reject to the wrong paragraph or the
/// wrong text.
///
/// After this fix: they find every paragraph that carries the revision id
/// (runs, paragraph mark, paragraph-format change) and remove runs by id.
final class Issue280RevisionResolutionByIdTests: XCTestCase {

    // MARK: - Fixture

    private func documentXML(_ paragraphs: [String]) -> String {
        let body = paragraphs.map { "<w:p><w:r><w:t xml:space=\"preserve\">\($0)</w:t></w:r></w:p>" }.joined()
        return """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
            \(body)\
            <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>\
            </w:body></w:document>
            """
    }

    private func buildFixture(_ paragraphs: [String]) throws -> URL {
        let staging = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue280-staging-\(UUID().uuidString)")
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
        try write(documentXML(paragraphs), to: "word/document.xml")

        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue280-\(UUID().uuidString).docx")
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

    private func openTracked(_ paragraphs: [String] = ["Alpha.", "Beta.", "Gamma."]) throws -> (WordDocument, URL) {
        let url = try buildFixture(paragraphs)
        var doc = try DocxReader.read(from: url)
        doc.enableTrackChanges(author: "Editor B")
        return (doc, url)
    }

    private func savedDocumentXML(_ doc: WordDocument) throws -> String {
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue280-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(doc, to: outURL)
        let archive = try Archive(url: outURL, accessMode: .read)
        let entry = try XCTUnwrap(archive["word/document.xml"])
        var data = Data()
        _ = try archive.extract(entry) { data.append($0) }
        return String(decoding: data, as: UTF8.self)
    }

    private func texts(_ doc: WordDocument) -> [String] {
        doc.getParagraphs().map { $0.getText() }
    }

    // MARK: - Tracked paragraph insertions (the #275 shape)

    /// A second tracked paragraph above the first one used to leave the
    /// first revision's stored index pointing at an original paragraph.
    func testRejectAfterAnotherTrackedInsertAboveTargetsTheRightParagraph() throws {
        var (doc, url) = try openTracked()
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        let first = try doc.insertParagraphAsRevision(Paragraph(text: "Inserted one."), at: .afterText("Gamma.", instance: 1))
        _ = try doc.insertParagraphAsRevision(Paragraph(text: "Inserted two."), at: .afterText("Alpha.", instance: 1))
        try doc.rejectRevision(revisionId: first)

        XCTAssertEqual(texts(doc), ["Alpha.", "Inserted two.", "Beta.", "Gamma."])
        let xml = try savedDocumentXML(doc)
        XCTAssertFalse(xml.contains("Inserted one."), "rejected paragraph must be gone — got: \(xml)")
        XCTAssertTrue(xml.contains(">Gamma.</w:t>"), "original paragraph must be untouched — got: \(xml)")
    }

    func testAcceptAllAfterAShiftLeavesNoRevisionMarkup() throws {
        var (doc, url) = try openTracked()
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        _ = try doc.insertParagraphAsRevision(Paragraph(text: "Inserted one."), at: .afterText("Gamma.", instance: 1))
        _ = try doc.insertParagraphAsRevision(Paragraph(text: "Inserted two."), at: .afterText("Alpha.", instance: 1))
        try doc.tryAcceptAllRevisions()

        XCTAssertEqual(texts(doc), ["Alpha.", "Inserted two.", "Beta.", "Gamma.", "Inserted one."])
        let xml = try savedDocumentXML(doc)
        XCTAssertFalse(xml.contains("<w:ins"), "every insertion was accepted — got: \(xml)")
    }

    /// Rejection removes the runs carrying the id. Matching the text snapshot
    /// failed as soon as the paragraph had more than one run.
    func testRejectRemovesTheInsertedRunsByIdNotByText() throws {
        var (doc, url) = try openTracked()
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        let id = try doc.insertParagraphAsRevision(
            Paragraph(runs: [Run(text: "Hello "), Run(text: "world")]), at: .afterText("Alpha.", instance: 1))
        try doc.rejectRevision(revisionId: id)

        XCTAssertEqual(texts(doc), ["Alpha.", "Beta.", "Gamma."])
    }

    // MARK: - Text revisions (the original #280 shape)

    /// `insertTextAsRevision` used to record paragraph 0; after any paragraph
    /// insert above, reject also missed. The revision's own runs say where it is.
    func testRejectTextInsertionAfterAParagraphIsInsertedAbove() throws {
        var (doc, url) = try openTracked()
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        let id = try doc.insertTextAsRevision(text: "NEW ", atParagraph: 2, position: 0)
        try doc.insertParagraph(Paragraph(text: "Plain."), at: .paragraphIndex(0))
        try doc.rejectRevision(revisionId: id)

        XCTAssertEqual(texts(doc), ["Plain.", "Alpha.", "Beta.", "Gamma."])
    }

    /// Accepting a deletion removes the deleted runs — not every occurrence
    /// of the deleted text in the paragraph.
    func testAcceptDeletionRemovesOnlyTheDeletedRuns() throws {
        var (doc, url) = try openTracked(["the cat and the dog"])
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        let id = try doc.deleteTextAsRevision(atParagraph: 0, start: 0, end: 4)
        try doc.acceptRevision(revisionId: id)

        XCTAssertEqual(texts(doc), ["cat and the dog"])
    }

    func testAcceptDeletionInALaterParagraphAfterAParagraphIsInsertedAbove() throws {
        var (doc, url) = try openTracked()
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        let id = try doc.deleteTextAsRevision(atParagraph: 2, start: 0, end: 3)
        try doc.insertParagraph(Paragraph(text: "Plain."), at: .paragraphIndex(0))
        try doc.acceptRevision(revisionId: id)

        XCTAssertEqual(texts(doc), ["Plain.", "Alpha.", "Beta.", "ma."])
    }
}
