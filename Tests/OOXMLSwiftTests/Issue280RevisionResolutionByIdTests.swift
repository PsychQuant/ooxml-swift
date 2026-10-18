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

    private func documentXML(_ paragraphs: [String], rawBody: String? = nil) -> String {
        let body = rawBody ?? paragraphs.map { "<w:p><w:r><w:t xml:space=\"preserve\">\($0)</w:t></w:r></w:p>" }.joined()
        return """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
            \(body)\
            <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>\
            </w:body></w:document>
            """
    }

    private func buildFixture(_ paragraphs: [String], rawBody: String? = nil) throws -> URL {
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
        try write(documentXML(paragraphs, rawBody: rawBody), to: "word/document.xml")

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

    private func openTracked(_ paragraphs: [String] = ["Alpha.", "Beta.", "Gamma."], rawBody: String? = nil) throws -> (WordDocument, URL) {
        let url = try buildFixture(paragraphs, rawBody: rawBody)
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

    // MARK: - Verify round 2 (che-word-mcp#275 round 3)

    /// A revision whose paragraph was deleted has no carrier left. It must
    /// not fall back to matching its text in whatever paragraph now sits at
    /// its recorded index — that deleted text from an untouched paragraph.
    func testAcceptingAnOrphanedDeletionTouchesNoOtherParagraph() throws {
        var (doc, url) = try openTracked(["Beta Alpha.", "Beta paragraph.", "Gamma."])
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        let id = try doc.deleteTextAsRevision(atParagraph: 1, start: 0, end: 4)
        try doc.deleteParagraph(at: 1)
        try doc.acceptRevision(revisionId: id)

        XCTAssertEqual(texts(doc), ["Beta Alpha.", "Gamma."])
        XCTAssertFalse(doc.revisions.revisions.contains { $0.id == id }, "the orphaned entry is removed")
    }

    func testRejectingAnOrphanedInsertionTouchesNoOtherParagraph() throws {
        var (doc, url) = try openTracked(["Alpha.", "Yes. Yes.", "Omega."])
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        let id = try doc.insertParagraphAsRevision(Paragraph(text: "Yes."), at: .afterText("Alpha.", instance: 1))
        try doc.deleteParagraph(at: 1)
        try doc.rejectRevision(revisionId: id)

        XCTAssertEqual(texts(doc), ["Alpha.", "Yes. Yes.", "Omega."])
    }

    /// A move takes the allocated id and the next one; near the ceiling the
    /// fallback must leave both free.
    func testCeilingFallbackLeavesRoomForAMovePair() {
        var doc = WordDocument()
        for id in [2, Int.max] { doc.revisions.revisions.append(Revision(id: id, type: .insertion, author: "X")) }
        let id = doc.allocateRevisionId()
        let used = Set(doc.revisions.revisions.map(\.id))
        XCTAssertFalse(used.contains(id))
        XCTAssertFalse(used.contains(id + 1), "the move pair's second id \(id + 1) collides")
    }

    /// Ids are not unique in every producer's files. One accept must act on
    /// the revision it names — matched by id AND type — not on another
    /// revision that happens to share the id.
    func testDuplicateIdsAreResolvedByType() throws {
        let body = """
            <w:p><w:del w:id="1" w:author="A" w:date="2026-10-01T00:00:00Z"><w:r><w:delText xml:space="preserve">old </w:delText></w:r></w:del><w:r><w:t>kept</w:t></w:r></w:p>\
            <w:p><w:ins w:id="1" w:author="B" w:date="2026-10-01T00:00:00Z"><w:r><w:t>IMPORTANT</w:t></w:r></w:ins></w:p>
            """
        var (doc, url) = try openTracked(rawBody: body)
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        try doc.acceptRevision(revisionId: 1)   // the first entry: the deletion
        XCTAssertEqual(texts(doc), ["kept", "IMPORTANT"])
        let xml = try savedDocumentXML(doc)
        XCTAssertTrue(xml.contains("IMPORTANT"), "B's insertion must survive — got: \(xml)")
        XCTAssertFalse(xml.contains("old"), "A's deletion was accepted — got: \(xml)")
    }
}
