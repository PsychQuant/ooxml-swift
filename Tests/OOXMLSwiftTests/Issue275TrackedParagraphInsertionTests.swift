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
    private func documentXML(withLeadingTable: Bool, bodyXML: String? = nil) -> String {
        if let bodyXML {
            return """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
                \(bodyXML)\
                <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>\
                </w:body></w:document>
                """
        }
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

    private func buildFixture(withLeadingTable: Bool = false, bodyXML: String? = nil) throws -> URL {
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
        try write(documentXML(withLeadingTable: withLeadingTable, bodyXML: bodyXML), to: "word/document.xml")

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

    private func openTracked(withLeadingTable: Bool = false, bodyXML: String? = nil) throws -> (WordDocument, URL) {
        let fixture = try buildFixture(withLeadingTable: withLeadingTable, bodyXML: bodyXML)
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

    // MARK: - Verify round 1 (che-word-mcp#275 round 2)

    /// Saves `doc` and reads the result back as a fresh document.
    private func reopened(_ doc: WordDocument) throws -> (WordDocument, URL) {
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue275-reopen-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: url)
        var again = try DocxReader.read(from: url)
        again.enableTrackChanges(author: "Editor B")
        return (again, url)
    }

    /// The paragraph-mark marker is read back, so a tracked paragraph can be
    /// rejected in a later session (it used to stay as raw XML on the mark).
    func testTrackedParagraphCanBeRejectedAfterSaveAndReopen() throws {
        var (doc, fixture) = try openTracked()
        defer { try? FileManager.default.removeItem(at: fixture); doc.close() }
        let id = try doc.insertParagraphAsRevision(
            Paragraph(text: "Inserted."), at: .afterText("Alpha.", instance: 1, options: .exact))

        var (again, url) = try reopened(doc)
        defer { try? FileManager.default.removeItem(at: url); again.close() }
        let marked = try XCTUnwrap(again.getParagraphs().first { $0.getText() == "Inserted." })
        XCTAssertEqual(marked.markRevisionId, id)

        try again.rejectRevision(revisionId: id)
        XCTAssertEqual(again.getParagraphs().map { $0.getText() }, ["Alpha.", "Omega."])
        XCTAssertFalse(try savedDocumentXML(again).contains("<w:ins"))
    }

    func testTrackedParagraphCanBeAcceptedAfterSaveAndReopen() throws {
        var (doc, fixture) = try openTracked()
        defer { try? FileManager.default.removeItem(at: fixture); doc.close() }
        let id = try doc.insertParagraphAsRevision(
            Paragraph(text: "Inserted."), at: .afterText("Alpha.", instance: 1, options: .exact))

        var (again, url) = try reopened(doc)
        defer { try? FileManager.default.removeItem(at: url); again.close() }
        try again.acceptRevision(revisionId: id)
        XCTAssertEqual(again.getParagraphs().map { $0.getText() }, ["Alpha.", "Inserted.", "Omega."])
        XCTAssertFalse(try savedDocumentXML(again).contains("<w:ins"))
    }

    /// Word gives the mark its own id. Both are read; re-serializing writes
    /// the mark once (not once typed and once from the raw copy).
    func testWordAuthoredParagraphMarkInsertionIsReadAndWrittenOnce() throws {
        let word = """
            <w:p><w:r><w:t>Alpha.</w:t></w:r></w:p>\
            <w:p><w:pPr><w:rPr><w:ins w:id="5" w:author="Reviewer A" w:date="2026-10-01T00:00:00Z"/></w:rPr></w:pPr>\
            <w:ins w:id="6" w:author="Reviewer A" w:date="2026-10-01T00:00:00Z"><w:r><w:t>Word paragraph.</w:t></w:r></w:ins></w:p>\
            <w:p><w:r><w:t>Omega.</w:t></w:r></w:p>
            """
        var (doc, fixture) = try openTracked(bodyXML: word)
        defer { try? FileManager.default.removeItem(at: fixture); doc.close() }
        let marked = try XCTUnwrap(doc.getParagraphs().first { $0.getText() == "Word paragraph." })
        XCTAssertEqual(marked.markRevisionId, 5)
        XCTAssertTrue(doc.revisions.revisions.contains { $0.id == 5 && $0.type == .insertion })

        try doc.updateParagraph(at: 0, text: "ALPHA.")   // force typed re-serialization
        let p = try XCTUnwrap(paragraph(containing: "Word paragraph.", in: try savedDocumentXML(doc)))
        XCTAssertEqual(p.components(separatedBy: "<w:ins w:id=\"5\"").count - 1, 1, "mark written once — got: \(p)")

        try doc.rejectRevision(revisionId: 6)
        try doc.rejectRevision(revisionId: 5)
        XCTAssertEqual(doc.getParagraphs().map { $0.getText() }, ["ALPHA.", "Omega."])
    }

    /// Rejecting never deletes content that is not part of the insertion. A
    /// hyperlink added later would have to merge into the next paragraph,
    /// which is not supported yet: the reject refuses and nothing changes
    /// (round 3 — round 2 kept the paragraph and dropped only the mark, a
    /// silent accept).
    func testRejectRefusesWhenTheParagraphGainedUntrackedNonRunContent() throws {
        var (doc, fixture) = try openTracked()
        defer { try? FileManager.default.removeItem(at: fixture); doc.close() }
        let id = try doc.insertParagraphAsRevision(
            Paragraph(text: "Inserted."), at: .afterText("Alpha.", instance: 1, options: .exact))
        // add a hyperlink afterwards, untracked
        guard let bodyIdx = doc.body.children.firstIndex(where: {
            if case .paragraph(let p) = $0 { return p.markRevisionId == id } else { return false }
        }), case .paragraph(var inserted) = doc.body.children[bodyIdx] else { return XCTFail("tracked paragraph not found") }
        inserted.hyperlinks.append(Hyperlink(id: "h1", text: "LINK", anchor: "target"))
        doc.body.children[bodyIdx] = .paragraph(inserted)

        XCTAssertThrowsError(try doc.rejectRevision(revisionId: id))
        let xml = try savedDocumentXML(doc)
        XCTAssertTrue(xml.contains("LINK"), "untracked hyperlink must survive — got: \(xml)")
        XCTAssertTrue(xml.contains("Inserted."), "nothing changes on a refused reject — got: \(xml)")
        XCTAssertTrue(doc.revisions.revisions.contains { $0.id == id })
    }

    /// A join carries the other revisions' definitions with their runs, and
    /// the joined text comes first in both the model and the saved file.
    func testJoinCarriesOtherRevisionsAndKeepsOrder() throws {
        var (doc, fixture) = try openTracked()
        defer { try? FileManager.default.removeItem(at: fixture); doc.close() }
        let first = try doc.insertParagraphAsRevision(
            Paragraph(text: "Inserted."), at: .afterText("Alpha.", instance: 1, options: .exact))
        doc.enableTrackChanges(author: "Editor C")
        let second = try doc.insertTextAsRevision(text: "More ", atParagraph: 1, position: 0)

        try doc.rejectRevision(revisionId: first)
        XCTAssertEqual(doc.getParagraphs().map { $0.getText() }, ["Alpha.", "More Omega."])
        let omega = try XCTUnwrap(paragraph(containing: "Omega.", in: try savedDocumentXML(doc)))
        XCTAssertTrue(omega.range(of: "<w:ins w:id=\"\(second)\"[^>]*w:author=\"Editor C\"[^>]*><w:r><w:t[^>]*>More </w:t></w:r></w:ins>.*Omega\\.",
                                  options: .regularExpression) != nil,
                      "Editor C's insertion must stay tracked and come first — got: \(omega)")
    }

    /// A crafted id at Int.max must not trap the allocator.
    func testRevisionIdAllocationAtTheCeilingDoesNotTrap() {
        var doc = WordDocument()
        doc.revisions.revisions.append(Revision(id: Int.max, type: .insertion, author: "X"))
        let id = doc.allocateRevisionId()
        XCTAssertGreaterThan(id, 0)
        XCTAssertNotEqual(id, Int.max)
    }

    /// Content that cannot carry a revision id on its runs is refused rather
    /// than inserted half-tracked.
    func testUntaggableContentIsRefused() throws {
        var (doc, fixture) = try openTracked()
        defer { try? FileManager.default.removeItem(at: fixture); doc.close() }
        var para = Paragraph(text: "Inserted.")
        para.hyperlinks.append(Hyperlink(id: "h1", text: "LINK", anchor: "target"))
        XCTAssertThrowsError(try doc.insertParagraphAsRevision(para, at: .afterText("Alpha.", instance: 1, options: .exact)))
        XCTAssertEqual(doc.getParagraphs().map { $0.getText() }, ["Alpha.", "Omega."])
    }

    /// Appending a tracked paragraph keeps the graft path, so the rest of
    /// the document is not re-serialized: a mid-body section break survives.
    func testTrackedAppendDoesNotReSerializeTheRestOfTheDocument() throws {
        let twoSections = """
            <w:p><w:pPr><w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr></w:pPr>\
            <w:r><w:t>Landscape section.</w:t></w:r></w:p>\
            <w:p><w:r><w:t>Portrait section.</w:t></w:r></w:p>
            """
        var (doc, fixture) = try openTracked(bodyXML: twoSections)
        defer { try? FileManager.default.removeItem(at: fixture); doc.close() }
        _ = try doc.insertParagraphAsRevision(Paragraph(text: "Appended."), at: .paragraphIndex(doc.body.children.count))

        let xml = try savedDocumentXML(doc)
        XCTAssertTrue(xml.contains("w:orient=\"landscape\""), "mid-body section break must survive — got: \(xml)")
        let appended = try XCTUnwrap(paragraph(containing: "Appended.", in: xml))
        XCTAssertTrue(appended.contains("<w:ins"), "appended paragraph must be tracked — got: \(appended)")
    }

    // MARK: - Verify round 2 (che-word-mcp#275 round 3)

    private static func wordInsertedParagraph(_ content: String) -> String {
        """
        <w:p><w:r><w:t>Alpha.</w:t></w:r></w:p>\
        <w:p><w:pPr><w:rPr><w:ins w:id="5" w:author="Reviewer A" w:date="2026-10-01T00:00:00Z"/></w:rPr></w:pPr>\
        \(content)</w:p>\
        <w:p><w:r><w:t>Omega.</w:t></w:r></w:p>
        """
    }

    /// Word's own `_GoBack` bookmark (left where the last edit happened) and
    /// proofing markers do not keep a rejected paragraph alive.
    func testRejectAllRemovesAWordParagraphHoldingTheGoBackBookmark() throws {
        let body = Self.wordInsertedParagraph("""
            <w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/>\
            <w:ins w:id="6" w:author="Reviewer A" w:date="2026-10-01T00:00:00Z"><w:r><w:t>Word paragraph.</w:t></w:r></w:ins>
            """)
        var (doc, fixture) = try openTracked(bodyXML: body)
        defer { try? FileManager.default.removeItem(at: fixture); doc.close() }

        try doc.tryRejectAllRevisions()
        XCTAssertEqual(doc.getParagraphs().map { $0.getText() }, ["Alpha.", "Omega."])
    }

    /// Content a rejected mark cannot move (here a comment range) makes the
    /// reject refuse before anything changes — never a silent accept.
    func testRejectingAMarkWhoseParagraphHoldsACommentRangeRefuses() throws {
        let body = Self.wordInsertedParagraph("""
            <w:commentRangeStart w:id="9"/>\
            <w:ins w:id="6" w:author="Reviewer A" w:date="2026-10-01T00:00:00Z"><w:r><w:t>Word paragraph.</w:t></w:r></w:ins>\
            <w:commentRangeEnd w:id="9"/>
            """)
        var (doc, fixture) = try openTracked(bodyXML: body)
        defer { try? FileManager.default.removeItem(at: fixture); doc.close() }

        try doc.rejectRevision(revisionId: 6)
        XCTAssertThrowsError(try doc.rejectRevision(revisionId: 5))
        XCTAssertTrue(doc.revisions.revisions.contains { $0.id == 5 }, "the refused revision stays pending")
        XCTAssertEqual(doc.getParagraphs().count, 3)
    }

    /// A comment added in the session (`commentIds`) is content too.
    func testRejectingATrackedParagraphThatGainedACommentRefuses() throws {
        var (doc, fixture) = try openTracked()
        defer { try? FileManager.default.removeItem(at: fixture); doc.close() }
        let id = try doc.insertParagraphAsRevision(
            Paragraph(text: "Inserted."), at: .afterText("Alpha.", instance: 1, options: .exact))
        _ = try doc.insertComment(text: "Why?", author: "Editor C", paragraphIndex: 1)

        XCTAssertThrowsError(try doc.rejectRevision(revisionId: id))
        XCTAssertTrue(doc.getParagraphs().contains { $0.getText() == "Inserted." })
    }

    /// Reject-all handles a Word paragraph's content before its mark, so a
    /// paragraph whose content is a mixed-content wrapper (a hyperlink inside
    /// `<w:ins>`) is removed, not left as an empty `<w:p>`.
    func testRejectAllRemovesAWordParagraphWhoseContentIsAMixedWrapper() throws {
        let body = Self.wordInsertedParagraph("""
            <w:ins w:id="6" w:author="Reviewer A" w:date="2026-10-01T00:00:00Z"><w:hyperlink w:anchor="x"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:ins>
            """)
        var (doc, fixture) = try openTracked(bodyXML: body)
        defer { try? FileManager.default.removeItem(at: fixture); doc.close() }

        try doc.tryRejectAllRevisions()
        XCTAssertEqual(doc.getParagraphs().count, 2, "got: \(doc.getParagraphs().map { $0.getText() })")
    }

    /// An inserted mark sharing its id with a run deletion is still written
    /// as an inserted mark.
    func testAMarkSharingItsIdWithARunDeletionStaysAnInsertion() throws {
        let body = """
            <w:p><w:r><w:t>Alpha.</w:t></w:r></w:p>\
            <w:p><w:pPr><w:rPr><w:ins w:id="7" w:author="A" w:date="2026-10-01T00:00:00Z"/></w:rPr></w:pPr>\
            <w:del w:id="7" w:author="A" w:date="2026-10-01T00:00:00Z"><w:r><w:delText>x</w:delText></w:r></w:del><w:r><w:t>y</w:t></w:r></w:p>
            """
        var (doc, fixture) = try openTracked(bodyXML: body)
        defer { try? FileManager.default.removeItem(at: fixture); doc.close() }
        try doc.updateParagraph(at: 0, text: "ALPHA.")   // force typed re-serialization

        let p = try XCTUnwrap(paragraph(containing: ">y<", in: try savedDocumentXML(doc)))
        XCTAssertTrue(p.contains("<w:pPr><w:rPr><w:ins w:id=\"7\""), "got: \(p)")
    }

    /// Outside the body (here a header) a rejected mark is resolved the same
    /// way: an emptied paragraph is removed, not left behind.
    func testRejectingAMarkInAHeaderRemovesTheEmptiedParagraph() throws {
        var doc = WordDocument()
        doc.enableTrackChanges(author: "Editor B")
        var marked = Paragraph(text: "Header line.")
        marked.runs[0].revisionId = 11
        marked.markRevisionId = 11
        var rev = Revision(id: 11, type: .insertion, author: "Editor B")
        rev.source = .header(id: "rIdH1")
        marked.revisions.append(rev)
        var header = Header(id: "rIdH1")
        header.bodyChildren = [.paragraph(Paragraph(text: "Keep.")), .paragraph(marked)]
        doc.headers = [header]
        doc.revisions.revisions.append(rev)

        try doc.rejectRevision(revisionId: 11)
        XCTAssertEqual(doc.headers[0].bodyChildren.count, 1)
    }

    func testAnyNonRunContentIsRefused() throws {
        var (doc, fixture) = try openTracked()
        defer { try? FileManager.default.removeItem(at: fixture); doc.close() }
        var para = Paragraph(text: "Inserted.")
        para.commentIds = [3]
        XCTAssertThrowsError(try doc.insertParagraphAsRevision(para, at: .afterText("Alpha.", instance: 1, options: .exact)))
    }

    func testOutOfRangeIndexThrowsLikeTheUntrackedPath() throws {
        var (doc, fixture) = try openTracked()
        defer { try? FileManager.default.removeItem(at: fixture); doc.close() }
        XCTAssertThrowsError(try doc.insertParagraphAsRevision(Paragraph(text: "X"), at: .paragraphIndex(doc.body.children.count + 5)))
    }
}
