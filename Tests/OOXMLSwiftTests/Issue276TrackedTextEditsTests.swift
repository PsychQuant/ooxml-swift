import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Regression tests for [PsychQuant/che-word-mcp#276](https://github.com/PsychQuant/che-word-mcp/issues/276), Stage 1.
///
/// Before this fix: `insertTextAsRevision` placed the new run correctly in
/// `paragraph.runs`, but a paragraph read from a file is written in source
/// order and the new run had no source position — it was emitted at the end
/// of the paragraph, outside the revision grouping, so it was untracked and
/// the revision vanished on save. Offsets counted deleted text. Edits inside
/// another author's insertion were written wrongly, and there was no tracked
/// replacement at all.
///
/// After this fix (Stage 1): inserted runs are emitted in place inside
/// `<w:ins>`; offsets count inserted text and skip deleted text; edits that
/// would have to nest inside another revision (#293), cross non-text
/// content, or split a run holding non-text children are refused. Tracked
/// replacement was withdrawn from stage 1 (#292).
final class Issue276TrackedTextEditsTests: XCTestCase {

    // MARK: - Fixture

    /// The issue's paragraph: plain text, Reviewer A's insertion, a period.
    static let withInsertion = """
        <w:p><w:r><w:t xml:space="preserve">Data came from Alphaland, Betaland, and Gammaland. </w:t></w:r>\
        <w:ins w:id="1" w:author="Reviewer A" w:date="2026-10-01T00:00:00Z"><w:r><w:t>Table 1, 2, 3, and 4 list the sample sizes</w:t></w:r></w:ins>\
        <w:r><w:t>.</w:t></w:r></w:p>
        """
    /// Same text with Reviewer A's insertion as a plain run (no revisions).
    static let control = """
        <w:p><w:r><w:t xml:space="preserve">Data came from Alphaland, Betaland, and Gammaland. </w:t></w:r>\
        <w:r><w:t>Table 1, 2, 3, and 4 list the sample sizes</w:t></w:r><w:r><w:t>.</w:t></w:r></w:p>
        """

    private func documentXML(_ body: String) -> String {
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
        \(body)\
        <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>\
        </w:body></w:document>
        """
    }

    private func buildFixture(_ body: String) throws -> URL {
        let staging = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue276-staging-\(UUID().uuidString)")
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
        try write(documentXML(body), to: "word/document.xml")

        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue276-\(UUID().uuidString).docx")
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

    private func open(_ body: String) throws -> (WordDocument, URL) {
        let url = try buildFixture(body)
        var doc = try DocxReader.read(from: url)
        doc.enableTrackChanges(author: "Editor B")
        return (doc, url)
    }

    private func savedDocumentXML(_ doc: WordDocument) throws -> String {
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue276-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(doc, to: outURL)
        let archive = try Archive(url: outURL, accessMode: .read)
        let entry = try XCTUnwrap(archive["word/document.xml"])
        var data = Data()
        _ = try archive.extract(entry) { data.append($0) }
        return String(decoding: data, as: UTF8.self)
    }

    private func matches(_ xml: String, _ pattern: String) -> Bool {
        xml.range(of: pattern, options: .regularExpression) != nil
    }

    // MARK: - Case 1: insert_text_as_revision

    func testInsertedTextIsWrittenInPlaceInsideAnInsertion() throws {
        var (doc, url) = try open(Self.control)
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        _ = try doc.insertTextAsRevision(text: "INSERTED ", atParagraph: 0, position: 15)

        let xml = try savedDocumentXML(doc)
        XCTAssertTrue(matches(xml, #"Data came from </w:t></w:r><w:ins w:id="\d+" w:author="Editor B"[^>]*><w:r><w:t[^>]*>INSERTED </w:t></w:r></w:ins><w:r><w:t[^>]*>Alphaland"#),
                      "got: \(xml)")
    }

    /// Position 51 is the boundary before Reviewer A's insertion: allowed,
    /// and the new insertion is a separate `<w:ins>` before it.
    func testInsertAtTheBoundaryOfAnotherAuthorsInsertion() throws {
        var (doc, url) = try open(Self.withInsertion)
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        _ = try doc.insertTextAsRevision(text: "INSERTED ", atParagraph: 0, position: 51)

        let xml = try savedDocumentXML(doc)
        XCTAssertTrue(matches(xml, #"Gammaland\. </w:t></w:r><w:ins w:id="\d+" w:author="Editor B"[^>]*><w:r><w:t[^>]*>INSERTED </w:t></w:r></w:ins><w:ins w:id="1" w:author="Reviewer A""#),
                      "got: \(xml)")
    }

    /// Strictly inside another author's insertion needs nesting (Stage 2):
    /// refused, nothing changed.
    func testInsertInsideAnotherAuthorsInsertionIsRefused() throws {
        var (doc, url) = try open(Self.withInsertion)
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }
        let before = doc.getParagraphs()[0].runs.map(\.text)

        XCTAssertThrowsError(try doc.insertTextAsRevision(text: "X", atParagraph: 0, position: 60))
        XCTAssertEqual(doc.getParagraphs()[0].runs.map(\.text), before)
    }

    // MARK: - Offsets: inserted text counts, deleted text does not

    static let withDeletion = """
        <w:p><w:r><w:t xml:space="preserve">Keep </w:t></w:r>\
        <w:del w:id="3" w:author="Reviewer A" w:date="2026-10-01T00:00:00Z"><w:r><w:delText xml:space="preserve">gone </w:delText></w:r></w:del>\
        <w:r><w:t>tail</w:t></w:r></w:p>
        """

    func testInsertOffsetsSkipDeletedText() throws {
        var (doc, url) = try open(Self.withDeletion)
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        _ = try doc.insertTextAsRevision(text: "NEW ", atParagraph: 0, position: 5)

        let xml = try savedDocumentXML(doc)
        XCTAssertTrue(matches(xml, #"gone </w:delText></w:r></w:del><w:ins [^>]*w:author="Editor B"[^>]*><w:r><w:t[^>]*>NEW </w:t></w:r></w:ins><w:r><w:t[^>]*>tail"#),
                      "visible offset 5 is after the deleted text, before \"tail\" — got: \(xml)")
    }

    func testDeleteOffsetsSkipDeletedText() throws {
        var (doc, url) = try open(Self.withDeletion)
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        _ = try doc.deleteTextAsRevision(atParagraph: 0, start: 5, end: 9)

        let xml = try savedDocumentXML(doc)
        XCTAssertTrue(matches(xml, #"<w:del [^>]*w:author="Editor B"[^>]*><w:r><w:delText[^>]*>tail</w:delText>"#),
                      "visible [5, 9) is \"tail\", not the already-deleted text — got: \(xml)")
    }

    func testDeleteOverlappingAnotherAuthorsInsertionIsRefused() throws {
        var (doc, url) = try open(Self.withInsertion)
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }
        let before = doc.getParagraphs()[0].runs.map(\.text)

        XCTAssertThrowsError(try doc.deleteTextAsRevision(atParagraph: 0, start: 51, end: 71))
        XCTAssertEqual(doc.getParagraphs()[0].runs.map(\.text), before)
    }

    /// The Revision records the paragraph's flat index (table-cell paragraphs
    /// count), not the default 0.
    func testRevisionsRecordTheFlatParagraphIndex() throws {
        let body = "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" + Self.control
        var (doc, url) = try open(body)
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        let ins = try doc.insertTextAsRevision(text: "X ", atParagraph: 0, position: 0)
        let del = try doc.deleteTextAsRevision(atParagraph: 0, start: 2, end: 6)
        XCTAssertEqual(doc.revisions.revisions.first { $0.id == ins }?.paragraphIndex, 1)
        XCTAssertEqual(doc.revisions.revisions.first { $0.id == del }?.paragraphIndex, 1)
    }

    // MARK: - Expected 5: every returned id is in the saved file

    func testEveryReturnedRevisionIdSurvivesSaveAndReopen() throws {
        var (doc, url) = try open(Self.control)
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        let ins = try doc.insertTextAsRevision(text: "NEW ", atParagraph: 0, position: 0)
        let del = try doc.deleteTextAsRevision(atParagraph: 0, start: 4, end: 8)

        let out = FileManager.default.temporaryDirectory.appendingPathComponent("issue276-reopen-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: out) }
        try DocxWriter.write(doc, to: out)
        var again = try DocxReader.read(from: out)
        defer { again.close() }
        let ids = Set(again.revisions.revisions.map(\.id))
        for id in [ins, del] {
            XCTAssertTrue(ids.contains(id), "revision \(id) missing after reopen — have \(ids)")
        }
    }

    // MARK: - Stage 1 round 2 (verify findings)

    /// A complex field (begin / instr / separate / result / end) inside a
    /// deleted span used to lose its structure runs untracked.
    func testDeletingASpanWithAFieldInsideRefuses() throws {
        let body = """
            <w:p><w:r><w:t xml:space="preserve">see Table </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r>\
            <w:r><w:instrText xml:space="preserve"> REF t1 </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>\
            <w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> above</w:t></w:r></w:p>
            """
        var (doc, url) = try open(body)
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }
        let before = doc.getParagraphs()[0].runs.count

        XCTAssertThrowsError(try doc.deleteTextAsRevision(atParagraph: 0, start: 4, end: 11))
        XCTAssertEqual(doc.getParagraphs()[0].runs.count, before)
        XCTAssertTrue(doc.revisions.revisions.isEmpty)
    }

    func testDeletingAcrossAFootnoteReferenceRefuses() throws {
        let body = """
            <w:p><w:r><w:t>the effect.</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>\
            <w:r><w:t xml:space="preserve"> However</w:t></w:r></w:p>
            """
        var (doc, url) = try open(body)
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        XCTAssertThrowsError(try doc.deleteTextAsRevision(atParagraph: 0, start: 4, end: 15))
        XCTAssertTrue(try savedDocumentXML(doc).contains("w:footnoteReference"))
    }

    /// A zero-length run at the edge of the span is outside it: kept.
    func testAZeroLengthRunAtTheSpanEdgeIsKept() throws {
        let body = "<w:p><w:r><w:t>Gone</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>kept</w:t></w:r></w:p>"
        var (doc, url) = try open(body)
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        _ = try doc.deleteTextAsRevision(atParagraph: 0, start: 0, end: 4)
        XCTAssertTrue(matches(try savedDocumentXML(doc), #"<w:r><w:tab(/>|></w:tab>)</w:r><w:r><w:t[^>]*>kept"#))
    }

    func testAnEmptySpanIsRefused() throws {
        var (doc, url) = try open(Self.control)
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }
        XCTAssertThrowsError(try doc.deleteTextAsRevision(atParagraph: 0, start: 3, end: 3))
        XCTAssertTrue(doc.revisions.revisions.isEmpty)
    }

    /// Splitting a run that also holds a tab / break would copy those into
    /// every piece; refused.
    func testSplittingARunWithNonTextChildrenIsRefused() throws {
        let body = "<w:p><w:r><w:t>ab</w:t><w:tab/><w:t>cd</w:t></w:r></w:p>"
        var (doc, url) = try open(body)
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        XCTAssertThrowsError(try doc.insertTextAsRevision(text: "X", atParagraph: 0, position: 1))
        XCTAssertThrowsError(try doc.deleteTextAsRevision(atParagraph: 0, start: 1, end: 2))
        let xml = try savedDocumentXML(doc)
        XCTAssertEqual(xml.components(separatedBy: "<w:tab").count - 1, 1, "got: \(xml)")
    }

    /// A paragraph with no runs (only a bookmark) and a run without a source
    /// position: the new text is still written tracked.
    func testInsertIsTrackedInAParagraphWithoutRuns() throws {
        let body = #"<w:p><w:bookmarkStart w:id="0" w:name="here"/><w:bookmarkEnd w:id="0"/></w:p>"#
        var (doc, url) = try open(body)
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        _ = try doc.insertTextAsRevision(text: "New", atParagraph: 0, position: 0)
        XCTAssertTrue(matches(try savedDocumentXML(doc), #"<w:ins [^>]*w:author="Editor B"[^>]*><w:r><w:t[^>]*>New</w:t></w:r></w:ins>"#))
    }

    func testInsertNextToAnUnpositionedRunIsTracked() throws {
        var (doc, url) = try open(Self.control)
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }
        guard case .paragraph(var para) = doc.body.children[0] else { return XCTFail() }
        para.runs.append(Run(text: " Tail."))   // API-added, no source position
        doc.body.children[0] = .paragraph(para)

        let total = doc.getParagraphs()[0].getText().count
        _ = try doc.insertTextAsRevision(text: "END", atParagraph: 0, position: total)
        XCTAssertTrue(matches(try savedDocumentXML(doc), #"<w:ins [^>]*w:author="Editor B"[^>]*><w:r><w:t[^>]*>END</w:t></w:r></w:ins>"#))
    }

    /// The read tools show the accept-all view in the session too, so the
    /// offsets a caller reads match the offsets the generators use.
    func testReadToolsHideTextDeletedInTheSession() throws {
        var (doc, url) = try open(Self.control)
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        _ = try doc.deleteTextAsRevision(atParagraph: 0, start: 0, end: 5)   // "Data "
        XCTAssertTrue(doc.getParagraphs()[0].getText().hasPrefix("came from"), "got: \(doc.getParagraphs()[0].getText())")
    }

    /// Text inside a mixed-content revision wrapper (captured raw) is
    /// invisible to offsets; edits in such a paragraph are refused.
    func testAParagraphWithAMixedRevisionWrapperRefusesEdits() throws {
        let body = """
            <w:p><w:r><w:t xml:space="preserve">Start </w:t></w:r>\
            <w:ins w:id="4" w:author="Reviewer A" w:date="2026-10-01T00:00:00Z"><w:hyperlink w:anchor="x"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:ins>\
            <w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p>
            """
        var (doc, url) = try open(body)
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        XCTAssertThrowsError(try doc.insertTextAsRevision(text: "X", atParagraph: 0, position: 2))
        XCTAssertThrowsError(try doc.deleteTextAsRevision(atParagraph: 0, start: 0, end: 2))
    }

    /// Word caches page breaks in runs (`<w:lastRenderedPageBreak/>`); that
    /// is not content, so it does not block a split, and only the first
    /// piece keeps it.
    func testARenderingHintDoesNotBlockASplit() throws {
        let body = "<w:p><w:r><w:lastRenderedPageBreak/><w:t>abcd</w:t></w:r></w:p>"
        var (doc, url) = try open(body)
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        _ = try doc.insertTextAsRevision(text: "X", atParagraph: 0, position: 2)
        let xml = try savedDocumentXML(doc)
        XCTAssertEqual(xml.components(separatedBy: "<w:lastRenderedPageBreak").count - 1, 1, "got: \(xml)")
        XCTAssertTrue(matches(xml, #"ab</w:t>.*</w:r><w:ins [^>]*><w:r><w:t[^>]*>X</w:t></w:r></w:ins><w:r><w:t[^>]*>cd"#), "got: \(xml)")
    }

    /// Text inserted right after a field result goes after the field end —
    /// inside the result it would be wiped when the field updates.
    func testInsertAfterAFieldResultLandsAfterTheField() throws {
        let body = """
            <w:p><w:r><w:t xml:space="preserve">Table </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r>\
            <w:r><w:instrText xml:space="preserve"> REF t1 </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>\
            <w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
            """
        var (doc, url) = try open(body)
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        _ = try doc.insertTextAsRevision(text: " shows", atParagraph: 0, position: 7)
        XCTAssertTrue(matches(try savedDocumentXML(doc), #"fldCharType="end"(/>|></w:fldChar>)</w:r><w:ins [^>]*><w:r><w:t[^>]*> shows</w:t>"#))
    }

    /// An inline equation counts as its visible text, as in the read tools:
    /// an offset after it lands after it.
    func testAnInlineEquationCountsTowardOffsets() throws {
        let body = """
            <w:p><w:r><w:t xml:space="preserve">see </w:t></w:r>\
            <w:r><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath></w:r>\
            <w:r><w:t xml:space="preserve"> here</w:t></w:r></w:p>
            """
        var (doc, url) = try open(body)
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }
        let text = doc.getParagraphs()[0].getText()
        XCTAssertEqual(text, "see x here")

        _ = try doc.insertTextAsRevision(text: " also", atParagraph: 0, position: 5)
        let xml = try savedDocumentXML(doc)
        XCTAssertTrue(matches(xml, #"</m:oMath>(</w:r>)?<w:ins [^>]*><w:r><w:t[^>]*> also</w:t></w:r></w:ins><w:r><w:t[^>]*> here"#), "got: \(xml)")
    }

    // MARK: - move_text_as_revision on the same rules

    func testMoveIsWrittenInPlaceAndItsIdsSurviveAReopen() throws {
        let body = "<w:p><w:r><w:t xml:space=\"preserve\">Move this.</w:t></w:r></w:p><w:p><w:r><w:t xml:space=\"preserve\">Target here</w:t></w:r></w:p>"
        var (doc, url) = try open(body)
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        let ids = try doc.moveTextAsRevision(fromParagraph: 0, fromStart: 0, fromEnd: 4, toParagraph: 1, toPosition: 7)
        let xml = try savedDocumentXML(doc)
        XCTAssertTrue(matches(xml, #"Target </w:t></w:r><w:moveTo [^>]*><w:r><w:t[^>]*>Move</w:t></w:r></w:moveTo><w:r><w:t[^>]*>here"#), "got: \(xml)")

        let out = FileManager.default.temporaryDirectory.appendingPathComponent("issue276-move-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: out) }
        try DocxWriter.write(doc, to: out)
        var again = try DocxReader.read(from: out)
        defer { again.close() }
        let have = Set(again.revisions.revisions.map(\.id))
        XCTAssertTrue(have.contains(ids.fromId) && have.contains(ids.toId), "have \(have)")
    }

    func testMovedTextKeepsItsFormatting() throws {
        let body = "<w:p><w:r><w:rPr><w:i/></w:rPr><w:t xml:space=\"preserve\">Move this.</w:t></w:r></w:p><w:p><w:r><w:t>Target</w:t></w:r></w:p>"
        var (doc, url) = try open(body)
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        _ = try doc.moveTextAsRevision(fromParagraph: 0, fromStart: 0, fromEnd: 4, toParagraph: 1, toPosition: 0)
        XCTAssertTrue(matches(try savedDocumentXML(doc), #"<w:moveTo [^>]*><w:r><w:rPr><w:i/></w:rPr><w:t[^>]*>Move</w:t>"#))
    }

    func testMoveOutOfAnotherAuthorsInsertionIsRefused() throws {
        var (doc, url) = try open(Self.withInsertion + "<w:p><w:r><w:t>Target</w:t></w:r></w:p>")
        defer { doc.close(); try? FileManager.default.removeItem(at: url) }

        XCTAssertThrowsError(try doc.moveTextAsRevision(fromParagraph: 0, fromStart: 51, fromEnd: 58, toParagraph: 1, toPosition: 0))
        XCTAssertEqual(doc.revisions.revisions.count, 1)
    }
}
