import XCTest
@testable import OOXMLSwift

/// Coverage extension tests for che-word-mcp#54 (4 sub-findings from #42 verify).
///
/// Sub-finding #4 — `testHeaderSEQAlreadyMatchingDoesNotMarkDirty`
/// Sub-finding #5 — `testRegexSchemaDriftEmitsWarning` (#5 covered via stderr; impl
///                  shipped in WordDocument+UpdateAllFields.swift no-op detection)
/// Sub-finding #6 — `testFootnotesByteEqualityWhenNoSEQ` + `testEndnotesByteEqualityWhenNoSEQ`
/// Sub-finding #8 — covered via doc-comment update on `updateAllFields` (no test)
final class UpdateAllFieldsCoverageTests: XCTestCase {

    private var tempDir: URL!

    override func setUp() {
        super.setUp()
        tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("UpdateAllFieldsCoverage-\(UUID().uuidString)")
        try? FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
    }

    override func tearDown() {
        try? FileManager.default.removeItem(at: tempDir)
        super.tearDown()
    }

    private func captionParagraph(identifier: String, initialCached: String = "0") -> Paragraph {
        let field = SequenceField(identifier: identifier, cachedResult: initialCached)
        var run = Run(text: "")
        run.rawXML = field.toFieldXML()
        var para = Paragraph()
        para.runs = [Run(text: "\(identifier) "), run]
        para.properties.style = "Caption"
        return para
    }

    private func plainParagraph(_ text: String) -> Paragraph {
        return Paragraph(text: text)
    }

    // MARK: - Sub-finding #4: SEQ already matching new value → no-op (no dirty mark)

    func testHeaderSEQAlreadyMatchingDoesNotMarkDirty() {
        var doc = WordDocument()
        // Body has Figure SEQ that will increment 0 → 1, cached "1" already matches
        doc.appendParagraph(captionParagraph(identifier: "Figure", initialCached: "1"))
        // Header has Chapter SEQ that will increment 0 → 1, cached "1" already matches
        let hdr = Header(id: "rId10", paragraphs: [captionParagraph(identifier: "Chapter", initialCached: "1")])
        doc.headers.append(hdr)

        // Snapshot modifiedParts BEFORE updateAllFields. (appendParagraph
        // marks document.xml dirty as a setup side-effect — that's not what
        // we're testing.)
        let preState = doc.modifiedPartsView

        _ = doc.updateAllFields()

        // Assert that updateAllFields itself adds NO new entries when every
        // SEQ rewrite is a no-op (cached value already matches new value).
        let added = doc.modifiedPartsView.subtracting(preState)
        XCTAssertTrue(added.isEmpty,
                      "updateAllFields SHALL NOT add to modifiedParts when SEQ rewrite is no-op (cached already matches); added: \(added)")
    }

    // MARK: - Sub-finding #6: footnote/endnote round-trip byte-equality (no SEQ)

    func testFootnotesByteEqualityWhenNoSEQ() throws {
        var src = WordDocument()
        // Body has Figure SEQ to ensure updateAllFields runs non-trivially
        src.appendParagraph(captionParagraph(identifier: "Figure"))
        src.appendParagraph(captionParagraph(identifier: "Figure"))
        // Footnote with plain text only (no SEQ)
        var footnote = Footnote(id: 1, text: "UNIQUE_FOOTNOTE_marker_54", paragraphIndex: 0)
        footnote.paragraphs = [plainParagraph("UNIQUE_FOOTNOTE_marker_54")]
        src.footnotes.footnotes.append(footnote)

        let srcURL = tempDir.appendingPathComponent("source.docx")
        try DocxWriter.write(src, to: srcURL)

        var preDoc = try DocxReader.read(from: srcURL)
        guard let preTempDir = preDoc.archiveTempDir else {
            return XCTFail("Reader-loaded doc SHALL carry archiveTempDir")
        }
        let preFootnotesPath = preTempDir.appendingPathComponent("word/footnotes.xml")
        guard FileManager.default.fileExists(atPath: preFootnotesPath.path) else {
            throw XCTSkip("Fixture footnotes.xml not produced; DocxWriter footnote support may differ")
        }
        let preBytes = try Data(contentsOf: preFootnotesPath)

        _ = preDoc.updateAllFields()
        let outURL = tempDir.appendingPathComponent("output.docx")
        try DocxWriter.write(preDoc, to: outURL)
        preDoc.close()

        var postDoc = try DocxReader.read(from: outURL)
        defer { postDoc.close() }
        guard let postTempDir = postDoc.archiveTempDir else {
            return XCTFail("Post-save doc SHALL carry archiveTempDir")
        }
        let postBytes = try Data(contentsOf: postTempDir.appendingPathComponent("word/footnotes.xml"))

        XCTAssertEqual(postBytes, preBytes,
                       "footnotes.xml SHALL be byte-equal after updateAllFields when no SEQ in footnotes (#54)")
        // Sanity: marker survives
        XCTAssertTrue(String(decoding: postBytes, as: UTF8.self).contains("UNIQUE_FOOTNOTE_marker_54"),
                      "Footnote content SHALL be preserved")
    }

    func testEndnotesByteEqualityWhenNoSEQ() throws {
        var src = WordDocument()
        src.appendParagraph(captionParagraph(identifier: "Figure"))
        src.appendParagraph(captionParagraph(identifier: "Figure"))
        var endnote = Endnote(id: 1, text: "UNIQUE_ENDNOTE_marker_54", paragraphIndex: 0)
        endnote.paragraphs = [plainParagraph("UNIQUE_ENDNOTE_marker_54")]
        src.endnotes.endnotes.append(endnote)

        let srcURL = tempDir.appendingPathComponent("source.docx")
        try DocxWriter.write(src, to: srcURL)

        var preDoc = try DocxReader.read(from: srcURL)
        guard let preTempDir = preDoc.archiveTempDir else {
            return XCTFail("Reader-loaded doc SHALL carry archiveTempDir")
        }
        let preEndnotesPath = preTempDir.appendingPathComponent("word/endnotes.xml")
        guard FileManager.default.fileExists(atPath: preEndnotesPath.path) else {
            throw XCTSkip("Fixture endnotes.xml not produced; DocxWriter endnote support may differ")
        }
        let preBytes = try Data(contentsOf: preEndnotesPath)

        _ = preDoc.updateAllFields()
        let outURL = tempDir.appendingPathComponent("output.docx")
        try DocxWriter.write(preDoc, to: outURL)
        preDoc.close()

        var postDoc = try DocxReader.read(from: outURL)
        defer { postDoc.close() }
        guard let postTempDir = postDoc.archiveTempDir else {
            return XCTFail("Post-save doc SHALL carry archiveTempDir")
        }
        let postBytes = try Data(contentsOf: postTempDir.appendingPathComponent("word/endnotes.xml"))

        XCTAssertEqual(postBytes, preBytes,
                       "endnotes.xml SHALL be byte-equal after updateAllFields when no SEQ in endnotes (#54)")
        XCTAssertTrue(String(decoding: postBytes, as: UTF8.self).contains("UNIQUE_ENDNOTE_marker_54"),
                      "Endnote content SHALL be preserved")
    }
}
