import XCTest
@testable import OOXMLSwift

/// Regression test for che-word-mcp#42 (P0 silent data loss).
///
/// `updateAllFields` previously inserted EVERY header/footer/footnote/endnote
/// path into `modifiedParts` regardless of whether any SEQ field was actually
/// found there. In overlay mode that triggers `Header.toXML()` re-emission,
/// which only knows about typed `paragraphs[]` and silently strips VML
/// watermarks, drawings, and any non-paragraph raw XML.
///
/// Fix: track per-container dirty bits during the scan; only mark
/// `modifiedParts` for containers where a SEQ rewrite actually occurred.
final class WordDocumentUpdateAllFieldsHeaderPreservationTests: XCTestCase {

    private func captionParagraph(identifier: String, initialCached: String = "0") -> Paragraph {
        // Default initialCached = "0" so any successful counter increment
        // (which starts at 1) causes a real string change in rewriteCachedResult.
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

    // MARK: - Scenario 1: header WITHOUT SEQ stays clean

    func testHeaderWithoutSEQNotMarkedDirty() {
        var doc = WordDocument()
        // Body has a SEQ that needs updating.
        doc.appendParagraph(captionParagraph(identifier: "Figure"))
        // Headers carry only paragraphs, no SEQ.
        let hdr = Header(id: "rId10", paragraphs: [plainParagraph("page header")])
        doc.headers.append(hdr)
        XCTAssertEqual(hdr.fileName, "header1.xml", "default-type header SHALL fileName == header1.xml")

        _ = doc.updateAllFields()

        XCTAssertTrue(
            doc.modifiedPartsView.contains("word/document.xml"),
            "document.xml SHALL be marked dirty (body had a SEQ rewrite)"
        )
        XCTAssertFalse(
            doc.modifiedPartsView.contains("word/header1.xml"),
            "header1.xml SHALL NOT be marked dirty when no SEQ field is present in any header paragraph (regression for #42)"
        )
    }

    // MARK: - Scenario 2: header WITH SEQ does mark itself dirty

    func testHeaderWithSEQIsMarkedDirty() {
        var doc = WordDocument()
        doc.appendParagraph(plainParagraph("body without SEQ"))
        // Header contains a SEQ-bearing paragraph (e.g., chapter caption running header).
        let hdr = Header(id: "rId10", paragraphs: [captionParagraph(identifier: "Chapter")])
        doc.headers.append(hdr)

        _ = doc.updateAllFields()

        XCTAssertTrue(
            doc.modifiedPartsView.contains("word/header1.xml"),
            "header1.xml SHALL be marked dirty when it contains a SEQ field whose cached result was rewritten"
        )
    }

    // MARK: - Scenario 3: footer dirty-bit logic mirrors header

    func testFooterWithoutSEQNotMarkedDirty() {
        var doc = WordDocument()
        doc.appendParagraph(captionParagraph(identifier: "Figure"))
        let ftr = Footer(id: "rId20", paragraphs: [plainParagraph("page footer")])
        doc.footers.append(ftr)
        XCTAssertEqual(ftr.fileName, "footer1.xml", "default-type footer SHALL fileName == footer1.xml")

        _ = doc.updateAllFields()

        XCTAssertFalse(
            doc.modifiedPartsView.contains("word/footer1.xml"),
            "footer1.xml SHALL NOT be marked dirty when no SEQ field is present in any footer paragraph"
        )
    }

    // MARK: - Scenario 4: complete no-op when no SEQ anywhere

    func testUpdateAllFieldsNoSEQAnywhereDoesNotAddToModifiedParts() {
        var doc = WordDocument()
        doc.appendParagraph(plainParagraph("p1"))
        doc.appendParagraph(plainParagraph("p2"))
        let hdr = Header(id: "rId10", paragraphs: [plainParagraph("h")])
        doc.headers.append(hdr)
        let ftr = Footer(id: "rId20", paragraphs: [plainParagraph("f")])
        doc.footers.append(ftr)

        // Snapshot modifiedParts BEFORE updateAllFields. (appendParagraph
        // already inserted document.xml — that's not the bug we're testing.)
        let preState = doc.modifiedPartsView

        let counters = doc.updateAllFields()

        XCTAssertTrue(counters.isEmpty, "Counters SHALL be empty when no SEQ field exists anywhere")

        // Critical regression assertion: updateAllFields itself must not add
        // anything to modifiedParts when no SEQ rewrite happened.
        let added = doc.modifiedPartsView.subtracting(preState)
        XCTAssertTrue(
            added.isEmpty,
            "updateAllFields SHALL NOT add any new entries to modifiedParts in a true no-op (no SEQ anywhere); added: \(added)"
        )
    }

    // MARK: - Scenario 5 (verify-driven): end-to-end byte-equality round-trip
    //
    // Per #42 Acceptance: open → update_all_fields → save → re-read; assert
    // header XML byte-equal to source. The dirty-bit unit tests above prove
    // modifiedPartsView correctness (the proxy); this test proves the proxy
    // actually implies byte-equality through the DocxWriter overlay path.

    func testHeaderByteEqualityAfterUpdateAllFieldsRoundTrip() throws {
        // Build a fixture doc with a header carrying ONLY plain paragraphs (no SEQ).
        // The body has SEQ to ensure updateAllFields runs non-trivially.
        var src = WordDocument()
        src.appendParagraph(captionParagraph(identifier: "Figure"))
        src.appendParagraph(captionParagraph(identifier: "Figure"))
        let hdr = Header(id: "rId10", paragraphs: [plainParagraph("UNIQUE_HEADER_TEXT_marker_42")])
        src.headers.append(hdr)
        let ftr = Footer(id: "rId20", paragraphs: [plainParagraph("UNIQUE_FOOTER_TEXT_marker_42")])
        src.footers.append(ftr)

        let tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("UpdateAllFieldsRoundTrip-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: tempDir) }

        let srcURL = tempDir.appendingPathComponent("source.docx")
        try DocxWriter.write(src, to: srcURL)

        // Capture original header/footer bytes (via reader-loaded archiveTempDir).
        var preDoc = try DocxReader.read(from: srcURL)
        guard let preTempDir = preDoc.archiveTempDir else {
            return XCTFail("Reader-loaded doc SHALL carry archiveTempDir")
        }
        let preHeaderBytes = try Data(contentsOf: preTempDir.appendingPathComponent("word/header1.xml"))
        let preFooterBytes = try Data(contentsOf: preTempDir.appendingPathComponent("word/footer1.xml"))
        XCTAssertGreaterThan(preHeaderBytes.count, 0, "Pre-condition: header1.xml must exist with non-zero bytes")

        // Run the workflow that previously stripped headers: updateAllFields → save → re-read.
        _ = preDoc.updateAllFields()
        let outURL = tempDir.appendingPathComponent("output.docx")
        try DocxWriter.write(preDoc, to: outURL)
        preDoc.close()

        var postDoc = try DocxReader.read(from: outURL)
        defer { postDoc.close() }
        guard let postTempDir = postDoc.archiveTempDir else {
            return XCTFail("Post-save doc SHALL carry archiveTempDir")
        }
        let postHeaderBytes = try Data(contentsOf: postTempDir.appendingPathComponent("word/header1.xml"))
        let postFooterBytes = try Data(contentsOf: postTempDir.appendingPathComponent("word/footer1.xml"))

        // Critical acceptance assertion (issue #42):
        //   inputHeader.bytes == outputHeader.bytes after open → updateAllFields → save
        XCTAssertEqual(
            postHeaderBytes, preHeaderBytes,
            "header1.xml SHALL be byte-equal after updateAllFields round-trip when header has no SEQ (closes #42 acceptance)"
        )
        XCTAssertEqual(
            postFooterBytes, preFooterBytes,
            "footer1.xml SHALL be byte-equal after updateAllFields round-trip when footer has no SEQ"
        )
        // Sanity: the unique marker text survives (not stripped).
        let postHeaderString = String(decoding: postHeaderBytes, as: UTF8.self)
        XCTAssertTrue(
            postHeaderString.contains("UNIQUE_HEADER_TEXT_marker_42"),
            "Header content SHALL be preserved (not stripped to <w:p/> stub)"
        )
    }
}
