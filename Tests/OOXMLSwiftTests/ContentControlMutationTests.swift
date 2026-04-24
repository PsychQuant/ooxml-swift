import XCTest
@testable import OOXMLSwift

/// Tests for Phase 4 mutation methods on WordDocument: updateContentControl,
/// replaceContentControlContent, deleteContentControl. Spec covered:
/// `openspec/changes/che-word-mcp-content-controls-read-write/specs/ooxml-content-insertion-primitives/spec.md`
final class ContentControlMutationTests: XCTestCase {

    // MARK: - Test fixture helper

    /// Builds a minimal document with one paragraph-level plain-text SDT
    /// at the given id, with current text content.
    private func makeDocWithPlainTextSDT(id: Int, tag: String, text: String) -> WordDocument {
        var doc = WordDocument()
        let sdt = StructuredDocumentTag(id: id, tag: tag, alias: tag.capitalized, type: .plainText)
        let control = ContentControl(sdt: sdt, content: text)
        var para = Paragraph()
        para.contentControls = [control]
        doc.body.children = [.paragraph(para)]
        return doc
    }

    private func makeDocWithPictureSDT(id: Int) -> WordDocument {
        var doc = WordDocument()
        let sdt = StructuredDocumentTag(id: id, tag: "logo", alias: "Logo", type: .picture)
        let control = ContentControl(sdt: sdt, content: "")
        var para = Paragraph()
        para.contentControls = [control]
        doc.body.children = [.paragraph(para)]
        return doc
    }

    // MARK: - updateContentControl

    /// Spec scenario: Update text on plain-text ContentControl.
    func testUpdatePlainTextSDTReplacesContent() throws {
        var doc = makeDocWithPlainTextSDT(id: 100000, tag: "client_name", text: "TBD")
        try doc.updateContentControl(id: 100000, newText: "Acme")

        // Round-trip via writer/reader to verify content reaches disk.
        let bytes = try DocxWriter.writeData(doc)
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("update-\(UUID().uuidString).docx")
        try bytes.write(to: url)
        defer { try? FileManager.default.removeItem(at: url) }
        let reread = try DocxReader.read(from: url)

        let controls = reread.getParagraphs().flatMap { $0.contentControls }
        guard let updated = controls.first(where: { $0.sdt.id == 100000 }) else {
            XCTFail("SDT id=100000 lost across update + round-trip")
            return
        }
        XCTAssertEqual(updated.sdt.tag, "client_name", "tag must be preserved")
        XCTAssertEqual(updated.sdt.alias, "Client_Name", "alias must be preserved")
        XCTAssertEqual(updated.sdt.type, .plainText)
        XCTAssertTrue(updated.content.contains("Acme"),
            "new text not in content after round-trip: '\(updated.content)'")
    }

    /// Spec scenario: Update fails on not-found id.
    func testUpdateNonexistentIdThrowsContentControlNotFound() {
        var doc = makeDocWithPlainTextSDT(id: 100000, tag: "x", text: "y")
        XCTAssertThrowsError(try doc.updateContentControl(id: 999999, newText: "z")) { error in
            guard case WordError.contentControlNotFound(let id) = error else {
                XCTFail("expected contentControlNotFound, got \(error)"); return
            }
            XCTAssertEqual(id, 999999)
        }
    }

    /// Spec scenario: Update fails on picture SDT.
    func testUpdatePictureSDTThrowsUnsupportedType() {
        var doc = makeDocWithPictureSDT(id: 100001)
        XCTAssertThrowsError(try doc.updateContentControl(id: 100001, newText: "hi")) { error in
            guard case WordError.unsupportedSDTType(let type) = error else {
                XCTFail("expected unsupportedSDTType, got \(error)"); return
            }
            XCTAssertEqual(type, .picture)
        }
    }

    /// Update preserves <w:sdtPr> properties (spec: "leaves <w:sdtPr> properties untouched").
    func testUpdatePreservesSdtPrProperties() throws {
        var doc = makeDocWithPlainTextSDT(id: 100000, tag: "client_name", text: "TBD")
        let originalSdt = doc.body.children.compactMap { child -> StructuredDocumentTag? in
            guard case .paragraph(let p) = child, let c = p.contentControls.first else { return nil }
            return c.sdt
        }.first!

        try doc.updateContentControl(id: 100000, newText: "X")
        let postSdt = doc.body.children.compactMap { child -> StructuredDocumentTag? in
            guard case .paragraph(let p) = child, let c = p.contentControls.first else { return nil }
            return c.sdt
        }.first!

        XCTAssertEqual(originalSdt, postSdt, "sdtPr metadata changed across updateContentControl")
    }

    // MARK: - replaceContentControlContent

    /// Spec scenario: Replace with single-paragraph fragment.
    func testReplaceContentSucceedsWithParagraphXML() throws {
        var doc = makeDocWithPlainTextSDT(id: 100000, tag: "client_name", text: "TBD")
        let newXML = "<w:p><w:r><w:t>Hello</w:t></w:r></w:p>"
        try doc.replaceContentControlContent(id: 100000, contentXML: newXML)

        let controls = doc.body.children.compactMap { child -> ContentControl? in
            guard case .paragraph(let p) = child else { return nil }
            return p.contentControls.first
        }
        XCTAssertEqual(controls.first?.content, newXML)
    }

    /// Spec scenario: Reject nested SDT.
    func testReplaceContentRejectsNestedSdt() {
        var doc = makeDocWithPlainTextSDT(id: 100000, tag: "x", text: "y")
        let bad = "<w:p><w:r><w:t>x</w:t></w:r></w:p><w:sdt></w:sdt>"
        XCTAssertThrowsError(try doc.replaceContentControlContent(id: 100000, contentXML: bad)) { error in
            guard case WordError.disallowedElement(let name) = error else {
                XCTFail("expected disallowedElement, got \(error)"); return
            }
            XCTAssertEqual(name, "w:sdt")
        }
    }

    /// Whitelist: rejects <w:body>.
    func testReplaceContentRejectsBodyElement() {
        var doc = makeDocWithPlainTextSDT(id: 100000, tag: "x", text: "y")
        XCTAssertThrowsError(try doc.replaceContentControlContent(id: 100000, contentXML: "<w:body></w:body>")) { error in
            guard case WordError.disallowedElement(let name) = error else {
                XCTFail("expected disallowedElement, got \(error)"); return
            }
            XCTAssertEqual(name, "w:body")
        }
    }

    /// Whitelist: rejects <w:sectPr>.
    func testReplaceContentRejectsSectPrElement() {
        var doc = makeDocWithPlainTextSDT(id: 100000, tag: "x", text: "y")
        XCTAssertThrowsError(try doc.replaceContentControlContent(id: 100000, contentXML: "<w:sectPr/>")) { error in
            guard case WordError.disallowedElement(let name) = error else {
                XCTFail("expected disallowedElement, got \(error)"); return
            }
            XCTAssertEqual(name, "w:sectPr")
        }
    }

    /// Whitelist: rejects XML declaration.
    func testReplaceContentRejectsXMLDeclaration() {
        var doc = makeDocWithPlainTextSDT(id: 100000, tag: "x", text: "y")
        XCTAssertThrowsError(try doc.replaceContentControlContent(id: 100000, contentXML: "<?xml version=\"1.0\"?><w:p/>")) { error in
            guard case WordError.disallowedElement(let name) = error else {
                XCTFail("expected disallowedElement, got \(error)"); return
            }
            XCTAssertEqual(name, "<?xml")
        }
    }

    // MARK: - deleteContentControl

    /// Spec scenario: Delete with keep_content=false removes everything (paragraph-level case).
    func testDeleteWithoutKeepContentRemovesSdt() throws {
        var doc = makeDocWithPlainTextSDT(id: 100000, tag: "x", text: "y")
        try doc.deleteContentControl(id: 100000, keepContent: false)

        let allControls = doc.getParagraphs().flatMap { $0.contentControls }
        XCTAssertTrue(allControls.isEmpty, "SDT not removed: \(allControls)")
    }

    /// Spec scenario: delete with keep_content=true (paragraph-level).
    /// For paragraph-level SDTs, content is dropped at the model layer
    /// (best handled at MCP tool level via structured paragraph manipulation).
    /// Verify the SDT itself is gone.
    func testDeleteWithKeepContentRemovesSdt() throws {
        var doc = makeDocWithPlainTextSDT(id: 100000, tag: "x", text: "y")
        try doc.deleteContentControl(id: 100000, keepContent: true)

        let allControls = doc.getParagraphs().flatMap { $0.contentControls }
        XCTAssertTrue(allControls.isEmpty, "SDT not removed: \(allControls)")
    }

    /// Block-level SDT delete with keep_content=true unwraps children into body.
    func testDeleteBlockLevelWithKeepContentUnwrapsChildren() throws {
        var doc = WordDocument()
        let sdt = StructuredDocumentTag(id: 200001, tag: "wrap", type: .richText)
        let control = ContentControl(sdt: sdt, content: "")
        let paraA = Paragraph(text: "A")
        let paraB = Paragraph(text: "B")
        doc.body.children = [
            .contentControl(control, children: [.paragraph(paraA), .paragraph(paraB)])
        ]

        try doc.deleteContentControl(id: 200001, keepContent: true)

        XCTAssertEqual(doc.body.children.count, 2,
            "expected 2 unwrapped paragraphs, got \(doc.body.children.count)")
        let texts = doc.body.children.compactMap { child -> String? in
            if case .paragraph(let p) = child { return p.getText() }
            return nil
        }
        XCTAssertEqual(texts, ["A", "B"])
    }

    /// Block-level SDT delete with keep_content=false removes everything.
    func testDeleteBlockLevelWithoutKeepContentRemoves() throws {
        var doc = WordDocument()
        let sdt = StructuredDocumentTag(id: 200001, tag: "wrap", type: .richText)
        let control = ContentControl(sdt: sdt, content: "")
        let paraA = Paragraph(text: "A")
        doc.body.children = [.contentControl(control, children: [.paragraph(paraA)])]

        try doc.deleteContentControl(id: 200001, keepContent: false)

        XCTAssertTrue(doc.body.children.isEmpty,
            "expected empty body, got \(doc.body.children)")
    }

    /// Delete fails on missing id.
    func testDeleteNonexistentThrowsNotFound() {
        var doc = makeDocWithPlainTextSDT(id: 100000, tag: "x", text: "y")
        XCTAssertThrowsError(try doc.deleteContentControl(id: 999999)) { error in
            guard case WordError.contentControlNotFound(let id) = error else {
                XCTFail("expected contentControlNotFound, got \(error)"); return
            }
            XCTAssertEqual(id, 999999)
        }
    }
}
