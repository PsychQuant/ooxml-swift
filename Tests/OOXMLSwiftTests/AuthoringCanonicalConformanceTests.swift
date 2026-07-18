// AuthoringCanonicalConformanceTests.swift
// authoring-canonical-conformance (ooxml-swift#85) — the authoring path
// (DocxWriter + typed models) must emit transcoder-canonical document.xml so
// self-authored documents upgrade to the DSL channel. Spec: delta on
// `ooxml-script-transcode` ("Authoring path emits transcoder-canonical
// document.xml" + "Authoring chokepoints stamp w14:paraId").

import XCTest
@testable import OOXMLSwift

final class AuthoringCanonicalConformanceTests: XCTestCase {

    private func paraIdPattern(_ id: String?) -> Bool {
        guard let id else { return false }
        return id.range(of: "^[0-9A-F]{8}$", options: .regularExpression) != nil
    }

    // MARK: - Task 2.1 — append / insert(Int) stamping

    func testAppendStampsDistinctConformingParaIds() {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(runs: [Run(text: "one")]))
        doc.appendParagraph(Paragraph(runs: [Run(text: "two")]))

        let paras = doc.getAllParagraphs()
        XCTAssertEqual(paras.count, 2)
        XCTAssertTrue(paraIdPattern(paras[0].w14ParaId), "first appended paragraph must carry a conforming paraId")
        XCTAssertTrue(paraIdPattern(paras[1].w14ParaId), "second appended paragraph must carry a conforming paraId")
        XCTAssertNotEqual(paras[0].w14ParaId, paras[1].w14ParaId, "stamped paraIds must be unique in the document")

        for para in paras {
            let xml = para.toXML()
            XCTAssertTrue(xml.contains("w14:paraId=\"\(para.w14ParaId!)\""),
                          "serialized <w:p> must carry the stamped paraId; got: \(xml.prefix(120))")
        }
    }

    func testInsertAtIndexStampsParaId() {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(runs: [Run(text: "existing")]))
        doc.insertParagraph(Paragraph(runs: [Run(text: "inserted")]), at: 0)

        let paras = doc.getAllParagraphs()
        XCTAssertEqual(paras.count, 2)
        XCTAssertTrue(paraIdPattern(paras[0].w14ParaId))
        XCTAssertNotEqual(paras[0].w14ParaId, paras[1].w14ParaId)
    }

    func testPresetParaIdIsPreservedVerbatim() {
        var doc = WordDocument()
        var preset = Paragraph(runs: [Run(text: "keep me")])
        preset.w14ParaId = "3F2A0001"
        doc.appendParagraph(preset)

        let para = doc.getAllParagraphs()[0]
        XCTAssertEqual(para.w14ParaId, "3F2A0001")
        XCTAssertTrue(para.toXML().contains("w14:paraId=\"3F2A0001\""))
    }

    // MARK: - Task 2.2 — insert(InsertLocation) stamping

    func testInsertAtLocationStampsParaId() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(runs: [Run(text: "anchor")]))
        try doc.insertParagraph(Paragraph(runs: [Run(text: "located")]),
                                at: .paragraphIndex(1))

        let paras = doc.getAllParagraphs()
        XCTAssertEqual(paras.count, 2)
        XCTAssertTrue(paraIdPattern(paras[1].w14ParaId),
                      "InsertLocation-inserted paragraph must carry a conforming paraId")
        XCTAssertNotEqual(paras[0].w14ParaId, paras[1].w14ParaId,
                          "generated paraId must be unique against existing paragraph IDs")
    }

    func testInsertAtLocationPreservesPresetParaId() throws {
        var doc = WordDocument()
        var preset = Paragraph(runs: [Run(text: "preset")])
        preset.w14ParaId = "3F2A0002"
        try doc.insertParagraph(preset, at: .paragraphIndex(0))
        XCTAssertEqual(doc.getAllParagraphs()[0].w14ParaId, "3F2A0002")
    }

    // MARK: - Task 2.3 — no backfill on legacy round-trip

    func testLegacyParagraphsAreNotBackfilledOnResave() throws {
        // Build a package whose paragraphs lack paraIds by bypassing the
        // chokepoints (direct body mutation — the legacy-document shape).
        var source = WordDocument()
        source.body.children.append(.paragraph(Paragraph(runs: [Run(text: "legacy one")])))
        source.body.children.append(.paragraph(Paragraph(runs: [Run(text: "legacy two")])))

        let dir = FileManager.default.temporaryDirectory
        let sourceURL = dir.appendingPathComponent("acc-legacy-\(UUID().uuidString).docx")
        let resavedURL = dir.appendingPathComponent("acc-resaved-\(UUID().uuidString).docx")
        defer {
            try? FileManager.default.removeItem(at: sourceURL)
            try? FileManager.default.removeItem(at: resavedURL)
        }
        try DocxWriter.write(source, to: sourceURL)

        // Open → force typed re-materialization of document.xml (the risky
        // path — a verbatim copy would trivially pass) → save.
        var reopened = try DocxReader.read(from: sourceURL)
        reopened.markTypedDirty("word/document.xml")
        try DocxWriter.write(reopened, to: resavedURL)

        let parts = try RawPartChannel.readAllParts(from: resavedURL)
        let documentXML = try XCTUnwrap(parts["word/document.xml"])
        let xml = String(decoding: documentXML, as: UTF8.self)
        XCTAssertFalse(xml.contains("w14:paraId"),
                       "parsed paragraphs without paraId must not gain one on re-save")
    }

    // MARK: - Task 3.1 — no inter-element whitespace in authoring output

    func testAuthoringDocumentXMLIsElementOnly() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(runs: [Run(text: "compact")]))

        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("acc-compact-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try DocxWriter.write(doc, to: url)

        let parts = try RawPartChannel.readAllParts(from: url)
        let tree = try XmlTreeReader.parse(try XCTUnwrap(parts["word/document.xml"]))

        XCTAssertTrue(tree.root.children.allSatisfy { $0.kind == .element },
                      "w:document children must be element-only; found kinds: \(tree.root.children.map(\.kind))")
        let body = try XCTUnwrap(tree.root.children.first { $0.localName == "body" })
        XCTAssertTrue(body.children.allSatisfy { $0.kind == .element },
                      "w:body children must be element-only; found kinds: \(body.children.map(\.kind))")
    }

    // MARK: - Task 3.2 — full Word-canonical root cloud (create-from-scratch)

    /// Expected root open tag — an independent copy of the `90_template_ja.docx`
    /// baseline root (word-canonical-forms source of truth). Kept verbatim in
    /// the test so drift in the writer constant is caught.
    private static let expectedWordCanonicalRootOpenTag = "<w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:oel=\"http://schemas.microsoft.com/office/2019/extlst\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" xmlns:w16du=\"http://schemas.microsoft.com/office/word/2023/wordml/word16du\" xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" xmlns:w16sdtfl=\"http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14\">"

    func testCreateFromScratchRootEmitsFullWordCanonicalCloud() throws {
        let tag = try DocxWriter.renderDocumentRootOpenTag([:])
        XCTAssertEqual(tag, Self.expectedWordCanonicalRootOpenTag,
                       "empty documentRootAttributes must emit the full Word-canonical cloud")
        XCTAssertTrue(tag.contains("xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\""),
                      "w14 namespace declaration is required for stamped w14:paraId validity")
    }

    func testSavedCreateFromScratchDocumentCarriesFullCloud() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(runs: [Run(text: "cloud")]))
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("acc-cloud-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try DocxWriter.write(doc, to: url)
        let xml = String(decoding: try XCTUnwrap(
            RawPartChannel.readAllParts(from: url)["word/document.xml"]), as: UTF8.self)
        XCTAssertTrue(xml.contains(Self.expectedWordCanonicalRootOpenTag),
                      "saved document.xml root must survive the tree round-trip byte-for-byte")
    }

    /// Provenance: the writer constant mirrors the real fixture, not memory.
    /// Env-gated (MACDOC_TEMPLATE_DIR) — skips loudly when absent.
    func testRootCloudProvenanceAgainstBaselineFixture() throws {
        let url = try TemplateFixtureGate.requireTemplate(TemplateFixtureGate.baselineTemplateName)
        let data = try XCTUnwrap(RawPartChannel.readAllParts(from: url)["word/document.xml"])
        let xml = String(decoding: data, as: UTF8.self)
        let start = try XCTUnwrap(xml.range(of: "<w:document"))
        let end = try XCTUnwrap(xml.range(of: ">", range: start.upperBound..<xml.endIndex))
        let fixtureRoot = String(xml[start.lowerBound..<end.upperBound])
        XCTAssertEqual(try DocxWriter.renderDocumentRootOpenTag([:]), fixtureRoot,
                       "writer root cloud must equal the real-Word baseline root")
    }

    // MARK: - Tasks 4.1 + 4.2 — reverse-to-DSL + export→execute→byte-equal

    /// Builds the acceptance fixture: create-from-scratch, plain + formatted
    /// paragraphs through the chokepoints, saved by DocxWriter.
    private func makeAuthoringFixture() throws -> (url: URL, cleanup: () -> Void) {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(runs: [Run(text: "plain paragraph")]))
        var boldProps = RunProperties()
        boldProps.bold = true
        doc.appendParagraph(Paragraph(runs: [Run(text: "bold lead", properties: boldProps),
                                             Run(text: " plain tail")]))
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("acc-roundtrip-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: url)
        return (url, { try? FileManager.default.removeItem(at: url) })
    }

    func testAuthoringDocumentUpgradesToDSLChannel() throws {
        let fixture = try makeAuthoringFixture()
        defer { fixture.cleanup() }
        let parts = try RawPartChannel.readAllParts(from: fixture.url)

        let result = try ReverseExtractor.reverse(parts: parts)
        XCTAssertTrue(result.dslParts.contains("word/document.xml"),
                      "document.xml must upgrade to the DSL channel; form gaps: "
                      + result.formGaps.map { "\($0.contentClass)@\($0.xmlPath)" }.joined(separator: ", "))

        let coverage = RawPartChannel.partLevelCoverage(parts: parts, dslParts: result.dslParts)
        let docPart = try XCTUnwrap(coverage.parts.first { $0.partPath == "word/document.xml" })
        XCTAssertEqual(docPart.coverageRatio, 1.0, accuracy: 1e-12,
                       "document.xml per-part DSL coverage must be 100%")
    }

    func testExportedScriptRebuildsByteEqual() throws {
        let fixture = try makeAuthoringFixture()
        defer { fixture.cleanup() }
        let parts = try RawPartChannel.readAllParts(from: fixture.url)
        let result = try ReverseExtractor.reverse(parts: parts)
        // Guard against a vacuous pass: a raw-carried document.xml would also
        // rebuild byte-equal. This test is about the DSL-channel rebuild.
        XCTAssertTrue(result.dslParts.contains("word/document.xml"),
                      "precondition: fixture must upgrade to the DSL channel")

        let script = ScriptExporter.exportSwift(log: result.log)
        let parsed = try ScriptImporter.parse(source: script)
        var rebuilt = WordDocument.emptyAuthoringDocument()
        try rebuilt.apply(operations: parsed.entries.map(\.op))
        let rebuiltURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("acc-rebuilt-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: rebuiltURL) }
        try rebuilt.writeAuthoringPackage(to: rebuiltURL)
        let rebuiltParts = try RawPartChannel.readAllParts(from: rebuiltURL)

        let source = try XCTUnwrap(parts["word/document.xml"])
        let output = try XCTUnwrap(rebuiltParts["word/document.xml"])
        XCTAssertEqual(source, output,
                       "rebuilt document.xml must be byte-equal; "
                       + ReverseExtractor.byteMismatchLocator(source: source, rebuilt: output))
    }

    // MARK: - Verify R1 fix — legacy root gains xmlns:w14 when stamping occurs

    /// Verify #85 R1 blocking finding 1: a legacy document whose captured root
    /// lacks `xmlns:w14` (old minimal w+r create-from-scratch output) must not
    /// emit unbound `w14:paraId` prefixes when a paragraph is appended through
    /// the authoring API — the writer augments the root with the w14
    /// declaration. Namespace augmentation only; no paragraph backfill.
    func testStampedParagraphAugmentsLegacyRootWithW14() throws {
        var doc = WordDocument()
        doc.documentRootAttributes = [
            "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ]
        doc.appendParagraph(Paragraph(runs: [Run(text: "legacy edit")]))

        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("acc-legacyw14-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try DocxWriter.write(doc, to: url)

        let xml = String(decoding: try XCTUnwrap(
            RawPartChannel.readAllParts(from: url)["word/document.xml"]), as: UTF8.self)
        XCTAssertTrue(xml.contains("w14:paraId="), "appended paragraph must carry the stamped paraId")
        XCTAssertTrue(xml.contains("xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\""),
                      "root must declare xmlns:w14 when body carries w14 attributes (unbound-prefix guard)")
    }

    /// Companion: a legacy document with NO stamped paragraphs keeps its
    /// captured root untouched (no drive-by w14 injection).
    func testLegacyRootWithoutW14ContentStaysUntouched() throws {
        var doc = WordDocument()
        doc.documentRootAttributes = [
            "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ]
        doc.body.children.append(.paragraph(Paragraph(runs: [Run(text: "untouched")])))

        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("acc-legacynone-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try DocxWriter.write(doc, to: url)

        let xml = String(decoding: try XCTUnwrap(
            RawPartChannel.readAllParts(from: url)["word/document.xml"]), as: UTF8.self)
        XCTAssertFalse(xml.contains("xmlns:w14"),
                       "captured root must stay verbatim when no w14 content exists")
    }

    // MARK: - Verify R1 fix — table-cell insertion is not stamped

    /// Verify #85 R1 blocking finding 2: `plainCellText` requires cell
    /// paragraphs to be attribute-free (cells use positional addressing), so
    /// `.intoTableCell` insertion must NOT stamp a paraId — stamping there
    /// would knock the whole part off the canonical plain-table form.
    func testInsertIntoTableCellDoesNotStampParaId() throws {
        var doc = WordDocument()
        var cell = TableCell()
        cell.paragraphs = [Paragraph(runs: [Run(text: "seed")])]
        let table = Table(rows: [TableRow(cells: [cell])])
        doc.body.children.append(.table(table))

        try doc.insertParagraph(Paragraph(runs: [Run(text: "cell insert")]),
                                at: .intoTableCell(tableIndex: 0, row: 0, col: 0))

        guard case .table(let updated) = doc.body.children[0] else {
            return XCTFail("expected table at body index 0")
        }
        let cellParas = updated.rows[0].cells[0].paragraphs
        XCTAssertEqual(cellParas.count, 2)
        XCTAssertNil(cellParas.last?.w14ParaId,
                     "cell paragraphs must stay attribute-free (positional addressing)")
    }

    // MARK: - Task 4.3 — bypass path stays raw with attribution

    func testBypassPathStaysRawWithParaIdAttribution() throws {
        var doc = WordDocument()
        // Direct body mutation skips the chokepoints — no paraId stamped.
        doc.body.children.append(.paragraph(Paragraph(runs: [Run(text: "bypass")])))
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("acc-bypass-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try DocxWriter.write(doc, to: url)

        let result = try ReverseExtractor.reverse(parts: RawPartChannel.readAllParts(from: url))
        XCTAssertFalse(result.dslParts.contains("word/document.xml"),
                       "paraId-less paragraph must keep document.xml on the raw channel")
        let gap = result.formGaps.first { $0.contentClass == "paragraph-no-paraId" }
        XCTAssertNotNil(gap, "form-gap report must name paragraph-no-paraId; got: "
                        + result.formGaps.map(\.contentClass).joined(separator: ", "))
        XCTAssertFalse(gap?.xmlPath.isEmpty ?? true, "the gap must carry a located path")
    }
}
