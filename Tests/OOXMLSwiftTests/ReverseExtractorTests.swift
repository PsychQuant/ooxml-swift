// ReverseExtractorTests.swift
// format-alignment-engine Phase B tasks 2.2 (run-level) + 2.3 (paragraph-
// level) — typed reverse extraction with the byte-equal upgrade rule
// (`ooxml-script-transcode`, «Reverse extraction covers the five format
// layers»; Decision 3: upgrades are accepted ONLY when the trial rebuild
// reproduces the source bytes; otherwise the part stays on the raw channel).
//
// Sources are built with the authoring API so their XML forms are the
// serializer's own — the honest self-round-trip that defines "imitation":
// the system re-derives typed ops from its own output and regenerates
// identical bytes. Foreign forms (e.g. the handwritten CJK fixture) fail the
// trial and stay raw — coverage reflects that truthfully.

import XCTest
@testable import OOXMLSwift

final class ReverseExtractorTests: XCTestCase {

    /// Builds a docx package from authoring ops and returns its part map.
    private func buildParts(_ ops: [OOXMLSwift.Operation]) throws -> [String: Data] {
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: ops)
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("rex-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try doc.writeAuthoringPackage(to: url)
        return try RawPartChannel.readAllParts(from: url)
    }

    /// Full pipeline: reverse (with upgrades) → script → parse → apply →
    /// write → part map of the rebuilt package.
    private func rebuild(_ result: ReverseExtractor.Result) throws -> [String: Data] {
        let script = ScriptExporter.exportSwift(log: result.log)
        let parsed = try ScriptImporter.parse(source: script)
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: parsed.entries.map(\.op))
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("rex-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try doc.writeAuthoringPackage(to: url)
        return try RawPartChannel.readAllParts(from: url)
    }

    // MARK: - Task 2.2: run-level (spec scenario)

    /// Spec scenario: CJK run formatting survives the DSL channel —
    /// eastAsia font ＭＳ ゴシック and size 21 half-points, rebuilt rPr
    /// byte-equal; Stage A/B stay green; coverage strictly increases.
    func testCJKRunFormattingSurvivesDSLChannel() throws {
        let reference = try buildParts([
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "", styleId: "Body", paraId: "P1")),
            .setRuns(target: ElementID(rawString: "w14:paraId=P1"), runs: [RunPayload(
                text: "ゴシック体テスト",
                fontEastAsia: "ＭＳ ゴシック",
                sizeHalfPoints: 21)]),
        ])

        let result = try ReverseExtractor.reverse(parts: reference)
        XCTAssertTrue(result.dslParts.contains("word/document.xml"),
                      "document.xml must upgrade to the DSL channel")

        // Byte-equal floor holds through the full script round-trip.
        let rebuilt = try rebuild(result)
        XCTAssertTrue(PartFidelity.stageB(reference: reference, rebuilt: rebuilt),
                      "Stage B must stay green after the DSL upgrade")

        // Coverage strictly increases versus the all-raw baseline.
        let before = RawPartChannel.partLevelCoverage(parts: reference, dslParts: [])
        let after = RawPartChannel.partLevelCoverage(parts: reference, dslParts: result.dslParts)
        XCTAssertEqual(before.aggregateRatio, 0.0, accuracy: 1e-12)
        XCTAssertGreaterThan(after.aggregateRatio, before.aggregateRatio,
                             "aggregate DSL coverage must strictly increase")
    }

    /// Multiple runs with mixed formatting (bold/color/underline/vertAlign)
    /// upgrade and round-trip byte-equal.
    func testMixedRunFormattingUpgrades() throws {
        let reference = try buildParts([
            .appendParagraph(in: nil, paragraph: ParagraphPayload(text: "", paraId: "P1")),
            .setRuns(target: ElementID(rawString: "w14:paraId=P1"), runs: [
                RunPayload(text: "plain "),
                RunPayload(text: "bold", bold: true),
                RunPayload(text: "red", color: "FF0000"),
                RunPayload(text: "note", sizeHalfPoints: 18, underline: "single",
                           vertAlign: "superscript"),
            ]),
        ])
        let result = try ReverseExtractor.reverse(parts: reference)
        XCTAssertTrue(result.dslParts.contains("word/document.xml"))
        XCTAssertTrue(PartFidelity.stageB(reference: reference, rebuilt: try rebuild(result)))
    }

    // MARK: - Task 2.3: paragraph-level

    /// pPr fields (spacing / indent / alignment / numPr) extract, upgrade,
    /// and rebuild byte-equal on the synthetic source.
    func testParagraphFormattingUpgradesByteEqual() throws {
        let reference = try buildParts([
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "整列された段落", styleId: "Body", paraId: "P1",
                alignment: "both",
                spacingBefore: 120, spacingAfter: 240,
                spacingLine: 360, spacingLineRule: "auto",
                indentLeft: 720, indentFirstLine: 420)),
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "箇条書き", paraId: "P2",
                numId: 3, numLevel: 1)),
        ])
        let result = try ReverseExtractor.reverse(parts: reference)
        XCTAssertTrue(result.dslParts.contains("word/document.xml"),
                      "pPr-bearing document must upgrade")
        XCTAssertTrue(PartFidelity.stageB(reference: reference, rebuilt: try rebuild(result)))
    }

    // MARK: - Task 2.4: section-level (spec scenario)

    /// Spec scenario: two-column section round-trips — a second section
    /// carrying `w:cols num="2"` rebuilds byte-equal through the DSL channel
    /// (mid-body sectPr inside pPr + trailing body sectPr).
    func testTwoColumnSectionRoundTrips() throws {
        let reference = try buildParts([
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "第一節の終わり", styleId: "Body", paraId: "P1")),
            .setSectionProperties(at: ElementID(rawString: "w14:paraId=P1"), section: SectionPayload(
                pageWidth: 11906, pageHeight: 16838,
                marginTop: 1985, marginRight: 1701, marginBottom: 1701, marginLeft: 1701,
                columnSpace: 708)),
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "二段組の本文", paraId: "P2")),
            .setSectionProperties(at: nil, section: SectionPayload(
                pageWidth: 11906, pageHeight: 16838,
                marginTop: 1985, marginRight: 1701, marginBottom: 1701, marginLeft: 1701,
                marginHeader: 851, marginFooter: 992, marginGutter: 0,
                columnCount: 2, columnSpace: 425)),
        ])
        // Sanity: the reference really carries the two-column marker.
        let refDoc = String(decoding: reference["word/document.xml"] ?? Data(), as: UTF8.self)
        XCTAssertTrue(refDoc.contains("w:num=\"2\""), "fixture must carry w:cols num=2")

        let result = try ReverseExtractor.reverse(parts: reference)
        XCTAssertTrue(result.dslParts.contains("word/document.xml"),
                      "two-section document must upgrade to the DSL channel")
        let rebuilt = try rebuild(result)
        XCTAssertTrue(PartFidelity.stageB(reference: reference, rebuilt: rebuilt),
                      "rebuilt sectPr must be byte-equal (Stage B)")
    }

    // MARK: - Task 2.5: tables

    /// A table-bearing synthetic fixture round-trips through the DSL channel
    /// (appendTable with the full cell grid), byte-equal.
    func testTableBearingFixtureRoundTrips() throws {
        let reference = try buildParts([
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "表の前の段落", paraId: "P1")),
            .appendTable(in: nil, table: TablePayload(rows: 2, columns: 3, cells: [
                ["項目", "数量", "備考"],
                ["りんご", "3", ""],
            ])),
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "表の後の段落", paraId: "P2")),
        ])
        let refDoc = String(decoding: reference["word/document.xml"] ?? Data(), as: UTF8.self)
        XCTAssertTrue(refDoc.contains("<w:tbl>"), "fixture must carry a table")

        let result = try ReverseExtractor.reverse(parts: reference)
        XCTAssertTrue(result.dslParts.contains("word/document.xml"),
                      "canonical-form table must upgrade to the DSL channel")
        XCTAssertTrue(PartFidelity.stageB(reference: reference, rebuilt: try rebuild(result)))
    }

    /// A foreign-form table (tblPr present — beyond the canonical minimal
    /// shape) stays raw, and the coverage attribution names the "table"
    /// content class (spec: "records which content classes remain on the
    /// raw channel").
    func testForeignTableStaysRawWithClassAttribution() throws {
        // Start from a canonical package, then perturb the table into a
        // foreign form by inserting a tblPr child the extractor rejects.
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: [
            .appendTable(in: nil, table: TablePayload(rows: 1, columns: 1, cells: [["x"]])),
        ])
        let tree = try XCTUnwrap(doc.xmlTrees["word/document.xml"])
        let body = try XCTUnwrap(tree.root.children.first {
            $0.kind == .element && $0.localName == "body"
        })
        let tbl = try XCTUnwrap(body.children.first {
            $0.kind == .element && $0.localName == "tbl"
        })
        tbl.children.insert(XmlNode.element(prefix: "w", localName: "tblPr"), at: 0)

        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("rex-tbl-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try doc.writeAuthoringPackage(to: url)
        let reference = try RawPartChannel.readAllParts(from: url)

        let result = try ReverseExtractor.reverse(parts: reference)
        XCTAssertFalse(result.dslParts.contains("word/document.xml"))
        XCTAssertEqual(result.rawReasons["word/document.xml"], "table",
                       "coverage attribution must name the table class")
        XCTAssertTrue(PartFidelity.stageB(reference: reference, rebuilt: try rebuild(result)))
    }

    // MARK: - Honest fallback

    /// Foreign XML forms (the handwritten CJK template fixture) fail the
    /// trial rebuild and stay raw — Stage B still green via carryPart, and
    /// document.xml is NOT claimed as DSL.
    func testForeignFormsStayRawAndStageBHolds() throws {
        let url = try CJKTemplateFixtureGenerator.generate()
        defer { try? FileManager.default.removeItem(at: url) }
        let reference = try RawPartChannel.readAllParts(from: url)

        let result = try ReverseExtractor.reverse(parts: reference)
        XCTAssertFalse(result.dslParts.contains("word/document.xml"),
                       "handwritten forms must NOT be claimed as DSL")
        XCTAssertTrue(PartFidelity.stageB(reference: reference, rebuilt: try rebuild(result)),
                      "raw fallback keeps Stage B green")
    }

    /// An unsupported feature inside an otherwise-clean document (hyperlink)
    /// forces the raw fallback rather than a lossy upgrade.
    func testUnsupportedContentFallsBackToRaw() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(text: "link", paraId: "P1")),
        ])
        // Wrap the paragraph's run in a hyperlink — not extractable.
        try doc.apply(operations: [
            .wrapWithHyperlink(target: ElementID(rawString: "w14:paraId=P1"), rId: "rId9"),
        ])
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("rex-h-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try doc.writeAuthoringPackage(to: url)
        let reference = try RawPartChannel.readAllParts(from: url)

        let result = try ReverseExtractor.reverse(parts: reference)
        XCTAssertFalse(result.dslParts.contains("word/document.xml"))
        XCTAssertTrue(PartFidelity.stageB(reference: reference, rebuilt: try rebuild(result)))
    }
}
