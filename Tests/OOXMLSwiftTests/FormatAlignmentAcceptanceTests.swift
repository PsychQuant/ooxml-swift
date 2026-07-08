// FormatAlignmentAcceptanceTests.swift
// format-alignment-engine Phase D task 4.4 — end-to-end acceptance:
// synthetic + real (env-gated) templates through reverse → rebuild →
// Stage B, recording final coverage numbers (`format-alignment-pipeline`,
// «Single-path rebuild pipeline»; Decision 4). The printed numbers feed
// docs/format-alignment-baselines.md (task 3.2).

import XCTest
@testable import OOXMLSwift

final class FormatAlignmentAcceptanceTests: XCTestCase {

    /// reverse → script → execute → part map + coverage, asserting Stage B.
    @discardableResult
    private func acceptancePipeline(reference: [String: Data],
                                    label: String) throws -> PartFidelity.CoverageReport {
        let result = try ReverseExtractor.reverse(parts: reference)
        let script = ScriptExporter.exportSwift(log: result.log)
        let parsed = try ScriptImporter.parse(source: script)
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: parsed.entries.map(\.op))
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("faa-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try doc.writeAuthoringPackage(to: outURL)
        let rebuilt = try RawPartChannel.readAllParts(from: outURL)

        let verdicts = PartFidelity.compareParts(reference: reference, rebuilt: rebuilt)
        let broken = verdicts.filter { !$0.isEqual }
        XCTAssertTrue(broken.isEmpty, "[\(label)] Stage B failed on: "
            + broken.map { "\($0.partPath) (\($0.status))" }.joined(separator: ", "))

        let coverage = RawPartChannel.partLevelCoverage(parts: reference, dslParts: result.dslParts)
        let rawClasses = Set(result.rawReasons.values).sorted().joined(separator: ", ")
        print("[format-alignment acceptance] \(label): "
            + "\(coverage.parts.count) XML parts, \(coverage.aggregateTotalBytes) bytes, "
            + String(format: "aggregate DSL %.1f%%", coverage.aggregateRatio * 100)
            + " | raw classes: [\(rawClasses)]")
        return coverage
    }

    /// Authoring-built synthetic exercising all five upgraded layers —
    /// full pipeline Stage B + a strictly positive DSL coverage share.
    func testSyntheticFiveLayerAcceptance() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "", styleId: "Body", paraId: "P1", alignment: "both",
                spacingBefore: 120, spacingAfter: 240, indentLeft: 720)),
            .setRuns(target: ElementID(rawString: "w14:paraId=P1"), runs: [
                RunPayload(text: "ゴシック体", bold: true, fontEastAsia: "ＭＳ ゴシック",
                           sizeHalfPoints: 21),
                RunPayload(text: " 本文", underline: "single"),
            ]),
            .setSectionProperties(at: ElementID(rawString: "w14:paraId=P1"), section: SectionPayload(
                pageWidth: 11906, pageHeight: 16838, columnSpace: 708)),
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "二段組の節", paraId: "P2", numId: 1, numLevel: 0)),
            .appendTable(in: nil, table: TablePayload(rows: 2, columns: 2, cells: [
                ["項目", "値"], ["係数", "0.85"],
            ])),
            .setSectionProperties(at: nil, section: SectionPayload(
                pageWidth: 11906, pageHeight: 16838, marginTop: 1985, marginRight: 1701,
                marginBottom: 1701, marginLeft: 1701, columnCount: 2, columnSpace: 425)),
        ])
        let refURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("faa-ref-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: refURL) }
        try doc.writeAuthoringPackage(to: refURL)
        let reference = try RawPartChannel.readAllParts(from: refURL)

        let coverage = try acceptancePipeline(reference: reference,
                                              label: "synthetic five-layer (authoring-built)")
        XCTAssertGreaterThan(coverage.aggregateRatio, 0,
                             "five-layer synthetic must show DSL coverage")
    }

    /// Committed CJK template fixture (handwritten foreign forms): the raw
    /// channel carries it Stage B byte-equal; coverage honestly reports 0%.
    func testCommittedCJKTemplateAcceptance() throws {
        let url = try CJKTemplateFixtureGenerator.generate()
        defer { try? FileManager.default.removeItem(at: url) }
        let reference = try RawPartChannel.readAllParts(from: url)
        let coverage = try acceptancePipeline(reference: reference,
                                              label: "synthetic CJK template (handwritten)")
        XCTAssertEqual(coverage.aggregateRatio, 0.0, accuracy: 1e-12,
                       "foreign forms stay raw — honest 0% DSL")
    }

    /// Real-world template (env-gated, maintainer machine): full pipeline
    /// Stage B + recorded coverage. Skips loudly on CI.
    func testRealTemplateAcceptance() throws {
        let url = try TemplateFixtureGate.requireTemplate(TemplateFixtureGate.baselineTemplateName)
        let reference = try RawPartChannel.readAllParts(from: url)
        try acceptancePipeline(reference: reference,
                               label: TemplateFixtureGate.baselineTemplateName)
    }
}
