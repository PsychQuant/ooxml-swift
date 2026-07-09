// RealTemplateUpgradeTests.swift
// word-canonical-forms Phase 3 task 3.2 — env-gated acceptance (design
// Decision 5 two-track). (a) 90_template_ja's document.xml upgrades to the
// DSL channel with per-part coverage 100%; (c) thesis-fixture does not
// regress (still Stage B green on the raw channel). Skips loudly on CI.
//
//   MACDOC_TEMPLATE_DIR=/path/to/private/docx swift test --filter RealTemplateUpgradeTests

import XCTest
@testable import OOXMLSwift

final class RealTemplateUpgradeTests: XCTestCase {

    /// reverse → script → execute → rebuilt part map.
    private func rebuild(_ result: ReverseExtractor.Result) throws -> [String: Data] {
        let parsed = try ScriptImporter.parse(source: ScriptExporter.exportSwift(log: result.log))
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: parsed.entries.map(\.op))
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("rtu-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try doc.writeAuthoringPackage(to: url)
        return try RawPartChannel.readAllParts(from: url)
    }

    /// (a) The JPA real template document.xml upgrades to the DSL channel and
    /// its per-part coverage is 100%; the full reverse→rebuild is Stage B
    /// byte-equal.
    func testRealTemplateDocumentUpgradesFully() throws {
        let url = try TemplateFixtureGate.requireTemplate(TemplateFixtureGate.baselineTemplateName)
        let reference = try RawPartChannel.readAllParts(from: url)

        let result = try ReverseExtractor.reverse(parts: reference)
        XCTAssertTrue(result.dslParts.contains("word/document.xml"),
                      "90_template_ja document.xml must upgrade to the DSL channel")
        XCTAssertTrue(result.formGaps.isEmpty,
                      "no residual form-gaps expected; got \(result.formGaps.map(\.xmlPath))")

        // Per-part coverage for document.xml is 100% (dslBytes == totalBytes).
        let coverage = RawPartChannel.partLevelCoverage(parts: reference, dslParts: result.dslParts)
        let docPart = try XCTUnwrap(coverage.parts.first { $0.partPath == "word/document.xml" })
        XCTAssertEqual(docPart.coverageRatio, 1.0, accuracy: 1e-12,
                       "document.xml per-part coverage must be 100%")

        // Stage B byte-equal through the full pipeline.
        let rebuilt = try rebuild(result)
        let broken = PartFidelity.compareParts(reference: reference, rebuilt: rebuilt)
            .filter { !$0.isEqual }
        XCTAssertTrue(broken.isEmpty, "Stage B failed on: "
            + broken.map { "\($0.partPath) (\($0.status))" }.joined(separator: ", "))
    }

    /// (c) thesis-fixture no-regress: its out-of-scope structures keep
    /// document.xml on the raw channel (this change must NOT falsely upgrade
    /// it), and every XML part that the raw channel carries round-trips
    /// byte-equal. (Full Stage B is unachievable for thesis-fixture because
    /// it embeds binary media — images — which the UTF-8 carryPart channel
    /// cannot carry; that is a documented pre-existing raw-channel limitation,
    /// not a regression from this change.)
    func testThesisFixtureNoRegress() throws {
        let dir = ProcessInfo.processInfo.environment["MACDOC_TEMPLATE_DIR"]
        guard let dir else { throw XCTSkip("set MACDOC_TEMPLATE_DIR") }
        let url = URL(fileURLWithPath: dir).appendingPathComponent("thesis-fixture.docx")
        guard FileManager.default.fileExists(atPath: url.path) else {
            throw XCTSkip("thesis-fixture.docx not present under MACDOC_TEMPLATE_DIR")
        }
        let reference = try RawPartChannel.readAllParts(from: url)
        let result = try ReverseExtractor.reverse(parts: reference)
        // The key no-regress: document.xml must NOT be falsely claimed as DSL.
        XCTAssertFalse(result.dslParts.contains("word/document.xml"),
                       "thesis-fixture document.xml must stay raw (out of scope)")

        // Every XML part the raw channel actually carries rebuilds byte-equal;
        // binary media parts are the only ones dropped (documented limitation).
        let rebuilt = try rebuild(result)
        for (path, bytes) in reference {
            guard String(data: bytes, encoding: .utf8) != nil else { continue }  // skip binary
            XCTAssertEqual(rebuilt[path], bytes,
                           "carried XML part \(path) must round-trip byte-equal")
        }
    }
}
