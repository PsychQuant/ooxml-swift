// TemplateFixtureGateTests.swift
// format-alignment-engine Phase A task 1.5 — the env-gate for private real-world
// template fixtures (`format-alignment-pipeline`, «Template fixture policy»;
// Decision 5). Verifies the suite passes both with and without the gate:
// without → XCTSkip (CI has no private documents); with → the raw-channel
// pipeline runs and reports baseline coverage.

import XCTest
@testable import OOXMLSwift

final class TemplateFixtureGateTests: XCTestCase {

    /// Real baseline template (90_template_ja) coverage — runs only on a
    /// maintainer machine with `MACDOC_TEMPLATE_DIR` set; skips loudly on CI.
    func testRealTemplateBaselineCoverage() throws {
        let url = try TemplateFixtureGate.requireTemplate(TemplateFixtureGate.baselineTemplateName)
        let parts = try RawPartChannel.readAllParts(from: url)
        // Phase A carries every part on the raw channel → 0% DSL coverage (the
        // honest-copy baseline; later phases raise it).
        let coverage = RawPartChannel.partLevelCoverage(parts: parts, dslParts: [])
        print("[format-alignment baseline] \(TemplateFixtureGate.baselineTemplateName): "
              + "\(coverage.parts.count) XML parts, "
              + "\(coverage.aggregateTotalBytes) XML bytes, aggregate DSL coverage "
              + String(format: "%.1f%%", coverage.aggregateRatio * 100))
        XCTAssertGreaterThan(parts.count, 0)
    }

    /// Without the gate the test skips (does not fail) — this is the CI path.
    func testGateSkipsWhenNoDirOverrideAndEnvUnset() throws {
        // Only meaningful when the env var is actually unset (the CI case).
        guard ProcessInfo.processInfo.environment["MACDOC_TEMPLATE_DIR"] == nil else {
            throw XCTSkip("MACDOC_TEMPLATE_DIR is set in this environment; skip the unset-path check")
        }
        XCTAssertThrowsError(try TemplateFixtureGate.requireTemplate("anything.docx")) { error in
            XCTAssertTrue(error is XCTSkip, "missing gate must throw XCTSkip, not a hard failure")
        }
    }

    /// With the gate present (simulated via dirOverride, no process-env
    /// mutation), the resolver returns the file and the raw-channel pipeline
    /// runs with measurable coverage.
    func testGateResolvesAndReportsCoverageWithOverride() throws {
        let dir = FileManager.default.temporaryDirectory
            .appendingPathComponent("tmpl-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: dir) }
        let synthetic = try CJKTemplateFixtureGenerator.generate()
        let named = dir.appendingPathComponent("sample.docx")
        try FileManager.default.moveItem(at: synthetic, to: named)

        let resolved = try TemplateFixtureGate.requireTemplate("sample.docx", dirOverride: dir.path)
        XCTAssertEqual(resolved.path, named.path)

        let parts = try RawPartChannel.readAllParts(from: resolved)
        let coverage = RawPartChannel.partLevelCoverage(parts: parts, dslParts: [])
        XCTAssertEqual(coverage.aggregateRatio, 0.0, accuracy: 1e-9,
                       "Phase A is all-raw, so DSL coverage is 0%")
        XCTAssertEqual(coverage.parts.count, parts.count,
                       "synthetic template is all-XML — every part counts toward coverage")
    }

    /// A missing file under the gate directory also skips loudly.
    func testGateSkipsWhenFileMissingUnderOverride() throws {
        let dir = FileManager.default.temporaryDirectory
            .appendingPathComponent("tmpl-empty-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: dir) }
        XCTAssertThrowsError(try TemplateFixtureGate.requireTemplate("absent.docx", dirOverride: dir.path)) { error in
            XCTAssertTrue(error is XCTSkip)
        }
    }
}
