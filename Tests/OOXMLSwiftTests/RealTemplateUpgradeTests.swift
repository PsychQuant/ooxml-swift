// RealTemplateUpgradeTests.swift
// word-canonical-forms Phase 3 task 3.2 — env-gated acceptance (design
// Decision 5 two-track). (a) 90_template_ja's document.xml upgrades to the
// DSL channel with per-part coverage 100%; (c) thesis-fixture does not
// regress (still Stage B green on the raw channel). Skips loudly on CI.
//
//   MACDOC_TEMPLATE_DIR=/path/to/private/docx swift test --filter RealTemplateUpgradeTests

import XCTest
import PDFKit
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

    /// reverse → slotted script → substitute → execute → rebuilt part map.
    private func executeScript(_ script: String) throws -> [String: Data] {
        let parsed = try ScriptImporter.parse(source: script)
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: parsed.entries.map(\.op))
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("rtu-slot-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try doc.writeAuthoringPackage(to: url)
        return try RawPartChannel.readAllParts(from: url)
    }

    /// (b) task 3.2(b) acceptance: `--slot` works end-to-end on the real JPA
    /// template. A formatted paragraph (raw-form, single-run setRuns) is
    /// designated as a slot; substituting new text at the call site lands the
    /// new content in document.xml while every non-document XML part stays
    /// byte-equal to the reference.
    func testRealTemplateSlotSubstitution() throws {
        let url = try TemplateFixtureGate.requireTemplate(TemplateFixtureGate.baselineTemplateName)
        let reference = try RawPartChannel.readAllParts(from: url)
        let log = try ReverseExtractor.reverse(parts: reference).log

        // Find a slottable formatted paragraph: a single-run setRuns with
        // non-empty text whose paragraph is op-level substitutable.
        var slotParaId: String?
        for entry in log.entries {
            guard case .setRuns(let target, let runs) = entry.op,
                  runs.count == 1, !runs[0].text.isEmpty,
                  target.raw.hasPrefix("w14:paraId=") else { continue }
            let pid = String(target.raw.dropFirst("w14:paraId=".count))
            if ScriptExporter.opLevelSlotDefault(log: log, paraId: pid) != nil {
                slotParaId = pid
                break
            }
        }
        let paraId = try XCTUnwrap(slotParaId,
            "expected at least one slottable formatted paragraph in the JPA template")

        var script = try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "slot0", paraId: paraId),
        ])
        XCTAssertTrue(script.contains("// @slot slot0 \(paraId)"),
                      "op-level slot directive expected")

        let sentinel = "＿＿差し込みテスト＿＿"
        let original = try XCTUnwrap(ScriptExporter.opLevelSlotDefault(log: log, paraId: paraId))
        script = script.replacingOccurrences(
            of: "    slot0: \(ScriptExporter.quote(original))",
            with: "    slot0: \(ScriptExporter.quote(sentinel))")

        let rebuilt = try executeScript(script)
        let docXML = String(decoding: rebuilt["word/document.xml"] ?? Data(), as: UTF8.self)
        XCTAssertTrue(docXML.contains(sentinel),
                      "the slot must land the new text in document.xml")

        // Every non-document XML part stays byte-equal (the slot only touches
        // document.xml); binary media is skipped (documented raw-channel limit).
        for (path, bytes) in reference where path != "word/document.xml" {
            guard String(data: bytes, encoding: .utf8) != nil else { continue }
            XCTAssertEqual(rebuilt[path], bytes,
                           "non-slot XML part \(path) must stay byte-equal")
        }
    }

    /// (d) render-effect-semantics task 3.1 — render-level acceptance for
    /// slotted rebuilds (`docx-render-semantics`, «Slotted rebuilds carry
    /// render-level acceptance»; design Decision 7). DUAL-gated: needs both
    /// `RUN_WORD_INTEGRATION=1` (live Word render) and `MACDOC_TEMPLATE_DIR`
    /// (real template) — each absence skips loudly naming the missing gate.
    /// Executing the slotted script with sentinel content of comparable
    /// length must preserve rendered structure: equal page count, equal page
    /// box on every page, substituted-page median line pitch within Decision
    /// 5 tolerance of the reference page; every untouched page stays under
    /// the pixel-ratio threshold.
    func testRealTemplateSlotSubstitutionRendersEquivalently() throws {
        try VisualDiffHarness.requireGate()
        let url = try TemplateFixtureGate.requireTemplate(TemplateFixtureGate.baselineTemplateName)
        let reference = try RawPartChannel.readAllParts(from: url)
        let log = try ReverseExtractor.reverse(parts: reference).log

        // Same slot selection as scenario (b): first op-level-substitutable
        // formatted paragraph.
        var slotParaId: String?
        for entry in log.entries {
            guard case .setRuns(let target, let runs) = entry.op,
                  runs.count == 1, !runs[0].text.isEmpty,
                  target.raw.hasPrefix("w14:paraId=") else { continue }
            let pid = String(target.raw.dropFirst("w14:paraId=".count))
            if ScriptExporter.opLevelSlotDefault(log: log, paraId: pid) != nil {
                slotParaId = pid
                break
            }
        }
        let paraId = try XCTUnwrap(slotParaId, "expected a slottable paragraph in the JPA template")
        let original = try XCTUnwrap(ScriptExporter.opLevelSlotDefault(log: log, paraId: paraId))

        // Sentinel of comparable length: distinctive prefix, padded to the
        // original's character count so line-fill stays comparable.
        let prefix = "差込検証"
        let sentinel = original.count > prefix.count
            ? prefix + String(repeating: "査", count: original.count - prefix.count)
            : prefix
        var script = try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "slot0", paraId: paraId),
        ])
        script = script.replacingOccurrences(
            of: "    slot0: \(ScriptExporter.quote(original))",
            with: "    slot0: \(ScriptExporter.quote(sentinel))")

        // Execute → rebuilt docx on disk; render both through Word. All I/O
        // rides the TCC-granted scratch dir (the reference docx is copied in
        // so Word never touches an unauthorized path).
        let scratch = try RenderProbeHarness.scratchDir()
        let refDocx = scratch.appendingPathComponent("slot-render-ref.docx")
        let rebuiltDocx = scratch.appendingPathComponent("slot-render-new.docx")
        let refPDFURL = scratch.appendingPathComponent("slot-render-ref.pdf")
        let rebuiltPDFURL = scratch.appendingPathComponent("slot-render-new.pdf")
        for f in [refDocx, rebuiltDocx, refPDFURL, rebuiltPDFURL] {
            try? FileManager.default.removeItem(at: f)
        }
        try FileManager.default.copyItem(at: url, to: refDocx)
        let parsed = try ScriptImporter.parse(source: script)
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: parsed.entries.map(\.op))
        try doc.writeAuthoringPackage(to: rebuiltDocx)

        try VisualDiffHarness.exportPDF(docx: refDocx, to: refPDFURL)
        try VisualDiffHarness.exportPDF(docx: rebuiltDocx, to: rebuiltPDFURL)
        let refPDF = try XCTUnwrap(PDFDocument(url: refPDFURL))
        let newPDF = try XCTUnwrap(PDFDocument(url: rebuiltPDFURL))

        // Structure: page count + page box equal on every page.
        XCTAssertEqual(RenderGeometry.pageCount(pdf: refPDF),
                       RenderGeometry.pageCount(pdf: newPDF),
                       "new slot content must not change the page count")
        for page in 0..<RenderGeometry.pageCount(pdf: refPDF) {
            XCTAssertEqual(RenderGeometry.pageBox(pdf: refPDF, page: page),
                           RenderGeometry.pageBox(pdf: newPDF, page: page),
                           "page box must be unchanged on page \(page + 1)")
        }

        // Locate the substituted page via the sentinel prefix.
        var substitutedPage: Int?
        for page in 0..<RenderGeometry.pageCount(pdf: newPDF)
        where (newPDF.page(at: page)?.string ?? "").contains(prefix) {
            substitutedPage = page
            break
        }
        let slotPage = try XCTUnwrap(substitutedPage,
                                     "sentinel text must appear in the rendered rebuild")

        // Substituted page: line pitch within Decision 5 tolerance of the
        // reference page (nil pitch on a text page is a failure, not a skip).
        if let refPitch = RenderGeometry.medianLinePitch(pdf: refPDF, page: slotPage) {
            let newPitch = try XCTUnwrap(
                RenderGeometry.medianLinePitch(pdf: newPDF, page: slotPage),
                "substituted page must keep measurable text lines")
            XCTAssertEqual(newPitch, refPitch,
                           accuracy: RenderProbeHarness.tolerance(forPredicted: refPitch),
                           "substituted-page line pitch must stay within tolerance")
        }

        // Untouched pages: pixel-equal within the harness threshold.
        let ratios = try VisualDiffHarness.pageDiffRatios(refPDFURL, rebuiltPDFURL)
        for (page, ratio) in ratios.enumerated() where page != slotPage {
            XCTAssertLessThanOrEqual(ratio, VisualDiffHarness.threshold,
                "untouched page \(page + 1) drifted (ratio \(ratio))")
        }
        print("[render-effect evidence] slot render acceptance: \(ratios.count) pages, "
              + "substituted page \(slotPage + 1), ratios \(ratios)")
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
