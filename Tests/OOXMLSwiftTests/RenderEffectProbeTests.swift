// RenderEffectProbeTests.swift
// render-effect-semantics tasks 2.1–2.7 — perturbation probes for the
// render-effect registry (`docx-render-semantics`, «Perturbation probes
// verify predicted rendering effects»). One probe per registry row in
// macdoc's docs/render-effect-registry.md: build baseline + perturbed docx
// via typed operations differing in exactly ONE payload field, render both
// through live Word, measure the observable with RenderGeometry, assert the
// predicted direction exactly and the magnitude within tolerance
// (design Decision 5: ±10% or ±1.0 pt, whichever is larger).
//
//     RUN_WORD_INTEGRATION=1 swift test --filter RenderEffectProbeTests
//
// All I/O rides the FIXED TCC-granted directory
// `~/.cache/ooxml-swift-visual-diff` (design Decision 4; see the
// VisualDiffTests scratch-path note — a fresh directory re-triggers Word's
// blocking Grant Access sheet).

import XCTest
import PDFKit
@testable import OOXMLSwift

/// Shared probe plumbing (task 2.1).
enum RenderProbeHarness {

    /// Decision 5 tolerance: magnitude within ±10% of the prediction or
    /// ±1.0 pt, whichever is larger.
    static func tolerance(forPredicted magnitude: CGFloat) -> CGFloat {
        max(abs(magnitude) * 0.10, 1.0)
    }

    /// The granted scratch directory. Files are cleared per call for the
    /// given base name only (probes run in one process; distinct base names
    /// avoid collisions); the directory itself is NEVER deleted.
    static func scratchDir() throws -> URL {
        let dir = FileManager.default.homeDirectoryForCurrentUser
            .appendingPathComponent(".cache/ooxml-swift-visual-diff")
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        return dir
    }

    /// ops → docx → Word render → PDFDocument. Skips (not fails) on
    /// environment problems, per the harness contract.
    static func renderedPDF(ops: [OOXMLSwift.Operation], baseName: String) throws -> PDFDocument {
        let dir = try scratchDir()
        let docxURL = dir.appendingPathComponent("\(baseName).docx")
        let pdfURL = dir.appendingPathComponent("\(baseName).pdf")
        try? FileManager.default.removeItem(at: docxURL)
        try? FileManager.default.removeItem(at: pdfURL)

        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: ops)
        try doc.writeAuthoringPackage(to: docxURL)

        try VisualDiffHarness.exportPDF(docx: docxURL, to: pdfURL)
        guard let pdf = PDFDocument(url: pdfURL) else {
            throw XCTSkip("Word produced an unreadable PDF at \(pdfURL.lastPathComponent)")
        }
        return pdf
    }

    /// Body text long enough to wrap into several lines at 10.5pt on A4.
    static func wrappingText(sentences: Int = 14) -> String {
        Array(repeating: "The quick brown fox jumps over the lazy dog near the riverbank.",
              count: sentences).joined(separator: " ")
    }
}

final class RenderEffectProbeTests: XCTestCase {

    /// Task 2.1 plumbing smoke test: the ops→docx→Word→PDF→geometry channel
    /// works end-to-end before any effect probe relies on it. Gated; skips
    /// loudly without RUN_WORD_INTEGRATION=1 or Word.
    func testProbeChannelPlumbing() throws {
        try VisualDiffHarness.requireGate()
        let pdf = try RenderProbeHarness.renderedPDF(ops: [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: RenderProbeHarness.wrappingText(), paraId: "P1")),
        ], baseName: "probe-plumbing")

        XCTAssertGreaterThanOrEqual(RenderGeometry.pageCount(pdf: pdf), 1)
        let lines = RenderGeometry.lineBoxes(pdf: pdf, page: 0)
        XCTAssertGreaterThanOrEqual(lines.count, 2,
            "wrapping text must produce multiple measurable lines")
        XCTAssertNotNil(RenderGeometry.medianLinePitch(pdf: pdf, page: 0),
            "multi-line page must yield a measured pitch — nil is a failure, not a skip")
    }

    // MARK: - Registry row 1 (task 2.2)

    /// Registry #1 — `SectionPayload.docGridLinePitch` (type `lines`),
    /// 360 → 480 twips. Spec example values exactly: predicted pitch
    /// increase = (480 − 360) / 20 = 6.0 pt.
    func testProbeDocGridLinePitch() throws {
        try VisualDiffHarness.requireGate()

        func ops(linePitch: Int) -> [OOXMLSwift.Operation] {
            [
                .appendParagraph(in: nil, paragraph: ParagraphPayload(text: "", paraId: "P1")),
                .setRuns(target: ElementID(rawString: "w14:paraId=P1"), runs: [RunPayload(
                    text: RenderProbeHarness.wrappingText(),
                    fontAscii: "Times New Roman", sizeHalfPoints: 21)]),
                .setSectionProperties(at: nil, section: SectionPayload(
                    pageWidth: 11906, pageHeight: 16838,
                    docGridType: "lines", docGridLinePitch: linePitch)),
            ]
        }
        let baseline = try RenderProbeHarness.renderedPDF(
            ops: ops(linePitch: 360), baseName: "probe-grid-360")
        let perturbed = try RenderProbeHarness.renderedPDF(
            ops: ops(linePitch: 480), baseName: "probe-grid-480")

        let basePitch = try XCTUnwrap(
            RenderGeometry.medianLinePitch(pdf: baseline, page: 0),
            "[registry #1 docGridLinePitch] baseline page must have measurable lines")
        let pertPitch = try XCTUnwrap(
            RenderGeometry.medianLinePitch(pdf: perturbed, page: 0),
            "[registry #1 docGridLinePitch] perturbed page must have measurable lines")

        // Direction exact, magnitude within Decision 5 tolerance.
        XCTAssertGreaterThan(pertPitch, basePitch,
            "[registry #1 docGridLinePitch] pitch must increase with linePitch")
        let predicted: CGFloat = (480 - 360) / 20  // 6.0 pt
        XCTAssertEqual(pertPitch - basePitch, predicted,
                       accuracy: RenderProbeHarness.tolerance(forPredicted: predicted),
                       "[registry #1 docGridLinePitch] measured Δpitch vs predicted 6.0 pt")
        print("[render-effect evidence] docGridLinePitch: baseline \(basePitch) pt, "
              + "perturbed \(pertPitch) pt, Δ \(pertPitch - basePitch) pt (predicted \(predicted))")
    }

    // MARK: - Registry rows 2 + 3 (task 2.3)

    /// Two short single-line paragraphs with explicit font/size. ALL spacing
    /// fields are pinned to explicit values (default 0) so a probe's
    /// perturbation is the only difference between baseline and perturbed —
    /// the first probe run proved that leaving them to style defaults lets
    /// Word's before/after collapse (max, not sum) against the default
    /// style's spacingAfter, breaking the single-field isolation.
    private func twoParagraphOps(p1After: Int = 0, p2Before: Int = 0) -> [OOXMLSwift.Operation] {
        [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "", paraId: "P1", spacingBefore: 0, spacingAfter: p1After)),
            .setRuns(target: ElementID(rawString: "w14:paraId=P1"), runs: [RunPayload(
                text: "First paragraph line.",
                fontAscii: "Times New Roman", sizeHalfPoints: 21)]),
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "", paraId: "P2", spacingBefore: p2Before, spacingAfter: 0)),
            .setRuns(target: ElementID(rawString: "w14:paraId=P2"), runs: [RunPayload(
                text: "Second paragraph line.",
                fontAscii: "Times New Roman", sizeHalfPoints: 21)]),
        ]
    }

    /// Gap between the two paragraphs' line boxes, in pt.
    private func paragraphGap(_ pdf: PDFDocument, registryEntry: String) throws -> CGFloat {
        let boxes = RenderGeometry.lineBoxes(pdf: pdf, page: 0)
        guard boxes.count == 2 else {
            XCTFail("[\(registryEntry)] expected exactly 2 line boxes, got \(boxes.count)")
            throw XCTSkip("fixture shape broken — see failure above")
        }
        return boxes[0].midY - boxes[1].midY
    }

    /// Registry #2 — `ParagraphPayload.spacingBefore` +240 twips on ¶2:
    /// the inter-paragraph gap grows by 240/20 = 12.0 pt.
    func testProbeSpacingBefore() throws {
        try VisualDiffHarness.requireGate()
        let baseline = try RenderProbeHarness.renderedPDF(
            ops: twoParagraphOps(),
            baseName: "probe-spacing-before-base")
        let perturbed = try RenderProbeHarness.renderedPDF(
            ops: twoParagraphOps(p2Before: 240),
            baseName: "probe-spacing-before-pert")

        let baseGap = try paragraphGap(baseline, registryEntry: "registry #2 spacingBefore")
        let pertGap = try paragraphGap(perturbed, registryEntry: "registry #2 spacingBefore")
        XCTAssertGreaterThan(pertGap, baseGap,
            "[registry #2 spacingBefore] gap must increase")
        let predicted: CGFloat = 240 / 20  // 12.0 pt
        XCTAssertEqual(pertGap - baseGap, predicted,
                       accuracy: RenderProbeHarness.tolerance(forPredicted: predicted),
                       "[registry #2 spacingBefore] measured Δgap vs predicted 12.0 pt")
        print("[render-effect evidence] spacingBefore: gap \(baseGap) → \(pertGap) pt, "
              + "Δ \(pertGap - baseGap) pt (predicted \(predicted))")
    }

    /// Registry #3 — `ParagraphPayload.spacingAfter` +240 twips on ¶1:
    /// the same observable via the other paragraph's property.
    func testProbeSpacingAfter() throws {
        try VisualDiffHarness.requireGate()
        let baseline = try RenderProbeHarness.renderedPDF(
            ops: twoParagraphOps(),
            baseName: "probe-spacing-after-base")
        let perturbed = try RenderProbeHarness.renderedPDF(
            ops: twoParagraphOps(p1After: 240),
            baseName: "probe-spacing-after-pert")

        let baseGap = try paragraphGap(baseline, registryEntry: "registry #3 spacingAfter")
        let pertGap = try paragraphGap(perturbed, registryEntry: "registry #3 spacingAfter")
        XCTAssertGreaterThan(pertGap, baseGap,
            "[registry #3 spacingAfter] gap must increase")
        let predicted: CGFloat = 240 / 20  // 12.0 pt
        XCTAssertEqual(pertGap - baseGap, predicted,
                       accuracy: RenderProbeHarness.tolerance(forPredicted: predicted),
                       "[registry #3 spacingAfter] measured Δgap vs predicted 12.0 pt")
        print("[render-effect evidence] spacingAfter: gap \(baseGap) → \(pertGap) pt, "
              + "Δ \(pertGap - baseGap) pt (predicted \(predicted))")
    }

    // MARK: - Registry row 4 (task 2.4)

    /// Registry #4 — `ParagraphPayload.spacingLine` with rule `auto`,
    /// 240 (single) → 360 (1.5×): median line pitch scales by ×1.5.
    func testProbeSpacingLineAuto() throws {
        try VisualDiffHarness.requireGate()

        func ops(line: Int) -> [OOXMLSwift.Operation] {
            [
                .appendParagraph(in: nil, paragraph: ParagraphPayload(
                    text: "", paraId: "P1",
                    spacingLine: line, spacingLineRule: "auto")),
                .setRuns(target: ElementID(rawString: "w14:paraId=P1"), runs: [RunPayload(
                    text: RenderProbeHarness.wrappingText(),
                    fontAscii: "Times New Roman", sizeHalfPoints: 21)]),
            ]
        }
        let baseline = try RenderProbeHarness.renderedPDF(
            ops: ops(line: 240), baseName: "probe-linespacing-240")
        let perturbed = try RenderProbeHarness.renderedPDF(
            ops: ops(line: 360), baseName: "probe-linespacing-360")

        let basePitch = try XCTUnwrap(
            RenderGeometry.medianLinePitch(pdf: baseline, page: 0),
            "[registry #4 spacingLine-auto] baseline page must have measurable lines")
        let pertPitch = try XCTUnwrap(
            RenderGeometry.medianLinePitch(pdf: perturbed, page: 0),
            "[registry #4 spacingLine-auto] perturbed page must have measurable lines")

        XCTAssertGreaterThan(pertPitch, basePitch,
            "[registry #4 spacingLine-auto] pitch must increase with the multiplier")
        // Predicted perturbed pitch = baseline × (360/240); assert the delta
        // within Decision 5 tolerance of the predicted delta.
        let predictedDelta = basePitch * 0.5
        XCTAssertEqual(pertPitch - basePitch, predictedDelta,
                       accuracy: RenderProbeHarness.tolerance(forPredicted: predictedDelta),
                       "[registry #4 spacingLine-auto] Δpitch vs predicted ×1.5 scaling")
        print("[render-effect evidence] spacingLine-auto: pitch \(basePitch) → \(pertPitch) pt, "
              + "ratio \(pertPitch / basePitch) (predicted 1.5)")
    }

    // MARK: - Registry row 5 (task 2.5)

    /// Registry #5 — `ParagraphPayload.indentFirstLineChars` 0 → 100
    /// (1 char): the first line box shifts right relative to the following
    /// lines by approximately one character advance (≈ run font size in pt
    /// for a CJK full-width character).
    func testProbeIndentFirstLineChars() throws {
        try VisualDiffHarness.requireGate()

        // CJK body text that wraps into several lines at 10.5pt on A4.
        let cjkText = String(repeating: "統計學的基本概念與心理測驗的信效度分析方法論。", count: 12)
        func ops(firstLineChars: Int?) -> [OOXMLSwift.Operation] {
            [
                .appendParagraph(in: nil, paragraph: ParagraphPayload(
                    text: "", paraId: "P1",
                    indentFirstLineChars: firstLineChars)),
                .setRuns(target: ElementID(rawString: "w14:paraId=P1"), runs: [RunPayload(
                    text: cjkText,
                    fontEastAsia: "ＭＳ 明朝", sizeHalfPoints: 21)]),
            ]
        }
        let baseline = try RenderProbeHarness.renderedPDF(
            ops: ops(firstLineChars: nil), baseName: "probe-flchars-0")
        let perturbed = try RenderProbeHarness.renderedPDF(
            ops: ops(firstLineChars: 100), baseName: "probe-flchars-100")

        /// First-line x-offset relative to the body lines (median minX of
        /// the following lines).
        func firstLineOffset(_ pdf: PDFDocument, label: String) throws -> CGFloat {
            let boxes = RenderGeometry.lineBoxes(pdf: pdf, page: 0)
            guard boxes.count >= 3 else {
                XCTFail("[\(label)] expected ≥3 line boxes for a stable body reference, got \(boxes.count)")
                throw XCTSkip("fixture shape broken — see failure above")
            }
            let bodyMinXs = boxes.dropFirst().map(\.minX).sorted()
            let medianBodyMinX = bodyMinXs[bodyMinXs.count / 2]
            return boxes[0].minX - medianBodyMinX
        }

        let baseOffset = try firstLineOffset(baseline, label: "registry #5 indentFirstLineChars baseline")
        let pertOffset = try firstLineOffset(perturbed, label: "registry #5 indentFirstLineChars perturbed")

        XCTAssertGreaterThan(pertOffset, baseOffset,
            "[registry #5 indentFirstLineChars] first line must shift right")
        // One full-width character advance at sz 21 (10.5 pt font).
        let predicted: CGFloat = 10.5
        XCTAssertEqual(pertOffset - baseOffset, predicted,
                       accuracy: RenderProbeHarness.tolerance(forPredicted: predicted),
                       "[registry #5 indentFirstLineChars] Δoffset vs predicted one char advance")
        print("[render-effect evidence] indentFirstLineChars: first-line offset "
              + "\(baseOffset) → \(pertOffset) pt, Δ \(pertOffset - baseOffset) pt (predicted \(predicted))")
    }

    // MARK: - Registry row 6 (task 2.6)

    /// Registry #6 — `RunPayload.sizeHalfPoints` 21 → 42 (10.5 → 21 pt):
    /// line-box height and median line pitch both increase; pitch
    /// approximately doubles.
    func testProbeSizeHalfPoints() throws {
        try VisualDiffHarness.requireGate()

        func ops(sz: Int) -> [OOXMLSwift.Operation] {
            [
                .appendParagraph(in: nil, paragraph: ParagraphPayload(
                    text: "", paraId: "P1",
                    spacingLine: 240, spacingLineRule: "auto")),
                .setRuns(target: ElementID(rawString: "w14:paraId=P1"), runs: [RunPayload(
                    text: RenderProbeHarness.wrappingText(sentences: 20),
                    fontAscii: "Times New Roman", sizeHalfPoints: sz)]),
            ]
        }
        let baseline = try RenderProbeHarness.renderedPDF(
            ops: ops(sz: 21), baseName: "probe-sz-21")
        let perturbed = try RenderProbeHarness.renderedPDF(
            ops: ops(sz: 42), baseName: "probe-sz-42")

        let baseBoxes = RenderGeometry.lineBoxes(pdf: baseline, page: 0)
        let pertBoxes = RenderGeometry.lineBoxes(pdf: perturbed, page: 0)
        let baseHeight = baseBoxes.map(\.height).sorted()[baseBoxes.count / 2]
        let pertHeight = pertBoxes.map(\.height).sorted()[pertBoxes.count / 2]
        XCTAssertGreaterThan(pertHeight, baseHeight,
            "[registry #6 sizeHalfPoints] line-box height must increase with font size")

        let basePitch = try XCTUnwrap(
            RenderGeometry.medianLinePitch(pdf: baseline, page: 0),
            "[registry #6 sizeHalfPoints] baseline page must have measurable lines")
        let pertPitch = try XCTUnwrap(
            RenderGeometry.medianLinePitch(pdf: perturbed, page: 0),
            "[registry #6 sizeHalfPoints] perturbed page must have measurable lines")
        XCTAssertGreaterThan(pertPitch, basePitch,
            "[registry #6 sizeHalfPoints] pitch must increase with font size")
        // Doubling the font size doubles the auto line pitch.
        let predictedDelta = basePitch
        XCTAssertEqual(pertPitch - basePitch, predictedDelta,
                       accuracy: RenderProbeHarness.tolerance(forPredicted: predictedDelta),
                       "[registry #6 sizeHalfPoints] Δpitch vs predicted ×2 scaling")
        print("[render-effect evidence] sizeHalfPoints: pitch \(basePitch) → \(pertPitch) pt "
              + "(ratio \(pertPitch / basePitch), predicted 2.0); "
              + "line height \(baseHeight) → \(pertHeight) pt")
    }

    // MARK: - Registry row 7 (task 2.7)

    /// Registry #7 — `SectionPayload` page size + margins: `marginTop`
    /// +567 twips (+1 cm). The page box must NOT change (595.3 × 841.9 pt
    /// for A4, exact); the first line box moves down by 567/20 = 28.35 pt.
    func testProbePageMargins() throws {
        try VisualDiffHarness.requireGate()

        func ops(marginTop: Int) -> [OOXMLSwift.Operation] {
            [
                .appendParagraph(in: nil, paragraph: ParagraphPayload(text: "", paraId: "P1")),
                .setRuns(target: ElementID(rawString: "w14:paraId=P1"), runs: [RunPayload(
                    text: RenderProbeHarness.wrappingText(),
                    fontAscii: "Times New Roman", sizeHalfPoints: 21)]),
                .setSectionProperties(at: nil, section: SectionPayload(
                    pageWidth: 11906, pageHeight: 16838, marginTop: marginTop)),
            ]
        }
        let baseline = try RenderProbeHarness.renderedPDF(
            ops: ops(marginTop: 1440), baseName: "probe-margin-1440")
        let perturbed = try RenderProbeHarness.renderedPDF(
            ops: ops(marginTop: 2007), baseName: "probe-margin-2007")

        // Core claim: margins must NOT change the page box (exact equality).
        // Absolute size ≈ nominal twips/20 within 0.15 pt — probe-measured:
        // Word renders w:pgSz 11906×16838 as 595.2×841.92 pt, a small
        // quantization off the raw arithmetic 595.3×841.9 (registry #7
        // evidence; recorded, not hidden by a loosened core assertion).
        let baseBox = try XCTUnwrap(RenderGeometry.pageBox(pdf: baseline, page: 0))
        let pertBox = try XCTUnwrap(RenderGeometry.pageBox(pdf: perturbed, page: 0))
        XCTAssertEqual(baseBox, pertBox,
                       "[registry #7 pageMargins] margins must not change the page box")
        XCTAssertEqual(baseBox.width, 11906.0 / 20, accuracy: 0.15,
                       "[registry #7 pageMargins] rendered width ≈ nominal twips/20")
        XCTAssertEqual(baseBox.height, 16838.0 / 20, accuracy: 0.15,
                       "[registry #7 pageMargins] rendered height ≈ nominal twips/20")

        // First line moves DOWN (smaller y in bottom-left origin) by 28.35 pt.
        let baseFirst = try XCTUnwrap(RenderGeometry.lineBoxes(pdf: baseline, page: 0).first,
                                      "[registry #7 pageMargins] baseline must have text lines")
        let pertFirst = try XCTUnwrap(RenderGeometry.lineBoxes(pdf: perturbed, page: 0).first,
                                      "[registry #7 pageMargins] perturbed must have text lines")
        XCTAssertLessThan(pertFirst.midY, baseFirst.midY,
            "[registry #7 pageMargins] larger top margin must push the first line down")
        let predicted: CGFloat = (2007 - 1440) / 20  // 28.35 pt
        XCTAssertEqual(baseFirst.midY - pertFirst.midY, predicted,
                       accuracy: RenderProbeHarness.tolerance(forPredicted: predicted),
                       "[registry #7 pageMargins] Δy vs predicted 28.35 pt")
        print("[render-effect evidence] pageMargins: page box \(baseBox.width)×\(baseBox.height) pt "
              + "unchanged; first-line midY \(baseFirst.midY) → \(pertFirst.midY), "
              + "Δ \(baseFirst.midY - pertFirst.midY) pt (predicted \(predicted))")
    }
}
