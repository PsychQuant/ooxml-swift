import XCTest
@testable import OOXMLSwift

/// Tests for `WordDocument.wrapCaptionSequenceFields(...)` POSITIONAL placement
/// of SEQ field runs (PsychQuant/che-word-mcp#93, post-#62 follow-up).
///
/// Issue #93 reports SEQ field placed at END of paragraph instead of mid-caption:
///   - Expected: `「圖 4-」 + SEQ(1) + 「：caption」`
///   - Reported: `「圖 4-：caption」 + SEQ(1)`
///
/// Existing `WrapCaptionSequenceFieldsTests` only asserts `containsSEQ`
/// (presence anywhere in any run's rawXML), not run-position. These tests
/// pin the positional contract.
final class Issue93WrapCaptionSeqPlacementTests: XCTestCase {

    // MARK: - Helpers

    /// Reconstruct the user-visible text of a paragraph by inlining SEQ field
    /// `cachedResult` from rawXML where present. Mirrors the lib's internal
    /// `renderedTextWithSEQ` semantics so the test asserts what users actually
    /// see in Word, not just `Run.text` joined.
    private func renderedText(of para: Paragraph, sequenceName: String) -> String {
        let needle = "SEQ \(sequenceName)"
        var parts: [String] = []
        for run in para.runs {
            if let raw = run.rawXML, raw.contains(needle) {
                if let cached = extractCachedResult(raw) {
                    parts.append(cached)
                }
            } else {
                parts.append(run.text)
            }
        }
        return parts.joined()
    }

    private func extractCachedResult(_ rawXML: String) -> String? {
        guard let sepRange = rawXML.range(of: "fldCharType=\"separate\"") else { return nil }
        let after = rawXML[sepRange.upperBound...]
        guard let openRange = after.range(of: "<w:t") else { return nil }
        let afterOpen = after[openRange.upperBound...]
        guard let gtRange = afterOpen.range(of: ">") else { return nil }
        let body = afterOpen[gtRange.upperBound...]
        guard let closeRange = body.range(of: "</w:t>") else { return nil }
        return String(body[..<closeRange.lowerBound])
    }

    // MARK: - 93.1 Single-run paragraph: SEQ must replace digit in-place

    func testSingleRunCaptionPlacesSEQInPlaceOfDigit() throws {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "圖 4-1：架構圖"))
        ]
        let pattern = try NSRegularExpression(pattern: "圖 4-(\\d+)：")
        _ = try doc.wrapCaptionSequenceFields(pattern: pattern, sequenceName: "Figure")

        guard case .paragraph(let p) = doc.body.children[0] else {
            return XCTFail("expected paragraph")
        }

        // Visual rendering must be identical to original — SEQ replaces the
        // digit at the same position, not appended at end.
        let rendered = renderedText(of: p, sequenceName: "Figure")
        XCTAssertEqual(rendered, "圖 4-1：架構圖",
            "SEQ field must render in original position; got '\(rendered)' (#93 placement bug)")

        // Structural pin: SEQ run must come BEFORE the trailing "：架構圖" text run.
        let seqRunIdx = p.runs.firstIndex { ($0.rawXML ?? "").contains("SEQ Figure") }
        XCTAssertNotNil(seqRunIdx, "SEQ Figure run must exist")

        let trailingRunIdx = p.runs.firstIndex { $0.text.contains("：") || $0.text.contains("架構圖") }
        XCTAssertNotNil(trailingRunIdx, "trailing text run with caption body must exist")

        if let seq = seqRunIdx, let trailing = trailingRunIdx {
            XCTAssertLessThan(seq, trailing,
                "SEQ run (idx \(seq)) must precede trailing text run (idx \(trailing))")
        }
    }

    // MARK: - 93.2 Multi-run paragraph (digit and trailing in same/different runs)

    func testMultiRunCaptionPlacesSEQAtMatchPosition() throws {
        var doc = WordDocument()
        // Pre-built multi-run paragraph mimicking what Word emits when the
        // user types caption with mid-stream formatting changes:
        //   Run 0: "圖 4-1"  (Heading style)
        //   Run 1: "："      (default style)
        //   Run 2: "前後期報酬率分布直方圖"  (caption body)
        let para = Paragraph(runs: [
            Run(text: "圖 4-1"),
            Run(text: "："),
            Run(text: "前後期報酬率分布直方圖")
        ])
        doc.body.children = [.paragraph(para)]

        let pattern = try NSRegularExpression(pattern: "圖 4-(\\d+)：")
        _ = try doc.wrapCaptionSequenceFields(pattern: pattern, sequenceName: "Figure")

        guard case .paragraph(let p) = doc.body.children[0] else {
            return XCTFail("expected paragraph")
        }

        let rendered = renderedText(of: p, sequenceName: "Figure")
        XCTAssertEqual(rendered, "圖 4-1：前後期報酬率分布直方圖",
            "SEQ field must render in mid-caption position; got '\(rendered)' (#93 placement bug)")

        // Structural pin: order must be [..., "圖 4-", SEQ, "：", ...]
        let seqRunIdx = p.runs.firstIndex { ($0.rawXML ?? "").contains("SEQ Figure") }
        XCTAssertNotNil(seqRunIdx, "SEQ Figure run must exist")

        let colonRunIdx = p.runs.firstIndex { $0.text.contains("：") }
        XCTAssertNotNil(colonRunIdx, "「：」 run must exist post-fix")

        let trailingRunIdx = p.runs.firstIndex { $0.text.contains("前後期") }
        XCTAssertNotNil(trailingRunIdx, "trailing caption body run must exist")

        if let seq = seqRunIdx, let colon = colonRunIdx {
            XCTAssertLessThan(seq, colon,
                "SEQ run (idx \(seq)) must precede 「：」 run (idx \(colon)) — #93 reproducer")
        }
    }

    // MARK: - 93.5 Source-loaded paragraph with explicit Run.position > 0

    /// Reproduces the user's actual bug pattern: real Word documents have
    /// runs with explicit `position > 0` (assigned by the Reader from
    /// document order). When the splice logic copies that position to both
    /// pre/post halves but creates the SEQ run with default `position = nil`,
    /// the writer's bifurcated emit (positioned-list vs legacy post-content)
    /// puts SEQ at the END of the paragraph. Pre-fix bug; post-fix expects
    /// SEQ to inherit the source position.
    func testWrapWithExplicitRunPositionPreservesSplicePosition() throws {
        var doc = WordDocument()
        // Construct a Run with explicit position > 0, mimicking what the
        // Reader does for source-loaded Word docs.
        var sourceRun = Run(text: "圖 4-1：前後期報酬率分布直方圖")
        sourceRun.position = 1
        let para = Paragraph(runs: [sourceRun])
        doc.body.children = [.paragraph(para)]

        let pattern = try NSRegularExpression(pattern: "圖 4-(\\d+)：")
        _ = try doc.wrapCaptionSequenceFields(pattern: pattern, sequenceName: "Figure")

        guard case .paragraph(let p) = doc.body.children[0] else {
            return XCTFail("expected paragraph")
        }

        // Locate the runs after splice — order must be prefix → SEQ → suffix.
        let prefixIdx = p.runs.firstIndex { $0.text == "圖 4-" }
        let seqIdx = p.runs.firstIndex { ($0.rawXML ?? "").contains("SEQ Figure") }
        let suffixIdx = p.runs.firstIndex { $0.text.hasPrefix("：") }

        XCTAssertNotNil(prefixIdx, "prefix run '圖 4-' must exist")
        XCTAssertNotNil(seqIdx, "SEQ run with rawXML containing 'SEQ Figure' must exist")
        XCTAssertNotNil(suffixIdx, "suffix run '：...' must exist")

        if let pre = prefixIdx, let seq = seqIdx, let suf = suffixIdx {
            XCTAssertLessThan(pre, seq,
                "prefix idx (\(pre)) must precede SEQ idx (\(seq))")
            XCTAssertLessThan(seq, suf,
                "SEQ idx (\(seq)) must precede suffix idx (\(suf)) — array order check")
        }

        // Critical positional invariant — SEQ run must inherit position from
        // source so it lands in the SAME emit section (positioned-list vs
        // legacy post-content) as preText/postText. If SEQ has nil position
        // while pre/post have position > 0, the writer's bifurcated emit
        // puts SEQ at end of paragraph (the actual #93 bug).
        if let seq = seqIdx, let pre = prefixIdx {
            let seqPos = p.runs[seq].position
            let prePos = p.runs[pre].position
            XCTAssertEqual(seqPos, prePos,
                "SEQ run position (\(String(describing: seqPos))) must match prefix position (\(String(describing: prePos))) so they emit in the same section")
        }
    }

    // MARK: - 24.1 Bookmark wrapper must stay with positioned SEQ splice

    func testInsertBookmarkMarkersInheritSourceRunPosition() throws {
        var doc = WordDocument()
        var sourceRun = Run(text: "圖 4-1：前後期報酬率分布直方圖")
        sourceRun.position = 1
        doc.body.children = [.paragraph(Paragraph(runs: [sourceRun]))]

        let pattern = try NSRegularExpression(pattern: "圖 4-(\\d+)：")
        _ = try doc.wrapCaptionSequenceFields(
            pattern: pattern,
            sequenceName: "Figure",
            insertBookmark: true,
            bookmarkTemplate: "fig_${number}"
        )

        guard case .paragraph(let p) = doc.body.children[0] else {
            return XCTFail("expected paragraph")
        }

        let bookmarkStartIdx = p.runs.firstIndex { ($0.rawXML ?? "").contains("bookmarkStart") }
        let seqIdx = p.runs.firstIndex { ($0.rawXML ?? "").contains("SEQ Figure") }
        let bookmarkEndIdx = p.runs.firstIndex { ($0.rawXML ?? "").contains("bookmarkEnd") }
        let suffixIdx = p.runs.firstIndex { $0.text.hasPrefix("：") }

        XCTAssertNotNil(bookmarkStartIdx, "bookmarkStart run must exist")
        XCTAssertNotNil(seqIdx, "SEQ run must exist")
        XCTAssertNotNil(bookmarkEndIdx, "bookmarkEnd run must exist")
        XCTAssertNotNil(suffixIdx, "suffix run must exist")

        if let start = bookmarkStartIdx, let seq = seqIdx, let end = bookmarkEndIdx, let suffix = suffixIdx {
            XCTAssertLessThan(start, seq, "bookmarkStart must immediately precede the SEQ run")
            XCTAssertLessThan(seq, end, "bookmarkEnd must follow the SEQ run")
            XCTAssertLessThan(end, suffix, "bookmark wrapper must remain before trailing caption text")
            XCTAssertEqual(p.runs[start].position, sourceRun.position)
            XCTAssertEqual(p.runs[end].position, sourceRun.position)
        }
    }

    // MARK: - 93.4 ROUND-TRIP test: write to docx, read back, verify position

    /// The user's reproducer (kiki830621/collaboration_guo_analysis#5) observed
    /// the bug AFTER `save_document` + re-`open_document` cycle. The in-memory
    /// `tryWrapParagraph` splice may produce correct runs, but the round-trip
    /// through DocxWriter → DocxReader could reorder them. This test exercises
    /// the full round-trip path.
    func testSingleRunCaptionRoundTripPreservesSEQPosition() throws {
        let tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("Issue93RoundTrip-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: tempDir) }

        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "圖 4-1：前後期報酬率分布直方圖"))
        ]
        let pattern = try NSRegularExpression(pattern: "圖 4-(\\d+)：")
        _ = try doc.wrapCaptionSequenceFields(pattern: pattern, sequenceName: "Figure")

        // Write + Read round-trip
        let docxURL = tempDir.appendingPathComponent("test.docx")
        try DocxWriter.write(doc, to: docxURL)
        let reloaded = try DocxReader.read(from: docxURL)

        guard case .paragraph(let p) = reloaded.body.children[0] else {
            return XCTFail("expected paragraph after roundtrip")
        }

        // Build a position-ordered sequence of "logical units" from the
        // round-tripped runs. After roundtrip, the SEQ field's 5-run block
        // re-parses as 5 separate Runs (begin/instrText/separate/result/end).
        // The cachedResult lives in the run with text="1". Locate the
        // boundary runs by their text content.
        let prefixRun = p.runs.firstIndex { $0.text == "圖 4-" || $0.text.hasPrefix("圖 4-") }
        let cachedResultRun = p.runs.firstIndex { run in
            // The SEQ result run after roundtrip has text="1" and
            // sits between the separate/end fldChars. It does NOT have
            // "SEQ Figure" in its rawXML (that's in the instrText sibling).
            run.text == "1"
        }
        let suffixRun = p.runs.firstIndex { $0.text.hasPrefix("：") || $0.text.contains("：前後期") }

        XCTAssertNotNil(prefixRun, "prefix run '圖 4-' must exist after roundtrip")
        XCTAssertNotNil(cachedResultRun, "SEQ cachedResult run (text='1') must exist after roundtrip")
        XCTAssertNotNil(suffixRun, "suffix run '：前後期...' must exist after roundtrip")

        // PRIMARY ASSERTION: the run ORDER must be prefix → cachedResult → suffix.
        // Pre-fix bug: prefix → suffix → cachedResult (SEQ at end).
        if let pre = prefixRun, let cr = cachedResultRun, let suf = suffixRun {
            XCTAssertLessThan(pre, cr,
                "prefix run (\(pre)) must precede SEQ cachedResult run (\(cr))")
            XCTAssertLessThan(cr, suf,
                "ROUNDTRIP BUG: SEQ cachedResult run (\(cr)) must precede suffix '：前後期...' run (\(suf)). " +
                "Pre-fix: SEQ goes to end of paragraph because seqRun.position was nil while " +
                "preText/postText inherited source's position>0, splitting them across emit sections.")
        }

        // Confirm the SEQ instrText survives round-trip somewhere — Word
        // needs this to recognize the field on F9 / first-open update.
        // The parser stores `<w:instrText>` in `Run.rawElements` rather
        // than `rawXML` or `text`, so check all three surfaces.
        let hasSeqInstrText = p.runs.contains { run in
            if (run.rawXML ?? "").contains("SEQ Figure") { return true }
            if run.text.contains("SEQ Figure") { return true }
            if let raws = run.rawElements,
               raws.contains(where: { $0.xml.contains("SEQ Figure") }) { return true }
            return false
        }
        let hasSeqInFieldSimples = p.fieldSimples.contains { ($0.instr).contains("SEQ Figure") }
        XCTAssertTrue(hasSeqInstrText || hasSeqInFieldSimples,
            "SEQ Figure instrText must survive round-trip somewhere — runs (rawXML/text/rawElements) or fieldSimples")
    }

    // MARK: - 93.3 Run with prefix-only digit: split must preserve prefix

    func testCaptionWithDigitInsideMixedRunPreservesSurroundingText() throws {
        var doc = WordDocument()
        // Single run carrying entire caption text including digit:
        //   Run 0: "圖 4-1：時序圖"
        // Post-wrap should split into:
        //   Run 0: "圖 4-"   (prefix)
        //   Run 1: SEQ field with cachedResult "1"
        //   Run 2: "：時序圖"  (suffix)
        let para = Paragraph(runs: [
            Run(text: "圖 4-1：時序圖")
        ])
        doc.body.children = [.paragraph(para)]

        let pattern = try NSRegularExpression(pattern: "圖 4-(\\d+)：")
        _ = try doc.wrapCaptionSequenceFields(pattern: pattern, sequenceName: "Figure")

        guard case .paragraph(let p) = doc.body.children[0] else {
            return XCTFail("expected paragraph")
        }

        let rendered = renderedText(of: p, sequenceName: "Figure")
        XCTAssertEqual(rendered, "圖 4-1：時序圖",
            "Visual render must be unchanged; got '\(rendered)'")

        // The digit "1" must NOT appear as plain Run.text anywhere — it should
        // live ONLY inside the SEQ field's cachedResult.
        let plainTextHasDigit = p.runs.contains { run in
            run.rawXML == nil && run.text.contains("1") && !run.text.contains("圖 4-1：")
        }
        XCTAssertFalse(plainTextHasDigit,
            "Digit '1' must NOT remain in plain Run.text — it should be replaced by the SEQ field. " +
            "If this fails, the splice didn't happen (digit was kept) or wasn't done at all.")

        // Structural pin: prefix "圖 4-" must come before SEQ which must come before "：時序圖"
        let prefixIdx = p.runs.firstIndex { $0.text == "圖 4-" || $0.text.hasPrefix("圖 4-") }
        let seqIdx = p.runs.firstIndex { ($0.rawXML ?? "").contains("SEQ Figure") }
        let suffixIdx = p.runs.firstIndex { $0.text.hasPrefix("：") || $0.text == "：時序圖" }

        XCTAssertNotNil(prefixIdx, "prefix run '圖 4-' must exist after split")
        XCTAssertNotNil(seqIdx, "SEQ run must exist")
        XCTAssertNotNil(suffixIdx, "suffix run '：時序圖' must exist after split")

        if let pre = prefixIdx, let seq = seqIdx, let suf = suffixIdx {
            XCTAssertLessThan(pre, seq, "prefix (\(pre)) must precede SEQ (\(seq))")
            XCTAssertLessThan(seq, suf, "SEQ (\(seq)) must precede suffix (\(suf))")
        }
    }
}
