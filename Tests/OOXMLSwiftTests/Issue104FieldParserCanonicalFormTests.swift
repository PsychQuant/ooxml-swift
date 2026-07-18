import XCTest
@testable import OOXMLSwift

/// Tests for `FieldParser.parse(paragraph:)` canonical 5-run fldChar form
/// detection (PsychQuant/che-word-mcp#104, post-#94 follow-up).
///
/// **Pre-fix bug**: `FieldParser.parse(paragraph:)` (FieldParser.swift:97-110)
/// only handled the v2.0.0 baked form where ALL 5 `<w:r>` elements (begin /
/// instrText / separate / cachedValue / end) live inside ONE `Run.rawXML`. For
/// the canonical 5-run form (each fldChar element in its own `<w:r>` — which is
/// what DocxReader produces after disk roundtrip and what native Word emits),
/// the line-101 guard `rawXML.contains("fldChar")` filtered out the instrText
/// run (whose rawXML has NO fldChar element). Result: parser returned 0 fields
/// even when SEQ structure was clearly present in the document.
///
/// **Post-fix**: two-phase parse — Phase-1 baked form (existing path) → Phase-2
/// `parseFiveRunSpan(paragraph:)` state machine fallback. State machine walks
/// runs as: idle → seenBegin → seenInstrText → seenSeparate → seenCached →
/// emit on `end`.
///
/// **Sibling gaps** (out of scope for #104):
/// - PsychQuant/ooxml-swift#25 — header/footer/footnote/endnote walker
/// - PsychQuant/ooxml-swift#26 — paragraph wrapper-path coverage (inline SDT /
///   hyperlink / fieldSimple / alternateContent)
final class Issue104FieldParserCanonicalFormTests: XCTestCase {

    // MARK: - 104.1 Primary RED reproducer: roundtrip via Writer → Reader

    /// The exact reproducer from che-word-mcp#104: build a paragraph via
    /// `wrapCaptionSequenceFields`, write through DocxWriter, read back through
    /// DocxReader, then call `FieldParser.parse(paragraph:)`. Pre-fix returns
    /// `[]`; post-fix returns one `.sequence` ParsedField with identifier
    /// "Figure".
    func testFieldParserDetectsCanonical5RunSEQAfterRoundTrip() throws {
        let tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("Issue104RoundTrip-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: tempDir) }

        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "圖 4-1：前後期報酬率分布直方圖"))
        ]
        let pattern = try NSRegularExpression(pattern: "圖 4-(\\d+)：")
        _ = try doc.wrapCaptionSequenceFields(pattern: pattern, sequenceName: "Figure")

        // Write + Read round-trip — this is what surfaces the canonical 5-run form
        let docxURL = tempDir.appendingPathComponent("test.docx")
        try DocxWriter.write(doc, to: docxURL)
        let reloaded = try DocxReader.read(from: docxURL)

        guard case .paragraph(let p) = reloaded.body.children[0] else {
            return XCTFail("expected paragraph after roundtrip")
        }

        let fields = FieldParser.parse(paragraph: p)
        XCTAssertEqual(fields.count, 1,
            "Pre-fix returns 0 (canonical 5-run instrText run filtered out by line-101 guard); post-fix should return exactly 1 ParsedField")

        guard let field = fields.first else {
            return XCTFail("no field parsed")
        }
        guard case .sequence(let seq) = field.field else {
            return XCTFail("expected .sequence field, got \(field.field)")
        }
        XCTAssertEqual(seq.identifier, "Figure",
            "ParsedField should expose identifier 'Figure' from instrText 'SEQ Figure \\* ARABIC'")
    }

    // MARK: - 104.1.1 Baked-vs-canonical discriminator invariant (#33)

    /// `wrapCaptionSequenceFields` emits the baked form before save: the parsed
    /// cached-result run is the same run whose rawXML embeds the full field
    /// block, including fldChar markers. `updateAllFields` depends on this
    /// invariant to choose the baked rewrite path (#33).
    func testBakedFormCachedRunRawXMLAlwaysContainsFldCharBeforeSave() throws {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "Figure 7. Distribution"))
        ]
        let pattern = try NSRegularExpression(pattern: "Figure (\\d+)\\.")
        _ = try doc.wrapCaptionSequenceFields(pattern: pattern, sequenceName: "Figure")

        guard case .paragraph(let paragraph) = doc.body.children[0] else {
            return XCTFail("expected paragraph")
        }

        let fields = FieldParser.parse(paragraph: paragraph)
        XCTAssertEqual(fields.count, 1)
        guard let field = fields.first,
              let cachedRunIdx = field.cachedResultRunIdx else {
            return XCTFail("expected parsed field with cached result run")
        }

        XCTAssertEqual(field.startRunIdx, cachedRunIdx,
            "baked form should report the same run for field start and cached result")
        let cachedRawXML = paragraph.runs[cachedRunIdx].rawXML
        XCTAssertNotNil(cachedRawXML)
        XCTAssertTrue(cachedRawXML?.contains("fldChar") ?? false,
            "baked-form cached run must carry fldChar markers for the discriminator")
    }

    /// After Writer -> Reader roundtrip, the same field is canonical 5-run
    /// form: cached-result run is a dedicated value run and must not contain
    /// fldChar. This prevents the baked-form discriminator from accidentally
    /// routing canonical fields through the rawXML field-block regex (#33).
    func testRoundTripCanonicalFormCachedRunRawXMLDoesNotContainFldChar() throws {
        let tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("Issue33RoundTripCanonical-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: tempDir) }

        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "Figure 7. Distribution"))
        ]
        let pattern = try NSRegularExpression(pattern: "Figure (\\d+)\\.")
        _ = try doc.wrapCaptionSequenceFields(pattern: pattern, sequenceName: "Figure")

        let docxURL = tempDir.appendingPathComponent("test.docx")
        try DocxWriter.write(doc, to: docxURL)
        let reloaded = try DocxReader.read(from: docxURL)

        guard case .paragraph(let paragraph) = reloaded.body.children[0] else {
            return XCTFail("expected paragraph after roundtrip")
        }

        let fields = FieldParser.parse(paragraph: paragraph)
        XCTAssertEqual(fields.count, 1)
        guard let field = fields.first,
              let cachedRunIdx = field.cachedResultRunIdx else {
            return XCTFail("expected parsed field with cached result run")
        }

        XCTAssertNotEqual(field.startRunIdx, cachedRunIdx,
            "canonical form should use a dedicated cached-result run")
        XCTAssertFalse(paragraph.runs[cachedRunIdx].rawXML?.contains("fldChar") ?? false,
            "canonical cached run must not look like a baked field block")
    }

    // MARK: - 104.1.2 Defensive large-paragraph scan (#32)

    /// DoS regression guard for #32: large paragraphs with many runs should not
    /// allocate a joined fragment string per run. The parser should complete a
    /// 100k-run empty paragraph within a conservative local-test budget.
    func testFieldParserHandlesHundredThousandEmptyRunsWithinBudget() {
        var paragraph = Paragraph()
        paragraph.runs = Array(repeating: Run(text: ""), count: 100_000)

        let started = Date()
        let fields = FieldParser.parse(paragraph: paragraph)
        let elapsed = Date().timeIntervalSince(started)

        XCTAssertTrue(fields.isEmpty)
        XCTAssertLessThan(elapsed, 5.0,
            "FieldParser.parse should not do expensive per-run fragment joining for empty runs")
    }

    // MARK: - 104.2 End-to-end: updateAllFields finds and updates SEQ after roundtrip

    /// Drives the actual MCP user scenario: roundtrip → `updateAllFields()`
    /// should return populated identifier dict, not an empty one. Pre-fix
    /// returns `[:]` (silent no-op); post-fix returns `["Figure": 1]`.
    func testUpdateAllFieldsHandlesCanonical5RunFormAfterRoundTrip() throws {
        let tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("Issue104UpdateAfterRoundTrip-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: tempDir) }

        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "圖 4-1：前後期報酬率分布直方圖"))
        ]
        let pattern = try NSRegularExpression(pattern: "圖 4-(\\d+)：")
        _ = try doc.wrapCaptionSequenceFields(pattern: pattern, sequenceName: "Figure")

        let docxURL = tempDir.appendingPathComponent("test.docx")
        try DocxWriter.write(doc, to: docxURL)
        var reloaded = try DocxReader.read(from: docxURL)

        let result = reloaded.updateAllFields()

        XCTAssertEqual(result, ["Figure": 1],
            "Pre-fix returns [:] because FieldParser misses canonical 5-run form; post-fix must return the populated counter dict (#104)")
    }

    // MARK: - 104.3 Native-Word 5-run form (no roundtrip dependency)

    /// Construct a paragraph by hand with 5 separate Run objects each carrying
    /// ONE fldChar / instrText / `<w:t>` element in rawXML — mimicking how
    /// native Word emits SEQ fields without going through
    /// `wrapCaptionSequenceFields`. Pins that the fix handles arbitrary
    /// 5-run paragraphs, not just our own wrap output.
    func testFieldParserHandlesNativeWord5RunSEQ() {
        // Mirror the exact run structure from che-word-mcp#104 reproducer:
        //
        //   <w:r><w:t>圖 4-</w:t></w:r>
        //   <w:r><w:t>：前後期</w:t></w:r>
        //   <w:r><w:t>報酬率分布直方圖與常態分配曲線比較</w:t></w:r>
        //   <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        //   <w:r><w:instrText xml:space="preserve"> SEQ Figure </w:instrText></w:r>
        //   <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        //   <w:r><w:t xml:space="preserve">1</w:t></w:r>
        //   <w:r><w:fldChar w:fldCharType="end"/></w:r>
        //
        // Each fldChar / instrText / `<w:t>` lives in its own Run.rawXML, which
        // is the canonical roundtrip / native-Word emission form.

        var caption1 = Run(text: "圖 4-")
        var caption2 = Run(text: "：前後期")
        var caption3 = Run(text: "報酬率分布直方圖與常態分配曲線比較")

        var beginRun = Run(text: "")
        beginRun.rawXML = "<w:fldChar w:fldCharType=\"begin\"/>"

        var instrRun = Run(text: "")
        instrRun.rawXML = "<w:instrText xml:space=\"preserve\"> SEQ Figure </w:instrText>"

        var separateRun = Run(text: "")
        separateRun.rawXML = "<w:fldChar w:fldCharType=\"separate\"/>"

        var cachedRun = Run(text: "1")
        cachedRun.rawXML = "<w:t xml:space=\"preserve\">1</w:t>"

        var endRun = Run(text: "")
        endRun.rawXML = "<w:fldChar w:fldCharType=\"end\"/>"

        var para = Paragraph()
        para.runs = [caption1, caption2, caption3, beginRun, instrRun, separateRun, cachedRun, endRun]

        let fields = FieldParser.parse(paragraph: para)
        XCTAssertEqual(fields.count, 1,
            "Native 5-run SEQ form must be detected (post-#104 fix). Pre-fix: 0 results; post-fix: 1 ParsedField")

        guard let field = fields.first else {
            return XCTFail("no field parsed")
        }
        guard case .sequence(let seq) = field.field else {
            return XCTFail("expected .sequence field")
        }
        XCTAssertEqual(seq.identifier, "Figure")

        // Span boundaries should point to the fldChar runs (not the caption text)
        XCTAssertEqual(field.startRunIdx, 3,
            "startRunIdx should point to the run containing fldCharType=begin (run index 3)")
        XCTAssertEqual(field.endRunIdx, 7,
            "endRunIdx should point to the run containing fldCharType=end (run index 7)")
        XCTAssertEqual(field.cachedResultRunIdx, 6,
            "cachedResultRunIdx should point to the standalone <w:t> run between separate and end (run index 6)")
    }

    // MARK: - 104.4 Native-Word 5-run + updateAllFields rewrite + emit roundtrip

    /// Hand-built native-Word 5-run paragraph where the cached run carries a
    /// non-nil `Run.rawXML` (mirroring native Word emission and any upstream
    /// tool that preserves the `<w:t>` raw form). Pre-fix-of-P1: the canonical
    /// branch in `processParagraph` only mutated `Run.text`, but `Run.toXML()`
    /// short-circuits on non-nil rawXML (Run.swift:246-248) so the emitted
    /// XML retained the **stale** cached value while `updateAllFields()` still
    /// reported a populated counter dict — silent desync between counter and
    /// disk content.
    ///
    /// **Pre-P1-fix**: `cachedRun.rawXML == "<w:t xml:space=\"preserve\">999</w:t>"`
    /// after `updateAllFields()` returns `["Figure": 1]`.
    /// **Post-P1-fix**: rawXML spliced to contain "1", emitted Paragraph XML
    /// shows "1" not "999". Counter dict and disk content stay in sync.
    ///
    /// Surfaced by Devil's Advocate during 6-AI verify of #104. This pins the
    /// invariant so future Run model changes can't silently regress this case.
    func testUpdateAllFieldsRewritesNativeWord5RunCachedRunRawXML() {
        var caption = Run(text: "圖 4-1：")

        var beginRun = Run(text: "")
        beginRun.rawXML = "<w:fldChar w:fldCharType=\"begin\"/>"

        var instrRun = Run(text: "")
        instrRun.rawXML = "<w:instrText xml:space=\"preserve\"> SEQ Figure </w:instrText>"

        var separateRun = Run(text: "")
        separateRun.rawXML = "<w:fldChar w:fldCharType=\"separate\"/>"

        // Stale cached value baked into rawXML — this is the P1 trigger
        var cachedRun = Run(text: "999")
        cachedRun.rawXML = "<w:t xml:space=\"preserve\">999</w:t>"

        var endRun = Run(text: "")
        endRun.rawXML = "<w:fldChar w:fldCharType=\"end\"/>"

        var para = Paragraph()
        para.runs = [caption, beginRun, instrRun, separateRun, cachedRun, endRun]

        var doc = WordDocument()
        doc.body.children = [.paragraph(para)]

        let result = doc.updateAllFields()
        XCTAssertEqual(result, ["Figure": 1],
            "counter dict should report Figure: 1 (this part already worked pre-P1-fix)")

        guard case .paragraph(let updatedPara) = doc.body.children[0] else {
            return XCTFail("expected paragraph after updateAllFields")
        }

        // Run.text mutated (this also worked pre-P1-fix)
        XCTAssertEqual(updatedPara.runs[4].text, "1",
            "Run.text should be updated to new counter value")

        // The P1 assertion — pre-fix this would still be "999" because
        // Run.toXML() short-circuits on non-nil rawXML.
        let cachedRawXML = updatedPara.runs[4].rawXML ?? ""
        XCTAssertTrue(cachedRawXML.contains(">1<"),
            "P1 fix: Run.rawXML must be spliced to contain new value '1', got: '\(cachedRawXML)'")
        XCTAssertFalse(cachedRawXML.contains(">999<"),
            "P1 fix: Run.rawXML must NOT contain stale value '999', got: '\(cachedRawXML)'")

        // Verify the emitted XML reflects the new value (the actual on-disk
        // surface a downstream consumer would see).
        let emittedRunXML = updatedPara.runs[4].toXML()
        XCTAssertTrue(emittedRunXML.contains(">1<"),
            "P1 fix: emitted run XML must contain new counter '1', got: '\(emittedRunXML)'")
        XCTAssertFalse(emittedRunXML.contains(">999<"),
            "P1 fix: emitted run XML must NOT contain stale '999', got: '\(emittedRunXML)'")

        // Preservation invariant: xml:space="preserve" attribute should survive
        // the splice (regression guard for the rawXML rewrite helper).
        XCTAssertTrue(cachedRawXML.contains("xml:space=\"preserve\""),
            "rawXML splice must preserve `xml:space=\"preserve\"` attribute")
    }
}
