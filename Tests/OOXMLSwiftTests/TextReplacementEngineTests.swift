import XCTest
@testable import OOXMLSwift

final class TextReplacementEngineTests: XCTestCase {

    // MARK: - Single-run replace

    func testSingleRunReplace() throws {
        var runs = [Run(text: "Hello world")]
        let count = try TextReplacementEngine.replace(runs: &runs, find: "world", with: "Swift")
        XCTAssertEqual(count, 1)
        XCTAssertEqual(runs.count, 1)
        XCTAssertEqual(runs[0].text, "Hello Swift")
    }

    func testNoMatchReturnsZero() throws {
        var runs = [Run(text: "Hello world")]
        let count = try TextReplacementEngine.replace(runs: &runs, find: "xyz", with: "abc")
        XCTAssertEqual(count, 0)
        XCTAssertEqual(runs[0].text, "Hello world")
    }

    func testMultipleMatchesInSingleRun() throws {
        var runs = [Run(text: "ab ab ab")]
        let count = try TextReplacementEngine.replace(runs: &runs, find: "ab", with: "X")
        XCTAssertEqual(count, 3)
        XCTAssertEqual(runs[0].text, "X X X")
    }

    // MARK: - Cross-run matches (the primary #7 fix)

    func testCrossRunMatchThreeRuns() throws {
        // Simulates the thesis-#3 failure: "均值方程式：r_t = ..." split into 3 runs
        // with a phantom empty run in the middle from LaTeX stripping.
        var runs = [
            Run(text: "均值方程式："),
            Run(text: ""),
            Run(text: "r_t = ...")
        ]
        let count = try TextReplacementEngine.replace(
            runs: &runs,
            find: "均值方程式：r_t",
            with: "Mean: r_t"
        )
        XCTAssertEqual(count, 1)
        // Flattened result
        let flat = runs.map { $0.text }.joined()
        XCTAssertEqual(flat, "Mean: r_t = ...")
    }

    func testCrossRunReplacementInheritsStartRunFormatting() throws {
        var boldProps = RunProperties()
        boldProps.bold = true
        var italicProps = RunProperties()
        italicProps.italic = true

        var runs = [
            Run(text: "old", properties: boldProps),
            Run(text: " word", properties: italicProps)
        ]
        let count = try TextReplacementEngine.replace(
            runs: &runs,
            find: "old word",
            with: "new"
        )
        XCTAssertEqual(count, 1)
        // After replace: the first run should contain "new" with bold props;
        // the second run's suffix "" (consumed entirely), so its text is empty
        // but italic props still on the run.
        XCTAssertEqual(runs[0].text, "new")
        XCTAssertTrue(runs[0].properties.bold, "Replacement text should inherit start run's bold formatting")
        XCTAssertEqual(runs[1].text, "")
    }

    // MARK: - Case insensitive

    func testCaseInsensitiveReplace() throws {
        var runs = [Run(text: "Hello HELLO hello")]
        var opts = ReplaceOptions()
        opts.matchCase = false
        let count = try TextReplacementEngine.replace(runs: &runs, find: "hello", with: "X", options: opts)
        XCTAssertEqual(count, 3)
        XCTAssertEqual(runs[0].text, "X X X")
    }

    // MARK: - Regex

    func testRegexWithCaptureGroup() throws {
        var runs = [Run(text: "Chapter 4 and Chapter 10")]
        var opts = ReplaceOptions()
        opts.regex = true
        let count = try TextReplacementEngine.replace(
            runs: &runs,
            find: #"Chapter (\d+)"#,
            with: "Ch. $1",
            options: opts
        )
        XCTAssertEqual(count, 2)
        XCTAssertEqual(runs[0].text, "Ch. 4 and Ch. 10")
    }

    func testRegexInvalidPatternThrows() {
        var runs = [Run(text: "text")]
        var opts = ReplaceOptions()
        opts.regex = true
        XCTAssertThrowsError(try TextReplacementEngine.replace(
            runs: &runs,
            find: "[unclosed",
            with: "x",
            options: opts
        )) { error in
            if case ReplaceError.invalidRegex(let pattern) = error {
                XCTAssertEqual(pattern, "[unclosed")
            } else {
                XCTFail("Expected ReplaceError.invalidRegex, got \(error)")
            }
        }
    }

    // MARK: - Field-run skipping

    func testFieldRunsAreSkipped() throws {
        // A run with rawXML (field-related) shouldn't be flattened into matching.
        var fieldRun = Run(text: "")
        fieldRun.rawXML = "<w:fldChar w:fldCharType=\"begin\"/>"

        var runs = [
            Run(text: "before "),
            fieldRun,
            Run(text: "after")
        ]
        // Searching "before after" should succeed (the field run adds no chars to flat).
        let count = try TextReplacementEngine.replace(
            runs: &runs,
            find: "before after",
            with: "replaced"
        )
        XCTAssertEqual(count, 1)
        XCTAssertEqual(runs[0].text, "replaced")
        // Field run preserved
        XCTAssertEqual(runs[1].rawXML, "<w:fldChar w:fldCharType=\"begin\"/>")
        // End run's suffix was empty (match went to end)
        XCTAssertEqual(runs[2].text, "")
    }

    // MARK: - flattenRuns direct

    func testFlattenRunsBuildsOffsetMap() {
        let runs = [
            Run(text: "ab"),
            Run(text: "cd")
        ]
        let (flat, map) = TextReplacementEngine.flattenRuns(runs)
        XCTAssertEqual(flat, "abcd")
        XCTAssertEqual(map.count, 4)
        XCTAssertEqual(map[0].runIdx, 0); XCTAssertEqual(map[0].offset, 0)
        XCTAssertEqual(map[1].runIdx, 0); XCTAssertEqual(map[1].offset, 1)
        XCTAssertEqual(map[2].runIdx, 1); XCTAssertEqual(map[2].offset, 0)
        XCTAssertEqual(map[3].runIdx, 1); XCTAssertEqual(map[3].offset, 1)
    }

    // MARK: - #65 Run with rawElements survives multi-run replacement

    /// Repro for PsychQuant/ooxml-swift#65 (kiki830621/collaboration_guo_analysis#20):
    /// when `replaceText` finds a match whose flat-string boundary crosses an
    /// empty-text Run that carries only `rawElements` (e.g. `<w:commentReference>`,
    /// `<w:bookmarkStart>`, `<w:smartTag>` legacy carrier), the Run was being
    /// silently REMOVED. `isTextRun` returns true for it (no rawXML, no drawing)
    /// but its rawElements payload was being treated as deletable.
    ///
    /// Concrete failure observed in NTPU thesis: `<w:r><w:commentReference w:id="23"/></w:r>`
    /// dropped between "適應性" and "，本研究…" runs, breaking the comment
    /// triplet schema (rangeStart + rangeEnd present, reference missing).
    /// Word strict validator rejects the resulting docx.
    func testReplaceMultiRunPreservesCommentReferenceRun() throws {
        // Simulate a paragraph parsed from:
        //   <w:r>適應性</w:r>
        //   <w:r><w:commentReference w:id="23"/></w:r>   ← empty text + rawElements
        //   <w:r>，本研究</w:r>
        // and replace text spanning the gap (idx[3] ⇒ "適應" only) with multi-run extent.
        var commentRefRun = Run(text: "")
        commentRefRun.rawElements = [
            RawElement(name: "commentReference", xml: "<w:commentReference w:id=\"23\"/>")
        ]

        var runs = [
            Run(text: "適應性"),
            commentRefRun,
            Run(text: "，本研究")
        ]

        // Force a multi-run match by replacing across the gap: "性，本" spans
        // runs[0] (last char) → runs[2] (first 2 chars). The empty Run sits
        // strictly between sRunIdx=0 and eRunIdx=2.
        let count = try TextReplacementEngine.replace(
            runs: &runs, find: "性，本", with: "X"
        )

        XCTAssertEqual(count, 1)
        // Critical: commentReference Run must NOT have been removed.
        XCTAssertEqual(runs.count, 3, "Run carrying only rawElements was deleted by multi-run remove pass")
        XCTAssertEqual(runs[1].rawElements?.count, 1)
        XCTAssertEqual(runs[1].rawElements?.first?.name, "commentReference")
        // Text content correct
        XCTAssertEqual(runs[0].text, "適應X")
        XCTAssertEqual(runs[1].text, "")
        XCTAssertEqual(runs[2].text, "研究")
    }

    /// Even when the "match" itself is single-run (start == end), but text
    /// boundaries align such that an empty rawElements Run sits between the
    /// start-run and a later text-bearing region the engine touches, that
    /// rawElements Run must survive. This guards against future regressions
    /// where the safety net might be applied only to multi-run paths.
    func testReplaceSingleRunWithAdjacentRawElementsRunIntact() throws {
        var refRun = Run(text: "")
        refRun.rawElements = [
            RawElement(name: "commentReference", xml: "<w:commentReference w:id=\"23\"/>")
        ]

        var runs = [
            Run(text: "適應性"),
            refRun,
            Run(text: "，本研究")
        ]

        // Pure single-run match: 適應性 → 配適度 (all in runs[0]).
        let count = try TextReplacementEngine.replace(
            runs: &runs, find: "適應性", with: "配適度"
        )

        XCTAssertEqual(count, 1)
        XCTAssertEqual(runs.count, 3)
        XCTAssertEqual(runs[0].text, "配適度")
        XCTAssertEqual(runs[1].rawElements?.first?.name, "commentReference")
        XCTAssertEqual(runs[2].text, "，本研究")
    }
}
