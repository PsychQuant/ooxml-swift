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
}
