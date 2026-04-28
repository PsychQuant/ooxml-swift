import XCTest
@testable import OOXMLSwift

/// Regression guard for [PsychQuant/ooxml-swift#4](https://github.com/PsychQuant/ooxml-swift/issues/4).
///
/// History: v0.19.0 (#56 Phase 4) added position-aware child walker to
/// `DocxReader.parseParagraph`, but the walker had no explicit `case "pPr"`,
/// so `<w:pPr>` fell into `default` and got captured into
/// `unrecognizedChildren` — even though `parseParagraphProperties` had
/// ALREADY consumed it into `paragraph.properties`. The sort-by-position
/// emit then wrote `<w:pPr>` twice per paragraph (once via the legacy pPr
/// block at the top of `Paragraph.toXMLSortedByPosition`, once verbatim
/// from `unrecognizedChildren`). xmllint accepts the duplicate; file size
/// grew ~1 KB per paragraph per round-trip.
///
/// v0.19.1 added `case "pPr": break` as a hot-fix. v0.21.1 (#4) hardens
/// the guard with two layers:
///
/// 1. **Walker entry whitelist** (`walkerPreConsumed: Set<String>`) — pPr
///    is filtered BEFORE entering the switch, structurally guaranteeing
///    it never reaches `unrecognizedChildren`.
/// 2. **`#if DEBUG` invariant assert** at the catch-all `default` append
///    site — runtime tripwire if anyone ever removes both the whitelist
///    AND the explicit `case` (defense in depth).
///
/// **Coverage claim (honest):** these observable-invariant tests will fail
/// only when *both* layers (the v0.19.1 explicit `case "pPr": break` AND
/// the v0.21.1 whitelist) are silently removed at the same time —
/// confirmed by counterfactual experiment in the verify report for #4.
/// Single-layer silent breakage is invisible to the observable-invariant
/// tests because the surviving layer absorbs the leak. The new positive
/// `testWalkerPreConsumedContainsPPr` below closes that gap by asserting
/// the whitelist constant directly, so a typo or accidental clear of
/// `walkerPreConsumed` is caught even when the explicit `case` still
/// holds.
final class Issue4PPrRegressionGuardTests: XCTestCase {

    // MARK: - Helpers

    private func parse(_ xml: String) throws -> Paragraph {
        let element = try XMLElement(xmlString: xml)
        return try DocxReader.parseParagraph(
            from: element,
            relationships: RelationshipsCollection(),
            styles: [],
            numbering: Numbering()
        )
    }

    /// Count occurrences of `<w:pPr` (open tag) in the emitted XML.
    /// Use the open-tag prefix (no `>` or `/>`) so it matches both empty
    /// (`<w:pPr/>`) and non-empty (`<w:pPr>...</w:pPr>`) forms.
    private func countPPrOpenTags(in xml: String) -> Int {
        var count = 0
        var searchRange = xml.startIndex..<xml.endIndex
        while let range = xml.range(of: "<w:pPr", range: searchRange) {
            count += 1
            searchRange = range.upperBound..<xml.endIndex
        }
        return count
    }

    // MARK: - 1. pPr only — no leak into unrecognizedChildren

    func testPPrOnlyParagraphProducesNoUnrecognizedPPr() throws {
        // Paragraph with only pPr (containing a distinctive style ref so
        // an "empty pPr drop" optimization can't silently swallow it).
        let xmlSrc = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:pPr>
            <w:pStyle w:val="Heading1"/>
            <w:jc w:val="center"/>
          </w:pPr>
        </w:p>
        """

        let paragraph = try parse(xmlSrc)

        let leaked = paragraph.unrecognizedChildren.filter { $0.name == "pPr" }
        XCTAssertTrue(
            leaked.isEmpty,
            "pPr must be consumed by parseParagraphProperties, NOT captured into unrecognizedChildren. Leaked: \(leaked.map { $0.rawXML })"
        )

        // Positive check: properties were actually populated (pStyle survived)
        XCTAssertEqual(
            paragraph.properties.style, "Heading1",
            "pStyle should be parsed into paragraph.properties.style"
        )
    }

    // MARK: - 2. Round-trip emits exactly one <w:pPr>

    func testPPrOnlyParagraphRoundTripsExactlyOnePPrBlock() throws {
        let xmlSrc = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:pPr>
            <w:pStyle w:val="Heading1"/>
            <w:jc w:val="center"/>
          </w:pPr>
        </w:p>
        """

        let paragraph = try parse(xmlSrc)
        let xml = paragraph.toXML()

        let pPrCount = countPPrOpenTags(in: xml)
        XCTAssertEqual(
            pPrCount, 1,
            "Round-trip must emit exactly ONE <w:pPr> block. Pre-v0.19.1 the bug emitted two (legacy pPr block + verbatim unrecognizedChildren copy). Output:\n\(xml)"
        )
    }

    // MARK: - 3. pPr + runs — runs preserved, no pPr leak

    func testPPrPlusRunsParagraphSkipsPPrPreservesRuns() throws {
        let xmlSrc = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:pPr>
            <w:pStyle w:val="Body"/>
          </w:pPr>
          <w:r><w:t>foo</w:t></w:r>
          <w:r><w:t>bar</w:t></w:r>
        </w:p>
        """

        let paragraph = try parse(xmlSrc)

        XCTAssertEqual(paragraph.runs.count, 2, "Both runs must be parsed")
        XCTAssertEqual(paragraph.runs[0].text, "foo")
        XCTAssertEqual(paragraph.runs[1].text, "bar")
        XCTAssertEqual(paragraph.properties.style, "Body")

        let leaked = paragraph.unrecognizedChildren.filter { $0.name == "pPr" }
        XCTAssertTrue(
            leaked.isEmpty,
            "pPr must not appear in unrecognizedChildren when sibling runs are present. Leaked: \(leaked.map { $0.rawXML })"
        )

        // Round-trip still emits exactly one pPr block
        let xml = paragraph.toXML()
        XCTAssertEqual(
            countPPrOpenTags(in: xml), 1,
            "Round-trip must emit exactly ONE <w:pPr> regardless of run count. Output:\n\(xml)"
        )
    }

    // MARK: - 4. Pre-v0.19.1 baseline — counterfactual sanity check

    /// If both the walker whitelist AND the explicit `case "pPr": break` were
    /// removed, this test class would fail (specifically tests 1-3 would
    /// detect pPr in unrecognizedChildren and double-emit on round-trip).
    /// This test is documentation only — it asserts a structural property of
    /// the codebase that can't be tested at runtime without a feature flag,
    /// but it makes the regression-guard intent explicit for a future reader
    /// who runs this file in isolation.
    func testRegressionGuardIntentDocumented() {
        // The 4 functional tests above collectively form the regression guard:
        //   #1 — pPr never leaks into unrecognizedChildren
        //   #2 — round-trip emits exactly one <w:pPr>
        //   #3 — invariant holds when sibling runs share the paragraph
        //   #4 — walker whitelist constant `walkerPreConsumed` contains "pPr"
        //
        // If any of #1-#3 fail, either:
        //   (a) `case "pPr": break` was removed from DocxReader.parseParagraph
        //   (b) `walkerPreConsumed: Set<String>` whitelist was removed
        //   (c) `parseParagraphProperties` is no longer being called
        //   (d) Paragraph.toXMLSortedByPosition emits unrecognizedChildren
        //       items whose name == "pPr" without filtering
        //
        // If #4 fails, the whitelist itself was silently broken (typo /
        // accidentally cleared) — caught directly even when the v0.19.1
        // explicit `case "pPr": break` still absorbs the observable leak.
        //
        // See: PsychQuant/ooxml-swift#4 + PsychQuant/che-word-mcp#56.
        XCTAssertTrue(true, "Documentation-only test — see comment above")
    }

    // MARK: - 5. Whitelist content sanity (closes single-layer regression gap)

    /// Direct assertion on the whitelist constant. Catches the case where the
    /// observable-invariant tests #1-#3 still pass (because the v0.19.1 explicit
    /// `case "pPr": break` absorbs the leak) but the new whitelist mechanism
    /// has been silently broken — typo `"pPR"`, accidentally cleared to `[]`,
    /// renamed without update, etc. Verify report for #4 documented this
    /// single-layer gap as Experiment B; this test closes it.
    func testWalkerPreConsumedContainsPPr() {
        XCTAssertTrue(
            DocxReader.walkerPreConsumed.contains("pPr"),
            "DocxReader.walkerPreConsumed must contain \"pPr\" to keep the structural pre-switch filter active. Without this, the v0.19.1 explicit `case` becomes the sole defense — defeating the v0.21.1 (#4) defense-in-depth design."
        )
    }
}
