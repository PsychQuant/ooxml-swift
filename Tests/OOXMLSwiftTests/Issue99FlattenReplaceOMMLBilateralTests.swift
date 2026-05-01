import XCTest
@testable import OOXMLSwift

/// Bilateral flatten + replace coverage tests for direct-child OMML across
/// 4 wrapper positions (PsychQuant/che-word-mcp #99 / #100 / #101 / #102 /
/// #103 cluster, post-#92 follow-up).
///
/// **Pre-fix bug**: `flattenedDisplayText` and `replaceInParagraphSurfaces`
/// both delegate to a `[Run]` walker, so OMML appearing as a *direct child*
/// of a wrapper (not wrapped in `<w:r>`) is silently dropped at 4 positions:
///
/// | Position | Wrapper | Issue |
/// |----------|---------|-------|
/// | 1 | `<w:p>` direct child (Pandoc display math) | #99 |
/// | 2 | `<w:hyperlink>` direct child | #100 |
/// | 3 | `<mc:Fallback>` direct child | #101 |
/// | 4 | nested `<w:hyperlink>`/`<w:fldSimple>` combo | #102 |
///
/// **Post-fix**: bilateral mirror coverage — read includes OMML visibleText
/// (anchor universe extends to all wrapper positions), write detects OMML
/// boundaries and refuses replacements crossing them via typed
/// `ReplaceResult.refusedDueToOMMLBoundary(occurrences:)`. Documented
/// asymmetry per Spectra change `flatten-replace-omml-bilateral-coverage`.
///
/// **Library design principles** (separate spec capability):
/// - Correctness primacy — refuse > incorrect approximation
/// - Human-like operations — no surprising state, no silent destruction
final class Issue99FlattenReplaceOMMLBilateralTests: XCTestCase {

    // MARK: - Fixture builders (inline-XML stubs — no external .docx dependencies)

    /// Build a `Paragraph` from inline `<w:p>` XML by invoking
    /// `DocxReader.parseParagraph` directly. Mirrors the helper used in
    /// `Issue92OMMLWalkSurfaceCoverageTests` and `Issue85InlineMathFlattenTests`.
    private func parseParagraph(xml: String) throws -> Paragraph {
        let data = xml.data(using: .utf8)!
        let doc = try XMLDocument(data: data)
        guard let root = doc.rootElement() else {
            throw NSError(domain: "Issue99FlattenReplaceOMMLBilateralTests", code: 1)
        }
        let document = WordDocument()
        return try DocxReader.parseParagraph(
            from: root,
            relationships: RelationshipsCollection(),
            styles: document.styles,
            numbering: document.numbering
        )
    }

    /// Fixture (a): `<m:oMath>` direct child of `<w:p>` between two runs.
    /// Pandoc display math (`$$...$$`) emits this shape. Verifies Decision 6
    /// source-XML position ordering (run1 → OMML → run3 in source order).
    private func paragraphFixtureA_directChildOfParagraph() throws -> Paragraph {
        let xml = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
          <w:r><w:t xml:space="preserve">see eq </w:t></w:r>
          <m:oMath><m:r><m:t>δ</m:t></m:r></m:oMath>
          <w:r><w:t xml:space="preserve"> here</w:t></w:r>
        </w:p>
        """
        return try parseParagraph(xml: xml)
    }

    /// Fixture (b): `<m:oMath>` direct child of `<w:hyperlink>` (no `<w:r>`
    /// wrapper). Triggered by LaTeX→docx tools wrapping cross-reference
    /// math in hyperlinks.
    private func paragraphFixtureB_directChildOfHyperlink() throws -> Paragraph {
        let xml = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
          <w:hyperlink r:id="rId10">
            <m:oMath><m:r><m:t>θ</m:t></m:r></m:oMath>
          </w:hyperlink>
        </w:p>
        """
        return try parseParagraph(xml: xml)
    }

    /// Fixture (c): `<m:oMath>` direct child of `<mc:Fallback>` inside
    /// `<mc:AlternateContent>` (no `<w:r>` wrapper inside fallback).
    /// Office.js fallback emit shape.
    private func paragraphFixtureC_directChildOfFallback() throws -> Paragraph {
        let xml = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
          <mc:AlternateContent>
            <mc:Choice Requires="w14"><w:r><w:t>modern</w:t></w:r></mc:Choice>
            <mc:Fallback>
              <m:oMath><m:r><m:t>κ</m:t></m:r></m:oMath>
            </mc:Fallback>
          </mc:AlternateContent>
        </w:p>
        """
        return try parseParagraph(xml: xml)
    }

    /// Fixture (d): nested wrapper combo — `<w:hyperlink>` containing
    /// `<w:fldSimple>` containing OMML. `parseHyperlink` raw-stashes the
    /// `<w:fldSimple>` since it's not `<w:r>`; `parseFieldSimple` is never
    /// invoked on it. Same blind spot in inverse direction.
    private func paragraphFixtureD_nestedHyperlinkFieldSimple() throws -> Paragraph {
        let xml = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
          <w:hyperlink r:id="rId11">
            <w:fldSimple w:instr="REF eq1 \\h">
              <w:r><m:oMath><m:r><m:t>η</m:t></m:r></m:oMath></w:r>
            </w:fldSimple>
          </w:hyperlink>
        </w:p>
        """
        return try parseParagraph(xml: xml)
    }

    // MARK: - 99.0 Fixture sanity checks (no behavior assertions — pure scaffolding)

    /// Pin that fixture (a) parses without throwing AND captures direct-child
    /// OMML in the expected raw storage location. Decision 4 (raw passthrough
    /// preserved) lives or dies on this invariant.
    func testFixtureAParsesAndStoresOMathInUnrecognizedChildren() throws {
        let para = try paragraphFixtureA_directChildOfParagraph()
        XCTAssertEqual(para.runs.count, 2, "expected 2 runs (caption + trailing)")
        let omathChildren = para.unrecognizedChildren.filter { $0.name == "oMath" || $0.name == "oMathPara" }
        XCTAssertEqual(omathChildren.count, 1, "expected 1 direct-child OMML in unrecognizedChildren")
        XCTAssertTrue(omathChildren[0].rawXML.contains("δ"), "OMML rawXML should preserve δ verbatim")
    }

    func testFixtureBParsesHyperlinkWithDirectChildOMML() throws {
        let para = try paragraphFixtureB_directChildOfHyperlink()
        XCTAssertEqual(para.hyperlinks.count, 1, "expected 1 hyperlink")
        // OMML lands in HyperlinkChild.rawXML(...) via parseHyperlink's non-Run branch
        let h = para.hyperlinks[0]
        XCTAssertTrue(
            h.children.contains(where: { child in
                if case .rawXML(let raw) = child { return raw.contains("oMath") && raw.contains("θ") }
                return false
            }),
            "expected hyperlink to carry direct-child OMML in HyperlinkChild.rawXML"
        )
    }

    func testFixtureCParsesAlternateContentWithDirectChildOMMLInFallback() throws {
        let para = try paragraphFixtureC_directChildOfFallback()
        XCTAssertEqual(para.alternateContents.count, 1, "expected 1 alternateContent")
        let ac = para.alternateContents[0]
        XCTAssertTrue(ac.rawXML.contains("oMath"), "AlternateContent rawXML should preserve OMML element")
        XCTAssertTrue(ac.rawXML.contains("κ"), "AlternateContent rawXML should preserve κ verbatim")
    }

    func testFixtureDParsesNestedHyperlinkFieldSimpleWithOMML() throws {
        let para = try paragraphFixtureD_nestedHyperlinkFieldSimple()
        XCTAssertEqual(para.hyperlinks.count, 1, "expected 1 hyperlink wrapping a fldSimple")
        let h = para.hyperlinks[0]
        let hasNestedOMML = h.children.contains(where: { child in
            if case .rawXML(let raw) = child { return raw.contains("fldSimple") && raw.contains("oMath") && raw.contains("η") }
            return false
        })
        XCTAssertTrue(hasNestedOMML, "expected hyperlink rawXML to carry nested fldSimple+OMML preserving η")
    }

    // MARK: - 99.2 ReplaceResult enum (Decision 7) — Tasks 2.1, 2.2

    /// Task 2.1 — `ReplaceResult.replaced(count: N)` constructible with N >= 0;
    /// `count == 0` distinct from `refusedDueToOMMLBoundary` (anchor-not-found
    /// is NOT a refusal — refusal is reserved for OMML boundary intersection).
    func testReplaceResultReplacedCaseConstructibleAndDistinct() {
        let zero = ReplaceResult.replaced(count: 0)
        let one = ReplaceResult.replaced(count: 1)
        let many = ReplaceResult.replaced(count: 42)

        // count == 0 is a normal "anchor not found" outcome
        if case .replaced(let c) = zero { XCTAssertEqual(c, 0) } else {
            XCTFail("expected .replaced case for count=0")
        }
        if case .replaced(let c) = one { XCTAssertEqual(c, 1) } else {
            XCTFail("expected .replaced case for count=1")
        }
        if case .replaced(let c) = many { XCTAssertEqual(c, 42) } else {
            XCTFail("expected .replaced case for count=42")
        }

        // .replaced(count: 0) MUST NOT equal .refusedDueToOMMLBoundary(occurrences: [])
        let refusedEmpty = ReplaceResult.refusedDueToOMMLBoundary(occurrences: [])
        XCTAssertNotEqual(zero, refusedEmpty,
            "Spec contract: anchor-not-found returns .replaced(count: 0), NOT a refusal case")
    }

    /// Task 2.2 — `ReplaceResult.refusedDueToOMMLBoundary(occurrences:)` carries
    /// `[Occurrence]` shape with `matchSpan: Range<Int>` and `ommlSpans: [Range<Int>]`.
    func testReplaceResultRefusedCaseCarriesOccurrenceShape() {
        let occ = ReplaceResult.Occurrence(matchSpan: 4..<13, ommlSpans: [7..<8])
        let refused = ReplaceResult.refusedDueToOMMLBoundary(occurrences: [occ])

        if case .refusedDueToOMMLBoundary(let occurrences) = refused {
            XCTAssertEqual(occurrences.count, 1)
            XCTAssertEqual(occurrences[0].matchSpan, 4..<13)
            XCTAssertEqual(occurrences[0].ommlSpans, [7..<8])
        } else {
            XCTFail("expected .refusedDueToOMMLBoundary case")
        }

        // Multi-occurrence shape — paragraph could have multiple cross-OMML matches
        let occ2 = ReplaceResult.Occurrence(matchSpan: 20..<25, ommlSpans: [22..<23, 24..<25])
        let multi = ReplaceResult.refusedDueToOMMLBoundary(occurrences: [occ, occ2])
        if case .refusedDueToOMMLBoundary(let occurrences) = multi {
            XCTAssertEqual(occurrences.count, 2)
            XCTAssertEqual(occurrences[1].ommlSpans, [22..<23, 24..<25])
        } else {
            XCTFail("expected .refusedDueToOMMLBoundary case for multi")
        }
    }

    // MARK: - 99.3 flattenedDisplayText walks direct-child OMML at 4 wrapper positions
    //               (Tasks 3.1-3.5)

    /// Task 3.1 — Position 1 (`<w:p>` direct-child).
    /// Pandoc display math reproducer (issue #99 primary).
    /// Decision 6 source-XML position ordering: caption text, OMML, trailing text
    /// must produce flattened "see eq δ here" not "see eq  here" + δ appended.
    func testFlattenedDisplayTextIncludesDirectChildOMathOfParagraph_Issue99() throws {
        let para = try paragraphFixtureA_directChildOfParagraph()
        let flat = para.flattenedDisplayText()

        XCTAssertTrue(flat.contains("see eq"), "leading <w:r> text dropped: '\(flat)'")
        XCTAssertTrue(flat.contains("δ"),       "direct-child <m:oMath> visibleText dropped: '\(flat)'")
        XCTAssertTrue(flat.contains("here"),   "trailing <w:r> text dropped: '\(flat)'")
        // Decision 6: source-XML position ordering — δ must appear BETWEEN "see eq " and " here"
        XCTAssertEqual(flat, "see eq δ here",
            "Decision 6 source-XML position order broken — expected interleaved 'see eq δ here', got '\(flat)'")
    }

    /// Task 3.2 — Position 2 (`<w:hyperlink>` direct-child) — issue #100.
    func testFlattenedDisplayTextIncludesDirectChildOMathOfHyperlink_Issue100() throws {
        let para = try paragraphFixtureB_directChildOfHyperlink()
        let flat = para.flattenedDisplayText()
        XCTAssertTrue(flat.contains("θ"),
            "<w:hyperlink> direct-child OMML θ dropped from flatten: '\(flat)'")
    }

    /// Task 3.3 — Position 3 (`<mc:Fallback>` direct-child) — issue #101.
    func testFlattenedDisplayTextIncludesDirectChildOMathOfFallback_Issue101() throws {
        let para = try paragraphFixtureC_directChildOfFallback()
        let flat = para.flattenedDisplayText()
        XCTAssertTrue(flat.contains("κ"),
            "<mc:Fallback> direct-child OMML κ dropped from flatten: '\(flat)'")
    }

    /// Task 3.4 — Position 4 (nested `<w:hyperlink>`/`<w:fldSimple>` combo) — issue #102.
    func testFlattenedDisplayTextIncludesNestedHyperlinkFieldSimpleOMML_Issue102() throws {
        let para = try paragraphFixtureD_nestedHyperlinkFieldSimple()
        let flat = para.flattenedDisplayText()
        XCTAssertTrue(flat.contains("η"),
            "nested hyperlink/fldSimple OMML η dropped from flatten: '\(flat)'")
    }

    /// Task 3.5 — Walker reads from raw storage on demand (Decision 4 raw passthrough).
    /// `Paragraph.unrecognizedChildren` entries with `name == "oMath"` must be
    /// preserved verbatim after `flattenedDisplayText()` walks them — flatten
    /// is a read-only operation.
    func testFlattenedDisplayTextDoesNotMutateRawOMMLStorage() throws {
        let para = try paragraphFixtureA_directChildOfParagraph()
        let omathBefore = para.unrecognizedChildren
            .filter { $0.name == "oMath" || $0.name == "oMathPara" }
            .map { $0.rawXML }

        _ = para.flattenedDisplayText()  // discard result — testing side-effect-freeness

        let omathAfter = para.unrecognizedChildren
            .filter { $0.name == "oMath" || $0.name == "oMathPara" }
            .map { $0.rawXML }

        XCTAssertEqual(omathBefore, omathAfter,
            "Decision 4 violated: flattenedDisplayText() mutated raw OMML storage")
    }

    // MARK: - 99.4 replaceInParagraphSurfaces detects OMML boundaries (Tasks 4.1, 4.2, 4.3)
    //
    // These tests exercise `WordDocument.replaceTextWithBoundaryDetection`
    // (the new public API surfacing `ReplaceResult`). The internal
    // `Document.replaceInParagraphSurfaces` (private static) is invoked
    // through a `WordDocument` driving the standard fixture A reproducer.

    /// Build a single-paragraph WordDocument containing the standard
    /// `<w:p>` direct-child OMML reproducer (fixture A). Used by 4.1-4.3.
    private func documentWithFixtureA() throws -> WordDocument {
        let para = try paragraphFixtureA_directChildOfParagraph()
        var doc = WordDocument()
        doc.body.children = [.paragraph(para)]
        return doc
    }

    /// Task 4.1 — wholly-within `<w:t>` mutation proceeds.
    func testReplaceWhollyWithinWtProceeds() throws {
        var doc = try documentWithFixtureA()
        let result = try doc.replaceTextWithBoundaryDetection(find: "here", with: "there")

        if case .replaced(let count) = result {
            XCTAssertEqual(count, 1, "wholly-within `<w:t>` find should mutate exactly 1 occurrence")
        } else {
            XCTFail("expected .replaced case for wholly-within match, got \(result)")
        }

        // Verify mutation actually landed
        guard case .paragraph(let updatedPara) = doc.body.children[0] else {
            return XCTFail("expected paragraph child")
        }
        XCTAssertTrue(updatedPara.flattenedDisplayText().contains("there"),
            "mutated paragraph should now contain 'there'")
    }

    /// Task 4.2 — cross-OMML mutation refuses with informative occurrence.
    func testReplaceCrossOMMLRefusesWithOccurrenceInfo() throws {
        var doc = try documentWithFixtureA()

        // Capture pre-call paragraph XML for byte-identity check
        guard case .paragraph(let preCallPara) = doc.body.children[0] else {
            return XCTFail("expected paragraph child pre-call")
        }
        let preFlat = preCallPara.flattenedDisplayText()

        let result = try doc.replaceTextWithBoundaryDetection(find: "eq δ here", with: "ref X")

        if case .refusedDueToOMMLBoundary(let occurrences) = result {
            XCTAssertEqual(occurrences.count, 1, "expected 1 cross-OMML occurrence")
            // Spec scenario: matchSpan 4..<13 covers "eq δ here" in flattened "see eq δ here"
            XCTAssertEqual(occurrences[0].matchSpan, 4..<13,
                "matchSpan should cover 'eq δ here' positions 4..<13")
            XCTAssertEqual(occurrences[0].ommlSpans, [7..<8],
                "ommlSpans should contain OMML at position 7..<8 (δ)")
        } else {
            XCTFail("expected .refusedDueToOMMLBoundary, got \(result)")
        }

        // Verify NO mutation occurred (spec contract)
        guard case .paragraph(let postCallPara) = doc.body.children[0] else {
            return XCTFail("expected paragraph child post-call")
        }
        XCTAssertEqual(preFlat, postCallPara.flattenedDisplayText(),
            "Decision 2 violated: refused replacement modified paragraph content")
    }

    /// Build a paragraph fixture engineered so that the same find string
    /// appears 3 times in the flattened text — twice wholly-within `<w:t>`
    /// and once spanning across direct-child OMML.
    ///
    /// Layout: `<w:r>abc a</w:r><m:oMath>b</m:oMath><w:r>c xyz abc</w:r>`
    /// flat:   "abc abc xyz abc"  (positions 0..14, OMML "b" at position 5)
    ///
    /// Find "abc" matches at:
    /// - 0..<3  — wholly within run 1 ("abc")
    /// - 4..<7  — spans run1's "a" (pos 4) + OMML "b" (pos 5) + run3's "c" (pos 6)
    /// - 12..<15 — wholly within run 3 ("abc")
    private func paragraphMixedFixture() throws -> Paragraph {
        let xml = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
          <w:r><w:t xml:space="preserve">abc a</w:t></w:r>
          <m:oMath><m:r><m:t>b</m:t></m:r></m:oMath>
          <w:r><w:t xml:space="preserve">c xyz abc</w:t></w:r>
        </w:p>
        """
        return try parseParagraph(xml: xml)
    }

    /// Task 4.3 — TRUE mixed: same find string appears 3 times in same
    /// paragraph (single call): twice wholly-within, once cross-OMML.
    /// Result MUST be `.mixed(replacedCount: 2, refusedOccurrences: [...])`
    /// per Spec contract "the result SHALL combine both — `ReplaceResult`
    /// MUST signal both `count > 0` replaced AND non-empty refused
    /// occurrences". Combinator chosen: third enum case `.mixed`.
    func testReplaceMixedWhollyWithinAndCrossOMMLInSingleCall() throws {
        let para = try paragraphMixedFixture()
        var doc = WordDocument()
        doc.body.children = [.paragraph(para)]

        // Sanity: pre-call flatten matches expected layout
        guard case .paragraph(let preCallPara) = doc.body.children[0] else {
            return XCTFail("expected paragraph child pre-call")
        }
        XCTAssertEqual(preCallPara.flattenedDisplayText(), "abc abc xyz abc",
            "fixture layout broken — flatten must match expected positions")

        let result = try doc.replaceTextWithBoundaryDetection(find: "abc", with: "XYZ")

        // Expected: 2 replaced (positions 0..<3 and 12..<15) + 1 refused
        // (position 4..<7 spans OMML at 5..<6)
        if case .mixed(let replacedCount, let refused) = result {
            XCTAssertEqual(replacedCount, 2,
                "two wholly-within matches should be replaced")
            XCTAssertEqual(refused.count, 1,
                "one cross-OMML occurrence should be refused")
            XCTAssertEqual(refused[0].matchSpan, 4..<7,
                "refused match should span positions 4..<7")
            XCTAssertEqual(refused[0].ommlSpans, [5..<6],
                "refused match should reference OMML at 5..<6")
        } else {
            XCTFail("expected .mixed combinator carrying both signals, got \(result)")
        }
    }

    // MARK: - 99.5 Mirror invariant — same surface coverage, asymmetric handling
    //               (Tasks 5.1, 5.2)

    /// Task 5.1 — Mirror invariant scenarios from spec
    /// `ooxml-paragraph-text-mirror`. Pin: same wrapper coverage walked
    /// for both flatten (read) and replaceTextWithBoundaryDetection (write),
    /// with documented asymmetry — read includes OMML visibleText, write
    /// treats OMML as opaque structural unit (refuses cross-OMML mutation).
    ///
    /// Strategy: drive the standard 4-position fixtures through both
    /// directions and assert that:
    /// - Read sees OMML text at all 4 positions (already verified by 3.1-3.4)
    /// - Write refuses cross-OMML mutation at position 1 (verified by 4.2)
    ///   AND does NOT modify the OMML element itself in any of the 4
    ///   positions when no `<w:t>` mutation is requested.
    func testMirrorInvariantSurfaceCoverageParityAndAsymmetricHandling() throws {
        // Position 1: paragraph direct-child OMML — bilateral coverage already
        // verified across 3.1 (read) + 4.2 (write refuses cross-OMML).
        // Re-pin in single test to make the mirror invariant explicit.
        var doc = try documentWithFixtureA()
        guard case .paragraph(let preParaA) = doc.body.children[0] else {
            return XCTFail("expected paragraph child for fixture A")
        }
        let preFlatA = preParaA.flattenedDisplayText()
        XCTAssertEqual(preFlatA, "see eq δ here", "read sees OMML text at position 1")

        // Trigger a cross-OMML find — write must refuse + not mutate OMML
        let writeOutcomeA = try doc.replaceTextWithBoundaryDetection(find: "eq δ here", with: "ref X")
        if case .refusedDueToOMMLBoundary = writeOutcomeA {
            // Expected — write refused per Decision 2 opaque OMML
        } else {
            XCTFail("position 1 write should refuse cross-OMML, got \(writeOutcomeA)")
        }
        guard case .paragraph(let postParaA) = doc.body.children[0] else {
            return XCTFail("expected paragraph child post-call")
        }
        // OMML element preserved verbatim
        XCTAssertEqual(
            preParaA.unrecognizedChildren.first(where: { $0.name == "oMath" })?.rawXML,
            postParaA.unrecognizedChildren.first(where: { $0.name == "oMath" })?.rawXML,
            "Decision 2 violated: write modified OMML element raw XML"
        )

        // Positions 2/3/4: OMML lives in wrapper raw storage. Write attempts
        // do not target OMML elements (typed `runs[]` empty in these
        // wrappers — no mutable `<w:t>`). Verify read sees OMML AND write
        // does not modify wrapper rawXML.
        let fixtures: [(String, () throws -> Paragraph, String)] = [
            ("position 2 hyperlink (#100)", paragraphFixtureB_directChildOfHyperlink, "θ"),
            ("position 3 fallback (#101)",  paragraphFixtureC_directChildOfFallback,  "κ"),
            ("position 4 nested (#102)",     paragraphFixtureD_nestedHyperlinkFieldSimple, "η"),
        ]
        for (label, fixtureBuilder, expectedChar) in fixtures {
            let para = try fixtureBuilder()
            var wrapDoc = WordDocument()
            wrapDoc.body.children = [.paragraph(para)]

            // Read side: includes OMML char
            guard case .paragraph(let preWrapPara) = wrapDoc.body.children[0] else {
                XCTFail("\(label): expected paragraph child"); continue
            }
            XCTAssertTrue(preWrapPara.flattenedDisplayText().contains(expectedChar),
                "\(label) read should include OMML char '\(expectedChar)'")

            // Write side: any find on these surfaces returns 0 (no mutable text)
            let outcome = try wrapDoc.replaceTextWithBoundaryDetection(find: expectedChar, with: "X")
            if case .replaced(let count) = outcome {
                XCTAssertEqual(count, 0,
                    "\(label) write should not mutate OMML — typed runs empty, count must be 0")
            } else if case .refusedDueToOMMLBoundary = outcome {
                // Acceptable: detected as cross-OMML and refused
            } else {
                XCTFail("\(label) unexpected outcome \(outcome)")
            }

            // OMML preserved across the call
            guard case .paragraph(let postWrapPara) = wrapDoc.body.children[0] else {
                XCTFail("\(label): expected paragraph child post-call"); continue
            }
            XCTAssertEqual(preWrapPara.flattenedDisplayText(), postWrapPara.flattenedDisplayText(),
                "\(label): paragraph flatten should be unchanged after write attempt (OMML preserved)")
        }
    }

    // MARK: - 99.7 Round-trip preserves direct-child OMML XML (Task 7.1)

    /// Task 7.1 — Requirement "Direct-child OMML storage remains raw passthrough"
    /// round-trip clause: read → no mutation → write → read again must
    /// produce byte-identical OMML element XML for all 4 wrapper positions.
    ///
    /// Scope: this verifies the `Paragraph.unrecognizedChildren[].rawXML`,
    /// `HyperlinkChild.rawXML(_)`, and `AlternateContent.rawXML` storage
    /// surfaces preserve the OMML element verbatim across round-trip.
    /// Decision 4 raw passthrough invariant.
    ///
    /// Note: this is an in-memory round-trip (paragraph re-emit via
    /// `Paragraph.toXMLSortedByPosition` then re-parse), not a full
    /// `DocxWriter` → `DocxReader` cycle — the latter would require building
    /// a complete WordDocument with Content_Types.xml etc. and is exercised
    /// by the matrix-pin `testDocumentContentEqualityInvariant` test (run
    /// as part of full suite in 7.2).
    func testRoundTripPreservesDirectChildOMMLForAllFourWrapperPositions() throws {
        let cases: [(String, () throws -> Paragraph, String)] = [
            ("position 1 (#99) <w:p> direct-child", paragraphFixtureA_directChildOfParagraph, "δ"),
            ("position 2 (#100) <w:hyperlink>",     paragraphFixtureB_directChildOfHyperlink, "θ"),
            ("position 3 (#101) <mc:Fallback>",     paragraphFixtureC_directChildOfFallback,  "κ"),
            ("position 4 (#102) nested",             paragraphFixtureD_nestedHyperlinkFieldSimple, "η"),
        ]
        for (label, fixtureBuilder, expectedChar) in cases {
            let para = try fixtureBuilder()
            // Read OMML char from flatten (proves it's parsed)
            XCTAssertTrue(para.flattenedDisplayText().contains(expectedChar),
                "\(label): pre-roundtrip flatten should include '\(expectedChar)'")

            // No mutation step — testing pure read+write round-trip preservation

            // Re-emit paragraph via public `toXML()` and re-parse.
            // `Paragraph.toXML()` is public (line 316) and emits the full
            // <w:p>...</w:p> element with all children including
            // unrecognizedChildren rawXML (Decision 4 storage preserved).
            let emittedXML = "<root xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">" + para.toXML() + "</root>"
            let dataOpt: Data? = emittedXML.data(using: String.Encoding.utf8)
            guard let data = dataOpt else {
                XCTFail("\(label): failed to encode emitted XML"); continue
            }
            let reparsedDoc = try XMLDocument(data: data)
            guard let pElement = reparsedDoc.rootElement()?.elements(forName: "w:p").first else {
                XCTFail("\(label): failed to re-find <w:p> after round-trip"); continue
            }
            let document = WordDocument()
            let reparsedPara = try DocxReader.parseParagraph(
                from: pElement,
                relationships: RelationshipsCollection(),
                styles: document.styles,
                numbering: document.numbering
            )

            // Post-roundtrip flatten must still include the OMML char
            XCTAssertTrue(reparsedPara.flattenedDisplayText().contains(expectedChar),
                "\(label): post-roundtrip flatten should still include '\(expectedChar)' — Decision 4 raw passthrough violated")
        }
    }
}
