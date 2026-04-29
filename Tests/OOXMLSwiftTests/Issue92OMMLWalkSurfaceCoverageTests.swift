import XCTest
@testable import OOXMLSwift

/// Surface coverage tests for `Paragraph.flattenedDisplayText()` OMML walk
/// across non-runs paths, per
/// [PsychQuant/che-word-mcp#92](https://github.com/PsychQuant/che-word-mcp/issues/92).
///
/// Pre-#92 (post-#85): the OMML walk added in v0.21.5 covered only the
/// top-level `runs` loop. The other 4 surface paths used
/// `runs.map { $0.text }.joined()` and silently dropped OMML inside their
/// wrappers:
///
/// | Surface path | OMML walked pre-#92 |
/// |--------------|---------------------|
/// | `paragraph.runs` | ✅ (#85 fix v0.21.5) |
/// | `paragraph.hyperlinks[].runs` | ❌ silent drop |
/// | `paragraph.fieldSimples[].runs` | ❌ silent drop |
/// | `paragraph.alternateContents[].fallbackRuns` | ❌ silent drop |
/// | `paragraph.contentControls` (typed children) | ✅ (covered via separate `flattenContentControlText` since #63) |
///
/// The docstring of `flattenedDisplayText` explicitly claimed it "mirrors the
/// surface coverage of `Document.replaceInParagraphSurfaces`" — that mirror
/// was broken on the OMML axis. Real-world impact: rare but real. Inline
/// math inside a hyperlink (e.g., `<w:hyperlink>...<m:oMath>...</m:oMath>...</w:hyperlink>`)
/// or inside a field simple (`<w:fldSimple>...<m:oMath>...</m:oMath>...</w:fldSimple>`)
/// silently 0-matched anchor lookups for math text — same diagnosis, same fix
/// pattern as #85's primary bug, just shifted to wrapper paths.
///
/// Post-#92: a `flattenRunsWithOMML(_ runs:)` helper unifies the walk across
/// 4 paths (top-level runs + hyperlinks + fieldSimples + AC fallbackRuns).
/// contentControls path stays separate (uses different recursion via
/// `flattenContentControlText` since #63).
final class Issue92OMMLWalkSurfaceCoverageTests: XCTestCase {

    // MARK: - Pre-#92 silent-drop scenarios (now covered)

    /// Spec: inline math inside a `<w:hyperlink>` flattens with math text.
    /// Pre-#92: `paragraph.hyperlinks[].runs.map { $0.text }.joined()` dropped
    /// the OMML's `<m:t>` content silently. Post-#92: helper walks OMML.
    func testHyperlinkRunsWithInlineMathFlattenIncludeMathText() throws {
        let xml = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
          <w:hyperlink r:id="rId10">
            <w:r><w:t>see eq </w:t></w:r>
            <w:r>
              <m:oMath>
                <m:r><m:t>α</m:t></m:r>
              </m:oMath>
            </w:r>
            <w:r><w:t> here</w:t></w:r>
          </w:hyperlink>
        </w:p>
        """
        let para = try parseParagraph(xml: xml)
        let flat = para.flattenedDisplayText()
        XCTAssertTrue(flat.contains("see eq"), "hyperlink leading run text dropped: '\(flat)'")
        XCTAssertTrue(flat.contains("α"), "hyperlink OMML <m:t> dropped: '\(flat)'")
        XCTAssertTrue(flat.contains("here"), "hyperlink trailing run text dropped: '\(flat)'")
    }

    /// Spec: inline math inside a `<w:fldSimple>` flattens with math text.
    /// Pre-#92: silent drop (same root cause as hyperlinks).
    func testFieldSimpleRunsWithInlineMathFlattenIncludeMathText() throws {
        let xml = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
          <w:fldSimple w:instr=" REF eq ">
            <w:r><w:t>(see </w:t></w:r>
            <w:r>
              <m:oMath>
                <m:r><m:t>β</m:t></m:r>
              </m:oMath>
            </w:r>
            <w:r><w:t>)</w:t></w:r>
          </w:fldSimple>
        </w:p>
        """
        let para = try parseParagraph(xml: xml)
        let flat = para.flattenedDisplayText()
        XCTAssertTrue(flat.contains("(see"), "fldSimple leading run text dropped: '\(flat)'")
        XCTAssertTrue(flat.contains("β"), "fldSimple OMML <m:t> dropped: '\(flat)'")
        XCTAssertTrue(flat.contains(")"), "fldSimple trailing run text dropped: '\(flat)'")
    }

    /// Spec: inline math inside `<mc:AlternateContent><mc:Fallback>` flattens
    /// with math text. Pre-#92: silent drop. The fallback runs path is the
    /// most likely surface to carry OMML in real docs (Word emits AC blocks
    /// when a feature has both a modern + legacy representation).
    func testAlternateContentFallbackRunsWithInlineMathFlattenIncludeMathText() throws {
        let xml = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
          <mc:AlternateContent>
            <mc:Choice Requires="w14">
              <w:r><w:t>modern</w:t></w:r>
            </mc:Choice>
            <mc:Fallback>
              <w:r><w:t>fallback </w:t></w:r>
              <w:r>
                <m:oMath>
                  <m:r><m:t>γ</m:t></m:r>
                </m:oMath>
              </w:r>
              <w:r><w:t> end</w:t></w:r>
            </mc:Fallback>
          </mc:AlternateContent>
        </w:p>
        """
        let para = try parseParagraph(xml: xml)
        let flat = para.flattenedDisplayText()
        XCTAssertTrue(flat.contains("fallback"), "AC fallback leading run dropped: '\(flat)'")
        XCTAssertTrue(flat.contains("γ"), "AC fallback OMML <m:t> dropped: '\(flat)'")
        XCTAssertTrue(flat.contains("end"), "AC fallback trailing run dropped: '\(flat)'")
    }

    // MARK: - Top-level runs regression guard (post-refactor)

    /// Regression: refactoring the top-level runs loop to use the new helper
    /// must NOT change its behavior. This pins the existing #85 contract
    /// (inline math in top-level runs flattens with math text) under the new
    /// helper-based implementation.
    func testTopLevelRunsRegressionAfterRefactor() throws {
        let xml = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
          <w:r><w:t>進行</w:t></w:r>
          <w:r>
            <m:oMath>
              <m:r><m:t>t</m:t></m:r>
            </m:oMath>
          </w:r>
          <w:r><w:t>檢定：</w:t></w:r>
        </w:p>
        """
        let para = try parseParagraph(xml: xml)
        let flat = para.flattenedDisplayText()
        XCTAssertTrue(flat.contains("進行"), "top-level leading dropped post-refactor: \(flat)")
        XCTAssertTrue(flat.contains("t"), "top-level OMML <m:t> dropped post-refactor: \(flat)")
        XCTAssertTrue(flat.contains("檢定"), "top-level trailing dropped post-refactor: \(flat)")
    }

    // MARK: - Plain runs regression (no OMML)

    /// Regression: paths without OMML must not be affected by the helper
    /// extraction. Pure plain text in hyperlink/fldSimple/AC fallbackRuns
    /// continues to flatten verbatim.
    func testPlainHyperlinkFlattensWithoutOMML() throws {
        let xml = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:hyperlink r:id="rId10">
            <w:r><w:t>plain link text</w:t></w:r>
          </w:hyperlink>
        </w:p>
        """
        let para = try parseParagraph(xml: xml)
        XCTAssertEqual(para.flattenedDisplayText(), "plain link text")
    }

    // MARK: - Helpers

    private func parseParagraph(xml: String) throws -> Paragraph {
        let data = xml.data(using: .utf8)!
        let doc = try XMLDocument(data: data)
        guard let root = doc.rootElement() else {
            throw NSError(domain: "Issue92OMMLWalkSurfaceCoverageTests", code: 1)
        }
        let document = WordDocument()
        return try DocxReader.parseParagraph(
            from: root,
            relationships: RelationshipsCollection(),
            styles: document.styles,
            numbering: document.numbering
        )
    }
}
