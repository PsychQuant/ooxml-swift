import XCTest
@testable import OOXMLSwift

/// Regression tests for `Paragraph.getText()` delegating to
/// `flattenedDisplayText()` so the two text-extraction paths return
/// identical strings (PsychQuant/ooxml-swift#43, mirror of
/// PsychQuant/che-word-mcp#155).
///
/// **Pre-fix bug**: Two divergent text-extraction paths existed:
///
/// | Method | Walks OMath? | Walks unrecognizedChildren? | Walks fieldSimples / AC? |
/// |--------|--------------|------------------------------|---------------------------|
/// | `Paragraph.getText()` (legacy)         | ❌ no | ❌ no | ❌ no |
/// | `Paragraph.flattenedDisplayText()` (#85+#92+#99-#103) | ✅ yes | ✅ yes | ✅ yes |
///
/// `che-word-mcp__search_text` calls `getText()`, so callers couldn't grep
/// for inline math symbols (α / β / γ / θ / λ / t) that lived inside
/// `<m:oMath>` — silent zero gaps in match positions.
///
/// **Post-fix**: `getText()` delegates to `flattenedDisplayText()`, so all
/// callers see consistent text including OMath visibleText, OMML inside
/// hyperlinks/fieldSimples/AC fallbacks, and content controls.
///
/// Test fixture lesson from #93 release notes: real Word output wraps
/// inline `<m:oMath>` with empty `<w:r><w:t></w:t></w:r>` runs before and
/// after. Synthetic `Paragraph(text:)` constructor doesn't reproduce that
/// shape and would have hidden the bug. These tests use real OOXML round-
/// trip via `DocxReader.parseParagraph` to ensure we actually catch the
/// `getText()` divergence.
final class Issue43GetTextDelegatesToFlattenedDisplayTextTests: XCTestCase {

    // MARK: - Fixture builder

    /// Parse inline `<w:p>` XML via `DocxReader.parseParagraph`. Mirrors
    /// the helper from `Issue99FlattenReplaceOMMLBilateralTests` so we
    /// build paragraphs the same way real `.docx` parsing does.
    private func parseParagraph(xml: String) throws -> Paragraph {
        let data = xml.data(using: .utf8)!
        let doc = try XMLDocument(data: data)
        guard let root = doc.rootElement() else {
            throw NSError(domain: "Issue43GetTextDelegatesToFlattenedDisplayTextTests", code: 1)
        }
        let document = WordDocument()
        return try DocxReader.parseParagraph(
            from: root,
            relationships: RelationshipsCollection(),
            styles: document.styles,
            numbering: document.numbering
        )
    }

    // MARK: - Real-Word-output fixture (the actual bug)

    /// Reproduces `碩士論文.docx` para 324 from the issue body:
    ///
    /// ```
    /// <w:r>模型所得出的參數進行</w:r>
    /// <w:r></w:r>                          <!-- empty wrapper before OMath -->
    /// <m:oMath><m:r><m:t>t</m:t></m:r></m:oMath>
    /// <w:r></w:r>                          <!-- empty wrapper after OMath -->
    /// <w:r>檢定：</w:r>
    /// ```
    ///
    /// Real Word's "Insert Equation" inline output emits the empty
    /// wrapper runs around `<m:oMath>` direct children. Synthetic
    /// `Paragraph(text:)` would NOT produce this shape, so this test
    /// is the load-bearing regression guard.
    private func paragraphFixtureRealWordOutput() throws -> Paragraph {
        let xml = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
          <w:r><w:t xml:space="preserve">模型所得出的參數進行</w:t></w:r>
          <w:r><w:t xml:space="preserve"></w:t></w:r>
          <m:oMath><m:r><m:t>t</m:t></m:r></m:oMath>
          <w:r><w:t xml:space="preserve"></w:t></w:r>
          <w:r><w:t xml:space="preserve">檢定：</w:t></w:r>
        </w:p>
        """
        return try parseParagraph(xml: xml)
    }

    // MARK: - Core invariant: the two methods MUST agree

    /// Primary regression test for #43: `getText()` and
    /// `flattenedDisplayText()` MUST return the same string. Pre-fix this
    /// failed because `getText()` only walked `runs.map { $0.text }` and
    /// missed direct-child OMML in `unrecognizedChildren`.
    func testGetTextEqualsFlattenedDisplayTextOnRealWordFixture_Issue43() throws {
        let para = try paragraphFixtureRealWordOutput()

        let getTextOutput = para.getText()
        let flattenedOutput = para.flattenedDisplayText()

        XCTAssertEqual(
            getTextOutput,
            flattenedOutput,
            """
            getText() and flattenedDisplayText() MUST return identical strings
            (post-#43 fix). Pre-fix divergence:
              getText():            \(String(reflecting: getTextOutput))
              flattenedDisplayText: \(String(reflecting: flattenedOutput))
            See PsychQuant/ooxml-swift#43 for the analysis of the two
            divergent paths and the unification fix.
            """
        )
    }

    /// Concrete content assertion (orthogonal to the equality check above).
    /// Verifies `getText()` actually includes the `t` from OMath, not just
    /// that it agrees with flattenedDisplayText. Position arithmetic from
    /// the issue body: `進行` ends at offset, `t` immediately follows,
    /// then `檢定：`.
    func testGetTextIncludesOMathContentInRealWordFixture_Issue43() throws {
        let para = try paragraphFixtureRealWordOutput()
        let text = para.getText()

        XCTAssertTrue(
            text.contains("進行t檢定"),
            """
            getText() should include OMath text content directly between
            adjacent run text — `進行` + OMath `t` + `檢定` should appear
            as the contiguous substring `進行t檢定`. Got: \(String(reflecting: text)).
            Pre-fix bug: OMath was silently dropped, producing `進行檢定`
            (no `t`), causing che-word-mcp__search_text to return positions
            inconsistent with anchor-matching paths.
            """
        )
    }

    /// Negative test: confirm the BUG signature — `getText()` was previously
    /// returning `進行檢定` (no `t`). After fix, this MUST NOT be the case.
    /// Documented as anti-regression: if anyone accidentally reverts
    /// `getText()` to legacy behavior, this test fires loud.
    func testGetTextDoesNotProducePreFixBugSignature_Issue43() throws {
        let para = try paragraphFixtureRealWordOutput()
        let text = para.getText()

        XCTAssertFalse(
            text.contains("進行檢定") && !text.contains("進行t檢定"),
            """
            ANTI-REGRESSION: pre-#43 bug signature detected — getText() returned
            text containing `進行檢定` WITHOUT `進行t檢定`, meaning it stripped
            the OMath `t`. This is the exact bug the fix collapses.
            """
        )
    }

    // MARK: - Plain text path (no regression for non-OMath paragraphs)

    /// Sanity check: paragraphs without OMath should still extract identically
    /// (no regression for the common case). Both methods should handle plain
    /// text the same way they did before the fix.
    func testGetTextEqualsFlattenedDisplayTextOnPlainTextParagraph_Issue43() throws {
        let xml = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r><w:t xml:space="preserve">Hello </w:t></w:r>
          <w:r><w:t xml:space="preserve">world</w:t></w:r>
        </w:p>
        """
        let para = try parseParagraph(xml: xml)

        XCTAssertEqual(para.getText(), para.flattenedDisplayText())
        XCTAssertEqual(para.getText(), "Hello world")
    }
}
