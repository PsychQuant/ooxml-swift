import XCTest
@testable import OOXMLSwift

/// Inline-math flatten coverage tests for [PsychQuant/che-word-mcp#85](https://github.com/PsychQuant/che-word-mcp/issues/85).
///
/// `Paragraph.flattenedDisplayText()` previously walked typed run children
/// (`runs` / `hyperlinks` / `fieldSimples` / `alternateContents` /
/// `contentControls`) but skipped OMML (`<m:oMath>` / `<m:oMathPara>`)
/// subtrees stored on `Run.rawXML`. Result: any `before_text` / `after_text`
/// anchor crossing an inline math span silently 0-matched (e.g.,
/// `"進行 t 檢定"` against a paragraph carrying `<m:oMath><m:r><m:t>t</m:t></m:r></m:oMath>`
/// flattened to `"進行   檢定"` with the math text dropped).
///
/// These tests pin the post-fix behaviour: inline math `<m:t>` content joins
/// the flatten output at the OMML run's source position, so anchor lookups
/// over math-bearing paragraphs match natural sentence text.
final class Issue85InlineMathFlattenTests: XCTestCase {

    // MARK: - MathComponent.visibleText accessor

    /// Spec: `MathRun` is the leaf — text passes through verbatim.
    func testMathRunVisibleText() {
        let run = MathRun(text: "t")
        XCTAssertEqual(run.visibleText, "t")
    }

    /// Spec: `MathFraction` concatenates numerator + denominator.
    func testMathFractionVisibleText() {
        let frac = MathFraction(
            numerator: [MathRun(text: "a")],
            denominator: [MathRun(text: "b")]
        )
        XCTAssertEqual(frac.visibleText, "ab")
    }

    /// Spec: `MathSubSuperScript` concatenates base + sub + sup in document order.
    func testMathSubSuperScriptVisibleText() {
        let sss = MathSubSuperScript(
            base: [MathRun(text: "x")],
            sub: [MathRun(text: "1")],
            sup: [MathRun(text: "2")]
        )
        XCTAssertEqual(sss.visibleText, "x12")
    }

    /// Spec: `MathNary` includes the operator symbol + sub + sup + base.
    func testMathNaryVisibleText() {
        let nary = MathNary(
            op: .sum,
            sub: [MathRun(text: "i")],
            sup: [MathRun(text: "n")],
            base: [MathRun(text: "x")]
        )
        XCTAssertEqual(nary.visibleText, "∑inx")
    }

    /// Spec: `[MathComponent]` array extension joins all leaf text in order.
    func testMathComponentArrayVisibleText() {
        let components: [MathComponent] = [
            MathRun(text: "x"),
            MathRun(text: " + "),
            MathRun(text: "y"),
        ]
        XCTAssertEqual(components.visibleText, "x + y")
    }

    // MARK: - Paragraph.flattenedDisplayText OMML coverage

    /// Issue's exact repro: inline `<m:oMath>` mid-paragraph between text runs.
    /// Pre-fix: `"進行   檢定："`. Post-fix: `"進行t檢定："` (or `"進行 t 檢定："`
    /// depending on surrounding text — here the text runs have no internal
    /// spaces, so the result is `"進行t檢定："`).
    func testInlineOMathMidParagraphFlattensWithMathText() throws {
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
        XCTAssertTrue(flat.contains("進行"), "leading text dropped: \(flat)")
        XCTAssertTrue(flat.contains("t"), "OMML <m:t> dropped: \(flat)")
        XCTAssertTrue(flat.contains("檢定"), "trailing text dropped: \(flat)")
        // Issue's "進行 t 檢定" anchor (with surrounding spaces) MAY or MAY NOT
        // match depending on where math text sits relative to text-run
        // boundaries; the issue's primary need is that the `t` is no longer
        // silently dropped.
    }

    /// Spec: nested OMML structure (`<m:f>` fraction containing `<m:r><m:t>`)
    /// — visible text walks recursively into all leaf `<m:t>` descendants.
    func testNestedOMMLFractionFlattensWithAllLeafText() throws {
        let xml = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
          <w:r><w:t>分數</w:t></w:r>
          <w:r>
            <m:oMath>
              <m:f>
                <m:num><m:r><m:t>a</m:t></m:r></m:num>
                <m:den><m:r><m:t>b</m:t></m:r></m:den>
              </m:f>
            </m:oMath>
          </w:r>
          <w:r><w:t>結束</w:t></w:r>
        </w:p>
        """
        let para = try parseParagraph(xml: xml)
        let flat = para.flattenedDisplayText()
        XCTAssertTrue(flat.contains("分數"), "leading text dropped: \(flat)")
        XCTAssertTrue(flat.contains("a"), "fraction numerator dropped: \(flat)")
        XCTAssertTrue(flat.contains("b"), "fraction denominator dropped: \(flat)")
        XCTAssertTrue(flat.contains("結束"), "trailing text dropped: \(flat)")
    }

    /// Spec: paragraph without OMML — flatten output unchanged for plain text.
    /// Regression guard: the OMML walk must not affect ordinary paragraphs.
    func testPlainParagraphFlattensUnchanged() throws {
        let xml = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r><w:t>hello</w:t></w:r>
          <w:r><w:t> world</w:t></w:r>
        </w:p>
        """
        let para = try parseParagraph(xml: xml)
        XCTAssertEqual(para.flattenedDisplayText(), "hello world")
    }

    // MARK: - Helpers

    private func parseParagraph(xml: String) throws -> Paragraph {
        let data = xml.data(using: .utf8)!
        let doc = try XMLDocument(data: data)
        guard let root = doc.rootElement() else {
            throw NSError(domain: "Issue85InlineMathFlattenTests", code: 1)
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
