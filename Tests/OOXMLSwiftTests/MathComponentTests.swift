import XCTest
@testable import OOXMLSwift

final class MathComponentTests: XCTestCase {

    // MARK: - MathRun

    func testMathRunPlain() {
        XCTAssertEqual(MathRun(text: "x").toOMML(), "<m:r><m:t>x</m:t></m:r>")
    }

    func testMathRunItalic() {
        XCTAssertEqual(
            MathRun(text: "x", style: .italic).toOMML(),
            "<m:r><m:rPr><m:sty m:val=\"i\"/></m:rPr><m:t>x</m:t></m:r>"
        )
    }

    func testMathRunXMLEscape() {
        XCTAssertEqual(
            MathRun(text: "<a&b>").toOMML(),
            "<m:r><m:t>&lt;a&amp;b&gt;</m:t></m:r>"
        )
    }

    // MARK: - MathFraction

    func testMathFractionSimple() {
        let frac = MathFraction(
            numerator: [MathRun(text: "a")],
            denominator: [MathRun(text: "b")]
        )
        let xml = frac.toOMML()
        XCTAssertTrue(xml.contains("<m:f>"))
        XCTAssertTrue(xml.contains("<m:num><m:r><m:t>a</m:t></m:r></m:num>"))
        XCTAssertTrue(xml.contains("<m:den><m:r><m:t>b</m:t></m:r></m:den>"))
    }

    func testMathFractionSkewedBar() {
        let frac = MathFraction(
            numerator: [MathRun(text: "a")],
            denominator: [MathRun(text: "b")],
            barStyle: .skewed
        )
        XCTAssertTrue(frac.toOMML().contains("<m:type m:val=\"skw\"/>"))
    }

    // MARK: - MathSubSuperScript

    func testMathSubSuperScriptBoth() {
        let ss = MathSubSuperScript(
            base: [MathRun(text: "x")],
            sub: [MathRun(text: "i")],
            sup: [MathRun(text: "2")]
        )
        XCTAssertTrue(ss.toOMML().hasPrefix("<m:sSubSup>"))
    }

    func testMathSubSuperScriptOnlySup() {
        let ss = MathSubSuperScript(
            base: [MathRun(text: "x")],
            sub: nil,
            sup: [MathRun(text: "2")]
        )
        XCTAssertTrue(ss.toOMML().hasPrefix("<m:sSup>"))
        XCTAssertFalse(ss.toOMML().contains("<m:sub>"))
    }

    func testMathSubSuperScriptOnlySub() {
        let ss = MathSubSuperScript(
            base: [MathRun(text: "x")],
            sub: [MathRun(text: "i")],
            sup: nil
        )
        XCTAssertTrue(ss.toOMML().hasPrefix("<m:sSub>"))
        XCTAssertFalse(ss.toOMML().contains("<m:sup>"))
    }

    // MARK: - MathRadical

    func testMathRadicalSquareRoot() {
        let rad = MathRadical(radicand: [MathRun(text: "2")])
        let xml = rad.toOMML()
        XCTAssertTrue(xml.contains("<m:rad>"))
        XCTAssertTrue(xml.contains("<m:degHide m:val=\"1\"/>"))
    }

    func testMathRadicalCubeRoot() {
        let rad = MathRadical(
            radicand: [MathRun(text: "x")],
            degree: [MathRun(text: "3")]
        )
        let xml = rad.toOMML()
        XCTAssertFalse(xml.contains("degHide"))
        XCTAssertTrue(xml.contains("<m:deg><m:r><m:t>3</m:t></m:r></m:deg>"))
    }

    // MARK: - MathNary

    func testMathNarySumWithBounds() {
        let nary = MathNary(
            op: .sum,
            sub: [MathRun(text: "i=1")],
            sup: [MathRun(text: "n")],
            base: [MathRun(text: "i")]
        )
        let xml = nary.toOMML()
        XCTAssertTrue(xml.contains("<m:nary>"))
        XCTAssertTrue(xml.contains("<m:chr m:val=\"∑\"/>"))
        XCTAssertTrue(xml.contains("<m:sub><m:r><m:t>i=1</m:t></m:r></m:sub>"))
        XCTAssertTrue(xml.contains("<m:sup><m:r><m:t>n</m:t></m:r></m:sup>"))
    }

    func testMathNaryIntegral() {
        let nary = MathNary(op: .integral, base: [MathRun(text: "f(x)dx")])
        XCTAssertTrue(nary.toOMML().contains("<m:chr m:val=\"∫\"/>"))
    }

    // MARK: - MathDelimiter

    func testMathDelimiterParens() {
        let d = MathDelimiter(
            open: "(",
            close: ")",
            elements: [[MathRun(text: "x")]]
        )
        let xml = d.toOMML()
        XCTAssertTrue(xml.contains("<m:d>"))
        XCTAssertTrue(xml.contains("<m:begChr m:val=\"(\"/>"))
        XCTAssertTrue(xml.contains("<m:endChr m:val=\")\"/>"))
    }

    func testMathDelimiterMultiElementWithSeparator() {
        let d = MathDelimiter(
            open: "(",
            close: ")",
            elements: [[MathRun(text: "a")], [MathRun(text: "b")]],
            separator: ","
        )
        let xml = d.toOMML()
        XCTAssertTrue(xml.contains("<m:sepChr m:val=\",\"/>"))
        // two <m:e> blocks for two elements
        XCTAssertEqual(xml.components(separatedBy: "<m:e>").count - 1, 2)
    }

    // MARK: - MathFunction

    func testMathFunctionSin() {
        let f = MathFunction(
            functionName: [MathRun(text: "sin")],
            argument: [MathRun(text: "x")]
        )
        let xml = f.toOMML()
        XCTAssertTrue(xml.contains("<m:func>"))
        XCTAssertTrue(xml.contains("<m:fName><m:r><m:t>sin</m:t></m:r></m:fName>"))
        XCTAssertTrue(xml.contains("<m:e><m:r><m:t>x</m:t></m:r></m:e>"))
    }

    // MARK: - MathLimit

    func testMathLimitLower() {
        let l = MathLimit(
            position: .lower,
            base: [MathRun(text: "lim")],
            limit: [MathRun(text: "x→0")]
        )
        XCTAssertTrue(l.toOMML().contains("<m:limLow>"))
        XCTAssertFalse(l.toOMML().contains("<m:limUpp>"))
    }

    func testMathLimitUpper() {
        let l = MathLimit(
            position: .upper,
            base: [MathRun(text: "x")],
            limit: [MathRun(text: "bar")]
        )
        XCTAssertTrue(l.toOMML().contains("<m:limUpp>"))
    }

    // MARK: - MathMatrix

    func testMathMatrix2x2() {
        let m = MathMatrix(rows: [
            [[MathRun(text: "a")], [MathRun(text: "b")]],
            [[MathRun(text: "c")], [MathRun(text: "d")]]
        ])
        let xml = m.toOMML()
        XCTAssertTrue(xml.contains("<m:m>"))
        XCTAssertEqual(xml.components(separatedBy: "<m:mr>").count - 1, 2, "expected 2 rows")
    }

    // MARK: - Nested composition

    func testFractionInsideRadical() {
        let expr = MathRadical(
            radicand: [
                MathFraction(
                    numerator: [MathRun(text: "a")],
                    denominator: [MathRun(text: "b")]
                )
            ]
        )
        let xml = expr.toOMML()
        XCTAssertTrue(xml.contains("<m:rad>"))
        XCTAssertTrue(xml.contains("<m:f>"))
        // Fraction should be inside the radicand `<m:e>` block
        XCTAssertTrue(xml.contains("<m:e><m:f>"))
    }

    func testSubscriptOnNaryBase() {
        // ∑_{i=1}^{n} x_i
        let expr = MathNary(
            op: .sum,
            sub: [MathRun(text: "i=1")],
            sup: [MathRun(text: "n")],
            base: [
                MathSubSuperScript(
                    base: [MathRun(text: "x")],
                    sub: [MathRun(text: "i")],
                    sup: nil
                )
            ]
        )
        let xml = expr.toOMML()
        XCTAssertTrue(xml.contains("<m:nary>"))
        XCTAssertTrue(xml.contains("<m:sSub>"))
    }

    // MARK: - MathAccent (added v0.11.0)

    func testMathAccentHatOverSingleRun() {
        let acc = MathAccent(base: [MathRun(text: "x")], accentChar: "\u{0302}")
        XCTAssertEqual(
            acc.toOMML(),
            "<m:acc><m:accPr><m:chr m:val=\"\u{0302}\"/></m:accPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:acc>"
        )
    }

    func testMathAccentBarOverGreekLetter() {
        let acc = MathAccent(base: [MathRun(text: "ρ")], accentChar: "\u{0304}")
        let xml = acc.toOMML()
        XCTAssertTrue(xml.contains("<m:chr m:val=\"\u{0304}\"/>"))
        XCTAssertTrue(xml.contains("<m:e><m:r><m:t>ρ</m:t></m:r></m:e>"))
        XCTAssertLessThan(
            xml.range(of: "<m:chr")!.lowerBound,
            xml.range(of: "<m:e>")!.lowerBound,
            "<m:chr> must precede <m:e>"
        )
    }

    func testMathAccentOverCompositeBase() {
        let inner = MathSubSuperScript(
            base: [MathRun(text: "ε")],
            sub: [MathRun(text: "t")],
            sup: nil
        )
        let acc = MathAccent(base: [inner], accentChar: "\u{0302}")
        XCTAssertTrue(
            acc.toOMML().contains(
                "<m:e><m:sSub><m:e><m:r><m:t>ε</m:t></m:r></m:e><m:sub><m:r><m:t>t</m:t></m:r></m:sub></m:sSub></m:e>"
            )
        )
    }

    func testMathAccentXMLEscapesAccentChar() {
        let acc = MathAccent(base: [MathRun(text: "y")], accentChar: "&")
        XCTAssertTrue(acc.toOMML().contains("<m:chr m:val=\"&amp;\"/>"))
    }
}
