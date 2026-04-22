import XCTest
@testable import OOXMLSwift

final class OMMLParserTests: XCTestCase {

    // MARK: MathRun round-trip

    func testMathRunRoundTrip() {
        let original = MathRun(text: "x")
        let xml = "<m:oMath>\(original.toOMML())</m:oMath>"
        let parsed = OMMLParser.parse(xml: xml)
        XCTAssertEqual(parsed.count, 1)
        guard let run = parsed.first as? MathRun else {
            XCTFail("Expected MathRun"); return
        }
        XCTAssertEqual(run.text, "x")
    }

    func testMathRunWithStyleRoundTrip() {
        let original = MathRun(text: "x", style: .italic)
        let xml = "<m:oMath>\(original.toOMML())</m:oMath>"
        let parsed = OMMLParser.parse(xml: xml)
        guard let run = parsed.first as? MathRun else {
            XCTFail("Expected MathRun"); return
        }
        XCTAssertEqual(run.text, "x")
        XCTAssertEqual(run.style, .italic)
    }

    func testMathRunXMLEscapeDecoded() {
        let original = MathRun(text: "<a&b>")
        let xml = "<m:oMath>\(original.toOMML())</m:oMath>"
        let parsed = OMMLParser.parse(xml: xml)
        guard let run = parsed.first as? MathRun else {
            XCTFail("Expected MathRun"); return
        }
        XCTAssertEqual(run.text, "<a&b>")
    }

    // MARK: MathFraction round-trip

    func testMathFractionRoundTrip() {
        let original = MathFraction(
            numerator: [MathRun(text: "a")],
            denominator: [MathRun(text: "b")]
        )
        let xml = "<m:oMath>\(original.toOMML())</m:oMath>"
        let parsed = OMMLParser.parse(xml: xml)
        guard let frac = parsed.first as? MathFraction else {
            XCTFail("Expected MathFraction"); return
        }
        guard let numRun = frac.numerator.first as? MathRun,
              let denRun = frac.denominator.first as? MathRun else {
            XCTFail("Expected nested MathRuns"); return
        }
        XCTAssertEqual(numRun.text, "a")
        XCTAssertEqual(denRun.text, "b")
    }

    // MARK: MathSubSuperScript round-trip

    func testMathSubSuperScriptBothSubAndSup() {
        let original = MathSubSuperScript(
            base: [MathRun(text: "x")],
            sub: [MathRun(text: "i")],
            sup: [MathRun(text: "2")]
        )
        let xml = "<m:oMath>\(original.toOMML())</m:oMath>"
        let parsed = OMMLParser.parse(xml: xml)
        guard let ss = parsed.first as? MathSubSuperScript else {
            XCTFail("Expected MathSubSuperScript"); return
        }
        XCTAssertNotNil(ss.sub)
        XCTAssertNotNil(ss.sup)
    }

    func testMathSubSuperScriptOnlySup() {
        let original = MathSubSuperScript(
            base: [MathRun(text: "x")],
            sub: nil,
            sup: [MathRun(text: "2")]
        )
        let xml = "<m:oMath>\(original.toOMML())</m:oMath>"
        let parsed = OMMLParser.parse(xml: xml)
        guard let ss = parsed.first as? MathSubSuperScript else {
            XCTFail("Expected MathSubSuperScript"); return
        }
        XCTAssertNil(ss.sub)
        XCTAssertNotNil(ss.sup)
    }

    // MARK: MathRadical round-trip

    func testMathRadicalSquareRootNoDegree() {
        let original = MathRadical(radicand: [MathRun(text: "2")])
        let xml = "<m:oMath>\(original.toOMML())</m:oMath>"
        let parsed = OMMLParser.parse(xml: xml)
        guard let rad = parsed.first as? MathRadical else {
            XCTFail("Expected MathRadical"); return
        }
        XCTAssertNil(rad.degree)
    }

    func testMathRadicalCubeRoot() {
        let original = MathRadical(
            radicand: [MathRun(text: "x")],
            degree: [MathRun(text: "3")]
        )
        let xml = "<m:oMath>\(original.toOMML())</m:oMath>"
        let parsed = OMMLParser.parse(xml: xml)
        guard let rad = parsed.first as? MathRadical else {
            XCTFail("Expected MathRadical"); return
        }
        XCTAssertNotNil(rad.degree)
    }

    // MARK: MathNary round-trip

    func testMathNarySumWithBounds() {
        let original = MathNary(
            op: .sum,
            sub: [MathRun(text: "i=1")],
            sup: [MathRun(text: "n")],
            base: [MathRun(text: "i")]
        )
        let xml = "<m:oMath>\(original.toOMML())</m:oMath>"
        let parsed = OMMLParser.parse(xml: xml)
        guard let nary = parsed.first as? MathNary else {
            XCTFail("Expected MathNary"); return
        }
        XCTAssertEqual(nary.op, .sum)
        XCTAssertNotNil(nary.sub)
        XCTAssertNotNil(nary.sup)
    }

    // MARK: oMathPara wrapper

    func testOMathParaWrapperStripped() {
        let xml = "<m:oMathPara><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara>"
        let parsed = OMMLParser.parse(xml: xml)
        XCTAssertEqual(parsed.count, 1)
        XCTAssertTrue(parsed.first is MathRun)
    }

    // MARK: Nested composition

    func testFractionInsideRadical() {
        let original = MathRadical(
            radicand: [
                MathFraction(
                    numerator: [MathRun(text: "a")],
                    denominator: [MathRun(text: "b")]
                )
            ]
        )
        let xml = "<m:oMath>\(original.toOMML())</m:oMath>"
        let parsed = OMMLParser.parse(xml: xml)
        guard let rad = parsed.first as? MathRadical else {
            XCTFail("Expected MathRadical"); return
        }
        XCTAssertTrue(rad.radicand.first is MathFraction)
    }

    // MARK: Unknown subtree fallback

    func testUnknownBorderBoxPreservedAsUnknownMath() {
        let xml = "<m:oMath><m:borderBox><m:e><m:r><m:t>x</m:t></m:r></m:e></m:borderBox></m:oMath>"
        let parsed = OMMLParser.parse(xml: xml)
        XCTAssertEqual(parsed.count, 1)
        guard let unknown = parsed.first as? UnknownMath else {
            XCTFail("Expected UnknownMath for m:borderBox"); return
        }
        XCTAssertTrue(unknown.rawXML.contains("<m:borderBox>"))
    }

    // MARK: Empty input

    func testEmptyOMMLReturnsEmpty() {
        let xml = "<m:oMath></m:oMath>"
        let parsed = OMMLParser.parse(xml: xml)
        XCTAssertEqual(parsed.count, 0)
    }
}
