import XCTest
@testable import OOXMLSwift

final class Issue90MathScriptAnchorLookupTests: XCTestCase {

    func testCanonicalizeMathScriptVariantsMapsPinnedGlyphs() {
        XCTAssertEqual(
            AnchorLookupOptions.canonicalizeMathScriptVariants("H₀ + xᵢ = y²"),
            "H0 + xi = y2"
        )
        XCTAssertEqual(
            AnchorLookupOptions.canonicalizeMathScriptVariants("A⁽¹⁾ ₊ B₍₂₎"),
            "A(1) + B(2)"
        )
    }

    func testCanonicalizeMathScriptVariantsPreservesUnsupportedCharacters() {
        XCTAssertEqual(
            AnchorLookupOptions.canonicalizeMathScriptVariants("ᾱ ∴ H₀"),
            "ᾱ ∴ H0"
        )
    }

    func testDefaultExactModeDoesNotMatchUnicodeSubscriptNeedleAgainstASCIIHaystack() {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "H0 is rejected")),
        ]

        XCTAssertNil(doc.findBodyChildContainingText("H₀"))
        XCTAssertFalse(WordDocument.bodyChildContainsText(
            .paragraph(Paragraph(text: "H0 is rejected")),
            needle: "H₀"
        ))
    }

    func testMathScriptInsensitiveMatchesUnicodeSubscriptNeedleAgainstASCIIHaystack() {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "H0 is rejected")),
        ]

        XCTAssertEqual(
            doc.findBodyChildContainingText(
                "H₀",
                options: AnchorLookupOptions(mathScriptInsensitive: true)
            ),
            0
        )
    }

    func testMathScriptInsensitiveMatchesASCIINeedleAgainstUnicodeSubscriptHaystack() {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "H₀ is rejected")),
        ]

        XCTAssertEqual(
            doc.findBodyChildContainingText(
                "H0",
                options: AnchorLookupOptions(mathScriptInsensitive: true)
            ),
            0
        )
    }

    func testMathScriptInsensitiveMatchesSubscriptAndSuperscriptLetters() {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "xᵢ + y²")),
        ]

        XCTAssertEqual(
            doc.findBodyChildContainingText(
                "xi + y2",
                options: AnchorLookupOptions(mathScriptInsensitive: true)
            ),
            0
        )
    }

    func testNthInstanceCountsNormalizedMatches() {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "H0 first")),
            .paragraph(Paragraph(text: "H₀ second")),
        ]

        XCTAssertEqual(
            doc.findBodyChildContainingText(
                "H₀",
                nthInstance: 2,
                options: AnchorLookupOptions(mathScriptInsensitive: true)
            ),
            1
        )
    }

    func testInsertLocationAfterTextAcceptsMathScriptInsensitiveOption() throws {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "H0 anchor")),
        ]

        try doc.insertParagraph(
            Paragraph(text: "inserted"),
            at: .afterText(
                "H₀",
                instance: 1,
                options: AnchorLookupOptions(mathScriptInsensitive: true)
            )
        )

        XCTAssertEqual(doc.body.children.count, 2)
        guard case .paragraph(let inserted) = doc.body.children[1] else {
            XCTFail("Expected inserted paragraph at index 1")
            return
        }
        XCTAssertEqual(inserted.runs.first?.text, "inserted")
    }

    func testMathSubSuperScriptVisibleTextRemainsASCII() {
        let sss = MathSubSuperScript(
            base: [MathRun(text: "H")],
            sub: [MathRun(text: "0")]
        )

        XCTAssertEqual(sss.visibleText, "H0")
    }
}
