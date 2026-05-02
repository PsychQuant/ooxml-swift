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

    func testCanonicalizeMathScriptVariantsPreservesUnmappedNonMarkCharacters() {
        // `∴`, `∮`, `Æ` are not in the math-script map and have no NFD
        // decomposition, so they pass through unchanged. (Note: `≠` is NOT
        // a good example here because NFD decomposes it to `=` + combining
        // long solidus overlay, and the canonicalizer strips the overlay.)
        XCTAssertEqual(
            AnchorLookupOptions.canonicalizeMathScriptVariants("Æ ∴ ∮ H₀"),
            "Æ ∴ ∮ H0"
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

    // MARK: - Combining macron / accent stripping (X̄, ŷ, etc.)

    func testCanonicalizeStripsCombiningMacronOnMathLetters() {
        // X̄ (NFC: U+0058 + U+0304) → X
        XCTAssertEqual(
            AnchorLookupOptions.canonicalizeMathScriptVariants("X̄ + Ȳ"),
            "X + Y"
        )
    }

    func testCanonicalizeHandlesDecomposedAndPrecomposedAccentsIdentically() {
        // ŷ precomposed (U+0177) and ŷ decomposed (y + U+0302 hat) must
        // canonicalize to the same string.
        let precomposed = "\u{0177}"             // ŷ
        let decomposed = "y\u{0302}"             // y + combining hat
        XCTAssertEqual(
            AnchorLookupOptions.canonicalizeMathScriptVariants(precomposed),
            AnchorLookupOptions.canonicalizeMathScriptVariants(decomposed)
        )
        XCTAssertEqual(
            AnchorLookupOptions.canonicalizeMathScriptVariants(precomposed),
            "y"
        )
    }

    func testMathScriptInsensitiveMatchesXBarAgainstX() {
        // Issue #90 body explicitly lists X̄ as an anchor pattern. OMML
        // `MathAccent.visibleText` emits just the base ("X"), so the user-
        // typed needle X̄ must match the haystack's plain X.
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "the sample mean X is significant")),
        ]

        XCTAssertEqual(
            doc.findBodyChildContainingText(
                "X̄",
                options: AnchorLookupOptions(mathScriptInsensitive: true)
            ),
            0
        )
    }

    func testMathScriptInsensitiveMatchesXAgainstXBar() {
        // Reverse direction: haystack contains the precomposed combining-mark
        // form, needle is plain ASCII.
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "the sample mean X̄ is significant")),
        ]

        XCTAssertEqual(
            doc.findBodyChildContainingText(
                "X",
                options: AnchorLookupOptions(mathScriptInsensitive: true)
            ),
            0
        )
    }

    // MARK: - Greek subscripts (U+1D66..U+1D6A)

    func testCanonicalizeMapsGreekSubscriptsToGreekLetters() {
        XCTAssertEqual(
            AnchorLookupOptions.canonicalizeMathScriptVariants("Σᵦ + Πᵧ + ρᵨ"),
            "Σβ + Πγ + ρρ"
        )
    }

    func testMathScriptInsensitiveMatchesGreekSubscriptVariants() {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "regression coefficient Σβ stable")),
        ]

        XCTAssertEqual(
            doc.findBodyChildContainingText(
                "Σᵦ",
                options: AnchorLookupOptions(mathScriptInsensitive: true)
            ),
            0
        )
    }

    // MARK: - U+2094 subscript schwa (completes the U+2090..U+209C block)

    func testCanonicalizeMapsSubscriptSchwa() {
        XCTAssertEqual(
            AnchorLookupOptions.canonicalizeMathScriptVariants("aₔb"),
            "aəb"
        )
    }

    // MARK: - Negative cases (over-fold prevention)

    func testMathScriptInsensitiveDoesNotMatchDifferentSubscriptDigits() {
        // H₂ must NOT match H₃ even with the flag enabled.
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "Hypothesis H₃ stated above")),
        ]

        XCTAssertNil(
            doc.findBodyChildContainingText(
                "H₂",
                options: AnchorLookupOptions(mathScriptInsensitive: true)
            )
        )
        XCTAssertNil(
            doc.findBodyChildContainingText(
                "H2",
                options: AnchorLookupOptions(mathScriptInsensitive: true)
            )
        )
    }

    func testMathScriptInsensitiveDoesNotConflateGreekAndLatinLetters() {
        // Greek alpha α (U+03B1) is NOT mapped to ASCII a; canonicalization
        // leaves it as-is so it does not over-fold to ASCII variables.
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "variable a defined")),
        ]

        XCTAssertNil(
            doc.findBodyChildContainingText(
                "α",
                options: AnchorLookupOptions(mathScriptInsensitive: true)
            )
        )
    }

    func testMathScriptInsensitiveDoesNotMatchAcrossBoundary() {
        // Needle "H₀p" requires a literal `p` immediately after the (folded)
        // subscript zero; haystack "H_0 p" with a space must NOT match.
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "H_0 p stands for posterior")),
        ]

        XCTAssertNil(
            doc.findBodyChildContainingText(
                "H₀p",
                options: AnchorLookupOptions(mathScriptInsensitive: true)
            )
        )
    }
}
