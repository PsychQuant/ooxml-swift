import XCTest
@testable import OOXMLSwift

final class RunTests: XCTestCase {

    // MARK: - RunProperties.toXML()

    func testEmptyProperties() {
        let props = RunProperties()
        XCTAssertEqual(props.toXML(), "")
    }

    func testBold() {
        var props = RunProperties()
        props.bold = true
        XCTAssertEqual(props.toXML(), "<w:b/>")
    }

    func testItalic() {
        var props = RunProperties()
        props.italic = true
        XCTAssertEqual(props.toXML(), "<w:i/>")
    }

    func testBoldItalic() {
        var props = RunProperties()
        props.bold = true
        props.italic = true
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:b/>"))
        XCTAssertTrue(xml.contains("<w:i/>"))
    }

    func testUnderlineSingle() {
        var props = RunProperties()
        props.underline = .single
        XCTAssertTrue(props.toXML().contains("<w:u w:val=\"single\"/>"))
    }

    func testUnderlineDouble() {
        var props = RunProperties()
        props.underline = .double
        XCTAssertTrue(props.toXML().contains("<w:u w:val=\"double\"/>"))
    }

    func testUnderlineDotted() {
        var props = RunProperties()
        props.underline = .dotted
        XCTAssertTrue(props.toXML().contains("<w:u w:val=\"dotted\"/>"))
    }

    func testUnderlineWave() {
        var props = RunProperties()
        props.underline = .wave
        XCTAssertTrue(props.toXML().contains("<w:u w:val=\"wave\"/>"))
    }

    func testStrikethrough() {
        var props = RunProperties()
        props.strikethrough = true
        XCTAssertEqual(props.toXML(), "<w:strike/>")
    }

    func testFontSize() {
        var props = RunProperties()
        props.fontSize = 24  // 12pt
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:sz w:val=\"24\"/>"))
        XCTAssertTrue(xml.contains("<w:szCs w:val=\"24\"/>"))
    }

    func testFontName() {
        var props = RunProperties()
        props.fontName = "Arial"
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("w:ascii=\"Arial\""))
        XCTAssertTrue(xml.contains("w:hAnsi=\"Arial\""))
        XCTAssertTrue(xml.contains("w:eastAsia=\"Arial\""))
        XCTAssertTrue(xml.contains("w:cs=\"Arial\""))
    }

    func testColor() {
        var props = RunProperties()
        props.color = "FF0000"
        XCTAssertTrue(props.toXML().contains("<w:color w:val=\"FF0000\"/>"))
    }

    func testHighlightYellow() {
        var props = RunProperties()
        props.highlight = .yellow
        XCTAssertTrue(props.toXML().contains("<w:highlight w:val=\"yellow\"/>"))
    }

    func testHighlightCyan() {
        var props = RunProperties()
        props.highlight = .cyan
        XCTAssertTrue(props.toXML().contains("<w:highlight w:val=\"cyan\"/>"))
    }

    func testVerticalAlignSuperscript() {
        var props = RunProperties()
        props.verticalAlign = .superscript
        XCTAssertTrue(props.toXML().contains("<w:vertAlign w:val=\"superscript\"/>"))
    }

    func testVerticalAlignSubscript() {
        var props = RunProperties()
        props.verticalAlign = .subscript
        XCTAssertTrue(props.toXML().contains("<w:vertAlign w:val=\"subscript\"/>"))
    }

    func testCharacterSpacing() {
        var props = RunProperties()
        props.characterSpacing = CharacterSpacing(spacing: 20)
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:spacing w:val=\"20\"/>"))
    }

    func testCharacterSpacingWithPosition() {
        var props = RunProperties()
        props.characterSpacing = CharacterSpacing(spacing: 10, position: 5)
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:spacing w:val=\"10\"/>"))
        XCTAssertTrue(xml.contains("<w:position w:val=\"5\"/>"))
    }

    func testCharacterSpacingWithKern() {
        var props = RunProperties()
        props.characterSpacing = CharacterSpacing(kern: 16)
        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:kern w:val=\"16\"/>"))
    }

    // MARK: - Run.toXML()

    func testSimpleRun() {
        let run = Run(text: "Hello")
        let xml = run.toXML()
        XCTAssertTrue(xml.hasPrefix("<w:r>"))
        XCTAssertTrue(xml.hasSuffix("</w:r>"))
        // PsychQuant/ooxml-swift#5 (F13): xml:space="preserve" is now autosense.
        // "Hello" has no leading / trailing / consecutive whitespace, so the
        // flag is omitted (XML normalises any single internal whitespace).
        XCTAssertTrue(xml.contains("<w:t>Hello</w:t>"))
    }

    func testRunWithProperties() {
        let run = Run(text: "Bold", properties: RunProperties(bold: true))
        let xml = run.toXML()
        XCTAssertTrue(xml.contains("<w:rPr><w:b/></w:rPr>"))
        XCTAssertTrue(xml.contains("Bold"))
    }

    func testRunXMLEscaping() {
        let run = Run(text: "A < B & C > D")
        let xml = run.toXML()
        XCTAssertTrue(xml.contains("A &lt; B &amp; C &gt; D"))
    }

    func testRunWithRawXML() {
        var run = Run(text: "ignored")
        run.rawXML = "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>"
        let xml = run.toXML()
        XCTAssertEqual(xml, "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>")
    }

    func testRunPropertiesRawXML() {
        var run = Run(text: "ignored")
        run.properties.rawXML = "<w:sdt><w:sdtContent/></w:sdt>"
        let xml = run.toXML()
        XCTAssertEqual(xml, "<w:sdt><w:sdtContent/></w:sdt>")
    }

    // MARK: - RunProperties Merge

    func testMergeProperties() {
        var base = RunProperties()
        base.bold = true

        var overlay = RunProperties()
        overlay.italic = true
        overlay.fontSize = 24

        base.merge(with: overlay)
        XCTAssertTrue(base.bold)
        XCTAssertTrue(base.italic)
        XCTAssertEqual(base.fontSize, 24)
    }

    // MARK: - Combined Properties

    func testFullyFormattedRun() {
        var props = RunProperties()
        props.bold = true
        props.italic = true
        props.underline = .single
        props.fontSize = 28
        props.fontName = "Times New Roman"
        props.color = "0000FF"

        let xml = props.toXML()
        XCTAssertTrue(xml.contains("<w:b/>"))
        XCTAssertTrue(xml.contains("<w:i/>"))
        XCTAssertTrue(xml.contains("<w:u w:val=\"single\"/>"))
        XCTAssertTrue(xml.contains("<w:sz w:val=\"28\"/>"))
        XCTAssertTrue(xml.contains("w:ascii=\"Times New Roman\""))
        XCTAssertTrue(xml.contains("<w:color w:val=\"0000FF\"/>"))
    }

    // MARK: - ECMA-376 §17.3.2.28 CT_RPr canonical child order
    //
    // PsychQuant/ooxml-swift#61 (kiki830621/collaboration_guo_analysis#20):
    // `RunProperties.toXML()` historically appended children in struct
    // declaration order, which violated the schema-mandated sequence in
    // CT_RPr. macOS Word's strict validator rejected docx files when more
    // than ~10% of `<w:rPr>` blocks had inverted ordering — thesis rescue
    // outputs hit ~65% rate and were completely refused.
    //
    // Canonical order (subset present in this struct, indexed per ECMA-376):
    //   1. rStyle, 2. rFonts, 3. b, 5. i, 9. strike, 15. noProof,
    //   19. color, 20. spacing (CharacterSpacing block: spacing→kern→position),
    //   22. kern (typed field), 24. sz, 25. szCs, 26. highlight, 27. u,
    //   28. effect (TextEffect), 32. vertAlign, 36. lang, then rawChildren.

    func testRunPropertiesEmitsChildrenInCanonicalOrder() {
        var props = RunProperties()
        // Set fields that span the canonical sequence so order matters.
        props.rStyle = "Hyperlink"            // 1
        props.rFonts = RFontsProperties(eastAsia: "DFKai-SB")  // 2
        props.bold = true                     // 3
        props.italic = true                   // 5
        props.strikethrough = true            // 9
        props.noProof = true                  // 15
        props.color = "FF0000"                // 19
        props.kern = 32                       // 22 (typed field)
        props.fontSize = 24                   // 24/25
        props.highlight = .yellow             // 26
        props.underline = .single             // 27
        props.verticalAlign = .superscript    // 32
        props.lang = LanguageProperties(val: "en-US")  // 36

        let xml = props.toXML()

        // Extract child element local names in document order.
        // Match self-closing `<w:foo .../>` and opening `<w:foo>`.
        let pattern = #"<(w(?:14)?:[A-Za-z][A-Za-z0-9]*)\b"#
        let regex = try! NSRegularExpression(pattern: pattern)
        let nsxml = xml as NSString
        let matches = regex.matches(in: xml, range: NSRange(location: 0, length: nsxml.length))
        let names = matches.map { nsxml.substring(with: $0.range(at: 1)) }

        // Canonical position lookup (subset relevant to this struct).
        let canonical: [String: Int] = [
            "w:rStyle":       1,
            "w:rFonts":       2,
            "w:b":            3,
            "w:bCs":          4,
            "w:i":            5,
            "w:iCs":          6,
            "w:strike":       9,
            "w:noProof":      15,
            "w:webHidden":    18,
            "w:color":        19,
            "w:spacing":      20,
            "w:kern":         22,
            "w:position":     23,
            "w:sz":           24,
            "w:szCs":         25,
            "w:highlight":    26,
            "w:u":            27,
            "w:effect":       28,
            "w:vertAlign":    32,
            "w:lang":         36
        ]

        // Walk emitted children pairwise; each pair must be in canonical order.
        let emitted = names.compactMap { canonical[$0] != nil ? $0 : nil }
        for i in 0..<(emitted.count - 1) {
            let a = emitted[i]
            let b = emitted[i + 1]
            // szCs immediately follows sz — same canonical slot pair, skip strict <
            if a == "w:sz" && b == "w:szCs" { continue }
            let posA = canonical[a]!
            let posB = canonical[b]!
            XCTAssertLessThan(
                posA, posB,
                "rPr child order violation: \(a) (canonical pos \(posA)) emitted before \(b) (canonical pos \(posB))\nFull child sequence: \(emitted)\nFull XML: \(xml)"
            )
        }
    }

    func testRunPropertiesRStyleFirstAndLangLastAmongTyped() {
        var props = RunProperties()
        props.lang = LanguageProperties(val: "zh-TW")
        props.rStyle = "Hyperlink"
        props.bold = true
        props.fontSize = 24

        let xml = props.toXML()
        // rStyle must be the first child
        XCTAssertTrue(xml.hasPrefix("<w:rStyle "), "rStyle must be first; got: \(xml)")
        // lang must come after sz/szCs and after b
        let langIdx = xml.range(of: "<w:lang ")!.lowerBound
        let bIdx = xml.range(of: "<w:b/>")!.lowerBound
        let szIdx = xml.range(of: "<w:sz ")!.lowerBound
        XCTAssertLessThan(bIdx, langIdx, "b must precede lang")
        XCTAssertLessThan(szIdx, langIdx, "sz must precede lang")
    }

    func testRunPropertiesRFontsBeforeSizeAndBold() {
        // The canonical-order regression: rFonts (pos 2) was being emitted
        // AFTER b (3), i (5), sz (24). This is the specific shape that broke
        // 4341/6675 rPr blocks in the thesis docx.
        var props = RunProperties()
        props.rFonts = RFontsProperties(eastAsia: "DFKai-SB")
        props.bold = true
        props.fontSize = 36

        let xml = props.toXML()
        let rFontsIdx = xml.range(of: "<w:rFonts ")!.lowerBound
        let bIdx = xml.range(of: "<w:b/>")!.lowerBound
        let szIdx = xml.range(of: "<w:sz ")!.lowerBound
        XCTAssertLessThan(rFontsIdx, bIdx, "rFonts must precede b")
        XCTAssertLessThan(rFontsIdx, szIdx, "rFonts must precede sz")
    }
}
