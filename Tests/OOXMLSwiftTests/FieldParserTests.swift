import XCTest
@testable import OOXMLSwift

final class FieldParserTests: XCTestCase {

    // Helper: build a paragraph with one field run containing the 5-run embedded XML
    private func paragraphWithFieldXML(_ fieldXML: String) -> Paragraph {
        var run = Run(text: "")
        run.rawXML = fieldXML
        var para = Paragraph()
        para.runs = [run]
        return para
    }

    // MARK: SEQ round-trip

    func testParseSequenceFieldFromToFieldXML() {
        let field = SequenceField(identifier: "Figure", format: .arabic, resetLevel: 1, cachedResult: "3")
        let xml = field.toFieldXML()
        let paragraph = paragraphWithFieldXML(xml)

        let parsed = FieldParser.parse(paragraph: paragraph)
        XCTAssertEqual(parsed.count, 1)
        guard let p = parsed.first, case .sequence(let seq) = p.field else {
            XCTFail("Expected .sequence case"); return
        }
        XCTAssertEqual(seq.identifier, "Figure")
        XCTAssertEqual(seq.format, .arabic)
        XCTAssertEqual(seq.resetLevel, 1)
        XCTAssertEqual(p.startRunIdx, 0)
        XCTAssertEqual(p.endRunIdx, 0)
        XCTAssertEqual(p.cachedResultRunIdx, 0)
    }

    func testParseChineseIdentifier() {
        let field = SequenceField(identifier: "圖", format: .arabic, resetLevel: 1, cachedResult: "1")
        let paragraph = paragraphWithFieldXML(field.toFieldXML())

        let parsed = FieldParser.parse(paragraph: paragraph)
        XCTAssertEqual(parsed.count, 1)
        if case .sequence(let seq) = parsed[0].field {
            XCTAssertEqual(seq.identifier, "圖")
        } else {
            XCTFail("Expected .sequence")
        }
    }

    // MARK: STYLEREF round-trip

    func testParseStyleRefField() {
        let field = StyleRefField(headingLevel: 1, suppressNonDelimiter: true, cachedResult: "4")
        let paragraph = paragraphWithFieldXML(field.toFieldXML())

        let parsed = FieldParser.parse(paragraph: paragraph)
        XCTAssertEqual(parsed.count, 1)
        if case .styleRef(let sr) = parsed[0].field {
            XCTAssertEqual(sr.headingLevel, 1)
            XCTAssertTrue(sr.suppressNonDelimiter)
        } else {
            XCTFail("Expected .styleRef")
        }
    }

    // MARK: REF round-trip

    func testParseReferenceField() {
        let field = ReferenceField(type: .ref, bookmarkName: "fig_returns", createHyperlink: true, cachedResult: "圖 4-1")
        let paragraph = paragraphWithFieldXML(field.toFieldXML())

        let parsed = FieldParser.parse(paragraph: paragraph)
        XCTAssertEqual(parsed.count, 1)
        if case .reference(let ref) = parsed[0].field {
            XCTAssertEqual(ref.type, .ref)
            XCTAssertEqual(ref.bookmarkName, "fig_returns")
            XCTAssertTrue(ref.createHyperlink)
        } else {
            XCTFail("Expected .reference")
        }
    }

    // MARK: Unknown field fallback

    func testUnknownFieldPreservedAsOpaque() {
        // Construct a field block manually using an unrecognized field type.
        let unknownXML = """
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>\
        <w:r><w:instrText xml:space="preserve"> TIME \\@ "hh:mm" </w:instrText></w:r>\
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>\
        <w:r><w:t>12:34</w:t></w:r>\
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
        """
        let paragraph = paragraphWithFieldXML(unknownXML)

        let parsed = FieldParser.parse(paragraph: paragraph)
        XCTAssertEqual(parsed.count, 1)
        if case .unknown(let instr) = parsed[0].field {
            XCTAssertTrue(instr.contains("TIME"))
        } else {
            XCTFail("Expected .unknown for TIME field")
        }
    }

    // MARK: Paragraph without any field runs

    func testEmptyParagraphReturnsEmptyArray() {
        let paragraph = Paragraph(text: "plain text no fields")
        XCTAssertEqual(FieldParser.parse(paragraph: paragraph), [])
    }

    // MARK: Multiple field blocks in one paragraph run (caption with chapter prefix)

    func testMultipleFieldsInOneRawXML() {
        let stylerRef = StyleRefField(headingLevel: 1, suppressNonDelimiter: true, cachedResult: "4")
        let seq = SequenceField(identifier: "Figure", resetLevel: 1, cachedResult: "2")
        let combined = stylerRef.toFieldXML() + "<w:r><w:t>-</w:t></w:r>" + seq.toFieldXML()
        let paragraph = paragraphWithFieldXML(combined)

        let parsed = FieldParser.parse(paragraph: paragraph)
        XCTAssertEqual(parsed.count, 2)
        // Order preserved by regex match order
        if case .styleRef = parsed[0].field {} else { XCTFail("Expected first to be styleRef") }
        if case .sequence = parsed[1].field {} else { XCTFail("Expected second to be sequence") }
    }

    // MARK: Issue 26 wrapper surfaces

    func testFieldParserHandlesFieldSimpleSEQ() {
        var para = Paragraph()
        para.fieldSimples = [
            FieldSimple(instr: " SEQ Figure \\* ARABIC ", runs: [Run(text: "1")])
        ]

        let parsed = FieldParser.parse(paragraph: para)
        XCTAssertEqual(parsed.count, 1)
        guard let first = parsed.first, case .sequence(let seq) = first.field else {
            return XCTFail("Expected .sequence")
        }
        XCTAssertEqual(seq.identifier, "Figure")
    }

    func testFieldParserHandlesInlineSDTSEQ() {
        let field = SequenceField(identifier: "Figure", cachedResult: "1")
        let sdt = StructuredDocumentTag(id: 2601, tag: "inline-seq")
        var para = Paragraph()
        para.contentControls = [
            ContentControl(sdt: sdt, content: field.toFieldXML())
        ]

        let parsed = FieldParser.parse(paragraph: para)
        XCTAssertEqual(parsed.count, 1)
        guard let first = parsed.first, case .sequence(let seq) = first.field else {
            return XCTFail("Expected .sequence")
        }
        XCTAssertEqual(seq.identifier, "Figure")
    }

    func testFieldParserHandlesHyperlinkWrappedSEQ() {
        let field = SequenceField(identifier: "Figure", cachedResult: "1")
        var run = Run(text: "")
        run.rawXML = field.toFieldXML()
        var link = Hyperlink.external(id: "h1", text: "", url: "https://example.com", relationshipId: "rId1")
        link.runs = [run]
        var para = Paragraph()
        para.hyperlinks = [link]

        let parsed = FieldParser.parse(paragraph: para)
        XCTAssertEqual(parsed.count, 1)
        guard let first = parsed.first, case .sequence(let seq) = first.field else {
            return XCTFail("Expected .sequence")
        }
        XCTAssertEqual(seq.identifier, "Figure")
    }

    func testFieldParserHandlesAlternateContentFallbackSEQ() {
        let field = SequenceField(identifier: "Figure", cachedResult: "1")
        var run = Run(text: "")
        run.rawXML = field.toFieldXML()
        var para = Paragraph()
        para.alternateContents = [
            AlternateContent(rawXML: "<mc:AlternateContent/>", fallbackRuns: [run])
        ]

        let parsed = FieldParser.parse(paragraph: para)
        XCTAssertEqual(parsed.count, 1)
        guard let first = parsed.first, case .sequence(let seq) = first.field else {
            return XCTFail("Expected .sequence")
        }
        XCTAssertEqual(seq.identifier, "Figure")
    }
}
