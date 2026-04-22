import XCTest
@testable import OOXMLSwift

final class UpdateAllFieldsTests: XCTestCase {

    private func captionParagraph(identifier: String, resetLevel: Int? = nil, initialCached: String = "1") -> Paragraph {
        let field = SequenceField(identifier: identifier, resetLevel: resetLevel, cachedResult: initialCached)
        var run = Run(text: "")
        run.rawXML = field.toFieldXML()
        var para = Paragraph()
        para.runs = [Run(text: "\(identifier) "), run]
        para.properties.style = "Caption"
        return para
    }

    private func headingParagraph(level: Int, text: String) -> Paragraph {
        var para = Paragraph(text: text)
        para.properties.style = "Heading \(level)"
        return para
    }

    // MARK: Single SEQ identifier

    func testSingleIdentifierIncrementsSequentially() {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(captionParagraph(identifier: "Figure", initialCached: "1")),
            .paragraph(captionParagraph(identifier: "Figure", initialCached: "1")),
            .paragraph(captionParagraph(identifier: "Figure", initialCached: "1")),
        ]

        let result = doc.updateAllFields()
        XCTAssertEqual(result, ["Figure": 3])

        // Verify rewritten cached values: 1, 2, 3 in order
        for (i, child) in doc.body.children.enumerated() {
            guard case .paragraph(let p) = child, let xml = p.runs.last?.rawXML else {
                XCTFail("Expected paragraph at \(i)"); continue
            }
            XCTAssertTrue(xml.contains("<w:t>\(i + 1)</w:t>"), "Expected cached result \(i + 1), got: \(xml)")
        }
    }

    // MARK: Distinct identifiers independent

    func testDistinctIdentifiersHaveIndependentCounters() {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(captionParagraph(identifier: "Figure")),  // F1
            .paragraph(captionParagraph(identifier: "Table")),   // T1
            .paragraph(captionParagraph(identifier: "Figure")),  // F2
            .paragraph(captionParagraph(identifier: "Figure")),  // F3
            .paragraph(captionParagraph(identifier: "Table")),   // T2
        ]
        let result = doc.updateAllFields()
        XCTAssertEqual(result, ["Figure": 3, "Table": 2])
    }

    // MARK: Chapter-reset captions

    func testChapterResetRestartsCountersAtHeading1() {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(headingParagraph(level: 1, text: "Chapter 1")),
            .paragraph(captionParagraph(identifier: "Figure", resetLevel: 1)),
            .paragraph(captionParagraph(identifier: "Figure", resetLevel: 1)),
            .paragraph(headingParagraph(level: 1, text: "Chapter 2")),
            .paragraph(captionParagraph(identifier: "Figure", resetLevel: 1)),
            .paragraph(captionParagraph(identifier: "Figure", resetLevel: 1)),
            .paragraph(captionParagraph(identifier: "Figure", resetLevel: 1)),
        ]
        let result = doc.updateAllFields()
        // Final counter reflects last chapter's final value
        XCTAssertEqual(result, ["Figure": 3])

        // Collect cached values from all 5 Figure paragraphs (not heading paragraphs)
        var cachedValues: [String] = []
        for child in doc.body.children {
            guard case .paragraph(let p) = child,
                  p.properties.style == "Caption",
                  let xml = p.runs.last?.rawXML,
                  let match = xml.range(of: #"<w:t>(\d+)</w:t>"#, options: .regularExpression) else {
                continue
            }
            let valStr = String(xml[match])
                .replacingOccurrences(of: "<w:t>", with: "")
                .replacingOccurrences(of: "</w:t>", with: "")
            cachedValues.append(valStr)
        }
        // Expected: chapter 1 (1, 2), chapter 2 (1, 2, 3)
        XCTAssertEqual(cachedValues, ["1", "2", "1", "2", "3"])
    }

    // MARK: Non-SEQ fields preserved

    func testNonSEQFieldsUnmutated() {
        var doc = WordDocument()
        // Insert a REF field (non-SEQ) — updateAllFields should leave it unchanged
        let ref = ReferenceField(type: .ref, bookmarkName: "fig1", cachedResult: "original")
        var refRun = Run(text: "")
        refRun.rawXML = ref.toFieldXML()
        var refPara = Paragraph()
        refPara.runs = [refRun]
        doc.body.children = [
            .paragraph(refPara),
            .paragraph(captionParagraph(identifier: "Figure")),
        ]
        let result = doc.updateAllFields()
        XCTAssertEqual(result, ["Figure": 1])

        // REF field's cached should still contain "original"
        if case .paragraph(let p) = doc.body.children[0],
           let xml = p.runs.first?.rawXML {
            XCTAssertTrue(xml.contains("<w:t>original</w:t>"), "REF cached should be preserved, got: \(xml)")
        }
    }

    // MARK: Empty document

    func testEmptyDocumentReturnsEmptyDict() {
        var doc = WordDocument()
        XCTAssertEqual(doc.updateAllFields(), [:])
    }
}
