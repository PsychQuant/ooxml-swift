import XCTest
@testable import OOXMLSwift

final class DocumentReplaceTextTests: XCTestCase {

    // MARK: - Default scope (bodyAndTables)

    func testDefaultScopeBodyOnly() throws {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "Body Draft"))
        ]
        doc.headers = [Header(id: "rId-h1", paragraphs: [Paragraph(text: "Header Draft")])]
        doc.footers = [Footer(id: "rId-f1", paragraphs: [Paragraph(text: "Footer Draft")])]

        let count = try doc.replaceText(find: "Draft", with: "Final")
        XCTAssertEqual(count, 1, "Default scope should only hit body — found \(count)")

        // Verify
        if case .paragraph(let p) = doc.body.children[0] {
            XCTAssertEqual(p.runs.first?.text, "Body Final")
        } else {
            XCTFail("Expected body paragraph")
        }
        XCTAssertEqual(doc.headers[0].paragraphs[0].runs.first?.text, "Header Draft", "Header should remain unchanged")
        XCTAssertEqual(doc.footers[0].paragraphs[0].runs.first?.text, "Footer Draft", "Footer should remain unchanged")
    }

    // MARK: - Scope .all

    func testScopeAllCoversHeadersAndFooters() throws {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "Body Draft"))
        ]
        doc.headers = [Header(id: "rId-h1", paragraphs: [Paragraph(text: "Header Draft")])]
        doc.footers = [Footer(id: "rId-f1", paragraphs: [Paragraph(text: "Footer Draft")])]

        let opts = ReplaceOptions(scope: .all)
        let count = try doc.replaceText(find: "Draft", with: "Final", options: opts)
        XCTAssertEqual(count, 3, "Expected body + header + footer = 3 replacements")

        XCTAssertEqual(doc.headers[0].paragraphs[0].runs.first?.text, "Header Final")
        XCTAssertEqual(doc.footers[0].paragraphs[0].runs.first?.text, "Footer Final")
    }

    func testScopeAllCoversFootnotes() throws {
        var doc = WordDocument()
        doc.body.children = [.paragraph(Paragraph(text: "Main text"))]

        var footnote = Footnote(id: 1, text: "see ref 1", paragraphIndex: 0)
        footnote.paragraphs = [Paragraph(text: "see ref 1")]
        doc.footnotes.footnotes.append(footnote)

        let opts = ReplaceOptions(scope: .all)
        let count = try doc.replaceText(find: "ref 1", with: "reference 1", options: opts)
        XCTAssertGreaterThanOrEqual(count, 1, "Footnote replacement should be counted")

        XCTAssertEqual(doc.footnotes.footnotes[0].paragraphs[0].runs.first?.text, "see reference 1")
    }

    // MARK: - Cross-run via Document API

    func testDocumentReplaceCrossRun() throws {
        var doc = WordDocument()
        var paragraph = Paragraph()
        paragraph.runs = [
            Run(text: "均值方程式："),
            Run(text: ""),
            Run(text: "r_t = ...")
        ]
        doc.body.children = [.paragraph(paragraph)]

        let count = try doc.replaceText(find: "均值方程式：r_t", with: "Mean: r_t")
        XCTAssertEqual(count, 1)

        if case .paragraph(let p) = doc.body.children[0] {
            let flat = p.runs.map { $0.text }.joined()
            XCTAssertEqual(flat, "Mean: r_t = ...")
        } else {
            XCTFail("Expected body paragraph")
        }
    }
}
