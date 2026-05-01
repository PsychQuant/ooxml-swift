import XCTest
@testable import OOXMLSwift

/// Tests for `DocumentWalker` (Issue #56 R5 stack-completion §1.3).
///
/// `DocumentWalker` is the unified abstraction that replaces ad-hoc walkers
/// across the codebase. Every part-spanning operation (calibration, revision
/// accept/reject, cross-part hyperlink ops) routes through it so adding a new
/// part is a single-site change instead of a 5-site grep.
final class DocumentWalkerTests: XCTestCase {

    // MARK: - walkAllParagraphs visits all parts

    func testWalkVisitsBodyParagraphs() {
        var doc = WordDocument()
        var p = Paragraph()
        p.runs = [Run(text: "body-para")]
        doc.body.children = [.paragraph(p)]

        var visited: [(text: String, partKey: String)] = []
        DocumentWalker.walkAllParagraphs(in: doc) { para, partKey in
            visited.append((para.runs.map { $0.text }.joined(), partKey))
        }

        XCTAssertEqual(visited.count, 1)
        XCTAssertEqual(visited[0].text, "body-para")
        XCTAssertEqual(visited[0].partKey, "word/document.xml")
    }

    func testWalkRecursesIntoBodyTableCellParagraphs() {
        var doc = WordDocument()
        var inner = Paragraph()
        inner.runs = [Run(text: "cell-para")]
        let cell = TableCell(paragraphs: [inner])
        let row = TableRow(cells: [cell])
        let table = Table(rows: [row])
        doc.body.children = [.table(table)]

        var visited: [(text: String, partKey: String)] = []
        DocumentWalker.walkAllParagraphs(in: doc) { para, partKey in
            visited.append((para.runs.map { $0.text }.joined(), partKey))
        }

        XCTAssertEqual(visited.count, 1)
        XCTAssertEqual(visited[0].text, "cell-para")
        XCTAssertEqual(visited[0].partKey, "word/document.xml")
    }

    func testWalkRecursesIntoNestedTables() {
        var doc = WordDocument()
        var nestedPara = Paragraph()
        nestedPara.runs = [Run(text: "nested-cell-para")]
        var outerCell = TableCell(paragraphs: [Paragraph()])
        outerCell.nestedTables = [Table(rows: [TableRow(cells: [TableCell(paragraphs: [nestedPara])])])]
        let outerTable = Table(rows: [TableRow(cells: [outerCell])])
        doc.body.children = [.table(outerTable)]

        var visited: [String] = []
        DocumentWalker.walkAllParagraphs(in: doc) { para, _ in
            visited.append(para.runs.map { $0.text }.joined())
        }

        XCTAssertTrue(visited.contains("nested-cell-para"), "Nested-table paragraph not visited; got: \(visited)")
    }

    func testWalkRecursesIntoBlockLevelContentControlChildren() {
        var doc = WordDocument()
        var inner = Paragraph()
        inner.runs = [Run(text: "sdt-inner")]
        let sdtControl = ContentControl.richText(tag: "T", alias: "A", content: "")
        let sdtChildren: [BodyChild] = [.paragraph(inner)]
        doc.body.children = [.contentControl(sdtControl, children: sdtChildren)]

        var visited: [String] = []
        DocumentWalker.walkAllParagraphs(in: doc) { para, _ in
            visited.append(para.runs.map { $0.text }.joined())
        }

        XCTAssertTrue(visited.contains("sdt-inner"), "SDT child paragraph not visited; got: \(visited)")
    }

    func testBodyChildWalkerHonorsExplicitRecursionPolicy() {
        struct TextVisitor: BodyChildVisitor {
            var initialState: [String] = []
            var recursesIntoTableCells: Bool = false
            var recursesIntoNestedTables: Bool = false
            var recursesIntoContentControls: Bool = true

            mutating func visitParagraph(_ paragraph: inout Paragraph, state: inout [String]) {
                state.append(paragraph.runs.map { $0.text }.joined())
            }

            mutating func visitSkippedTable(_ table: inout Table, state: inout [String]) {
                state.append("skipped-table")
            }
        }

        var top = Paragraph()
        top.runs = [Run(text: "top")]
        var tablePara = Paragraph()
        tablePara.runs = [Run(text: "table")]
        var sdtPara = Paragraph()
        sdtPara.runs = [Run(text: "sdt")]

        let table = Table(rows: [TableRow(cells: [TableCell(paragraphs: [tablePara])])])
        let contentControl = ContentControl.richText(tag: "T", alias: "A", content: "")
        let children: [BodyChild] = [
            .paragraph(top),
            .table(table),
            .contentControl(contentControl, children: [.paragraph(sdtPara)])
        ]

        var visitor = TextVisitor()
        let visited = BodyChildWalker.walk(children, visitor: &visitor)

        XCTAssertEqual(visited, ["top", "skipped-table", "sdt"])
    }

    func testWalkVisitsHeaderParagraphsWithHeaderPartKey() {
        var doc = WordDocument()
        var h = Header(id: "rId10", paragraphs: [], type: .default, originalFileName: "header3.xml")
        var p = Paragraph()
        p.runs = [Run(text: "header-text")]
        h.paragraphs = [p]
        doc.headers = [h]

        var visited: [(text: String, partKey: String)] = []
        DocumentWalker.walkAllParagraphs(in: doc) { para, partKey in
            visited.append((para.runs.map { $0.text }.joined(), partKey))
        }

        XCTAssertEqual(visited.count, 1)
        XCTAssertEqual(visited[0].text, "header-text")
        XCTAssertEqual(visited[0].partKey, "word/header3.xml")
    }

    func testWalkVisitsFooterParagraphsWithFooterPartKey() {
        var doc = WordDocument()
        var f = Footer(id: "rId11", paragraphs: [], type: .default, originalFileName: "footer2.xml")
        var p = Paragraph()
        p.runs = [Run(text: "footer-text")]
        f.paragraphs = [p]
        doc.footers = [f]

        var visited: [(text: String, partKey: String)] = []
        DocumentWalker.walkAllParagraphs(in: doc) { para, partKey in
            visited.append((para.runs.map { $0.text }.joined(), partKey))
        }

        XCTAssertEqual(visited.count, 1)
        XCTAssertEqual(visited[0].partKey, "word/footer2.xml")
    }

    func testWalkVisitsFootnoteParagraphsWithFootnotesPartKey() {
        var doc = WordDocument()
        var fn = Footnote(id: 1, text: "", paragraphIndex: 0)
        var p = Paragraph()
        p.runs = [Run(text: "fn-text")]
        fn.paragraphs = [p]
        doc.footnotes.footnotes = [fn]

        var visited: [(text: String, partKey: String)] = []
        DocumentWalker.walkAllParagraphs(in: doc) { para, partKey in
            visited.append((para.runs.map { $0.text }.joined(), partKey))
        }

        XCTAssertEqual(visited.count, 1)
        XCTAssertEqual(visited[0].partKey, "word/footnotes.xml")
    }

    func testWalkVisitsEndnoteParagraphsWithEndnotesPartKey() {
        var doc = WordDocument()
        var en = Endnote(id: 1, text: "", paragraphIndex: 0)
        var p = Paragraph()
        p.runs = [Run(text: "en-text")]
        en.paragraphs = [p]
        doc.endnotes.endnotes = [en]

        var visited: [(text: String, partKey: String)] = []
        DocumentWalker.walkAllParagraphs(in: doc) { para, partKey in
            visited.append((para.runs.map { $0.text }.joined(), partKey))
        }

        XCTAssertEqual(visited.count, 1)
        XCTAssertEqual(visited[0].partKey, "word/endnotes.xml")
    }

    // MARK: - findUnrecognizedChild locates wrappers across parts

    func testFindUnrecognizedChildInBodyReturnsBodyPartKey() {
        var doc = WordDocument()
        var p = Paragraph()
        p.unrecognizedChildren.append(UnrecognizedChild(name: "w:ins", rawXML: "<w:ins w:id=\"5\" w:author=\"Bob\"><w:r><w:t>x</w:t></w:r></w:ins>", position: 1))
        doc.body.children = [.paragraph(p)]

        let hit = DocumentWalker.findUnrecognizedChild(in: doc, name: "w:ins", idMarker: "w:id=\"5\"")

        XCTAssertNotNil(hit)
        XCTAssertEqual(hit?.partKey, "word/document.xml")
        XCTAssertEqual(hit?.indexInParagraph, 0)
    }

    func testFindUnrecognizedChildInHeaderReturnsHeaderPartKey() {
        var doc = WordDocument()
        var p = Paragraph()
        p.unrecognizedChildren.append(UnrecognizedChild(name: "w:ins", rawXML: "<w:ins w:id=\"9\" w:author=\"Alice\"><w:r><w:t>head</w:t></w:r></w:ins>", position: 1))
        let h = Header(id: "rId10", paragraphs: [p], type: .default, originalFileName: "header1.xml")
        doc.headers = [h]

        let hit = DocumentWalker.findUnrecognizedChild(in: doc, name: "w:ins", idMarker: "w:id=\"9\"")

        XCTAssertNotNil(hit)
        XCTAssertEqual(hit?.partKey, "word/header1.xml")
    }

    func testFindUnrecognizedChildReturnsNilWhenNoMatch() {
        var doc = WordDocument()
        var p = Paragraph()
        p.unrecognizedChildren.append(UnrecognizedChild(name: "w:ins", rawXML: "<w:ins w:id=\"5\"><w:r><w:t>x</w:t></w:r></w:ins>", position: 1))
        doc.body.children = [.paragraph(p)]

        let hit = DocumentWalker.findUnrecognizedChild(in: doc, name: "w:ins", idMarker: "w:id=\"99\"")

        XCTAssertNil(hit)
    }

    func testFindUnrecognizedChildOpeningTagOnlyMatch() {
        // Codex P1 from R3-NEW-4: substring match "w:id=\"5\"" must not false-hit
        // a nested element whose closing portion contains the same substring.
        var doc = WordDocument()
        var p = Paragraph()
        p.unrecognizedChildren.append(UnrecognizedChild(name: "w:ins", rawXML: "<w:ins w:id=\"5\"><w:bookmarkStart w:id=\"50\"/></w:ins>", position: 1))
        doc.body.children = [.paragraph(p)]

        let hitFive = DocumentWalker.findUnrecognizedChild(in: doc, name: "w:ins", idMarker: "w:id=\"5\"")
        XCTAssertNotNil(hitFive)

        let hitFifty = DocumentWalker.findUnrecognizedChild(in: doc, name: "w:ins", idMarker: "w:id=\"50\"")
        XCTAssertNil(hitFifty, "Nested w:bookmarkStart w:id=\"50\" must not match the outer w:ins lookup")
    }
}
