import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// PsychQuant/ooxml-swift#133 — RED tests for the shapes that the default
/// (non-tree-backed) typed model loses when a typed-dirty operation makes the
/// writer re-serialize the whole `word/document.xml`.
///
/// Every test follows the reproduction measured on a corpus of real documents:
/// read the default way → `insertParagraph(_:at: 0)` → `DocxWriter.write(_:to:)`
/// (overlay mode, the save path che-word-mcp uses) → inspect the saved
/// `document.xml`. A read → write with no edit keeps `document.xml`
/// byte-identical, so every loss below needs the typed re-serialization.
///
/// Shapes (one MARK each). Controls are the same fixture minus the trigger
/// and are expected to pass today; they pin the boundary of each shape.
///
/// 1. A run whose subtree holds `m:oMath` below another child (typically an
///    equation inside a text box of a shape) collapses to that bare `m:oMath`:
///    the drawing, every text-box paragraph and the fallback are gone.
/// 2. Whitespace-only `<w:t xml:space="preserve">` in typed runs is recovered
///    from `WhitespaceOverlay` by `<w:t>` sequence index. Two reader paths walk
///    `<w:t>` elements out of step with that index — a paragraph-level
///    `<w:sdt>` (its content is captured without advancing the counter) and a
///    table cell whose paragraphs are parsed before its nested table — so the
///    whitespace that follows comes back empty.
/// 3. A table cell with a nested table followed by a paragraph is written back
///    with the paragraph moved in front of the nested table.
/// 4. Whitespace-only `<w:t xml:space="preserve">` inside a subtree the reader
///    captures as raw XML (`XMLElement.xmlString`) — a run-level
///    `mc:AlternateContent`, a `w:ins` with a non-run child — is serialized
///    from Foundation's DOM, which has already dropped whitespace-only text.
/// 5. Whitespace-only `<w:t>` without `xml:space="preserve"` is never recorded
///    by the overlay (documented limitation) and comes back empty.
final class Issue133LossyShapeTests: XCTestCase {

    // MARK: - Fixture plumbing

    private let rootNamespaces = [
        #"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#,
        #"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#,
        #"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#,
        #"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006""#,
        #"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing""#,
        #"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#,
        #"xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape""#,
        #"xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math""#,
        #"xmlns:v="urn:schemas-microsoft-com:vml""#,
        #"mc:Ignorable="w14""#,
    ].joined(separator: " ")

    private func makeDocx(body: String) throws -> URL {
        let prolog = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        let root = "<w:document \(rootNamespaces)><w:body>\(body)<w:sectPr/></w:body></w:document>"
        let documentXML: String = prolog + root
        let parts: [String: Data] = [
            "[Content_Types].xml": Data(#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.utf8),
            "_rels/.rels": Data(#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.utf8),
            "word/document.xml": Data(documentXML.utf8),
            "word/_rels/document.xml.rels": Data(#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#.utf8),
        ]
        let archive = try Archive(accessMode: .create)
        for name in parts.keys.sorted() {
            let data = parts[name]!
            try archive.addEntry(with: name, type: .file, uncompressedSize: Int64(data.count), compressionMethod: .deflate) { position, size in
                let start = data.startIndex.advanced(by: Int(position))
                return data.subdata(in: start..<start.advanced(by: size))
            }
        }
        let url = FileManager.default.temporaryDirectory.appendingPathComponent("i133-\(UUID().uuidString).docx")
        try XCTUnwrap(archive.data).write(to: url)
        return url
    }

    /// Read the default way, insert one paragraph at 0, save in overlay mode,
    /// and return the saved `word/document.xml` as text.
    private func insertAndSave(body: String, file: StaticString = #filePath, line: UInt = #line) throws -> String {
        let source = try makeDocx(body: body)
        let saved = FileManager.default.temporaryDirectory.appendingPathComponent("i133-out-\(UUID().uuidString).docx")
        defer {
            try? FileManager.default.removeItem(at: source)
            try? FileManager.default.removeItem(at: saved)
        }
        var doc = try DocxReader.read(from: source)
        defer { doc.close() }
        doc.insertParagraph(Paragraph(text: "插入段"), at: 0)
        try DocxWriter.write(doc, to: saved)
        let archive = try Archive(url: saved, accessMode: .read)
        let entry = try XCTUnwrap(archive["word/document.xml"], file: file, line: line)
        var data = Data()
        _ = try archive.extract(entry) { data.append($0) }
        return String(decoding: data, as: UTF8.self)
    }

    /// Text content of every `<w:t>` in document order, read from the raw
    /// bytes (not through `XMLDocument`, which drops whitespace-only text).
    private func wtTexts(_ xml: String) -> [String] {
        let regex = try! NSRegularExpression(pattern: #"<w:t(?:\s[^>]*)?(?:/>|>([^<]*)</w:t>)"#)
        let ns = xml as NSString
        return regex.matches(in: xml, range: NSRange(location: 0, length: ns.length)).map { match in
            let r = match.range(at: 1)
            return r.location == NSNotFound ? "" : ns.substring(with: r)
        }
    }

    private func occurrences(of needle: String, in haystack: String) -> Int {
        haystack.components(separatedBy: needle).count - 1
    }

    // MARK: - Shape 1: run whose subtree holds m:oMath below another child

    /// The measured shape: a group/shape text box (DrawingML choice + VML
    /// fallback) inside a run, with an equation in the text box.
    private func textBoxRun(innerParagraph: String) -> String {
        "<w:r><mc:AlternateContent>"
            + #"<mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="1" name="TextBox 1"/><wp:cNvGraphicFramePr/>"#
            + #"<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>"#
            + "<wps:txbx><w:txbxContent>\(innerParagraph)</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>"
            + #"<mc:Fallback><w:pict><v:shape id="TextBox1" style="width:72pt;height:36pt"><v:textbox><w:txbxContent>"#
            + innerParagraph
            + "</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>"
            + "</mc:AlternateContent></w:r>"
    }

    private let mathInTextBox = #"<w:p><w:r><w:t>欄位甲</w:t></w:r><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:p>"#
    private let plainTextBox = #"<w:p><w:r><w:t>欄位甲</w:t></w:r></w:p>"#

    func testTextBoxRunWithMathKeepsTheTextBox() throws {
        let body = "<w:p><w:r><w:t>前文</w:t></w:r>\(textBoxRun(innerParagraph: mathInTextBox))</w:p>"
        let xml = try insertAndSave(body: body)
        XCTAssertEqual(occurrences(of: "<mc:AlternateContent", in: xml), 1, "the run's mc:AlternateContent must survive")
        XCTAssertEqual(occurrences(of: "<w:txbxContent", in: xml), 2, "both text boxes (choice + fallback) must survive")
        XCTAssertEqual(occurrences(of: "欄位甲", in: xml), 2, "text-box paragraph text must survive in choice and fallback")
        XCTAssertEqual(occurrences(of: "<m:oMath", in: xml), 2, "the equation stays inside each text box, not hoisted into the body paragraph")
    }

    func testPictRunWithMathKeepsTheTextBox() throws {
        // Minimal form of shape 1: no mc:AlternateContent, just VML. The
        // trigger is m:oMath anywhere below the run, not the wrapper.
        let run = #"<w:r><w:pict><v:shape id="TextBox1" style="width:72pt;height:36pt"><v:textbox><w:txbxContent>"#
            + mathInTextBox + "</w:txbxContent></v:textbox></v:shape></w:pict></w:r>"
        let xml = try insertAndSave(body: "<w:p><w:r><w:t>前文</w:t></w:r>\(run)</w:p>")
        XCTAssertEqual(occurrences(of: "<w:pict", in: xml), 1, "the run's w:pict must survive")
        XCTAssertEqual(occurrences(of: "欄位甲", in: xml), 1, "text-box paragraph text must survive")
    }

    func testControl_TextBoxRunWithoutMathSurvives() throws {
        let body = "<w:p><w:r><w:t>前文</w:t></w:r>\(textBoxRun(innerParagraph: plainTextBox))</w:p>"
        let xml = try insertAndSave(body: body)
        XCTAssertEqual(occurrences(of: "<w:txbxContent", in: xml), 2)
        XCTAssertEqual(occurrences(of: "欄位甲", in: xml), 2)
    }

    // MARK: - Shape 2a: whitespace after a paragraph-level w:sdt

    private let inlineCheckbox = #"<w:sdt><w:sdtPr><w:id w:val="101"/></w:sdtPr><w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt>"#
    private let spaceThenText = #"<w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:t>欄位甲</w:t></w:r>"#

    func testWhitespaceAfterInlineContentControlSurvives() throws {
        let xml = try insertAndSave(body: "<w:p>\(inlineCheckbox)\(spaceThenText)</w:p>")
        let texts = wtTexts(xml)
        let i = try XCTUnwrap(texts.firstIndex(of: "欄位甲"))
        XCTAssertEqual(texts[i - 1], " ", "the preserved space after the inline content control must survive")
    }

    func testControl_WhitespaceWithoutInlineContentControlSurvives() throws {
        let xml = try insertAndSave(body: "<w:p>\(spaceThenText)</w:p>")
        let texts = wtTexts(xml)
        let i = try XCTUnwrap(texts.firstIndex(of: "欄位甲"))
        XCTAssertEqual(texts[i - 1], " ")
    }

    // MARK: - Shapes 2b + 3: table cell with a nested table followed by a paragraph

    private func outerTable(cell: String) -> String {
        #"<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>"#
            + cell + "</w:tc></w:tr></w:tbl>"
    }

    private let nestedTable = #"<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">  </w:t></w:r><w:r><w:t>欄位乙</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#

    private var cellWithParagraphAfterNestedTable: String {
        outerTable(cell: #"<w:p><w:r><w:t>欄位甲</w:t></w:r></w:p>"# + nestedTable + #"<w:p><w:r><w:t>欄位丙</w:t></w:r></w:p>"#)
    }

    func testParagraphAfterNestedTableStaysAfterIt() throws {
        let xml = try insertAndSave(body: cellWithParagraphAfterNestedTable + "<w:p/>")
        let nestedEnd = try XCTUnwrap(xml.range(of: "欄位乙"), "nested table text must survive")
        let after = try XCTUnwrap(xml.range(of: "欄位丙"))
        XCTAssertGreaterThan(after.lowerBound, nestedEnd.lowerBound,
                             "the cell paragraph that follows the nested table must still follow it")
    }

    func testWhitespaceInsideNestedTableSurvivesWhenCellHasParagraphAfterIt() throws {
        let xml = try insertAndSave(body: cellWithParagraphAfterNestedTable + "<w:p/>")
        let texts = wtTexts(xml)
        let i = try XCTUnwrap(texts.firstIndex(of: "欄位乙"))
        XCTAssertEqual(texts[i - 1], "  ", "the preserved whitespace inside the nested table must survive")
    }

    func testControl_NestedTableWhitespaceWhenTrailingCellParagraphHasNoText() throws {
        let body = outerTable(cell: #"<w:p><w:r><w:t>欄位甲</w:t></w:r></w:p>"# + nestedTable + "<w:p/>") + "<w:p/>"
        let xml = try insertAndSave(body: body)
        let texts = wtTexts(xml)
        let i = try XCTUnwrap(texts.firstIndex(of: "欄位乙"))
        XCTAssertEqual(texts[i - 1], "  ")
    }

    // MARK: - Shape 4: whitespace inside a raw-captured subtree

    func testWhitespaceInsideRunLevelAlternateContentSurvives() throws {
        let inner = #"<w:p><w:r><w:t>欄位甲</w:t></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:t>欄位乙</w:t></w:r></w:p>"#
        let xml = try insertAndSave(body: "<w:p>\(textBoxRun(innerParagraph: inner))</w:p>")
        let texts = wtTexts(xml)
        let spaces = texts.indices.filter { texts[$0] == "欄位乙" }.map { texts[$0 - 1] }
        XCTAssertEqual(spaces, [" ", " "], "the preserved space in both text boxes must survive")
    }

    func testWhitespaceInsideInsertionWithNonRunChildSurvives() throws {
        let ins = #"<w:ins w:id="7" w:author="Tester" w:date="2026-01-01T00:00:00Z"><w:bookmarkStart w:id="3" w:name="mark"/><w:r><w:t>欄位甲</w:t></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:t>欄位乙</w:t></w:r><w:bookmarkEnd w:id="3"/></w:ins>"#
        let xml = try insertAndSave(body: "<w:p>\(ins)</w:p>")
        let texts = wtTexts(xml)
        let i = try XCTUnwrap(texts.firstIndex(of: "欄位乙"))
        XCTAssertEqual(texts[i - 1], " ", "the preserved space inside the tracked insertion must survive")
    }

    func testControl_WhitespaceInsideInsertionOfRunsOnlySurvives() throws {
        let ins = #"<w:ins w:id="7" w:author="Tester" w:date="2026-01-01T00:00:00Z"><w:r><w:t>欄位甲</w:t></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:t>欄位乙</w:t></w:r></w:ins>"#
        let xml = try insertAndSave(body: "<w:p>\(ins)</w:p>")
        let texts = wtTexts(xml)
        let i = try XCTUnwrap(texts.firstIndex(of: "欄位乙"))
        XCTAssertEqual(texts[i - 1], " ")
    }

    // MARK: - Shape 5: whitespace-only w:t without xml:space="preserve"

    /// Whether Word displays this whitespace is not established here (ECMA-376
    /// treats unpreserved leading/trailing whitespace as insignificant); what
    /// is established is that the saved bytes differ from the untouched source
    /// only because of the typed re-serialization.
    func testUnpreservedWhitespaceOnlyTextSurvives() throws {
        let xml = try insertAndSave(body: #"<w:p><w:r><w:t>欄位甲</w:t></w:r><w:r><w:t> </w:t></w:r><w:r><w:t>欄位乙</w:t></w:r></w:p>"#)
        let texts = wtTexts(xml)
        let i = try XCTUnwrap(texts.firstIndex(of: "欄位乙"))
        XCTAssertEqual(texts[i - 1], " ", "the untouched whitespace-only run text must survive")
    }
}
