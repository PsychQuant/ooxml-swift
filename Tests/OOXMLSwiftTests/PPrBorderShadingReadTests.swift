import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Regression tests for PsychQuant/ooxml-swift#176.
///
/// 修正前：`DocxReader.parseParagraphProperties` 從來沒把 `<w:pBdr>` /
/// `<w:shd>` 讀進 `ParagraphProperties.border` / `.shading`，#168 的 raw 捕捉
/// 又刻意排除這兩個元素（避免 setter 造成新舊兩份），結果是：任何 typed 編輯
/// 讓 `word/document.xml` 從 typed 模型重新產生時，**每一個**段落既有的框線與
/// 網底都消失。
///
/// typed 模型（`ParagraphBorderStyle` 只有 type/color/size/space，
/// `ParagraphBorderType` 只有 10 種 val，沒有 `bar` 邊；`CellShading` 只有
/// fill/color/pattern，`ShadingPattern` 只有 5 種 val）表達不了來源的全部屬性，
/// 所以修法是：typed 欄位存「讀取當下的投影」，同時保留來源元素原文；輸出時
/// typed 值仍等於投影（沒被改過）就原樣輸出原文，被 setter 或直接指派改過才用
/// typed 輸出。這組測試把「沒改就一個屬性都不少」與「改了只有一份」都釘住。
final class PPrBorderShadingReadTests: XCTestCase {

    static let wNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    /// 每個邊都帶著 typed 模型表達不了的東西：theme 色、themeTint/Shade、
    /// shadow、frame、`thinThickSmallGap` 這種 enum 外的 val、`bar` 邊；
    /// shd 帶 enum 外的 `pct10` 與全部 theme 屬性。
    static let sourcePBdr = "<w:pBdr>"
        + "<w:top w:val=\"single\" w:sz=\"4\" w:space=\"1\" w:color=\"auto\" w:themeColor=\"accent1\" w:themeTint=\"99\" w:shadow=\"1\"/>"
        + "<w:left w:val=\"thinThickSmallGap\" w:sz=\"12\" w:space=\"4\" w:color=\"FF0000\" w:frame=\"1\"/>"
        + "<w:bottom w:val=\"double\" w:sz=\"6\" w:space=\"1\" w:color=\"1F3864\" w:themeColor=\"accent1\" w:themeShade=\"80\"/>"
        + "<w:right w:val=\"single\" w:sz=\"4\" w:space=\"4\" w:color=\"auto\"/>"
        + "<w:between w:val=\"single\" w:sz=\"4\" w:space=\"1\" w:color=\"auto\"/>"
        + "<w:bar w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>"
        + "</w:pBdr>"
    static let sourceShd = "<w:shd w:val=\"pct10\" w:color=\"auto\" w:themeColor=\"text1\" w:themeTint=\"80\""
        + " w:themeShade=\"BF\" w:fill=\"DEEAF6\" w:themeFill=\"accent1\" w:themeFillTint=\"33\" w:themeFillShade=\"F2\"/>"

    /// pPr 同時有 typed（pStyle、numPr、jc）、raw（kinsoku）與 pBdr/shd，
    /// 而且整段是 schema 順序——#175 之後重新序列化應逐位元組等於來源。
    static let borderedPPr = "<w:pPr><w:pStyle w:val=\"Quote\"/>"
        + "<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>"
        + sourcePBdr + sourceShd
        + "<w:kinsoku w:val=\"0\"/><w:jc w:val=\"center\"/></w:pPr>"

    static let borderedParagraph = "<w:p>\(borderedPPr)<w:r><w:t>BORDERED</w:t></w:r></w:p>"

    // MARK: - Helpers

    private func parseParagraph(_ xml: String) throws -> Paragraph {
        let element = try XMLElement(xmlString: xml.replacingOccurrences(
            of: "<w:p>", with: "<w:p xmlns:w=\"\(Self.wNS)\">"))
        return try DocxReader.parseParagraph(
            from: element,
            relationships: RelationshipsCollection(),
            styles: [],
            numbering: Numbering()
        )
    }

    private func buildDocx(body: String) throws -> URL {
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue176-\(UUID().uuidString).docx")
        let archive = try Archive(url: url, accessMode: .create)
        let parts: [(String, String)] = [
            ("[Content_Types].xml", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\
                <Default Extension="xml" ContentType="application/xml"/>\
                <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\
                </Types>
                """),
            ("_rels/.rels", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>\
                </Relationships>
                """),
            ("word/document.xml", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <w:document xmlns:w="\(Self.wNS)"><w:body>\(body)\
                <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>
                """),
        ]
        for (path, text) in parts {
            let data = Data(text.utf8)
            try archive.addEntry(with: path, type: .file, uncompressedSize: Int64(data.count),
                                 compressionMethod: .deflate) { position, size in
                data.subdata(in: Int(position)..<Int(position) + size)
            }
        }
        return url
    }

    private func saveAndExtractDocumentXML(_ doc: WordDocument) throws -> XMLDocument {
        let out = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue176-out-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: out)
        defer { try? FileManager.default.removeItem(at: out) }
        let archive = try Archive(url: out, accessMode: .read)
        let entry = try XCTUnwrap(archive["word/document.xml"])
        var data = Data()
        _ = try archive.extract(entry) { data.append($0) }
        return try XMLDocument(data: data)
    }

    private func pPr(in document: XMLDocument, paragraphText marker: String) throws -> XMLElement {
        let nodes = try document.nodes(forXPath:
            "//*[local-name()='p'][.//*[local-name()='t' and text()='\(marker)']]/*[local-name()='pPr']")
        return try XCTUnwrap(nodes.first as? XMLElement, "no <w:pPr> on the paragraph containing \(marker)")
    }

    private func attributes(_ element: XMLElement) -> [String: String] {
        Dictionary(uniqueKeysWithValues: (element.attributes ?? []).map { ($0.name ?? "", $0.stringValue ?? "") })
    }

    /// 子元素名稱 → 屬性字典，保留順序（用來逐屬性比對 pBdr 的每一個邊）。
    private func childAttributes(_ element: XMLElement) -> [(String, [String: String])] {
        (element.children ?? []).compactMap { $0 as? XMLElement }.map { ($0.name ?? "", attributes($0)) }
    }

    private func sourceElement(_ xml: String) throws -> XMLElement {
        let wrapped = try XMLElement(xmlString: "<w:x xmlns:w=\"\(Self.wNS)\">\(xml)</w:x>")
        return try XCTUnwrap(wrapped.children?.first as? XMLElement)
    }

    private func assertBorderAndShadingIntact(_ pPr: XMLElement,
                                              file: StaticString = #filePath, line: UInt = #line) throws {
        let pBdrs = pPr.elements(forName: "w:pBdr")
        let shds = pPr.elements(forName: "w:shd")
        XCTAssertEqual(pBdrs.count, 1, "exactly one <w:pBdr>", file: file, line: line)
        XCTAssertEqual(shds.count, 1, "exactly one <w:shd>", file: file, line: line)
        guard let pBdr = pBdrs.first, let shd = shds.first else { return }

        let expectedSides = childAttributes(try sourceElement(Self.sourcePBdr))
        let actualSides = childAttributes(pBdr)
        XCTAssertEqual(actualSides.map(\.0), expectedSides.map(\.0),
                       "every side (incl. bar) survives, in CT_PBdr order", file: file, line: line)
        for ((name, expected), (_, actual)) in zip(expectedSides, actualSides) {
            XCTAssertEqual(actual, expected, "\(name): every attribute survives verbatim", file: file, line: line)
        }
        XCTAssertEqual(attributes(shd), attributes(try sourceElement(Self.sourceShd)),
                       "every <w:shd> attribute (theme*, pct10) survives verbatim", file: file, line: line)
    }

    // MARK: - Section A: reading into typed fields

    func testParseReadsPBdrAndShdIntoTypedFields() throws {
        let props = try parseParagraph(Self.borderedParagraph).properties
        let border = try XCTUnwrap(props.border, "<w:pBdr> must be read into `border`")
        XCTAssertEqual(border.top, ParagraphBorderStyle(type: .single, color: "auto", size: 4, space: 1))
        XCTAssertEqual(border.bottom, ParagraphBorderStyle(type: .double, color: "1F3864", size: 6, space: 1))
        XCTAssertEqual(border.right, ParagraphBorderStyle(type: .single, color: "auto", size: 4, space: 4))
        XCTAssertEqual(border.between, ParagraphBorderStyle(type: .single, color: "auto", size: 4, space: 1))
        XCTAssertNil(border.left, "`thinThickSmallGap` has no ParagraphBorderType case — not projected")

        let shading = try XCTUnwrap(props.shading, "<w:shd> must be read into `shading`")
        XCTAssertEqual(shading.fill, "DEEAF6")
        XCTAssertEqual(shading.color, "auto")
        XCTAssertNil(shading.pattern, "`pct10` has no ShadingPattern case — not projected")

        let rawNames = props.rawChildren.map(\.name)
        XCTAssertFalse(rawNames.contains("pBdr"), "pBdr is typed, never raw-captured too")
        XCTAssertFalse(rawNames.contains("shd"), "shd is typed, never raw-captured too")
        XCTAssertTrue(DocxReader.recognizedPPrChildNames.isSuperset(of: ["pBdr", "shd"]))
    }

    /// schema 預設值：`color` 缺席投影成 "auto"、`space` 缺席投影成 0；
    /// `sz` 在 schema 沒有預設值，缺席時投影成 0。shd 沒有 fill 時投影成 "auto"，
    /// 讓「來源有 shd」一定對應到非 nil 的 typed 值（`shading = nil` 才能表達刪除）。
    func testMissingAttributesProjectToSchemaDefaults() throws {
        let props = try parseParagraph("<w:p><w:pPr><w:pBdr><w:top w:val=\"dotted\"/></w:pBdr>"
            + "<w:shd w:val=\"clear\"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>").properties
        XCTAssertEqual(props.border?.top, ParagraphBorderStyle(type: .dotted, color: "auto", size: 0, space: 0))
        XCTAssertEqual(props.shading, CellShading(fill: "auto", color: nil, pattern: .clear))
    }

    /// 只有 enum 外的邊（或空的 `<w:pBdr/>`）時，typed 值仍是非 nil 的空
    /// `ParagraphBorder()`——「有這個元素」不能被讀成「沒有」。
    func testPBdrWithOnlyUnrepresentableSidesStillProjectsNonNil() throws {
        let props = try parseParagraph("<w:p><w:pPr><w:pBdr><w:bar w:val=\"single\" w:sz=\"4\"/>"
            + "</w:pBdr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>").properties
        XCTAssertEqual(props.border, ParagraphBorder())
        XCTAssertTrue(props.toXML().contains("<w:bar w:val=\"single\" w:sz=\"4\"/>"),
                      "untouched → the source element is emitted verbatim: \(props.toXML())")
    }

    // MARK: - Section B: untouched → verbatim; edited → one typed copy

    /// 沒被改過的 pBdr/shd 原樣輸出——連同 #175 的 schema 順序，整段 pPr 逐位元組等於來源。
    func testUntouchedParagraphReemitsPPrByteForByte() throws {
        let paragraph = try parseParagraph(Self.borderedParagraph)
        XCTAssertTrue(paragraph.toXML().contains(Self.borderedPPr),
                      "expected verbatim pPr, got: \(paragraph.toXML())")
    }

    /// Issue 的 Expected：含框線與網底的文件做一次**無關**的 typed 編輯後，兩者仍在
    /// ——而且是逐屬性（theme 色、shadow、frame、bar、pct10）。
    func testUnrelatedTypedEditKeepsEveryBorderAndShadingAttribute() throws {
        let url = try buildDocx(body: "<w:p><w:r><w:t>Intro</w:t></w:r></w:p>" + Self.borderedParagraph)
        defer { try? FileManager.default.removeItem(at: url) }
        var doc = try DocxReader.read(from: url)
        defer { doc.close() }

        doc.insertParagraph(Paragraph(text: "INSERTED"), at: 0)

        let document = try saveAndExtractDocumentXML(doc)
        try assertBorderAndShadingIntact(try pPr(in: document, paragraphText: "BORDERED"))
    }

    /// 同一段落改別的 typed 欄位（對齊）：框線與網底不受影響。
    func testTypedEditOfAnotherFieldOnTheSameParagraphKeepsBorderAndShading() throws {
        let url = try buildDocx(body: Self.borderedParagraph)
        defer { try? FileManager.default.removeItem(at: url) }
        var doc = try DocxReader.read(from: url)
        defer { doc.close() }

        var change = ParagraphProperties()
        change.alignment = .right
        try doc.setParagraphFormat(at: 0, properties: change)

        let pPrElement = try pPr(in: try saveAndExtractDocumentXML(doc), paragraphText: "BORDERED")
        try assertBorderAndShadingIntact(pPrElement)
        XCTAssertEqual(pPrElement.elements(forName: "w:jc").first?.attribute(forName: "w:val")?.stringValue, "right")
    }

    /// 呼叫 setter 之後只輸出一份，而且是新的 typed 值（不是來源舊值、也不是兩份）；
    /// 同時仍落在 #175 的 schema 位置（numPr 之後、kinsoku 之前）。
    func testSettersEmitExactlyOneTypedCopyAtTheSchemaSlot() throws {
        let url = try buildDocx(body: Self.borderedParagraph)
        defer { try? FileManager.default.removeItem(at: url) }
        var doc = try DocxReader.read(from: url)
        defer { doc.close() }

        let newSide = ParagraphBorderStyle(type: .dashed, color: "00B050", size: 8, space: 2)
        try doc.setParagraphBorder(at: 0, border: ParagraphBorder(top: newSide))
        try doc.setParagraphShading(at: 0, fill: "00FF00", pattern: .solid)

        let pPrElement = try pPr(in: try saveAndExtractDocumentXML(doc), paragraphText: "BORDERED")
        let pBdrs = pPrElement.elements(forName: "w:pBdr")
        let shds = pPrElement.elements(forName: "w:shd")
        XCTAssertEqual(pBdrs.count, 1)
        XCTAssertEqual(shds.count, 1)
        let sides = childAttributes(try XCTUnwrap(pBdrs.first))
        XCTAssertEqual(sides.map(\.0), ["w:top"], "setter replaces the whole border")
        XCTAssertEqual(sides.first?.1, ["w:val": "dashed", "w:sz": "8", "w:space": "2", "w:color": "00B050"])
        XCTAssertEqual(attributes(try XCTUnwrap(shds.first)), ["w:val": "solid", "w:fill": "00FF00"])

        let names = (pPrElement.children ?? []).compactMap { ($0 as? XMLElement)?.name }
        XCTAssertEqual(names, ["w:pStyle", "w:numPr", "w:pBdr", "w:shd", "w:kinsoku", "w:jc"])
    }

    /// 已知限制，刻意釘住：讀出 typed 值、改其中一個欄位再寫回（read-modify-write），
    /// 就走 typed 輸出——typed 模型表達不了的屬性（theme 色、bar 邊…）在這條路上
    /// 不保留。只有一份 pBdr，被改的欄位生效。
    func testReadModifyWriteOfTheProjectionFallsBackToTypedOutput() throws {
        var props = try parseParagraph(Self.borderedParagraph).properties
        props.border?.top?.color = "00FF00"
        let wrapper = try XMLElement(xmlString: "<w:pPr xmlns:w=\"\(Self.wNS)\">\(props.toXML())</w:pPr>")
        let pBdrs = wrapper.elements(forName: "w:pBdr")
        XCTAssertEqual(pBdrs.count, 1)
        let sides = childAttributes(try XCTUnwrap(pBdrs.first))
        XCTAssertEqual(sides.map(\.0), ["w:top", "w:bottom", "w:right", "w:between"])
        XCTAssertEqual(sides.first?.1["w:color"], "00FF00")
        XCTAssertNil(sides.first?.1["w:themeColor"], "typed output cannot carry theme attributes")
        XCTAssertTrue(props.toXML().contains(Self.sourceShd), "shading untouched → still verbatim")
    }

    /// 刪除：typed 值設成 nil 就不輸出（來源原文不能「復活」）。
    func testSettingTypedValueToNilRemovesTheElement() throws {
        var props = try parseParagraph(Self.borderedParagraph).properties
        props.border = nil
        props.shading = nil
        let xml = props.toXML()
        XCTAssertFalse(xml.contains("pBdr"), xml)
        XCTAssertFalse(xml.contains("<w:shd"), xml)
    }

    /// 樣式的 `<w:pPr>` 走同一個 `parseParagraphProperties`：styles.xml 被 typed
    /// 重新產生時（`addStyle` 等），既有樣式的框線與網底也要原樣保留。
    func testStyleParagraphPropertiesKeepBorderAndShadingVerbatim() throws {
        let stylesXML = """
            <w:styles xmlns:w="\(Self.wNS)"><w:style w:type="paragraph" w:styleId="Boxed">\
            <w:name w:val="Boxed"/><w:pPr>\(Self.sourcePBdr)\(Self.sourceShd)</w:pPr></w:style></w:styles>
            """
        let styles = try DocxReader.parseStyles(from: try XMLDocument(xmlString: stylesXML))
        let boxed = try XCTUnwrap(styles.first { $0.id == "Boxed" })
        XCTAssertNotNil(boxed.paragraphProperties?.border)
        let xml = boxed.toXML()
        XCTAssertTrue(xml.contains(Self.sourcePBdr + Self.sourceShd), xml)
    }

    /// 解析 → 輸出 → 再解析，`ParagraphProperties` 相等（`Issue168` 的真實範本測試
    /// 靠的正是這個不變量）。
    func testParseEmitReparseYieldsEqualProperties() throws {
        let first = try parseParagraph(Self.borderedParagraph)
        let second = try parseParagraph(first.toXML())
        XCTAssertEqual(second.properties, first.properties)
    }
}
