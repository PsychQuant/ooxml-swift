import Foundation
import XCTest
@testable import OOXMLSwift

/// Regression tests for PsychQuant/ooxml-swift#175（注意：不是 macdoc#175，
/// 那組是 `Issue175*Tests`）。
///
/// `ParagraphProperties.toXML()` 過去以固定、非 schema 的順序輸出
/// （pStyle → numPr → jc → spacing → ind → keepNext → … → rawChildren → rPr），
/// 已建模欄位與 #168 的 `rawChildren` 都沒有落在 ECMA-376 `CT_PPr` 規定的位置。
///
/// 期望順序直接抄自 ISO/IEC 29500-4:2016 transitional `wml.xsd`
/// （`CT_PPrBase` sequence + `CT_PPr` extension），並與 python-docx
/// `CT_PPr._tag_seq` 交叉核對過 —— 刻意寫成測試裡的字面常數，
/// 不引用實作端的表，這樣「表抄錯」會被測試抓到，而不是自己驗證自己。
final class PPrSchemaOrderTests: XCTestCase {

    /// ECMA-376 `CT_PPr` 子元素順序（CT_PPrBase 33 個 + rPr, sectPr, pPrChange）。
    static let schemaSequence: [String] = [
        "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
        "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
        "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
        "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
        "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
        "suppressOverlap", "jc", "textDirection", "textAlignment",
        "textboxTightWrap", "outlineLvl", "divId", "cnfStyle",
        "rPr", "sectPr", "pPrChange",
    ]

    /// `ParagraphProperties` 有 typed 欄位的元素；其餘 CT_PPrBase 元素只能經
    /// `rawChildren` 進來。
    static let typedNames: Set<String> = [
        "pStyle", "keepNext", "keepLines", "pageBreakBefore", "numPr",
        "pBdr", "shd", "spacing", "ind", "jc", "rPr",
    ]

    static let namespaces = [
        "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"",
        "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"",
        "xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\"",
        "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"",
    ].joined(separator: " ")

    // MARK: - Helpers

    /// `toXML()` 輸出的是 pPr 的內容（不含外層 `<w:pPr>`）；包一層帶命名空間
    /// 宣告的 `<w:pPr>` 後解析，回傳直接子元素的 qualified name。
    private func childNames(ofPPrInner inner: String,
                            file: StaticString = #filePath, line: UInt = #line) throws -> [String] {
        let wrapper = try XMLElement(xmlString: "<w:pPr \(Self.namespaces)>\(inner)</w:pPr>")
        return (wrapper.children ?? []).compactMap { ($0 as? XMLElement)?.name }
    }

    /// 取出 `<w:p>` 輸出中第一個 `<w:pPr>` 的直接子元素 qualified name。
    private func pPrChildNames(ofParagraphXML xml: String) throws -> [String] {
        let wrapped = xml.replacingOccurrences(of: "<w:p>", with: "<w:p \(Self.namespaces)>")
        let p = try XMLElement(xmlString: wrapped)
        guard let pPr = p.elements(forName: "w:pPr").first else { return [] }
        return (pPr.children ?? []).compactMap { ($0 as? XMLElement)?.name }
    }

    private func parseParagraph(_ xml: String) throws -> Paragraph {
        let element = try XMLElement(xmlString: xml)
        return try DocxReader.parseParagraph(
            from: element,
            relationships: RelationshipsCollection(),
            styles: [],
            numbering: Numbering()
        )
    }

    private func raw(_ localName: String) -> RawElement {
        RawElement(name: localName, xml: "<w:\(localName) w:val=\"0\"/>")
    }

    private func allTypedFieldsSet() -> ParagraphProperties {
        var props = ParagraphProperties()
        props.style = "Heading1"
        props.keepNext = true
        props.keepLines = true
        props.pageBreakBefore = true
        props.numbering = NumberingInfo(numId: 3, level: 1)
        props.border = ParagraphBorder(top: ParagraphBorderStyle())
        props.shading = CellShading(fill: "FFFF00")
        props.spacing = Spacing(before: 120, after: 240)
        props.indentation = Indentation(left: 720)
        props.alignment = .center
        props.markRunProperties = RunProperties()
        props.markRunProperties?.bold = true
        return props
    }

    // MARK: - Section A: schema-order table (typed + raw)

    /// 只設 typed 欄位：輸出順序必須是 schema 順序，不再是
    /// pStyle → numPr → jc → spacing → ind → keepNext → …
    func testTypedFieldsEmitInCTPPrSequence() throws {
        let names = try childNames(ofPPrInner: allTypedFieldsSet().toXML())
        let expected = Self.schemaSequence
            .filter { Self.typedNames.contains($0) }
            .map { "w:\($0)" }
        XCTAssertEqual(names, expected)
    }

    /// 表驅動：CT_PPrBase 的每一個元素（typed 的走欄位，其餘走 rawChildren，
    /// 而且 rawChildren 故意以**反向** schema 順序給）都要落在自己的 schema 位置。
    func testEveryCTPPrBaseChildLandsAtItsSchemaSlot() throws {
        var props = allTypedFieldsSet()
        let rawOnly = Self.schemaSequence.filter {
            !Self.typedNames.contains($0) && $0 != "sectPr" && $0 != "pPrChange"
        }
        props.rawChildren = rawOnly.reversed().map(raw)

        let names = try childNames(ofPPrInner: props.toXML())
        let expected = Self.schemaSequence
            .filter { $0 != "sectPr" && $0 != "pPrChange" }
            .map { "w:\($0)" }
        XCTAssertEqual(names, expected,
                       "typed 欄位與 rawChildren 必須依元素名稱交錯排入 CT_PPr 的 schema 位置")
    }

    /// Paragraph 層：`rPr` → `sectPr` → `pPrChange` 必須是 pPr 的最後三個子元素，
    /// 且在所有 CT_PPrBase 元素之後。
    func testParagraphEmitsRPrThenSectPrThenPPrChangeLast() throws {
        var paragraph = Paragraph(text: "x")
        var props = allTypedFieldsSet()
        props.rawChildren = [raw("outlineLvl"), raw("kinsoku")]
        props.sectionBreak = .nextPage
        paragraph.properties = props
        let revision = Revision(id: 9, type: .paragraphChange, author: "t", paragraphIndex: 0)
        paragraph.revisions = [revision]
        paragraph.paragraphFormatChangeRevisionId = 9
        paragraph.previousProperties = ParagraphProperties()

        let names = try pPrChildNames(ofParagraphXML: paragraph.toXML())
        XCTAssertEqual(Array(names.suffix(3)), ["w:rPr", "w:sectPr", "w:pPrChange"])
        let indices = names.compactMap { name in
            Self.schemaSequence.firstIndex(of: String(name.dropFirst(2)))
        }
        XCTAssertEqual(indices.count, names.count, "every emitted child must be a CT_PPr element")
        XCTAssertEqual(indices, indices.sorted(), "pPr children must be in CT_PPr order: \(names)")
    }

    /// 來源本來就是 schema 順序、同時有 typed 與未建模子元素的段落：
    /// 讀進來再 typed 重新序列化，pPr 應逐位元組等於來源（#175 Impact：
    /// 修正前 typed 重新序列化的段落「永遠無法逐位元組等於來源」）。
    func testSchemaOrderedSourcePPrRoundTripsByteForByte() throws {
        let sourcePPr = "<w:pPr><w:pStyle w:val=\"Body\"/><w:keepNext/>"
            + "<w:widowControl w:val=\"0\"/>"
            + "<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"4\"/></w:numPr>"
            + "<w:kinsoku w:val=\"0\"/><w:snapToGrid w:val=\"0\"/>"
            + "<w:spacing w:before=\"120\" w:after=\"60\"/><w:ind w:left=\"360\"/>"
            + "<w:jc w:val=\"both\"/><w:outlineLvl w:val=\"2\"/>"
            + "<w:rPr><w:b/></w:rPr></w:pPr>"
        let paragraph = try parseParagraph(
            "<w:p \(Self.namespaces)>\(sourcePPr)<w:r><w:t>body</w:t></w:r></w:p>")
        XCTAssertTrue(paragraph.toXML().contains(sourcePPr),
                      "expected verbatim pPr, got: \(paragraph.toXML())")
    }

    // MARK: - Section B: children outside the schema table

    /// 規則：不在 CT_PPr 表內的子元素（擴充命名空間、`mc:AlternateContent`…）
    /// 緊跟在來源中前一個 schema 已知元素之後；多個錨定在同一元素的依來源順序。
    func testUnknownChildKeepsPositionAfterPrecedingKnownSibling() throws {
        let paragraph = try parseParagraph("""
            <w:p \(Self.namespaces)><w:pPr><w:pStyle w:val="A"/>\
            <w15:collapsed w:val="1"/><w14:futureThing w14:val="x"/>\
            <w:snapToGrid w:val="0"/><w:jc w:val="center"/>\
            <mc:AlternateContent><mc:Fallback/></mc:AlternateContent>\
            <w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>
            """)
        var props = paragraph.properties
        // 讀進來之後再加一個 typed 欄位：spacing（slot 在 snapToGrid 與 jc 之間）。
        props.spacing = Spacing(before: 10)
        let names = try childNames(ofPPrInner: props.toXML())
        XCTAssertEqual(names, [
            "w:pStyle", "w15:collapsed", "w14:futureThing", "w:snapToGrid",
            "w:spacing", "w:jc", "mc:AlternateContent", "w:rPr",
        ])
    }

    /// 來源中沒有前一個已知元素的未知子元素 → 放在 pPr 最前面。
    func testUnknownChildWithNoPrecedingKnownSiblingStaysFirst() throws {
        let paragraph = try parseParagraph("""
            <w:p \(Self.namespaces)><w:pPr><w15:collapsed w:val="1"/>\
            <w:pStyle w:val="A"/><w:jc w:val="left"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>
            """)
        let names = try childNames(ofPPrInner: paragraph.properties.toXML())
        XCTAssertEqual(names, ["w15:collapsed", "w:pStyle", "w:jc"])
    }

    /// 錨點是 schema 的「位置」而不是實際元素：錨點元素被 typed 編輯移除後，
    /// 未知子元素仍停在該位置之後（不會漂到最前或最後）。
    func testUnknownChildAnchorSurvivesRemovalOfAnchorElement() throws {
        let paragraph = try parseParagraph("""
            <w:p \(Self.namespaces)><w:pPr><w:pStyle w:val="A"/><w:jc w:val="left"/>\
            <w15:collapsed w:val="1"/><w:rPr><w:i/></w:rPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>
            """)
        var props = paragraph.properties
        props.alignment = nil
        props.spacing = Spacing(after: 5)
        let names = try childNames(ofPPrInner: props.toXML())
        XCTAssertEqual(names, ["w:pStyle", "w:spacing", "w15:collapsed", "w:rPr"])
    }

    /// 呼叫端自己組的 rawChildren（沒有讀取時的來源錨點）：未知元素錨定在
    /// `rawChildren` 中前一個 schema 已知元素之後；前面沒有已知元素就放最前。
    func testCallerBuiltUnknownRawChildrenAnchorToPrecedingRawSibling() throws {
        var props = ParagraphProperties()
        props.style = "A"
        props.alignment = .right
        props.rawChildren = [
            RawElement(name: "leading", xml: "<w15:leading/>"),
            raw("kinsoku"),
            RawElement(name: "afterKinsoku", xml: "<w15:afterKinsoku/>"),
            raw("outlineLvl"),
        ]
        let names = try childNames(ofPPrInner: props.toXML())
        XCTAssertEqual(names, [
            "w15:leading", "w:pStyle", "w:kinsoku", "w15:afterKinsoku", "w:jc", "w:outlineLvl",
        ])
    }

    /// 擴充命名空間裡與 schema 同名的元素（`w14:jc`）不是 `w:jc`：
    /// 讀取時要保留（修正前 raw 捕捉只比 localName，會把它當成已建模的 jc 丟掉），
    /// 輸出時依未知元素規則處理，不能被排到 jc 的 schema 位置。
    func testForeignNamespaceElementWithSchemaLocalNameIsUnknown() throws {
        let paragraph = try parseParagraph("""
            <w:p \(Self.namespaces)><w:pPr><w:pStyle w:val="A"/>\
            <w14:jc w14:val="x"/><w:kinsoku w:val="0"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>
            """)
        XCTAssertNil(paragraph.properties.alignment, "w14:jc must not be read as w:jc")
        let names = try childNames(ofPPrInner: paragraph.properties.toXML())
        XCTAssertEqual(names, ["w:pStyle", "w14:jc", "w:kinsoku"])
    }

    // MARK: - Section C: CT_PBdr child order

    /// `<w:pBdr>` 的子元素順序是 top, left, bottom, right, between, bar
    /// （修正前 typed 輸出是 top, bottom, left, right, between）。
    func testParagraphBorderChildrenFollowCTPBdrSequence() throws {
        let style = ParagraphBorderStyle()
        let border = ParagraphBorder(top: style, bottom: style, left: style, right: style, between: style)
        let element = try XMLElement(xmlString: border.toXML()
            .replacingOccurrences(of: "<w:pBdr>", with: "<w:pBdr \(Self.namespaces)>"))
        let names = (element.children ?? []).compactMap { ($0 as? XMLElement)?.localName }
        XCTAssertEqual(names, ["top", "left", "bottom", "right", "between"])
    }
}
