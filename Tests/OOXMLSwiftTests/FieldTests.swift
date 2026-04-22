import XCTest
@testable import OOXMLSwift

final class FieldTests: XCTestCase {

    // MARK: - SequenceField (existing, verifying Chinese identifier support)

    func testSequenceFieldASCIIIdentifierArabic() {
        let field = SequenceField(identifier: "Figure", format: .arabic, resetLevel: nil, cachedResult: "1")
        let xml = field.toFieldXML()
        XCTAssertTrue(
            xml.contains("<w:instrText xml:space=\"preserve\"> SEQ Figure </w:instrText>"),
            "Expected instrText to contain ' SEQ Figure ', got: \(xml)"
        )
        XCTAssertTrue(xml.contains("<w:fldChar w:fldCharType=\"begin\"/>"))
        XCTAssertTrue(xml.contains("<w:fldChar w:fldCharType=\"separate\"/>"))
        XCTAssertTrue(xml.contains("<w:fldChar w:fldCharType=\"end\"/>"))
    }

    func testSequenceFieldChineseIdentifierWithChapterReset() {
        let field = SequenceField(identifier: "圖", format: .arabic, resetLevel: 1, cachedResult: "1")
        let xml = field.toFieldXML()
        XCTAssertTrue(
            xml.contains("<w:instrText xml:space=\"preserve\"> SEQ 圖 \\s 1 </w:instrText>"),
            "Expected instrText to contain ' SEQ 圖 \\s 1 ', got: \(xml)"
        )
    }

    func testSequenceFieldAllFormats() {
        XCTAssertEqual(SequenceField(identifier: "F", format: .arabic).fieldInstruction, "SEQ F")
        XCTAssertEqual(SequenceField(identifier: "F", format: .alphabetic).fieldInstruction, "SEQ F \\* ALPHABETIC")
        XCTAssertEqual(SequenceField(identifier: "F", format: .roman).fieldInstruction, "SEQ F \\* ROMAN")
    }

    // MARK: - ReferenceField (existing, verifying REF + PAGEREF paths)

    func testReferenceFieldRefWithHyperlink() {
        let field = ReferenceField(
            type: .ref,
            bookmarkName: "fig_returns",
            createHyperlink: true,
            cachedResult: "圖 4-1"
        )
        let xml = field.toFieldXML()
        XCTAssertTrue(
            xml.contains("<w:instrText xml:space=\"preserve\"> REF fig_returns \\h </w:instrText>"),
            "Expected REF with \\h, got: \(xml)"
        )
    }

    func testReferenceFieldPageRefConvenienceConstructor() {
        let field = ReferenceField.pageOf("fig_returns", hyperlink: true)
        XCTAssertEqual(field.fieldInstruction, "PAGEREF fig_returns \\h")
    }

    // MARK: - StyleRefField (new)

    func testStyleRefFieldLevel1SuppressNonDelimiter() {
        let field = StyleRefField(headingLevel: 1, suppressNonDelimiter: true, cachedResult: "4")
        let xml = field.toFieldXML()
        XCTAssertTrue(
            xml.contains("<w:instrText xml:space=\"preserve\"> STYLEREF 1 \\s </w:instrText>"),
            "Expected STYLEREF 1 \\s, got: \(xml)"
        )
    }

    func testStyleRefFieldLevel2NoFlag() {
        let field = StyleRefField(headingLevel: 2, suppressNonDelimiter: false, cachedResult: "1.1")
        XCTAssertEqual(field.fieldInstruction, "STYLEREF 2")
    }

    func testStyleRefFieldEmitsFiveRunStructure() {
        let field = StyleRefField(headingLevel: 1, suppressNonDelimiter: true, cachedResult: "Chapter")
        let xml = field.toFieldXML()
        XCTAssertTrue(xml.contains("<w:fldChar w:fldCharType=\"begin\"/>"))
        XCTAssertTrue(xml.contains("<w:fldChar w:fldCharType=\"separate\"/>"))
        XCTAssertTrue(xml.contains("<w:t>Chapter</w:t>"))
        XCTAssertTrue(xml.contains("<w:fldChar w:fldCharType=\"end\"/>"))
    }
}
