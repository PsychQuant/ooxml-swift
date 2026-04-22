import XCTest
@testable import OOXMLSwift

final class FieldCodeParseTests: XCTestCase {

    // MARK: SequenceField

    func testSequenceFieldParseArabic() {
        let f = SequenceField.parse(instrText: " SEQ Figure ")
        XCTAssertEqual(f?.identifier, "Figure")
        XCTAssertEqual(f?.format, .arabic)
        XCTAssertNil(f?.resetLevel)
    }

    func testSequenceFieldParseWithFormatAndReset() {
        let f = SequenceField.parse(instrText: " SEQ Figure \\* ARABIC \\s 1 ")
        XCTAssertEqual(f?.identifier, "Figure")
        XCTAssertEqual(f?.format, .arabic)
        XCTAssertEqual(f?.resetLevel, 1)
    }

    func testSequenceFieldParseChineseIdentifier() {
        let f = SequenceField.parse(instrText: " SEQ 圖 \\s 1 ")
        XCTAssertEqual(f?.identifier, "圖")
        XCTAssertEqual(f?.resetLevel, 1)
    }

    func testSequenceFieldParseHideResult() {
        let f = SequenceField.parse(instrText: " SEQ Figure \\h ")
        XCTAssertEqual(f?.identifier, "Figure")
        XCTAssertTrue(f?.hideResult ?? false)
    }

    func testSequenceFieldParseNonSEQ() {
        XCTAssertNil(SequenceField.parse(instrText: " STYLEREF 1 \\s "))
        XCTAssertNil(SequenceField.parse(instrText: " REF fig_returns "))
        XCTAssertNil(SequenceField.parse(instrText: ""))
    }

    func testSequenceFieldRoundTrip() {
        let original = SequenceField(identifier: "Figure", format: .roman, resetLevel: 2)
        let instrText = original.fieldInstruction
        let parsed = SequenceField.parse(instrText: " \(instrText) ")
        XCTAssertEqual(parsed?.identifier, original.identifier)
        XCTAssertEqual(parsed?.format, original.format)
        XCTAssertEqual(parsed?.resetLevel, original.resetLevel)
    }

    // MARK: StyleRefField

    func testStyleRefFieldParse() {
        let f = StyleRefField.parse(instrText: " STYLEREF 1 \\s ")
        XCTAssertEqual(f?.headingLevel, 1)
        XCTAssertTrue(f?.suppressNonDelimiter ?? false)
    }

    func testStyleRefFieldParseBareLevel() {
        let f = StyleRefField.parse(instrText: " STYLEREF 2 ")
        XCTAssertEqual(f?.headingLevel, 2)
        XCTAssertFalse(f?.suppressNonDelimiter ?? true)
        XCTAssertFalse(f?.insertPositionBeforeRef ?? true)
    }

    func testStyleRefFieldParseNonSTYLEREF() {
        XCTAssertNil(StyleRefField.parse(instrText: " SEQ Figure "))
        XCTAssertNil(StyleRefField.parse(instrText: " STYLEREF abc "))
    }

    func testStyleRefFieldRoundTrip() {
        let original = StyleRefField(headingLevel: 3, suppressNonDelimiter: true)
        let parsed = StyleRefField.parse(instrText: " \(original.fieldInstruction) ")
        XCTAssertEqual(parsed?.headingLevel, 3)
        XCTAssertTrue(parsed?.suppressNonDelimiter ?? false)
    }

    // MARK: ReferenceField

    func testReferenceFieldParseREF() {
        let f = ReferenceField.parse(instrText: " REF fig_returns \\h ")
        XCTAssertEqual(f?.type, .ref)
        XCTAssertEqual(f?.bookmarkName, "fig_returns")
        XCTAssertTrue(f?.createHyperlink ?? false)
    }

    func testReferenceFieldParsePAGEREF() {
        let f = ReferenceField.parse(instrText: " PAGEREF fig_returns \\h ")
        XCTAssertEqual(f?.type, .pageRef)
        XCTAssertEqual(f?.bookmarkName, "fig_returns")
    }

    func testReferenceFieldParseNOTEREF() {
        let f = ReferenceField.parse(instrText: " NOTEREF footnote1 ")
        XCTAssertEqual(f?.type, .noteRef)
    }

    func testReferenceFieldParseNonREF() {
        XCTAssertNil(ReferenceField.parse(instrText: " SEQ Figure "))
    }

    func testReferenceFieldRoundTrip() {
        let original = ReferenceField(type: .ref, bookmarkName: "b1", createHyperlink: true)
        let parsed = ReferenceField.parse(instrText: " \(original.fieldInstruction) ")
        XCTAssertEqual(parsed?.type, original.type)
        XCTAssertEqual(parsed?.bookmarkName, original.bookmarkName)
        XCTAssertEqual(parsed?.createHyperlink, original.createHyperlink)
    }
}
