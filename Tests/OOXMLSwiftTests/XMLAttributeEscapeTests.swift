import XCTest
@testable import OOXMLSwift

final class XMLAttributeEscapeTests: XCTestCase {

    func testEscapesAllFiveXMLSpecialCharacters() {
        let input = "a&b<c>d\"e'f"
        let output = escapeXMLAttribute(input)
        XCTAssertEqual(output, "a&amp;b&lt;c&gt;d&quot;e&apos;f")
    }

    func testPreservesBenignCharacters() {
        let input = "hello world 你好 123"
        XCTAssertEqual(escapeXMLAttribute(input), input)
    }

    func testHandlesEmptyString() {
        XCTAssertEqual(escapeXMLAttribute(""), "")
    }

    func testEscapesRepeatedSpecialChars() {
        XCTAssertEqual(escapeXMLAttribute("&&<<>>"), "&amp;&amp;&lt;&lt;&gt;&gt;")
    }

    func testUsesAposNotNumericReference() {
        XCTAssertEqual(escapeXMLAttribute("'"), "&apos;")
        XCTAssertNotEqual(escapeXMLAttribute("'"), "&#39;")
    }
}
