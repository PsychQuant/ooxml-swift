import XCTest
@testable import OOXMLSwift

/// Regression coverage for PsychQuant/ooxml-swift#19.
///
/// `parseRun` intentionally treats typed `<w:rPr>` lookup as prefix-qualified
/// while its raw-child skip list uses `localName`. This pins the documented
/// malformed-input policy: a foreign-namespace `<x:rPr>` lookalike is silently
/// dropped, not parsed as WordprocessingML run properties and not preserved as
/// raw XML.
final class Issue19NamespacePrefixAsymmetryTests: XCTestCase {

    func testForeignNamespaceRPrLookalikeIsDocumentedSilentDrop() throws {
        let element = try XMLElement(xmlString: """
        <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:x="urn:foreign">
          <x:rPr>
            <x:b/>
          </x:rPr>
          <w:t>body</w:t>
          <x:custom>keep me</x:custom>
        </w:r>
        """)

        let run = try DocxReader.parseRun(
            from: element,
            relationships: RelationshipsCollection()
        )

        XCTAssertEqual(run.text, "body")
        XCTAssertFalse(run.properties.bold, "foreign-namespace rPr must not be parsed as OOXML run properties")

        let rawElements = run.rawElements ?? []
        let rawNames = rawElements.map(\.name)
        XCTAssertFalse(rawNames.contains("rPr"), "foreign-namespace rPr is intentionally skipped by localName allowlist")
        XCTAssertEqual(rawNames, ["custom"], "unrelated foreign children should still be preserved as rawElements")
        XCTAssertEqual(rawElements.first?.xml.contains("<x:custom>keep me</x:custom>"), true)
    }
}
