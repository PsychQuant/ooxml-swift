import XCTest
@testable import OOXMLSwift

/// Phase 0 task 1.4: spec `ooxml-tree-io` requirement
/// "Identity-noise normalization for diff comparison".
final class XmlNodeFingerprintTests: XCTestCase {

    // MARK: - rsid noise

    func testRsidOnlyDifferencesFingerprintEqual() throws {
        let a = #"<?xml version="1.0" encoding="UTF-8"?><w:p xmlns:w="ns_w" w:rsidR="00ABC123"><w:r><w:t>Hello</w:t></w:r></w:p>"#
        let b = #"<?xml version="1.0" encoding="UTF-8"?><w:p xmlns:w="ns_w" w:rsidR="11DEF456"><w:r><w:t>Hello</w:t></w:r></w:p>"#
        let treeA = try XmlTreeReader.parse(Data(a.utf8))
        let treeB = try XmlTreeReader.parse(Data(b.utf8))
        XCTAssertEqual(treeA.root.normalizedFingerprint(), treeB.root.normalizedFingerprint())
    }

    func testAllRsidVariantsDropped() throws {
        let withRsids = #"<?xml version="1.0" encoding="UTF-8"?><w:p xmlns:w="ns_w" w:rsidR="A" w:rsidRPr="B" w:rsidP="C" w:rsidRDefault="D" w:rsidSect="E" w:rsidTr="F"><w:r><w:t>x</w:t></w:r></w:p>"#
        let without = #"<?xml version="1.0" encoding="UTF-8"?><w:p xmlns:w="ns_w"><w:r><w:t>x</w:t></w:r></w:p>"#
        let treeA = try XmlTreeReader.parse(Data(withRsids.utf8))
        let treeB = try XmlTreeReader.parse(Data(without.utf8))
        XCTAssertEqual(treeA.root.normalizedFingerprint(), treeB.root.normalizedFingerprint())
    }

    // MARK: - real content differences

    func testTextContentDifferenceFingerprintsUnequal() throws {
        let a = #"<?xml version="1.0" encoding="UTF-8"?><w:t xmlns:w="ns_w">Hello</w:t>"#
        let b = #"<?xml version="1.0" encoding="UTF-8"?><w:t xmlns:w="ns_w">World</w:t>"#
        let treeA = try XmlTreeReader.parse(Data(a.utf8))
        let treeB = try XmlTreeReader.parse(Data(b.utf8))
        XCTAssertNotEqual(treeA.root.normalizedFingerprint(), treeB.root.normalizedFingerprint())
    }

    func testAttributeValueDifferenceFingerprintsUnequal() throws {
        let a = #"<?xml version="1.0" encoding="UTF-8"?><w:p xmlns:w="ns_w" xmlns:w14="ns_w14" w14:paraId="0AB7C123"/>"#
        let b = #"<?xml version="1.0" encoding="UTF-8"?><w:p xmlns:w="ns_w" xmlns:w14="ns_w14" w14:paraId="DEADBEEF"/>"#
        let treeA = try XmlTreeReader.parse(Data(a.utf8))
        let treeB = try XmlTreeReader.parse(Data(b.utf8))
        XCTAssertNotEqual(treeA.root.normalizedFingerprint(), treeB.root.normalizedFingerprint())
    }

    func testChildOrderDifferenceFingerprintsUnequal() throws {
        // OOXML order is semantic — paragraph order matters.
        let a = #"<?xml version="1.0" encoding="UTF-8"?><w:body xmlns:w="ns_w"><w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p></w:body>"#
        let b = #"<?xml version="1.0" encoding="UTF-8"?><w:body xmlns:w="ns_w"><w:p><w:r><w:t>B</w:t></w:r></w:p><w:p><w:r><w:t>A</w:t></w:r></w:p></w:body>"#
        let treeA = try XmlTreeReader.parse(Data(a.utf8))
        let treeB = try XmlTreeReader.parse(Data(b.utf8))
        XCTAssertNotEqual(treeA.root.normalizedFingerprint(), treeB.root.normalizedFingerprint())
    }

    // MARK: - prefix-variation tolerance

    func testDifferentNamespacePrefixSameURIFingerprintsEqual() throws {
        // Same NS URI bound to different prefixes; semantic equality holds.
        let a = #"<?xml version="1.0" encoding="UTF-8"?><w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Hello</w:t></w:r></w:p>"#
        let b = #"<?xml version="1.0" encoding="UTF-8"?><x:p xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><x:r><x:t>Hello</x:t></x:r></x:p>"#
        let treeA = try XmlTreeReader.parse(Data(a.utf8))
        let treeB = try XmlTreeReader.parse(Data(b.utf8))
        XCTAssertEqual(treeA.root.normalizedFingerprint(), treeB.root.normalizedFingerprint())
    }

    // MARK: - attribute-order tolerance

    func testAttributeOrderPermutationFingerprintsEqual() throws {
        let a = #"<?xml version="1.0" encoding="UTF-8"?><w:p xmlns:w="ns_w" xmlns:w14="ns_w14" w:val="x" w14:paraId="0AB7C123"/>"#
        let b = #"<?xml version="1.0" encoding="UTF-8"?><w:p xmlns:w="ns_w" xmlns:w14="ns_w14" w14:paraId="0AB7C123" w:val="x"/>"#
        let treeA = try XmlTreeReader.parse(Data(a.utf8))
        let treeB = try XmlTreeReader.parse(Data(b.utf8))
        XCTAssertEqual(treeA.root.normalizedFingerprint(), treeB.root.normalizedFingerprint())
    }

    // MARK: - whitespace preservation

    func testPreservedWhitespaceInTextIsSemanticallyDifferent() throws {
        // OOXML preserve-space: leading whitespace is meaningful.
        let a = #"<?xml version="1.0" encoding="UTF-8"?><w:t xmlns:w="ns_w" xml:space="preserve"> Hello</w:t>"#
        let b = #"<?xml version="1.0" encoding="UTF-8"?><w:t xmlns:w="ns_w" xml:space="preserve">Hello</w:t>"#
        let treeA = try XmlTreeReader.parse(Data(a.utf8))
        let treeB = try XmlTreeReader.parse(Data(b.utf8))
        XCTAssertNotEqual(treeA.root.normalizedFingerprint(), treeB.root.normalizedFingerprint())
    }
}
