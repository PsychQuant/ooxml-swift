import XCTest
@testable import OOXMLSwift

/// word-aligned-state-sync 7.1 verify findings (v30-correctness panel) —
/// robustness gaps in the primary tree parse/serialize path:
/// 1. unbounded recursion → uncatchable SIGSEGV on hostile nesting (P1)
/// 2. UTF-8 BOM → total parse failure on BOM-emitting tools' output (P2)
/// 3. control chars in attribute values pass through unescaped — a
///    conformant reader (libxml2, which the v1.0 read projection feeds)
///    normalizes them to spaces, corrupting values (P2)
final class TreeRobustnessTests: XCTestCase {

    func testHostileNestingThrowsInsteadOfCrashing() throws {
        let depth = 60_000
        let xml = "<?xml version=\"1.0\"?>"
            + String(repeating: "<a>", count: depth)
            + String(repeating: "</a>", count: depth)
        XCTAssertThrowsError(try XmlTreeReader.parse(Data(xml.utf8)),
                             "hostile nesting must throw a catchable error, not overflow the stack") { error in
            guard case XmlTreeReaderError.nestingTooDeep = error else {
                return XCTFail("expected nestingTooDeep, got \(error)")
            }
        }
    }

    func testRealisticNestingStillParses() throws {
        // Word tables nest ≤5 levels; 200 leaves generous headroom while
        // staying far under the guard.
        let depth = 200
        let xml = "<?xml version=\"1.0\"?>"
            + String(repeating: "<a>", count: depth)
            + String(repeating: "</a>", count: depth)
        XCTAssertNoThrow(try XmlTreeReader.parse(Data(xml.utf8)))
    }

    func testUTF8BOMIsSkipped() throws {
        var data = Data([0xEF, 0xBB, 0xBF])
        data.append(Data("<?xml version=\"1.0\"?><w:document xmlns:w=\"http://x\"><w:body/></w:document>".utf8))
        let tree = try XmlTreeReader.parse(data)
        XCTAssertEqual(tree.root.localName, "document",
                       "a UTF-8 BOM before the prolog must be skipped, not fail the parse")
    }

    func testAttributeControlCharactersEscapedOnDirtySerialize() throws {
        let xml = "<?xml version=\"1.0\"?><r a=\"x&#10;y&#9;z\"/>"
        let tree = try XmlTreeReader.parse(Data(xml.utf8))
        tree.root.markDirty()   // force re-serialization from typed fields

        let out = try XmlTreeWriter.serialize(tree)
        let s = String(decoding: out, as: UTF8.self)
        XCTAssertFalse(s.contains("a=\"x\ny"),
                       "literal newline in attribute value corrupts under conformant readers")
        XCTAssertTrue(s.contains("&#10;") && s.contains("&#9;"),
                      "control chars in attribute values must re-escape as character references; got: \(s)")

        // Cross-parser confirmation: libxml2 (the v1.0 read-projection consumer)
        // must recover the exact original value.
        let doc = try XMLDocument(data: out)
        let attr = (doc.rootElement()?.attribute(forName: "a"))?.stringValue
        XCTAssertEqual(attr, "x\ny\tz", "value must survive a conformant re-read")
    }
}
