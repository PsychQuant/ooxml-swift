import XCTest
@testable import OOXMLSwift

/// Phase 0 task 1.8: byte-equal round-trip on the four-fixture
/// **golden corpus** required by spec `ooxml-tree-io` requirement
/// "Identity round-trip on untouched sub-trees".
///
/// For each fixture, walk the parts the spec scenario covers, run them through
/// `XmlTreeReader.parse → XmlTreeWriter.serialize` with no mutations, and
/// assert byte-equal.
///
final class TreeRoundTripCorpusTests: XCTestCase {

    func testRegenerateCommittedFixturesWhenExplicitlyRequested() throws {
        guard ProcessInfo.processInfo.environment["UPDATE_TREE_GOLDEN_CORPUS"] == "1" else {
            throw XCTSkip("Set UPDATE_TREE_GOLDEN_CORPUS=1 to regenerate committed fixtures")
        }
        try CorpusFixtureBuilder.regenerateCommittedFixtures()
    }

    func testMultiSectionThesisRoundTripsByteEqual() throws {
        let fixture = try CorpusFixtureBuilder.buildMultiSectionThesis()
        defer { try? FileManager.default.removeItem(at: fixture.url) }
        let document = String(decoding: try CorpusFixtureBuilder.readPart(
            "word/document.xml", from: fixture.url), as: UTF8.self)
        XCTAssertEqual(document.components(separatedBy: "<w:sectPr").count - 1, 3)
        XCTAssertTrue(document.contains("<w:headerReference"))
        XCTAssertTrue(document.contains("<w:footerReference"))
        try assertCorpusFixtureRoundTrips(fixture, ownsFixture: false)
    }

    func testVMLRichRoundTripsByteEqual() throws {
        let fixture = try CorpusFixtureBuilder.buildVMLRich()
        defer { try? FileManager.default.removeItem(at: fixture.url) }
        let document = String(decoding: try CorpusFixtureBuilder.readPart(
            "word/document.xml", from: fixture.url), as: UTF8.self)
        for required in ["<mc:AlternateContent", "<w:pict>", "<wps:wsp>", "<wpg:wgp>"] {
            XCTAssertTrue(document.contains(required), "Missing required corpus element \(required)")
        }
        try assertCorpusFixtureRoundTrips(fixture, ownsFixture: false)
    }

    func testCJKSettingsRoundTripsByteEqual() throws {
        try assertCorpusFixtureRoundTrips(CorpusFixtureBuilder.buildCJKSettings())
    }

    func testCommentAnchoredRoundTripsByteEqual() throws {
        try assertCorpusFixtureRoundTrips(CorpusFixtureBuilder.buildCommentAnchored())
    }

    // MARK: - Helper

    private func assertCorpusFixtureRoundTrips(
        _ fixture: CorpusFixtureBuilder.Fixture,
        ownsFixture: Bool = true,
        file: StaticString = #file,
        line: UInt = #line
    ) throws {
        defer {
            if ownsFixture { try? FileManager.default.removeItem(at: fixture.url) }
        }
        for partPath in fixture.partsToVerify {
            let partBytes = try CorpusFixtureBuilder.readPart(partPath, from: fixture.url)
            let tree = try XmlTreeReader.parse(partBytes)
            let output = try XmlTreeWriter.serialize(tree)
            if output != partBytes {
                let inputString = String(decoding: partBytes, as: UTF8.self)
                let outputString = String(decoding: output, as: UTF8.self)
                XCTFail(
                    """
                    Fixture \(fixture.name) part \(partPath) not byte-equal.
                    --- input ---
                    \(inputString)
                    --- output ---
                    \(outputString)
                    """,
                    file: file, line: line
                )
            }
        }
    }
}
