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
/// Note: fixtures are generated programmatically by `CorpusFixtureBuilder` at
/// test time rather than checked in as binary blobs. The Spectra requirement
/// is satisfied by the `(fixture-class, byte-equal-assertion)` pair, not by
/// blob-on-disk specifically; programmatic builders keep the test repo small
/// and the diff legible.
final class TreeRoundTripCorpusTests: XCTestCase {

    func testMultiSectionThesisRoundTripsByteEqual() throws {
        try assertCorpusFixtureRoundTrips(CorpusFixtureBuilder.buildMultiSectionThesis())
    }

    func testVMLRichRoundTripsByteEqual() throws {
        try assertCorpusFixtureRoundTrips(CorpusFixtureBuilder.buildVMLRich())
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
        file: StaticString = #file,
        line: UInt = #line
    ) throws {
        defer { try? FileManager.default.removeItem(at: fixture.url) }
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
