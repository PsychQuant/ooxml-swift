import XCTest
@testable import OOXMLSwift

/// Tests for the document-content-preservation Spectra change covering
/// PsychQuant/che-word-mcp#58 (body-level bookmark markers), #59 (whitespace
/// `<w:t>` runs), and #60 (RunProperties field-loss audit).
///
/// Each sub-stack adds tests + a matrix-pin assertion class to the
/// `testDocumentContentEqualityInvariant` test that lands here.
final class Issue58_60ContentPreservationTests: XCTestCase {

    // MARK: - Sub-stack A: #58 BodyChild block-level marker preservation

    /// §1.1 — Body-level `<w:bookmarkStart>` and `<w:bookmarkEnd>` SHALL survive
    /// open → modify → save round-trip. Pre-fix, `parseBodyChildren` switch's
    /// `default: continue` silently drops anything that isn't `<w:p>` / `<w:tbl>`
    /// / `<w:sdt>`, so body-level bookmarks vanish.
    ///
    /// Covers spec requirement: `BodyChild enum SHALL cover EG_BlockLevelElts
    /// members beyond paragraph and table` — Scenarios "Body-level bookmarkStart
    /// preserved through round-trip" and "Body-level bookmarkEnd preserved
    /// through round-trip" (specs/ooxml-paragraph-child-schema-coverage/spec.md).
    func testBodyLevelBookmarkRoundTripPreserved() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p><w:r><w:t>before</w:t></w:r></w:p>
        <w:bookmarkStart w:id="0" w:name="_TocTest"/>
        <w:p><w:r><w:t>middle</w:t></w:r></w:p>
        <w:bookmarkEnd w:id="0"/>
        <w:p><w:r><w:t>after</w:t></w:r></w:p>
        </w:body>
        </w:document>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-§1.1-in-\(UUID().uuidString).docx")
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-§1.1-out-\(UUID().uuidString).docx")
        defer {
            try? FileManager.default.removeItem(at: inURL)
            try? FileManager.default.removeItem(at: outURL)
        }
        try buildMinimalDocx(documentXML: documentXML, to: inURL)

        var doc = try DocxReader.read(from: inURL)
        // Force document.xml re-serialization (mirrors MCP body-mutating save).
        doc.modifiedParts.insert("word/document.xml")
        try DocxWriter.write(doc, to: outURL)

        let outXML = try Self.readDocumentXMLString(from: outURL)
        XCTAssertTrue(
            outXML.contains("<w:bookmarkStart") && outXML.contains("w:name=\"_TocTest\""),
            "body-level <w:bookmarkStart w:name=\"_TocTest\"/> SHALL survive round-trip; output:\n\(outXML)"
        )
        XCTAssertTrue(
            outXML.contains("<w:bookmarkEnd") && outXML.contains("w:id=\"0\""),
            "body-level <w:bookmarkEnd w:id=\"0\"/> SHALL survive round-trip; output:\n\(outXML)"
        )
    }

    /// §1.2 — Unknown body-level elements (e.g., `<w:moveFromRangeStart>`) SHALL
    /// be preserved as raw-XML carriers rather than silently dropped, so future
    /// EG_BlockLevelElts / vendor extensions byte-roundtrip even without typed
    /// parser branches.
    ///
    /// Covers spec requirement: `BodyChild enum SHALL cover EG_BlockLevelElts
    /// members beyond paragraph and table` — Scenario "Unknown body-level element
    /// preserved as raw element" (specs/ooxml-paragraph-child-schema-coverage/spec.md).
    func testBodyLevelUnknownElementPreservedAsRaw() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p><w:r><w:t>x</w:t></w:r></w:p>
        <w:moveFromRangeStart w:id="1" w:name="testMove"/>
        <w:p><w:r><w:t>y</w:t></w:r></w:p>
        </w:body>
        </w:document>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-§1.2-in-\(UUID().uuidString).docx")
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-§1.2-out-\(UUID().uuidString).docx")
        defer {
            try? FileManager.default.removeItem(at: inURL)
            try? FileManager.default.removeItem(at: outURL)
        }
        try buildMinimalDocx(documentXML: documentXML, to: inURL)

        var doc = try DocxReader.read(from: inURL)
        doc.modifiedParts.insert("word/document.xml")
        try DocxWriter.write(doc, to: outURL)

        let outXML = try Self.readDocumentXMLString(from: outURL)
        XCTAssertTrue(
            outXML.contains("moveFromRangeStart"),
            "unknown body-level <w:moveFromRangeStart> SHALL be preserved as raw element; output:\n\(outXML)"
        )
    }

    /// §1.6 — `nextBookmarkId` calibration walker SHALL include body-level
    /// `BookmarkRangeMarker` entries (in addition to paragraph-level
    /// `paragraph.bookmarkMarkers`). Otherwise a future API-built bookmark
    /// could collide with an existing body-level id.
    ///
    /// Covers spec requirement: `nextBookmarkId calibration SHALL include
    /// body-level bookmark markers` — Scenario "nextBookmarkId reflects
    /// body-level bookmarks after read" (specs/ooxml-paragraph-child-schema-coverage/spec.md).
    func testNextBookmarkIdReflectsBodyLevelBookmarksAfterRead() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p><w:bookmarkStart w:id="3" w:name="paraLevel"/><w:bookmarkEnd w:id="3"/><w:r><w:t>x</w:t></w:r></w:p>
        <w:bookmarkStart w:id="7" w:name="bodyLevel"/>
        <w:p><w:r><w:t>y</w:t></w:r></w:p>
        <w:bookmarkEnd w:id="7"/>
        </w:body>
        </w:document>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-§1.6-in-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try buildMinimalDocx(documentXML: documentXML, to: inURL)

        let doc = try DocxReader.read(from: inURL)
        XCTAssertEqual(
            doc.nextBookmarkId, 8,
            "nextBookmarkId SHALL be one greater than the global max bookmark id (7 from body-level + 3 from paragraph-level → 8)"
        )
    }

    // MARK: - Cross-cutting matrix-pin (incremental — sub-stack A initial version)

    /// §1.7 / §2.7 / §3.9 — Cross-cutting content-equality invariant against
    /// the thesis fixture. Asserts that for every preservation class covered
    /// by this Spectra change, the round-tripped `word/document.xml` content
    /// equals the source content.
    ///
    /// **Sub-stack A** (this version): `<w:bookmarkStart>` count parity (#58).
    /// **Sub-stack B** (lands with §2.7): + `<w:t>` total-character parity (#59).
    /// **Sub-stack C** (lands with §3.9): + `<w:rFonts>` / `<w:noProof>` /
    /// `<w:lang>` / `<w:kern>` / `w14:*` count parity (#60).
    ///
    /// The pin asserts CONTENT equality (counts and joined-strings), not BYTE
    /// equality — Word's own canonicalization (e.g., adjacent run consolidation)
    /// is allowed to differ.
    ///
    /// Covers spec requirement: `testDocumentContentEqualityInvariant matrix-pin
    /// SHALL assert content equality across preservation classes` — initial
    /// version covering preservation-class 1 of 3
    /// (specs/ooxml-roundtrip-fidelity/spec.md).
    func testDocumentContentEqualityInvariant() throws {
        let fixturePath = "/Users/che/Developer/macdoc/mcp/che-word-mcp/test-files/thesis-fixture.docx"
        guard FileManager.default.fileExists(atPath: fixturePath) else {
            throw XCTSkip("thesis fixture not present at \(fixturePath); skipping content-equality matrix-pin")
        }
        let srcURL = URL(fileURLWithPath: fixturePath)

        // Read source document.xml directly from the ZIP for ground-truth counts.
        let srcDocXML = try Self.readDocumentXMLString(from: srcURL)
        let srcBookmarkStartCount = Self.countBookmarkStartElements(in: srcDocXML)
        XCTAssertGreaterThan(srcBookmarkStartCount, 0,
            "fixture sanity: source has at least one <w:bookmarkStart>; got \(srcBookmarkStartCount)")

        // Round-trip: read → mark modified → write → re-read document.xml.
        var doc = try DocxReader.read(from: srcURL)
        doc.modifiedParts.insert("word/document.xml")
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("matrix-pin-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(doc, to: outURL)

        let outDocXML = try Self.readDocumentXMLString(from: outURL)
        let outBookmarkStartCount = Self.countBookmarkStartElements(in: outDocXML)

        // Preservation class 1 of 3 (#58): bookmarkStart count parity.
        XCTAssertEqual(
            outBookmarkStartCount, srcBookmarkStartCount,
            "<w:bookmarkStart> count SHALL be preserved across round-trip; src=\(srcBookmarkStartCount), out=\(outBookmarkStartCount)"
        )

        // Preservation class 2 of 3 (#59): <w:t> total character content parity.
        // Lands with §2.7 (sub-stack B). Until then, this assertion is documented
        // but skipped to keep the matrix-pin green for sub-stack A.

        // Preservation class 3 of 3 (#60): <w:rFonts>/<w:noProof>/<w:lang>/<w:kern>/w14 counts.
        // Lands with §3.9 (sub-stack C). Until then, this assertion is documented
        // but skipped to keep the matrix-pin green for sub-stack A.
    }

    /// Count `<w:bookmarkStart` elements in raw XML via simple substring scan.
    /// Avoids regex compilation overhead and matches both `<w:bookmarkStart`
    /// followed by space/attr or `/>`.
    static func countBookmarkStartElements(in xml: String) -> Int {
        var count = 0
        var searchRange = xml.startIndex..<xml.endIndex
        while let r = xml.range(of: "<w:bookmarkStart", range: searchRange) {
            count += 1
            searchRange = r.upperBound..<xml.endIndex
        }
        return count
    }

    // MARK: - Helpers

    /// Build a minimal valid `.docx` with the given `document.xml` content.
    /// Mirrors the `buildMinimalDocx` helper used elsewhere in the test suite
    /// (see `Issue56R4StackTests.buildMinimalDocx` for the original).
    private func buildMinimalDocx(documentXML: String, to url: URL) throws {
        let stagingURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-60-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: stagingURL, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: stagingURL) }

        try FileManager.default.createDirectory(at: stagingURL.appendingPathComponent("_rels"), withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: stagingURL.appendingPathComponent("word/_rels"), withIntermediateDirectories: true)

        let contentTypes = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>
        """
        try contentTypes.write(to: stagingURL.appendingPathComponent("[Content_Types].xml"), atomically: true, encoding: .utf8)

        let rels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """
        try rels.write(to: stagingURL.appendingPathComponent("_rels/.rels"), atomically: true, encoding: .utf8)

        let documentRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        </Relationships>
        """
        try documentRels.write(to: stagingURL.appendingPathComponent("word/_rels/document.xml.rels"), atomically: true, encoding: .utf8)

        try documentXML.write(to: stagingURL.appendingPathComponent("word/document.xml"), atomically: true, encoding: .utf8)

        try ZipHelper.zip(stagingURL, to: url)
    }

    /// Read `word/document.xml` from a saved `.docx` as a UTF-8 string.
    /// Mirrors `DocumentXmlLosslessRoundTripTests.readDocumentXMLString`.
    static func readDocumentXMLString(from docxURL: URL) throws -> String {
        let unzipped = try ZipHelper.unzip(docxURL)
        defer { ZipHelper.cleanup(unzipped) }
        let documentURL = unzipped.appendingPathComponent("word/document.xml")
        return try String(contentsOf: documentURL, encoding: .utf8)
    }
}
