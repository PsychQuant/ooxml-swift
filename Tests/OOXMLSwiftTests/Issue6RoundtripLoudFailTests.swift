import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Round-trip completeness tests for [PsychQuant/ooxml-swift#6](https://github.com/PsychQuant/ooxml-swift/issues/6).
///
/// Covers two sub-findings from the post-#56 verification bundle:
/// - F8 — `AlternateContent.fallbackRuns` typed-edit dirty-tracking + emit-time throw
/// - F9 — `Paragraph.commentIds` deprecation + computed-getter migration; comment marker round-trip invariant
///
/// Each test maps to one `Scenario:` block in the spec
/// `openspec/changes/roundtrip-loud-fail/specs/ooxml-roundtrip-completeness/spec.md`.
final class Issue6RoundtripLoudFailTests: XCTestCase {

    // MARK: - F8: AlternateContent.fallbackRunsModified didSet flag

    /// Spec: "Reader-loaded AlternateContent starts clean"
    func testReaderLoadedAlternateContentStartsClean() throws {
        let url = try buildAlternateContentFixture()
        defer { try? FileManager.default.removeItem(at: url) }

        let doc = try DocxReader.read(from: url)
        guard let firstPara = firstParagraph(in: doc),
              let ac = firstPara.alternateContents.first else {
            XCTFail("Reader did not surface an AlternateContent; check fixture")
            return
        }
        XCTAssertFalse(ac.fallbackRunsModified,
            "Reader-loaded AlternateContent must start with fallbackRunsModified == false")
    }

    /// Spec: "Reassigning fallbackRuns flips the flag"
    func testReassignmentFlipsFallbackRunsModified() {
        var ac = AlternateContent(rawXML: "<mc:AlternateContent/>", fallbackRuns: [], position: 1)
        XCTAssertFalse(ac.fallbackRunsModified)
        ac.fallbackRuns = [Run(text: "new")]
        XCTAssertTrue(ac.fallbackRunsModified)
    }

    /// Spec: "Mutating fallbackRuns through index flips the flag"
    func testIndexedMutationFlipsFallbackRunsModified() {
        var ac = AlternateContent(rawXML: "<mc:AlternateContent/>",
                                  fallbackRuns: [Run(text: "old")],
                                  position: 1)
        XCTAssertFalse(ac.fallbackRunsModified)
        ac.fallbackRuns[0].text = "changed"
        XCTAssertTrue(ac.fallbackRunsModified)
    }

    /// Spec example table: append() flips the flag.
    func testAppendFlipsFallbackRunsModified() {
        var ac = AlternateContent(rawXML: "<mc:AlternateContent/>", fallbackRuns: [], position: 1)
        XCTAssertFalse(ac.fallbackRunsModified)
        ac.fallbackRuns.append(Run(text: "added"))
        XCTAssertTrue(ac.fallbackRunsModified)
    }

    // MARK: - F8: Paragraph emit throw on dirty fallback

    /// Spec: "Modified fallbackRuns triggers throw on emit"
    func testEmitThrowsOnModifiedFallbackRuns() throws {
        let url = try buildAlternateContentFixture()
        defer { try? FileManager.default.removeItem(at: url) }

        var doc = try DocxReader.read(from: url)
        guard var para = firstParagraph(in: doc),
              !para.alternateContents.isEmpty else {
            XCTFail("Fixture missing AlternateContent")
            return
        }
        para.alternateContents[0].fallbackRuns = [Run(text: "x")]
        doc = replaceFirstParagraph(in: doc, with: para)
        let modifiedAC = para.alternateContents[0]

        XCTAssertThrowsError(try para.toXMLThrowing()) { error in
            guard case let RoundtripError.unserializedFallbackEdit(position) = error else {
                XCTFail("Expected RoundtripError.unserializedFallbackEdit, got \(error)")
                return
            }
            XCTAssertEqual(position, modifiedAC.position)
        }
    }

    /// Spec: "Unmodified fallbackRuns emits cleanly"
    func testEmitDoesNotThrowOnUnmodifiedFallbackRuns() throws {
        let url = try buildAlternateContentFixture()
        defer { try? FileManager.default.removeItem(at: url) }

        let doc = try DocxReader.read(from: url)
        guard let para = firstParagraph(in: doc),
              !para.alternateContents.isEmpty else {
            XCTFail("Fixture missing AlternateContent")
            return
        }

        // Both emit paths must succeed for unmutated input.
        XCTAssertNoThrow(try para.toXMLThrowing())
        let throwing = try para.toXMLThrowing()
        let nonThrowing = para.toXML()
        XCTAssertEqual(throwing, nonThrowing,
            "Throwing emit must match non-throwing emit byte-equivalent for clean input")
    }

    /// Spec: "Multiple AlternateContents — only modified one triggers throw"
    func testEmitThrowsForOnlyTheModifiedAlternateContent() {
        var para = Paragraph()
        let ac1 = AlternateContent(rawXML: "<mc:AlternateContent/>", fallbackRuns: [], position: 1)
        var ac2 = AlternateContent(rawXML: "<mc:AlternateContent/>", fallbackRuns: [], position: 2)
        ac2.fallbackRuns = [Run(text: "modified")]
        para.alternateContents = [ac1, ac2]

        XCTAssertThrowsError(try para.toXMLThrowing()) { error in
            guard case let RoundtripError.unserializedFallbackEdit(position) = error else {
                XCTFail("Expected RoundtripError.unserializedFallbackEdit, got \(error)")
                return
            }
            XCTAssertEqual(position, 2,
                "Throw must carry position of the modified AC (2), not the clean one (1)")
        }
    }

    // MARK: - F9: commentIds deprecation + computed getter

    /// Spec: "Reader-loaded paragraph commentIds matches markers"
    /// Updated per apply-time deviation: Reader stops populating `commentIds`
    /// (deprecated stored field); the new `commentRangeIds` computed property
    /// derives the canonical list from markers.
    func testCommentRangeIdsReflectsReaderLoadedMarkers() throws {
        let url = try buildCommentMarkerFixture()
        defer { try? FileManager.default.removeItem(at: url) }

        let doc = try DocxReader.read(from: url)
        guard let para = firstParagraph(in: doc) else {
            XCTFail("Fixture missing paragraph")
            return
        }
        XCTAssertEqual(para.commentRangeIds, [3], "commentRangeIds must derive [3] from markers")
        let starts = para.commentRangeMarkers.filter { $0.kind == .start && $0.id == 3 }
        let ends = para.commentRangeMarkers.filter { $0.kind == .end && $0.id == 3 }
        XCTAssertEqual(starts.count, 1, "exactly one commentRangeStart with id 3")
        XCTAssertEqual(ends.count, 1, "exactly one commentRangeEnd with id 3")
    }

    /// Spec: "commentRangeIds reflects markers added post-Reader"
    func testCommentRangeIdsReflectsLiveMarkerEdits() {
        var para = Paragraph()
        let start = CommentRangeMarker(kind: .start, id: 99, position: 1)
        let end = CommentRangeMarker(kind: .end, id: 99, position: 2)
        para.commentRangeMarkers = [start, end]
        XCTAssertTrue(para.commentRangeIds.contains(99),
            "commentRangeIds must reflect markers added after construction")
    }

    // MARK: - F9: Comment marker round-trip regression

    /// Spec example table: marker count invariant (start=1, commentReference=1, end=1).
    func testCommentRangeMarkersRoundTripPreservesAllThreeElements() throws {
        let url = try buildCommentMarkerFixture()
        defer { try? FileManager.default.removeItem(at: url) }

        let doc = try DocxReader.read(from: url)
        guard let para = firstParagraph(in: doc) else {
            XCTFail("Fixture missing paragraph")
            return
        }
        let xml = para.toXML()
        let startCount = countOccurrences(of: "<w:commentRangeStart", in: xml)
        let refCount = countOccurrences(of: "<w:commentReference", in: xml)
        let endCount = countOccurrences(of: "<w:commentRangeEnd", in: xml)
        XCTAssertEqual(startCount, 1, "commentRangeStart count: \(xml)")
        XCTAssertEqual(refCount, 1, "commentReference count: \(xml)")
        XCTAssertEqual(endCount, 1, "commentRangeEnd count: \(xml)")
    }

    // MARK: - Fixture builders

    private func firstParagraph(in doc: WordDocument) -> Paragraph? {
        for child in doc.body.children {
            if case let .paragraph(p) = child { return p }
        }
        return nil
    }

    private func replaceFirstParagraph(in doc: WordDocument, with para: Paragraph) -> WordDocument {
        var copy = doc
        for (i, child) in copy.body.children.enumerated() {
            if case .paragraph = child {
                copy.body.children[i] = .paragraph(para)
                return copy
            }
        }
        return copy
    }

    private func countOccurrences(of needle: String, in haystack: String) -> Int {
        var count = 0
        var search = haystack.startIndex..<haystack.endIndex
        while let r = haystack.range(of: needle, range: search) {
            count += 1
            search = r.upperBound..<haystack.endIndex
        }
        return count
    }

    /// Fixture with a single paragraph carrying one `<mc:AlternateContent>` block.
    private func buildAlternateContentFixture() throws -> URL {
        let alternateContentXML = """
        <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <mc:Choice Requires="wps14"><w:r><w:t>choice</w:t></w:r></mc:Choice>
          <mc:Fallback><w:r><w:t>fallback-text</w:t></w:r></mc:Fallback>
        </mc:AlternateContent>
        """
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
          <w:body>
            <w:p>
              <w:r><w:t>before</w:t></w:r>
              \(alternateContentXML)
              <w:r><w:t>after</w:t></w:r>
            </w:p>
          </w:body>
        </w:document>
        """
        return try buildMinimalDocx(documentXML: documentXML)
    }

    /// Fixture with a single paragraph containing comment range markers + reference.
    private func buildCommentMarkerFixture() throws -> URL {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:commentRangeStart w:id="3"/>
              <w:r><w:t>commented text</w:t></w:r>
              <w:commentRangeEnd w:id="3"/>
              <w:r><w:commentReference w:id="3"/></w:r>
            </w:p>
          </w:body>
        </w:document>
        """
        return try buildMinimalDocx(documentXML: documentXML)
    }

    /// Pattern lifted from `Issue7XMLHardeningTests.buildHardeningFixture`.
    private func buildMinimalDocx(documentXML: String) throws -> URL {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("loudfail-fixture-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: stagingDir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: stagingDir) }

        func write(_ s: String, to rel: String) throws {
            let url = stagingDir.appendingPathComponent(rel)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(),
                withIntermediateDirectories: true
            )
            try s.write(to: url, atomically: true, encoding: .utf8)
        }

        try write(#"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>
        """#, to: "[Content_Types].xml")

        try write(#"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """#, to: "_rels/.rels")

        try write(#"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        </Relationships>
        """#, to: "word/_rels/document.xml.rels")

        try write(documentXML, to: "word/document.xml")

        let outputURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("loudfail-fixture-\(UUID().uuidString).docx")
        let archive = try Archive(url: outputURL, accessMode: .create)
        let normalizedStaging = stagingDir.resolvingSymlinksInPath().path
        let basePathLen = normalizedStaging.count + 1
        let enumerator = FileManager.default.enumerator(at: stagingDir, includingPropertiesForKeys: [.isDirectoryKey])!
        for case let fileURL as URL in enumerator {
            let isDir = (try? fileURL.resourceValues(forKeys: [.isDirectoryKey]))?.isDirectory ?? false
            if isDir { continue }
            let normalizedFile = fileURL.resolvingSymlinksInPath().path
            let entryName = String(normalizedFile.dropFirst(basePathLen))
            try archive.addEntry(with: entryName, fileURL: fileURL, compressionMethod: .deflate)
        }
        return outputURL
    }
}
