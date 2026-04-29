import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Mutation surface safety tests for [PsychQuant/ooxml-swift#5](https://github.com/PsychQuant/ooxml-swift/issues/5).
///
/// Covers three sub-findings from the post-#56 verification bundle:
/// - F5 — `Hyperlink.text` setter deprecation (destructive; loses RunProperties + rawElements)
/// - F6 — `position: Int = 0` default → `Int? = nil` cascade across 13 typed-child sites
/// - F13 — `Run.toXML()` auto-emits `xml:space="preserve"` for semantically-significant whitespace
///
/// Each test maps to one `Scenario:` block in the spec
/// `openspec/changes/mutation-surface-fix/specs/ooxml-mutation-surface-safety/spec.md`.
final class Issue5MutationSurfaceTests: XCTestCase {

    // MARK: - F5: Hyperlink.text setter deprecation

    /// Spec: setter call site emits compile-time warning; runtime behaviour
    /// unchanged from v0.21.4 baseline (single Run carrying new text).
    func testHyperlinkTextSetterRuntimeBehaviorUnchanged() {
        var props1 = RunProperties()
        props1.bold = true
        var props2 = RunProperties()
        props2.italic = true
        var hyperlink = Hyperlink(
            id: "h1",
            runs: [
                Run(text: "before", properties: props1),
                Run(text: "after", properties: props2),
            ],
            relationshipId: "rId1"
        )
        XCTAssertEqual(hyperlink.runs.count, 2)
        // Trigger the (deprecated) setter — runtime semantics: collapse to
        // one Run with new text. Per D1, the deprecation does NOT change
        // behaviour for one minor.
        hyperlink.text = "replaced"
        XCTAssertEqual(hyperlink.runs.count, 1,
            "setter must collapse runs to a single Run (legacy behaviour preserved)")
        XCTAssertEqual(hyperlink.runs[0].text, "replaced",
            "single replacement Run carries the new text")
    }

    // MARK: - F6: Position type cascade (Int? = nil)

    /// Spec: default-constructed Hyperlink has `position == nil`.
    func testDefaultConstructedHyperlinkHasNilPosition() {
        let h = Hyperlink(id: "h1", runs: [Run(text: "x")], relationshipId: "rId1")
        XCTAssertNil(h.position,
            "default Hyperlink position should be nil (was Int = 0 pre-fix)")
    }

    /// Spec: default-constructed Run has `position == nil`.
    func testDefaultConstructedRunHasNilPosition() {
        let r = Run(text: "x")
        XCTAssertNil(r.position,
            "default Run position should be nil (was Int = 0 pre-fix)")
    }

    /// Spec: Reader-loaded typed children carry explicit position values.
    func testReaderLoadedTypedChildHasExplicitPosition() throws {
        let url = try buildSimpleParagraphFixture()
        defer { try? FileManager.default.removeItem(at: url) }

        let doc = try DocxReader.read(from: url)
        guard let para = firstParagraph(in: doc) else {
            XCTFail("Fixture missing paragraph")
            return
        }
        XCTAssertEqual(para.runs.count, 3, "fixture has 3 runs")
        for (i, run) in para.runs.enumerated() {
            XCTAssertNotNil(run.position,
                "Reader-assigned position for run \(i) should not be nil — got \(String(describing: run.position))")
        }
    }

    // MARK: - F6: Paragraph emit partition + max+1 heuristic

    /// Spec example: source [pos 1, pos 2, pos 3] + 1 append-mode run → appendee at pos 4.
    func testAppendRunLandsAfterSourceChildrenInSourceLoadedParagraph() throws {
        let url = try buildSimpleParagraphFixture()
        defer { try? FileManager.default.removeItem(at: url) }

        var doc = try DocxReader.read(from: url)
        guard var para = firstParagraph(in: doc) else {
            XCTFail("Fixture missing paragraph")
            return
        }
        // All Reader-loaded runs already have explicit positions (1/2/3).
        // Append a new API-built Run (no explicit position → nil).
        para.runs.append(Run(text: "appendee"))
        let xml = para.toXML()
        // The appendee text should appear AFTER the source runs in the
        // serialized XML — search for the text-position ordering as a proxy.
        guard let firstIdx = xml.range(of: "first")?.lowerBound,
              let appendeeIdx = xml.range(of: "appendee")?.lowerBound else {
            XCTFail("expected text not found in emit: \(xml)")
            return
        }
        XCTAssertLessThan(firstIdx, appendeeIdx,
            "appendee must emit after source runs — got XML: \(xml)")
        doc.body.children.removeAll() // silence unused write
    }

    /// Spec example: all-nil collection emits in array order (1/2/3).
    func testAllNilCollectionEmitsInArrayOrder() {
        var para = Paragraph()
        para.runs = [
            Run(text: "alpha"),
            Run(text: "beta"),
            Run(text: "gamma"),
        ]
        let xml = para.toXML()
        guard let alphaIdx = xml.range(of: "alpha")?.lowerBound,
              let betaIdx = xml.range(of: "beta")?.lowerBound,
              let gammaIdx = xml.range(of: "gamma")?.lowerBound else {
            XCTFail("expected text not found: \(xml)")
            return
        }
        XCTAssertLessThan(alphaIdx, betaIdx)
        XCTAssertLessThan(betaIdx, gammaIdx)
    }

    /// Spec example: sparse [pos 1, pos 100] + 1 append-mode → appendee at 101.
    func testSparseExplicitPositionsAppendCorrectly() {
        var para = Paragraph()
        var r1 = Run(text: "low")
        r1.position = 1
        var r2 = Run(text: "high")
        r2.position = 100
        let r3 = Run(text: "appendee") // position == nil
        para.runs = [r1, r2, r3]
        let xml = para.toXML()
        guard let lowIdx = xml.range(of: "low")?.lowerBound,
              let highIdx = xml.range(of: "high")?.lowerBound,
              let appIdx = xml.range(of: "appendee")?.lowerBound else {
            XCTFail("expected text not found: \(xml)")
            return
        }
        XCTAssertLessThan(lowIdx, highIdx)
        XCTAssertLessThan(highIdx, appIdx,
            "appendee must emit after the highest explicit position — got XML: \(xml)")
    }

    // MARK: - F13: Run.toXML xml:space="preserve" autosense

    /// Spec example table row: leading whitespace → preserve flag.
    func testRunWithLeadingWhitespaceEmitsPreserveFlag() {
        let xml = Run(text: " leading").toXML()
        XCTAssertTrue(xml.contains(#"xml:space="preserve""#),
            "leading whitespace must trigger preserve flag — got: \(xml)")
    }

    /// Spec example table row: trailing whitespace → preserve flag.
    func testRunWithTrailingWhitespaceEmitsPreserveFlag() {
        let xml = Run(text: "trailing ").toXML()
        XCTAssertTrue(xml.contains(#"xml:space="preserve""#),
            "trailing whitespace must trigger preserve flag — got: \(xml)")
    }

    /// Spec example table row: 2 consecutive internal spaces → preserve flag.
    func testRunWithConsecutiveInternalWhitespaceEmitsPreserveFlag() {
        let xml = Run(text: "two  spaces").toXML()
        XCTAssertTrue(xml.contains(#"xml:space="preserve""#),
            "double-space must trigger preserve flag — got: \(xml)")
    }

    /// Spec example table row: single internal space → no flag (XML normalises).
    func testRunWithSingleInternalWhitespaceDoesNotEmitPreserveFlag() {
        let xml = Run(text: "hello world").toXML()
        XCTAssertFalse(xml.contains(#"xml:space="preserve""#),
            "single internal space must NOT trigger preserve flag — got: \(xml)")
    }

    /// Spec example table row: empty text → no flag.
    func testRunWithEmptyTextDoesNotEmitPreserveFlag() {
        let xml = Run(text: "").toXML()
        XCTAssertFalse(xml.contains(#"xml:space="preserve""#),
            "empty text must NOT trigger preserve flag — got: \(xml)")
    }

    /// Spec example table rows: tab + newline patterns.
    func testRunWithTabAndNewlineWhitespacePatterns() {
        XCTAssertFalse(Run(text: "a\tb").toXML().contains(#"xml:space="preserve""#),
            "single internal tab → no flag")
        XCTAssertTrue(Run(text: "a\t\tb").toXML().contains(#"xml:space="preserve""#),
            "consecutive tabs → flag")
        XCTAssertTrue(Run(text: "\nleading-newline").toXML().contains(#"xml:space="preserve""#),
            "leading newline → flag")
    }

    // MARK: - Helpers

    private func firstParagraph(in doc: WordDocument) -> Paragraph? {
        for child in doc.body.children {
            if case let .paragraph(p) = child { return p }
        }
        return nil
    }

    /// Fixture: one paragraph with three runs ("first" / "second" / "third").
    private func buildSimpleParagraphFixture() throws -> URL {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:r><w:t>first</w:t></w:r>
              <w:r><w:t>second</w:t></w:r>
              <w:r><w:t>third</w:t></w:r>
            </w:p>
          </w:body>
        </w:document>
        """
        return try buildMinimalDocx(documentXML: documentXML)
    }

    /// Pattern lifted from `Issue6RoundtripLoudFailTests.buildMinimalDocx`.
    private func buildMinimalDocx(documentXML: String) throws -> URL {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("mutation-fixture-\(UUID().uuidString)")
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
            .appendingPathComponent("mutation-fixture-\(UUID().uuidString).docx")
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
