import XCTest
@testable import OOXMLSwift

/// v0.19.4+ regression tests for the round 3 verify findings on
/// PsychQuant/che-word-mcp#56 (https://github.com/PsychQuant/che-word-mcp/issues/56#issuecomment-4321007538).
///
/// R3 found 6 new P0 regressions caused by the R2 (v0.19.3) fixes themselves.
/// Each test in this file targets one R3-NEW finding and follows the per-task
/// verify discipline established in the spectra change
/// `che-word-mcp-issue-56-r3-stack-completion`:
///
/// - R3-NEW-1: Hyperlink mutation API SHALL round-trip on source-loaded hyperlinks
/// - R3-NEW-2: ContentControl SHALL expose a position: Int field and emit in source order
/// - R3-NEW-3: insertComment SHALL emit anchor markers on source paragraphs with existing comment markers
/// - R3-NEW-4: Mixed-content revision wrappers SHALL populate both raw and typed representations
/// - R3-NEW-5: nextBookmarkId calibration SHALL scan all bookmark-bearing document parts
/// - R3-NEW-6: Direct-emit XML attribute values SHALL be escaped to prevent injection
final class Issue56R3StackTests: XCTestCase {

    // MARK: - R3-NEW-1: Hyperlink mutation round-trip on source-loaded hyperlinks

    /// Spec scenario "replaceText on source-loaded hyperlink emits new text":
    /// load doc with `<w:hyperlink r:id="rId1"><w:r><w:t>old</w:t></w:r></w:hyperlink>`,
    /// call `replaceText("old", with: "new")`, re-emit, expect "new" in the
    /// hyperlink and no "old".
    ///
    /// Pre-v0.19.4: Reader populates BOTH `runs` AND `children`. Writer's
    /// "if !children.isEmpty use children else runs" priority makes `children`
    /// the source of truth, but `replaceText` only mutates `runs`. The saved
    /// XML re-emits the original "old" — silent edit-failure regression.
    func testReplaceText_OnSourceLoadedHyperlink_EmitsNewText() throws {
        let docxURL = try Self.buildHyperlinkSourceFixture(text: "old")
        defer { try? FileManager.default.removeItem(at: docxURL) }

        var doc = try DocxReader.read(from: docxURL)
        defer { doc.close() }

        let count = try doc.replaceText(find: "old", with: "new")
        XCTAssertEqual(count, 1, "replaceText should report 1 substitution")

        // Locate the paragraph carrying the hyperlink and re-emit.
        guard case .paragraph(let para) = doc.body.children[0] else {
            XCTFail("Expected paragraph at body[0]")
            return
        }
        let xml = para.toXML()

        // Bug: writer prefers `children` (still says "old"); fix: writer prefers `runs`.
        XCTAssertTrue(
            xml.contains(">new<"),
            "Re-emitted hyperlink XML must contain 'new' (the replacement). Output:\n\(xml)"
        )
        XCTAssertFalse(
            xml.contains(">old<"),
            "Re-emitted hyperlink XML must NOT contain 'old' (the original). Output:\n\(xml)"
        )
    }

    /// Spec scenario "updateHyperlink text setter on source-loaded hyperlink emits new text":
    /// same setup, but mutate via `updateHyperlink(hyperlinkId:text:)` instead of
    /// `replaceText`. Both APIs feed through `Hyperlink.text` setter / `runs`
    /// mutation, so both fail under the v0.19.3 priority order.
    func testUpdateHyperlinkText_OnSourceLoadedHyperlink_EmitsNewText() throws {
        let docxURL = try Self.buildHyperlinkSourceFixture(text: "old")
        defer { try? FileManager.default.removeItem(at: docxURL) }

        var doc = try DocxReader.read(from: docxURL)
        defer { doc.close() }

        // Reader assigns id from r:id + position; for our single hyperlink it's "rId1@0".
        guard case .paragraph(let para) = doc.body.children[0],
              let firstHyperlink = para.hyperlinks.first else {
            XCTFail("Expected source-loaded hyperlink in body[0]")
            return
        }
        try doc.updateHyperlink(hyperlinkId: firstHyperlink.id, text: "Updated")

        guard case .paragraph(let updatedPara) = doc.body.children[0] else {
            XCTFail("Expected paragraph at body[0] after update")
            return
        }
        let xml = updatedPara.toXML()

        XCTAssertTrue(
            xml.contains(">Updated<"),
            "Re-emitted hyperlink XML must contain 'Updated' after updateHyperlink. Output:\n\(xml)"
        )
        XCTAssertFalse(
            xml.contains(">old<"),
            "Re-emitted hyperlink XML must NOT contain 'old' after updateHyperlink. Output:\n\(xml)"
        )
    }

    // MARK: - R3-NEW-2: ContentControl SHALL expose a position: Int field and emit in source order

    /// Spec scenario "SDT between runs round-trips at its source position":
    /// load doc with `<w:p><w:r><w:t>A</w:t></w:r><w:sdt>X</w:sdt><w:r><w:t>B</w:t></w:r></w:p>`,
    /// re-emit, expect output order A, X (inside sdt), B.
    ///
    /// Pre-v0.19.4: ContentControl has no `position` field. DocxReader.parseParagraph
    /// appends source-loaded SDTs to `paragraph.contentControls` without
    /// passing `childPosition`, and `Paragraph.toXMLSortedByPosition` emits
    /// `contentControls` unconditionally in the post-content section after
    /// the sorted positioned-entry list. Result: SDT moves to end (A→B→SDT).
    func testParagraphLevelSDT_BetweenRuns_RoundTripsAtSourcePosition() throws {
        let docxURL = try Self.buildSDTBetweenRunsFixture()
        defer { try? FileManager.default.removeItem(at: docxURL) }

        var doc = try DocxReader.read(from: docxURL)
        defer { doc.close() }

        guard case .paragraph(let para) = doc.body.children[0] else {
            XCTFail("Expected paragraph at body[0]")
            return
        }
        let xml = para.toXML()

        guard let aPos = xml.range(of: ">A<")?.lowerBound,
              let sdtPos = xml.range(of: "<w:sdt")?.lowerBound,
              let bPos = xml.range(of: ">B<")?.lowerBound else {
            XCTFail("Output missing A / sdt / B markers. Output:\n\(xml)")
            return
        }
        XCTAssertLessThan(aPos, sdtPos,
                          "Run 'A' must precede <w:sdt> in source order. Output:\n\(xml)")
        XCTAssertLessThan(sdtPos, bPos,
                          "<w:sdt> must precede run 'B' in source order. Pre-fix the SDT moved to end (A→B→SDT). Output:\n\(xml)")
    }

    // MARK: - Fixture builders

    /// Builds a minimal valid `.docx` whose body contains exactly one paragraph
    /// with one source-loaded `<w:hyperlink>` wrapping the given text. Returns
    /// the path to the assembled `.docx` file (caller cleans up).
    private static func buildHyperlinkSourceFixture(text: String) throws -> URL {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("r3-hl-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(
            at: stagingDir.appendingPathComponent("_rels"),
            withIntermediateDirectories: true
        )
        try FileManager.default.createDirectory(
            at: stagingDir.appendingPathComponent("word/_rels"),
            withIntermediateDirectories: true
        )
        defer { try? FileManager.default.removeItem(at: stagingDir) }

        let contentTypes = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>
        """
        try contentTypes.write(
            to: stagingDir.appendingPathComponent("[Content_Types].xml"),
            atomically: true, encoding: .utf8
        )

        let rootRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """
        try rootRels.write(
            to: stagingDir.appendingPathComponent("_rels/.rels"),
            atomically: true, encoding: .utf8
        )

        let documentRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
        </Relationships>
        """
        try documentRels.write(
            to: stagingDir.appendingPathComponent("word/_rels/document.xml.rels"),
            atomically: true, encoding: .utf8
        )

        let documentXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <w:body>
                <w:p><w:hyperlink r:id="rId1"><w:r><w:t>\(text)</w:t></w:r></w:hyperlink></w:p>
            </w:body>
        </w:document>
        """
        try documentXml.write(
            to: stagingDir.appendingPathComponent("word/document.xml"),
            atomically: true, encoding: .utf8
        )

        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("r3-hl-fixture-\(UUID().uuidString).docx")
        try ZipHelper.zip(stagingDir, to: docxURL)
        return docxURL
    }

    /// Builds a minimal valid `.docx` whose body contains exactly one paragraph
    /// `<w:p><w:r><w:t>A</w:t></w:r><w:sdt><w:sdtContent><w:r><w:t>X</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t>B</w:t></w:r></w:p>`.
    /// Used by R3-NEW-2 source-position round-trip tests.
    private static func buildSDTBetweenRunsFixture() throws -> URL {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("r3-sdt-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(
            at: stagingDir.appendingPathComponent("_rels"),
            withIntermediateDirectories: true
        )
        try FileManager.default.createDirectory(
            at: stagingDir.appendingPathComponent("word/_rels"),
            withIntermediateDirectories: true
        )
        defer { try? FileManager.default.removeItem(at: stagingDir) }

        let contentTypes = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>
        """
        try contentTypes.write(
            to: stagingDir.appendingPathComponent("[Content_Types].xml"),
            atomically: true, encoding: .utf8
        )

        let rootRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """
        try rootRels.write(
            to: stagingDir.appendingPathComponent("_rels/.rels"),
            atomically: true, encoding: .utf8
        )

        let emptyDocRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        </Relationships>
        """
        try emptyDocRels.write(
            to: stagingDir.appendingPathComponent("word/_rels/document.xml.rels"),
            atomically: true, encoding: .utf8
        )

        let documentXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p><w:r><w:t>A</w:t></w:r><w:sdt><w:sdtPr><w:tag w:val="t1"/></w:sdtPr><w:sdtContent><w:r><w:t>X</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t>B</w:t></w:r></w:p>
            </w:body>
        </w:document>
        """
        try documentXml.write(
            to: stagingDir.appendingPathComponent("word/document.xml"),
            atomically: true, encoding: .utf8
        )

        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("r3-sdt-fixture-\(UUID().uuidString).docx")
        try ZipHelper.zip(stagingDir, to: docxURL)
        return docxURL
    }
}
