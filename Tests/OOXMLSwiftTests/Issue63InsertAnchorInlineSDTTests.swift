import XCTest
@testable import OOXMLSwift

/// PsychQuant/che-word-mcp#63 follow-up (verify F1 P1) — `findBodyChildContainingText`
/// (used by `InsertLocation.afterText` / `.beforeText` resolution) only flattens
/// `para.runs`, NOT `paragraph.contentControls` / `hyperlinks` / `fieldSimples` /
/// `alternateContents`. So `insert_image_from_path(after_text: "[tab:foo]")` against
/// an inline-SDT-wrapped anchor still throws `textNotFound` even after v3.14.4 fix.
///
/// Mirror surface coverage of `Document.replaceInParagraphSurfaces` so the read
/// (lookup) path matches the write (replace) path.
final class Issue63InsertAnchorInlineSDTTests: XCTestCase {

    private func injectDocumentXML(_ documentXML: String) throws -> URL {
        var base = WordDocument()
        base.body.children.append(.paragraph(Paragraph(runs: [Run(text: "x")])))
        let baseURL = URL(fileURLWithPath: NSTemporaryDirectory())
            .appendingPathComponent("anchor63_base_\(UUID().uuidString).docx")
        try DocxWriter.write(base, to: baseURL)
        defer { try? FileManager.default.removeItem(at: baseURL) }

        let outURL = URL(fileURLWithPath: NSTemporaryDirectory())
            .appendingPathComponent("anchor63_inj_\(UUID().uuidString).docx")
        try FileManager.default.copyItem(at: baseURL, to: outURL)

        let workDir = URL(fileURLWithPath: NSTemporaryDirectory())
            .appendingPathComponent("anchor63_work_\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: workDir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: workDir) }
        let docXMLDir = workDir.appendingPathComponent("word")
        try FileManager.default.createDirectory(at: docXMLDir, withIntermediateDirectories: true)
        try documentXML.write(to: docXMLDir.appendingPathComponent("document.xml"),
                              atomically: true, encoding: .utf8)

        let proc = Process()
        proc.executableURL = URL(fileURLWithPath: "/usr/bin/zip")
        proc.currentDirectoryURL = workDir
        proc.arguments = [outURL.path, "word/document.xml"]
        proc.standardOutput = Pipe()
        proc.standardError = Pipe()
        try proc.run()
        proc.waitUntilExit()

        return outURL
    }

    private func fldSimpleAnchorFixture() -> String {
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:r><w:t xml:space="preserve">prefix </w:t></w:r>
              <w:fldSimple w:instr=" REF foo \\h ">
                <w:r><w:t>[tab:foo]</w:t></w:r>
              </w:fldSimple>
              <w:r><w:t xml:space="preserve"> further description follows</w:t></w:r>
            </w:p>
          </w:body>
        </w:document>
        """
    }

    private func hyperlinkAnchorFixture() -> String {
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:body>
            <w:p>
              <w:r><w:t xml:space="preserve">prefix </w:t></w:r>
              <w:hyperlink w:anchor="foo">
                <w:r><w:t>[tab:foo]</w:t></w:r>
              </w:hyperlink>
              <w:r><w:t xml:space="preserve"> further description follows</w:t></w:r>
            </w:p>
          </w:body>
        </w:document>
        """
    }

    private func inlineSDTAnchorFixture() -> String {
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:r><w:t xml:space="preserve">prefix </w:t></w:r>
              <w:sdt>
                <w:sdtPr><w:tag w:val="ref"/><w:alias w:val="cross-ref"/></w:sdtPr>
                <w:sdtContent>
                  <w:r><w:t>[tab:foo]</w:t></w:r>
                </w:sdtContent>
              </w:sdt>
              <w:r><w:t xml:space="preserve"> further description follows</w:t></w:r>
            </w:p>
          </w:body>
        </w:document>
        """
    }

    /// Differential test: insertParagraph with `.afterText("[tab:foo]", instance: 1)`
    /// SHOULD succeed against any of the 3 wrapper kinds. Pre-fix only fldSimple
    /// and hyperlink work (because... wait, actually NEITHER works pre-fix because
    /// `findBodyChildContainingText` only inspects `para.runs`). Post-fix all 3 work.
    func testInsertAfterTextResolvesAcrossAllInlineWrappers() throws {
        let cases: [(name: String, xml: String)] = [
            ("fldSimple", fldSimpleAnchorFixture()),
            ("hyperlink", hyperlinkAnchorFixture()),
            ("inlineSDT", inlineSDTAnchorFixture()),
        ]
        for (name, xml) in cases {
            let url = try injectDocumentXML(xml)
            defer { try? FileManager.default.removeItem(at: url) }
            var doc = try DocxReader.read(from: url)

            // Try inserting a marker paragraph after the bracketed anchor.
            do {
                try doc.insertParagraph(
                    Paragraph(runs: [Run(text: "MARKER")]),
                    at: .afterText("[tab:foo]", instance: 1)
                )
            } catch InsertLocationError.textNotFound(let needle, let inst) {
                XCTFail("\(name) wrapper: anchor [tab:foo] not found by findBodyChildContainingText (needle=\(needle), instance=\(inst))")
                continue
            }

            // Verify MARKER was inserted at correct position (after the anchor para).
            // Body.children should have 2 paragraphs: original + MARKER.
            XCTAssertEqual(doc.body.children.count, 2,
                           "\(name) wrapper: expected 2 body children after insert")
            if case .paragraph(let markerPara) = doc.body.children[1] {
                XCTAssertEqual(markerPara.runs.first?.text, "MARKER",
                               "\(name) wrapper: MARKER paragraph at expected position")
            } else {
                XCTFail("\(name) wrapper: expected MARKER paragraph at index 1")
            }
        }
    }

    /// `.beforeText` should also work — symmetric to `.afterText`.
    func testInsertBeforeTextResolvesInsideInlineSDT() throws {
        let url = try injectDocumentXML(inlineSDTAnchorFixture())
        defer { try? FileManager.default.removeItem(at: url) }
        var doc = try DocxReader.read(from: url)

        try doc.insertParagraph(
            Paragraph(runs: [Run(text: "BEFORE")]),
            at: .beforeText("[tab:foo]", instance: 1)
        )
        XCTAssertEqual(doc.body.children.count, 2)
        if case .paragraph(let p0) = doc.body.children[0] {
            XCTAssertEqual(p0.runs.first?.text, "BEFORE",
                           "BEFORE paragraph at index 0 (before the anchor)")
        } else {
            XCTFail("expected BEFORE paragraph at index 0")
        }
    }

    /// insertImage variant should also use the new lookup. Smoke test by checking
    /// the image insertion doesn't throw textNotFound.
    func testInsertImageAfterTextResolvesInsideInlineSDT() throws {
        let url = try injectDocumentXML(inlineSDTAnchorFixture())
        defer { try? FileManager.default.removeItem(at: url) }
        var doc = try DocxReader.read(from: url)

        // Create a tiny PNG fixture (1x1 transparent pixel).
        let pngData = Data([
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
            0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41,
            0x54, 0x78, 0x9C, 0x62, 0x00, 0x01, 0x00, 0x00,
            0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
            0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
            0x42, 0x60, 0x82
        ])
        let pngURL = URL(fileURLWithPath: NSTemporaryDirectory())
            .appendingPathComponent("anchor63_pixel_\(UUID().uuidString).png")
        try pngData.write(to: pngURL)
        defer { try? FileManager.default.removeItem(at: pngURL) }

        XCTAssertNoThrow(try doc.insertImage(
            path: pngURL.path, widthPx: 100, heightPx: 100,
            at: .afterText("[tab:foo]", instance: 1)
        ), "insertImage with after_text=[tab:foo] should resolve inside inline SDT")
    }
}
