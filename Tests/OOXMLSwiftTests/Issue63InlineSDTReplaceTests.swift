import XCTest
@testable import OOXMLSwift

/// PsychQuant/che-word-mcp#63 — `replace_text` silently 0-matches on text wrapped
/// in inline `<w:sdt>` content controls (root cause: `Document.replaceInParagraphSurfaces`
/// covers `paragraph.runs` / `hyperlinks` / `fieldSimples` / `alternateContents` but
/// not `paragraph.contentControls`).
///
/// Differential test surface: build hand-crafted `document.xml` fixtures wrapping the
/// same payload (`[tab:foo]`) in 4 candidate wrappers (fldChar / fldSimple / hyperlink /
/// inlineSDT) and assert all 4 reach 1 match. Pre-fix only inlineSDT is 0; post-fix all 1.
final class Issue63InlineSDTReplaceTests: XCTestCase {

    // MARK: - Fixture builders

    private func injectDocumentXML(_ documentXML: String) throws -> URL {
        var base = WordDocument()
        base.body.children.append(.paragraph(Paragraph(runs: [Run(text: "x")])))
        let baseURL = URL(fileURLWithPath: NSTemporaryDirectory())
            .appendingPathComponent("issue63_base_\(UUID().uuidString).docx")
        try DocxWriter.write(base, to: baseURL)
        defer { try? FileManager.default.removeItem(at: baseURL) }

        let outURL = URL(fileURLWithPath: NSTemporaryDirectory())
            .appendingPathComponent("issue63_inj_\(UUID().uuidString).docx")
        try FileManager.default.copyItem(at: baseURL, to: outURL)

        let workDir = URL(fileURLWithPath: NSTemporaryDirectory())
            .appendingPathComponent("issue63_work_\(UUID().uuidString)")
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

    private func fldCharFixture() -> String {
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:r><w:t xml:space="preserve">prefix </w:t></w:r>
              <w:r><w:fldChar w:fldCharType="begin"/></w:r>
              <w:r><w:instrText xml:space="preserve"> REF foo \\h </w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="separate"/></w:r>
              <w:r><w:t>[tab:foo]</w:t></w:r>
              <w:r><w:fldChar w:fldCharType="end"/></w:r>
              <w:r><w:t xml:space="preserve"> further description follows</w:t></w:r>
            </w:p>
          </w:body>
        </w:document>
        """
    }

    private func fldSimpleFixture() -> String {
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

    private func hyperlinkFixture() -> String {
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

    private func inlineSDTFixture() -> String {
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

    /// Outer SDT wraps inner SDT wrapping the bracketed text. Closes the
    /// "nested SDT recursion" requirement.
    private func nestedInlineSDTFixture() -> String {
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:sdt>
                <w:sdtPr><w:tag w:val="outer"/></w:sdtPr>
                <w:sdtContent>
                  <w:sdt>
                    <w:sdtPr><w:tag w:val="inner"/></w:sdtPr>
                    <w:sdtContent>
                      <w:r><w:t>[tab:foo]</w:t></w:r>
                    </w:sdtContent>
                  </w:sdt>
                </w:sdtContent>
              </w:sdt>
            </w:p>
          </w:body>
        </w:document>
        """
    }

    // MARK: - Differential test: all 4 wrappers should match (post-fix)

    func testReplaceTextMatchesAcrossAllInlineWrappers() throws {
        let cases: [(name: String, xml: String)] = [
            ("fldChar", fldCharFixture()),
            ("fldSimple", fldSimpleFixture()),
            ("hyperlink", hyperlinkFixture()),
            ("inlineSDT", inlineSDTFixture()),
        ]
        for (name, xml) in cases {
            let url = try injectDocumentXML(xml)
            defer { try? FileManager.default.removeItem(at: url) }
            for needle in ["[tab:foo]", "tab:foo", "[tab:foo", "tab:foo]"] {
                var doc = try DocxReader.read(from: url)
                let n = try doc.replaceText(find: needle, with: "REPLACED")
                XCTAssertEqual(n, 1,
                    "\(name) wrapper: needle \(needle.debugDescription) expected 1 match, got \(n)")
            }
        }
    }

    // MARK: - Nested SDT recursion

    func testReplaceTextDescendsIntoNestedInlineSDT() throws {
        let url = try injectDocumentXML(nestedInlineSDTFixture())
        defer { try? FileManager.default.removeItem(at: url) }
        var doc = try DocxReader.read(from: url)
        let n = try doc.replaceText(find: "[tab:foo]", with: "REPLACED")
        XCTAssertEqual(n, 1, "nested SDT replacement count")
    }

    // MARK: - Round-trip: SDT wrapper preserved after replace

    func testInlineSDTReplaceRoundTripPreservesWrapper() throws {
        let url = try injectDocumentXML(inlineSDTFixture())
        defer { try? FileManager.default.removeItem(at: url) }
        var doc = try DocxReader.read(from: url)
        _ = try doc.replaceText(find: "[tab:foo]", with: "REPLACED")

        let savedURL = URL(fileURLWithPath: NSTemporaryDirectory())
            .appendingPathComponent("issue63_rt_\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: savedURL)
        defer { try? FileManager.default.removeItem(at: savedURL) }

        let reread = try DocxReader.read(from: savedURL)
        guard case .paragraph(let p) = reread.body.children[0] else {
            XCTFail("expected paragraph"); return
        }
        XCTAssertEqual(p.contentControls.count, 1, "SDT wrapper survived round-trip")
        XCTAssertEqual(p.contentControls.first?.sdt.tag, "ref",
                       "SDT tag preserved")
        // Content should now hold REPLACED, not [tab:foo].
        let cc = try XCTUnwrap(p.contentControls.first)
        XCTAssertTrue(cc.content.contains("REPLACED"),
                      "replaced text written into SDT content (got: \(cc.content))")
        XCTAssertFalse(cc.content.contains("[tab:foo]"),
                       "original bracketed text removed from SDT content")
    }
}
