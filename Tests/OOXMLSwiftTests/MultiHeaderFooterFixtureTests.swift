import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Phase 1 task 1.9 of the `che-word-mcp-true-byte-preservation` Spectra change.
///
/// Builds a 22-part .docx fixture in-process (no committed binary blob) covering:
/// - 6 default headers (header1.xml..header6.xml) each with a unique
///   recognizable text + watermark VML shape
///   `<v:shape id="PowerPlusWaterMarkObjectN" o:spt="136">` `<v:textpath string="浮水印N">`
/// - 4 default footers (footer1.xml..footer4.xml) where footer3.xml contains
///   the three-segment `<w:fldChar>` + `<w:instrText>PAGE</w:instrText>` +
///   `<w:fldChar>` field pattern that #33 reported as undetected by
///   `has_page_number`
/// - 1 `<w15:person>` with full presenceInfo:
///     userId="S::test@example.com::aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"
///     providerId="AD"
/// - fontTable with 13 entries (Calibri/Calibri Light/TNR/DFKai-SB/華康中楷體/
///   PMingLiU/Microsoft JhengHei/Cambria Math/Symbol/Wingdings/Courier New/
///   Cambria/Malgun Gothic)
/// - theme1.xml with custom name `MultiHeaderFooterFixtureTheme`
///
/// Validates the v0.13.0 byte-preservation contract end-to-end:
/// - 6 distinct header fileNames (no collapse to header1.xml)
/// - No-op round-trip preserves every typed-managed part byte-for-byte
/// - Modifying only one header preserves the other 5
/// - Modifying styles preserves headers/footers/fontTable/theme
/// - markPartDirty + direct write preserves all other parts
final class MultiHeaderFooterFixtureTests: XCTestCase {

    // MARK: - Fixture builder

    static let fontEntries = [
        "Calibri", "Calibri Light", "Times New Roman", "DFKai-SB", "華康中楷體",
        "PMingLiU", "Microsoft JhengHei", "Cambria Math", "Symbol", "Wingdings",
        "Courier New", "Cambria", "Malgun Gothic"
    ]

    /// Build the fixture .docx in /tmp. Returns the .docx URL. Caller is
    /// responsible for cleanup (or rely on /tmp eviction at reboot).
    static func buildFixture() throws -> URL {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("multi-hf-staging-\(UUID().uuidString)")
        defer { try? FileManager.default.removeItem(at: stagingDir) }

        try FileManager.default.createDirectory(at: stagingDir, withIntermediateDirectories: true)

        func write(_ content: String, to relativePath: String) throws {
            let url = stagingDir.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(),
                withIntermediateDirectories: true
            )
            try content.write(to: url, atomically: true, encoding: .utf8)
        }

        try write(contentTypesXML, to: "[Content_Types].xml")
        try write(packageRelsXML, to: "_rels/.rels")
        try write(documentRelsXML, to: "word/_rels/document.xml.rels")
        try write(documentXML, to: "word/document.xml")
        for n in 1...6 { try write(headerXML(n: n), to: "word/header\(n).xml") }
        for n in 1...4 { try write(footerXML(n: n), to: "word/footer\(n).xml") }
        try write(stylesXML, to: "word/styles.xml")
        try write(fontTableXML, to: "word/fontTable.xml")
        try write(settingsXML, to: "word/settings.xml")
        try write(themeXML, to: "word/theme/theme1.xml")
        try write(peopleXML, to: "word/people.xml")
        try write(coreXML, to: "docProps/core.xml")
        try write(appXML, to: "docProps/app.xml")

        let outputURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("multi-header-footer-\(UUID().uuidString).docx")
        let archive = try Archive(url: outputURL, accessMode: .create)
        // macOS resolves /var/folders/... to /private/var/folders/... in
        // enumerator URLs but stagingDir.path stays /var/folders/... — use
        // resolvingSymlinksInPath on both sides to compute consistent relative paths.
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

    // MARK: - Tests

    func testFixtureLoadsWithSixDistinctHeaderFileNames() throws {
        let fixture = try Self.buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }

        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        XCTAssertEqual(doc.headers.count, 6, "Fixture should expose 6 default headers")
        let names = Set(doc.headers.map { $0.fileName })
        XCTAssertEqual(names, Set((1...6).map { "header\($0).xml" }),
                       "All 6 headers must keep distinct fileName via originalFileName preservation")
    }

    func testFixtureLoadsWithFourDistinctFooterFileNames() throws {
        let fixture = try Self.buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }

        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        XCTAssertEqual(doc.footers.count, 4, "Fixture should expose 4 default footers")
        let names = Set(doc.footers.map { $0.fileName })
        XCTAssertEqual(names, Set((1...4).map { "footer\($0).xml" }))
    }

    func testNoOpRoundTripPreservesAllZipEntriesByteEqual() throws {
        let fixture = try Self.buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }

        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }
        XCTAssertTrue(doc.modifiedPartsView.isEmpty)

        let dest = FileManager.default.temporaryDirectory
            .appendingPathComponent("multi-hf-roundtrip-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: dest) }
        try DocxWriter.write(doc, to: dest)

        let srcDir = try unzip(fixture)
        let destDir = try unzip(dest)
        defer {
            try? FileManager.default.removeItem(at: srcDir)
            try? FileManager.default.removeItem(at: destDir)
        }

        // Verify every typed-managed part survives byte-equal
        for n in 1...6 { try assertByteEqual(part: "word/header\(n).xml", in: srcDir, dest: destDir) }
        for n in 1...4 { try assertByteEqual(part: "word/footer\(n).xml", in: srcDir, dest: destDir) }
        try assertByteEqual(part: "word/document.xml", in: srcDir, dest: destDir)
        try assertByteEqual(part: "word/styles.xml", in: srcDir, dest: destDir)
        try assertByteEqual(part: "word/fontTable.xml", in: srcDir, dest: destDir)
        try assertByteEqual(part: "word/settings.xml", in: srcDir, dest: destDir)
        try assertByteEqual(part: "word/theme/theme1.xml", in: srcDir, dest: destDir)
        try assertByteEqual(part: "word/people.xml", in: srcDir, dest: destDir)
        try assertByteEqual(part: "docProps/core.xml", in: srcDir, dest: destDir)
        try assertByteEqual(part: "docProps/app.xml", in: srcDir, dest: destDir)
    }

    func testEditingOneHeaderPreservesTheOtherFive() throws {
        let fixture = try Self.buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }

        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }
        guard let header2Id = doc.headers.first(where: { $0.fileName == "header2.xml" })?.id else {
            XCTFail("header2 must be present"); return
        }
        try doc.updateHeader(id: header2Id, text: "Header 2 — UPDATED")
        XCTAssertEqual(doc.modifiedPartsView, ["word/header2.xml"],
                       "Only header2.xml should be dirty after updateHeader")

        let dest = FileManager.default.temporaryDirectory
            .appendingPathComponent("multi-hf-h2edit-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: dest) }
        try DocxWriter.write(doc, to: dest)

        let srcDir = try unzip(fixture)
        let destDir = try unzip(dest)
        defer {
            try? FileManager.default.removeItem(at: srcDir)
            try? FileManager.default.removeItem(at: destDir)
        }

        // header2 changed
        let destHeader2 = try String(contentsOf: destDir.appendingPathComponent("word/header2.xml"), encoding: .utf8)
        XCTAssertTrue(destHeader2.contains("UPDATED"))
        // headers 1, 3, 4, 5, 6 unchanged
        for n in [1, 3, 4, 5, 6] {
            try assertByteEqual(part: "word/header\(n).xml", in: srcDir, dest: destDir)
        }
        // footers + fontTable + theme unchanged
        for n in 1...4 { try assertByteEqual(part: "word/footer\(n).xml", in: srcDir, dest: destDir) }
        try assertByteEqual(part: "word/fontTable.xml", in: srcDir, dest: destDir)
        try assertByteEqual(part: "word/theme/theme1.xml", in: srcDir, dest: destDir)
    }

    func testEditingStylePreservesHeadersFootersFontTableTheme() throws {
        let fixture = try Self.buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }

        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }
        try doc.addStyle(Style(id: "FixtureCustom", name: "Fixture Custom", type: .paragraph))
        XCTAssertEqual(doc.modifiedPartsView, ["word/styles.xml"])

        let dest = FileManager.default.temporaryDirectory
            .appendingPathComponent("multi-hf-styleedit-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: dest) }
        try DocxWriter.write(doc, to: dest)

        let srcDir = try unzip(fixture)
        let destDir = try unzip(dest)
        defer {
            try? FileManager.default.removeItem(at: srcDir)
            try? FileManager.default.removeItem(at: destDir)
        }

        for n in 1...6 { try assertByteEqual(part: "word/header\(n).xml", in: srcDir, dest: destDir) }
        for n in 1...4 { try assertByteEqual(part: "word/footer\(n).xml", in: srcDir, dest: destDir) }
        try assertByteEqual(part: "word/fontTable.xml", in: srcDir, dest: destDir)
        try assertByteEqual(part: "word/theme/theme1.xml", in: srcDir, dest: destDir)
    }

    func testMarkThemeDirtyPlusDirectWritePreservesAllOtherParts() throws {
        let fixture = try Self.buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }

        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }
        guard let tempDir = doc.archiveTempDir else {
            XCTFail("archiveTempDir required for direct write"); return
        }

        // External-writer pattern: edit theme1.xml directly + markPartDirty
        let themeURL = tempDir.appendingPathComponent("word/theme/theme1.xml")
        let newTheme = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="UpdatedTheme">
          <a:themeElements><a:clrScheme name="Office"/><a:fontScheme name="Office"/><a:fmtScheme name="Office"/></a:themeElements>
        </a:theme>
        """
        try newTheme.write(to: themeURL, atomically: true, encoding: .utf8)
        doc.markPartDirty("word/theme/theme1.xml")

        let dest = FileManager.default.temporaryDirectory
            .appendingPathComponent("multi-hf-themeedit-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: dest) }
        try DocxWriter.write(doc, to: dest)

        let srcDir = try unzip(fixture)
        let destDir = try unzip(dest)
        defer {
            try? FileManager.default.removeItem(at: srcDir)
            try? FileManager.default.removeItem(at: destDir)
        }

        // theme actually updated in dest
        let destTheme = try String(contentsOf: destDir.appendingPathComponent("word/theme/theme1.xml"), encoding: .utf8)
        XCTAssertTrue(destTheme.contains("UpdatedTheme"))

        // Everything else preserved — including all 13 fontTable entries
        for n in 1...6 { try assertByteEqual(part: "word/header\(n).xml", in: srcDir, dest: destDir) }
        for n in 1...4 { try assertByteEqual(part: "word/footer\(n).xml", in: srcDir, dest: destDir) }
        try assertByteEqual(part: "word/fontTable.xml", in: srcDir, dest: destDir)
        try assertByteEqual(part: "word/styles.xml", in: srcDir, dest: destDir)
        try assertByteEqual(part: "word/document.xml", in: srcDir, dest: destDir)

        // Verify all 13 fontTable entries actually present
        let destFontTable = try String(contentsOf: destDir.appendingPathComponent("word/fontTable.xml"), encoding: .utf8)
        for fontName in Self.fontEntries {
            XCTAssertTrue(destFontTable.contains("w:name=\"\(fontName)\""),
                          "fontTable must preserve entry for \(fontName)")
        }
    }

    // MARK: - Helpers

    private func unzip(_ docx: URL) throws -> URL {
        let dir = FileManager.default.temporaryDirectory
            .appendingPathComponent("multi-hf-verify-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        try FileManager.default.unzipItem(at: docx, to: dir)
        return dir
    }

    private func assertByteEqual(part: String, in srcDir: URL, dest destDir: URL,
                                 file: StaticString = #file, line: UInt = #line) throws {
        let srcURL = srcDir.appendingPathComponent(part)
        let destURL = destDir.appendingPathComponent(part)
        XCTAssertTrue(FileManager.default.fileExists(atPath: srcURL.path),
                      "\(part) missing from source", file: file, line: line)
        XCTAssertTrue(FileManager.default.fileExists(atPath: destURL.path),
                      "\(part) missing from dest", file: file, line: line)
        let srcBytes = try Data(contentsOf: srcURL)
        let destBytes = try Data(contentsOf: destURL)
        XCTAssertEqual(srcBytes, destBytes,
                       "\(part) must be byte-equal after round-trip",
                       file: file, line: line)
    }
}

// MARK: - XML body templates

private let contentTypesXML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/word/people.xml" ContentType="application/vnd.ms-word.people+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/header3.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/header4.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/header5.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/header6.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/footer3.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/footer4.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""

private let packageRelsXML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""

private let documentRelsXML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2011/relationships/people" Target="people.xml"/>
  <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>
  <Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header3.xml"/>
  <Relationship Id="rId13" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header4.xml"/>
  <Relationship Id="rId14" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header5.xml"/>
  <Relationship Id="rId15" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header6.xml"/>
  <Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
  <Relationship Id="rId21" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/>
  <Relationship Id="rId22" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer3.xml"/>
  <Relationship Id="rId23" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer4.xml"/>
</Relationships>
"""

private let documentXML: String = {
    var xml = """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <w:body>
    """
    for sectionNum in 1...6 {
        let headerRId = 9 + sectionNum
        let footerRId = sectionNum <= 4 ? 19 + sectionNum : 22
        xml += """
          <w:p>
            <w:r><w:t>Section \(sectionNum) body content</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr>
              <w:sectPr>
                <w:headerReference w:type="default" r:id="rId\(headerRId)"/>
                <w:footerReference w:type="default" r:id="rId\(footerRId)"/>
                <w:pgSz w:w="12240" w:h="15840"/>
                <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
              </w:sectPr>
            </w:pPr>
          </w:p>
        """
    }
    xml += """
      </w:body>
    </w:document>
    """
    return xml
}()

private func headerXML(n: Int) -> String {
    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
           xmlns:v="urn:schemas-microsoft-com:vml"
           xmlns:o="urn:schemas-microsoft-com:office:office">
      <w:p>
        <w:r>
          <w:pict>
            <v:shape id="PowerPlusWaterMarkObject\(n)" o:spt="136" type="#_x0000_t136"
                     style="position:absolute;margin-left:0;margin-top:0;width:200pt;height:50pt;rotation:-45;z-index:-251653120">
              <v:textpath style="font-family:'宋体';font-size:36pt" string="浮水印\(n)"/>
            </v:shape>
          </w:pict>
        </w:r>
        <w:r><w:t>Header \(n) — distinct content for section \(n)</w:t></w:r>
      </w:p>
    </w:hdr>
    """
}

private func footerXML(n: Int) -> String {
    let body: String
    if n == 3 {
        body = """
          <w:p>
            <w:r><w:t xml:space="preserve">Page </w:t></w:r>
            <w:r><w:fldChar w:fldCharType="begin"/></w:r>
            <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate"/></w:r>
            <w:r><w:t>1</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
          </w:p>
        """
    } else {
        body = """
          <w:p>
            <w:r><w:t>Footer \(n) — section \(n)</w:t></w:r>
          </w:p>
        """
    }
    return """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    \(body)
    </w:ftr>
    """
}

private let stylesXML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="DFKai-SB" w:hAnsi="Calibri" w:cs="Times New Roman"/></w:rPr></w:rPrDefault>
    <w:pPrDefault/>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>
"""

private let fontTableXML: String = {
    let entries = MultiHeaderFooterFixtureTests.fontEntries
    var xml = """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    """
    for font in entries {
        xml += """
          <w:font w:name="\(font)"><w:family w:val="auto"/></w:font>
        """
    }
    xml += """
    </w:fonts>
    """
    return xml
}()

private let settingsXML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>
"""

private let themeXML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="MultiHeaderFooterFixtureTheme">
  <a:themeElements>
    <a:clrScheme name="Office"/>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface="DFKai-SB"/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface="DFKai-SB"/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office"/>
  </a:themeElements>
</a:theme>
"""

private let peopleXML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w15:person w15:author="Test User">
    <w15:presenceInfo w15:providerId="AD" w15:userId="S::test@example.com::aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"/>
  </w15:person>
</w15:people>
"""

private let coreXML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>Multi-header-footer Fixture</dc:title>
  <dc:creator>che-word-mcp-true-byte-preservation</dc:creator>
</cp:coreProperties>
"""

private let appXML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>che-word-mcp-fixture</Application>
</Properties>
"""
