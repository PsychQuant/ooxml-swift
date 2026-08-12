import Foundation
import ZIPFoundation

/// Phase 0 task 1.8: programmatic builders for the four-fixture
/// **round-trip golden corpus** required by spec `ooxml-tree-io`:
///
/// - `multi-section-thesis.docx` — 3 `<w:sectPr>` blocks with header/footer refs
/// - `vml-rich.docx` — `<w:pict>` plus `<mc:AlternateContent>` with wps/wpg
/// - `cjk-settings.docx` — `word/settings.xml` with East-Asian theme refs
///   and CJK font fallback chain
/// - `comment-anchored.docx` — `<w:commentRangeStart>` / `<w:commentRangeEnd>`
///   in the body plus `word/comments.xml`
///
/// Each builder returns a committed .docx plus the list of OOXML part paths
/// the round-trip test should walk. The test layer
/// in `TreeRoundTripCorpusTests` reads each part via ZIPFoundation, runs
/// `XmlTreeReader.parse → XmlTreeWriter.serialize`, and asserts byte-equal.
///
enum CorpusFixtureBuilder {

    /// One built fixture: file URL plus the OOXML part paths whose XML the
    /// round-trip test should walk.
    struct Fixture {
        let name: String
        let url: URL
        let partsToVerify: [String]
    }

    // MARK: - Public entrypoints

    static func buildMultiSectionThesis() throws -> Fixture {
        try committedFixture(
            name: "multi-section-thesis",
            partsToVerify: [
                "word/document.xml",
                "word/header1.xml", "word/header2.xml",
                "word/footer1.xml", "word/footer2.xml",
                "[Content_Types].xml",
            ]
        )
    }

    static func buildVMLRich() throws -> Fixture {
        try committedFixture(
            name: "vml-rich",
            partsToVerify: ["word/document.xml", "[Content_Types].xml"]
        )
    }

    static func buildCJKSettings() throws -> Fixture {
        try committedFixture(
            name: "cjk-settings",
            partsToVerify: [
                "word/document.xml", "word/settings.xml", "[Content_Types].xml",
            ]
        )
    }

    static func buildCommentAnchored() throws -> Fixture {
        try committedFixture(
            name: "comment-anchored",
            partsToVerify: [
                "word/document.xml", "word/comments.xml", "[Content_Types].xml",
            ]
        )
    }

    private static func committedFixture(
        name: String, partsToVerify: [String]
    ) throws -> Fixture {
        let committedURL = URL(fileURLWithPath: #filePath)
            .deletingLastPathComponent()
            .appendingPathComponent("golden-corpus", isDirectory: true)
            .appendingPathComponent("\(name).docx")
        guard FileManager.default.fileExists(atPath: committedURL.path) else {
            throw NSError(
                domain: "CorpusFixtureBuilder",
                code: 2,
                userInfo: [NSLocalizedDescriptionKey: "Committed fixture missing: \(committedURL.path)"]
            )
        }
        // Existing callers own and delete the returned URL. Hand each one a
        // temporary copy so the checked-in source fixture can never be erased
        // by test cleanup.
        let temporaryURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("\(name)-\(UUID().uuidString).docx")
        try FileManager.default.copyItem(at: committedURL, to: temporaryURL)
        return Fixture(name: name, url: temporaryURL, partsToVerify: partsToVerify)
    }

    /// Maintainer-only deterministic source for regenerating the committed
    /// corpus. Tests consume the checked-in `.docx` files above.
    static func regenerateCommittedFixtures() throws {
        let destination = URL(fileURLWithPath: #filePath)
            .deletingLastPathComponent()
            .appendingPathComponent("golden-corpus", isDirectory: true)
        try FileManager.default.createDirectory(at: destination, withIntermediateDirectories: true)
        let fixtures = try [
            build(
                name: "multi-section-thesis",
                partsToVerify: [],
                extraParts: [
                    "word/document.xml": multiSectionThesisDocumentXML,
                    "word/header1.xml": thesisHeader1XML,
                    "word/header2.xml": thesisHeader2XML,
                    "word/footer1.xml": thesisFooter1XML,
                    "word/footer2.xml": thesisFooter2XML,
                ],
                contentTypesOverride: contentTypesWithThesisParts,
                documentRelsOverride: documentRelsWithThesisParts
            ),
            build(
                name: "vml-rich",
                partsToVerify: [],
                extraParts: ["word/document.xml": vmlRichDocumentXML]
            ),
            build(
                name: "cjk-settings",
                partsToVerify: [],
                extraParts: [
                    "word/document.xml": cjkDocumentXML,
                    "word/settings.xml": cjkSettingsXML,
                ],
                contentTypesOverride: contentTypesWithSettings
            ),
            build(
                name: "comment-anchored",
                partsToVerify: [],
                extraParts: [
                    "word/document.xml": commentAnchoredDocumentXML,
                    "word/comments.xml": commentAnchoredCommentsXML,
                ],
                contentTypesOverride: contentTypesWithComments,
                documentRelsOverride: documentRelsWithComments
            ),
        ]
        defer {
            for fixture in fixtures { try? FileManager.default.removeItem(at: fixture.url) }
        }
        for fixture in fixtures {
            let output = destination.appendingPathComponent("\(fixture.name).docx")
            if FileManager.default.fileExists(atPath: output.path) {
                try FileManager.default.removeItem(at: output)
            }
            try FileManager.default.copyItem(at: fixture.url, to: output)
        }
    }

    // MARK: - Generic build

    private static func build(
        name: String,
        partsToVerify: [String],
        extraParts: [String: String],
        contentTypesOverride: String? = nil,
        documentRelsOverride: String? = nil
    ) throws -> Fixture {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("\(name)-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: stagingDir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: stagingDir) }

        func write(_ content: String, to relativePath: String) throws {
            let url = stagingDir.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(),
                withIntermediateDirectories: true
            )
            try content.write(to: url, atomically: true, encoding: .utf8)
            try FileManager.default.setAttributes(
                [.modificationDate: Date(timeIntervalSince1970: 1_767_225_600)],
                ofItemAtPath: url.path
            )
        }

        try write(contentTypesOverride ?? contentTypesMinimal, to: "[Content_Types].xml")
        try write(packageRelsMinimal, to: "_rels/.rels")
        if extraParts.keys.contains(where: { $0.hasPrefix("word/") }) {
            try write(documentRelsOverride ?? documentRelsMinimal, to: "word/_rels/document.xml.rels")
        }
        for (path, content) in extraParts {
            try write(content, to: path)
        }

        let outputURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("\(name)-\(UUID().uuidString).docx")
        let archive = try Archive(url: outputURL, accessMode: .create)
        let normalizedStaging = stagingDir.resolvingSymlinksInPath().path
        let basePathLen = normalizedStaging.count + 1
        let enumerator = FileManager.default.enumerator(
            at: stagingDir, includingPropertiesForKeys: [.isDirectoryKey])!
        let files = enumerator.compactMap { $0 as? URL }.filter { fileURL in
            let isDir = (try? fileURL.resourceValues(forKeys: [.isDirectoryKey]))?.isDirectory ?? false
            return !isDir
        }.sorted { $0.path < $1.path }
        for fileURL in files {
            let normalizedFile = fileURL.resolvingSymlinksInPath().path
            let entryName = String(normalizedFile.dropFirst(basePathLen))
            try archive.addEntry(with: entryName, fileURL: fileURL, compressionMethod: .deflate)
        }
        return Fixture(name: name, url: outputURL, partsToVerify: partsToVerify)
    }

    /// Reads a single OOXML part out of a built .docx as raw XML bytes.
    static func readPart(_ partPath: String, from docxURL: URL) throws -> Data {
        let archive = try Archive(url: docxURL, accessMode: .read)
        guard let entry = archive[partPath] else {
            throw NSError(
                domain: "CorpusFixtureBuilder",
                code: 1,
                userInfo: [NSLocalizedDescriptionKey: "Part \(partPath) not found in \(docxURL.lastPathComponent)"]
            )
        }
        var buffer = Data()
        _ = try archive.extract(entry) { chunk in
            buffer.append(chunk)
        }
        return buffer
    }
}

// MARK: - multi-section-thesis fixture

private let multiSectionThesisDocumentXML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p w14:paraId="00000001"><w:r><w:t>Front matter section</w:t></w:r></w:p>
<w:p><w:pPr><w:sectPr w:rsidR="00ABC123"><w:headerReference w:type="default" r:id="rId1"/><w:footerReference w:type="default" r:id="rId3"/><w:type w:val="nextPage"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:pgNumType w:fmt="lowerRoman"/></w:sectPr></w:pPr></w:p>
<w:p w14:paraId="00000002"><w:r><w:t>Chapter 1 body</w:t></w:r></w:p>
<w:p w14:paraId="00000003"><w:pPr><w:sectPr w:rsidR="00BCD234"><w:headerReference w:type="default" r:id="rId2"/><w:footerReference w:type="default" r:id="rId4"/><w:type w:val="continuous"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:pPr><w:r><w:t>Chapter 2 transition</w:t></w:r></w:p>
<w:p w14:paraId="00000004"><w:r><w:t>Chapter 2 body</w:t></w:r></w:p>
<w:sectPr w:rsidR="00DEF456"><w:headerReference w:type="default" r:id="rId2"/><w:footerReference w:type="default" r:id="rId4"/><w:type w:val="nextPage"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:pgNumType w:fmt="decimal" w:start="1"/></w:sectPr>
</w:body>
</w:document>
"""

private let thesisHeader1XML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Front matter header</w:t></w:r></w:p></w:hdr>
"""

private let thesisHeader2XML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Chapter header</w:t></w:r></w:p></w:hdr>
"""

private let thesisFooter1XML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Roman footer</w:t></w:r></w:p></w:ftr>
"""

private let thesisFooter2XML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Arabic footer</w:t></w:r></w:p></w:ftr>
"""

// MARK: - vml-rich fixture

private let vmlRichDocumentXML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" mc:Ignorable="wps wpg">
<w:body>
<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps wpg"><w:drawing><wp:inline><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp><wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>DrawingML callout text</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape id="_x0000_s1026" type="#_x0000_t202" style="position:absolute;margin-left:36pt;margin-top:18pt;width:144pt;height:72pt;z-index:251660288" filled="t" fillcolor="#fff"><v:textbox style="mso-fit-shape-to-text:t"><w:txbxContent><w:p><w:r><w:t>VML callout text</w:t></w:r></w:p></w:txbxContent></v:textbox><w10:wrap type="square"/></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>
<w:p><w:r><w:t>Body text after VML</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>
</w:body>
</w:document>
"""

// MARK: - cjk-settings fixture

private let cjkDocumentXML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="標楷體" w:hAnsi="Times New Roman"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="標楷體" w:hAnsi="Times New Roman"/></w:rPr><w:t>假設 H₀ 為樣本平均數等於母體平均數。</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>
</w:body>
</w:document>
"""

private let cjkSettingsXML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
<w:zoom w:percent="100"/>
<w:defaultTabStop w:val="720"/>
<w:characterSpacingControl w:val="compressPunctuation"/>
<w:themeFontLang w:val="en-US" w:eastAsia="zh-TW"/>
<w:rsids><w:rsidRoot w:val="00ABC123"/><w:rsid w:val="00ABC123"/><w:rsid w:val="00DEF456"/><w:rsid w:val="00111222"/></w:rsids>
<m:mathPr><m:mathFont m:val="Cambria Math"/></m:mathPr>
<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/><w:compatSetting w:name="useFELayout" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/></w:compat>
</w:settings>
"""

// MARK: - comment-anchored fixture

private let commentAnchoredDocumentXML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:commentRangeStart w:id="0"/><w:r><w:t xml:space="preserve">This sentence is annotated. </w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r></w:p>
<w:p><w:r><w:t>Plain follow-up paragraph.</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>
</w:body>
</w:document>
"""

private let commentAnchoredCommentsXML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:comment w:id="0" w:author="Reviewer A" w:date="2026-05-05T18:00:00Z" w:initials="RA"><w:p><w:r><w:t>Consider rewording.</w:t></w:r></w:p></w:comment>
</w:comments>
"""

// MARK: - shared parts

private let contentTypesMinimal = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""

private let contentTypesWithSettings = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>
"""

private let contentTypesWithThesisParts = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>
"""

private let contentTypesWithComments = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
</Types>
"""

private let packageRelsMinimal = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

private let documentRelsMinimal = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>
"""

private let documentRelsWithComments = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>
"""

private let documentRelsWithThesisParts = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/>
</Relationships>
"""
