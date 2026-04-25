import Foundation
import ZIPFoundation
@testable import OOXMLSwift

/// Programmatic fixture builder for the
/// `che-word-mcp-document-xml-lossless-roundtrip` Spectra change
/// (PsychQuant/che-word-mcp#56) Phase 5.
///
/// Synthesizes a 50–100 KB `.docx` containing every code path the v0.19.0
/// Reader / Writer pair must preserve through `open → mark dirty → save → reload`:
/// - 5+ bookmarks (3 paragraphs each carrying a pair, plus 2 standalone paragraphs)
/// - 3 hyperlinks (one external URL with `r:id`, one internal anchor, one mailto)
/// - 2 `<w:fldSimple>` blocks (one SEQ Table caption, one REF cross-reference)
/// - 1 `<mc:AlternateContent>` math block with both `<mc:Choice>` and `<mc:Fallback>`
/// - 10+ `xmlns:*` declarations on `<w:document>` root
/// - mixed runs / wrappers across 3+ paragraphs to exercise position-index ordering
///
/// Mirrors the `SDTFixtureBuilder` pattern: build content as raw OOXML strings,
/// stage to a temp directory, ZIP via ZIPFoundation, return the .docx URL.
/// Caller is responsible for cleanup (or rely on `/tmp` eviction at reboot).
enum LosslessRoundTripFixtureBuilder {

    /// Namespaces declared on the `<w:document>` root. Includes the typical
    /// Word 2019+ extension set so the fixture exercises the namespace
    /// preservation path even after `markPartDirty("word/document.xml")`.
    static let rootNamespaces: [(String, String)] = [
        ("w",  "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
        ("r",  "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
        ("a",  "http://schemas.openxmlformats.org/drawingml/2006/main"),
        ("m",  "http://schemas.openxmlformats.org/officeDocument/2006/math"),
        ("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
        ("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"),
        ("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
        ("v",  "urn:schemas-microsoft-com:vml"),
        ("o",  "urn:schemas-microsoft-com:office:office"),
        ("w14", "http://schemas.microsoft.com/office/word/2010/wordml"),
        ("w15", "http://schemas.microsoft.com/office/word/2012/wordml"),
        ("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex"),
    ]

    static let ignorableValue = "w14 w15 w16se"

    /// Build the fixture .docx in /tmp. Returns the .docx URL.
    static func build() throws -> URL {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("lossless-rt-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: stagingDir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: stagingDir) }

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
        try write(buildDocumentXML(), to: "word/document.xml")

        let outputURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("lossless-roundtrip-\(UUID().uuidString).docx")
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

    // MARK: - document.xml

    static func buildDocumentXML() -> String {
        var lines: [String] = []
        lines.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")
        let nsAttrs = rootNamespaces
            .map { "xmlns:\($0.0)=\"\($0.1)\"" }
            .joined(separator: " ")
        lines.append("<w:document \(nsAttrs) mc:Ignorable=\"\(ignorableValue)\">")
        lines.append("<w:body>")

        // Paragraph 1: bookmark pair around a single run (anchor for hyperlink).
        lines.append("""
        <w:p><w:bookmarkStart w:id="1" w:name="anchor-1"/><w:r><w:t>Section 1</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>
        """)

        // Paragraph 2: bookmark pair, FieldSimple SEQ Table caption,
        // a hyperlink anchor reference, a comment range — interleaved.
        lines.append("""
        <w:p><w:bookmarkStart w:id="2" w:name="anchor-2"/><w:r><w:t xml:space="preserve">Table </w:t></w:r><w:fldSimple w:instr=" SEQ Table \\* ARABIC "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve">: see </w:t></w:r><w:hyperlink w:anchor="anchor-1"><w:r><w:t>Section 1</w:t></w:r></w:hyperlink><w:bookmarkEnd w:id="2"/></w:p>
        """)

        // Paragraph 3: external hyperlink + REF fldSimple + internal mailto link.
        lines.append("""
        <w:p><w:bookmarkStart w:id="3" w:name="anchor-3"/><w:hyperlink r:id="rId10" w:tooltip="External"><w:r><w:t>example.com</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> and </w:t></w:r><w:fldSimple w:instr=" REF anchor-1 \\h "><w:r><w:t>Section 1</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> or </w:t></w:r><w:hyperlink r:id="rId11"><w:r><w:t>email</w:t></w:r></w:hyperlink><w:bookmarkEnd w:id="3"/></w:p>
        """)

        // Paragraph 4: AlternateContent math block + comment range.
        lines.append("""
        <w:p><w:commentRangeStart w:id="50"/><w:r><w:t xml:space="preserve">Pearson correlation: </w:t></w:r><mc:AlternateContent><mc:Choice Requires="wps14"><w:r><w:t>choice-placeholder</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t>r(X,Y)</w:t></w:r></mc:Fallback></mc:AlternateContent><w:r><w:t xml:space="preserve"> reported.</w:t></w:r><w:commentRangeEnd w:id="50"/></w:p>
        """)

        // Paragraph 5: standalone bookmark + standalone reference.
        lines.append("""
        <w:p><w:bookmarkStart w:id="4" w:name="appendix-a"/><w:r><w:t>Appendix A</w:t></w:r><w:bookmarkEnd w:id="4"/><w:r><w:t xml:space="preserve"> sits at the end.</w:t></w:r></w:p>
        """)

        // Paragraph 6: bookmark + proofing error markers (not blocking, just exercise).
        lines.append("""
        <w:p><w:bookmarkStart w:id="5" w:name="conclusion"/><w:proofErr w:type="spellStart"/><w:r><w:t>teh</w:t></w:r><w:proofErr w:type="spellEnd"/><w:r><w:t xml:space="preserve"> is a typo.</w:t></w:r><w:bookmarkEnd w:id="5"/></w:p>
        """)

        lines.append("<w:sectPr></w:sectPr>")
        lines.append("</w:body>")
        lines.append("</w:document>")
        return lines.joined(separator: "\n")
    }
}

// MARK: - Static OOXML boilerplate

private let contentTypesXML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""

private let packageRelsXML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

/// Includes rId10 (https://example.com/) and rId11 (mailto:user@example.com)
/// so the fixture's two external hyperlinks resolve cleanly.
private let documentRelsXML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>
  <Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="mailto:user@example.com" TargetMode="External"/>
</Relationships>
"""
