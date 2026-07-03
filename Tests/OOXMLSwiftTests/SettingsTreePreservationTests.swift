import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// word-aligned-state-sync Phase 1 task 2.5 — settings.xml tree-backed
/// preservation (`docx-container-parsing` Requirement: "Reader does not
/// silently drop element classes", Scenario "parseSettings retains all
/// settings.xml children" + Example "settings.xml round-trip preservation").
///
/// Root cause under test (PsychQuant/ooxml-swift#69): `DocxWriter.writeSettings`
/// emits a hard-coded ~400-byte template whenever `word/settings.xml` is dirty,
/// discarding every child of the original `<w:settings>` (rsids, compat,
/// mathPr, themeFontLang, footnotePr, …) AND omitting the very flags
/// (`<w:trackChanges/>`, `<w:evenAndOddHeaders/>`) whose mutation APIs are the
/// only way settings.xml becomes dirty in the first place.
///
/// Contract established here:
/// - Reader populates typed flags from the settings tree on read
///   (`isTrackChangesEnabled()`, `evenAndOddHeaders`), so writer-side flag
///   sync never destroys flags that Word (not Swift) set.
/// - Writer serializes dirty settings from `xmlTrees["word/settings.xml"]`
///   with the typed flags synced into the tree, preserving all other children.
/// - Scratch mode (no source archive → no tree) falls back to the template
///   plus the typed flags.
final class SettingsTreePreservationTests: XCTestCase {

    // MARK: - Fixture

    /// Representative subset of a Word-authored 19KB settings.xml (#69):
    /// every element here must survive a settings-dirtying mutation.
    private static let richSettingsChildren = """
        <w:zoom w:percent="130"/>\
        <w:proofState w:spelling="clean" w:grammar="clean"/>\
        <w:defaultTabStop w:val="480"/>\
        <w:characterSpacingControl w:val="compressPunctuation"/>\
        <w:hdrShapeDefaults><o:shapedefaults v:ext="edit" spidmax="2049"/></w:hdrShapeDefaults>\
        <w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>\
        <w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>\
        <w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/><w:compatSetting w:name="useWord2013TrackBottomHyphenation" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/></w:compat>\
        <w:rsids><w:rsidRoot w:val="00AB12CD"/><w:rsid w:val="00AB12CD"/><w:rsid w:val="00DE34F5"/><w:rsid w:val="0012AB78"/></w:rsids>\
        <m:mathPr><m:mathFont m:val="Cambria Math"/><m:brkBin m:val="before"/></m:mathPr>\
        <w:themeFontLang w:val="en-US" w:eastAsia="zh-TW"/>\
        <w:clrSchemeMapping w:bg1="light1" w:t1="dark1"/>\
        <w:shapeDefaults><o:shapedefaults v:ext="edit" spidmax="2049"/></w:shapeDefaults>\
        <w:decimalSymbol w:val="."/>\
        <w:listSeparator w:val=","/>
        """

    /// Elements asserted to survive every settings-dirtying round-trip.
    /// (Substring probes into the serialized settings.xml — element presence,
    /// not byte positions.)
    private static let preservationProbes = [
        "<w:zoom w:percent=\"130\"/>",
        "<w:proofState w:spelling=\"clean\" w:grammar=\"clean\"/>",
        "<w:defaultTabStop w:val=\"480\"/>",
        "<w:characterSpacingControl w:val=\"compressPunctuation\"/>",
        "<w:hdrShapeDefaults>",
        "<w:footnotePr>",
        "<w:endnotePr>",
        "useWord2013TrackBottomHyphenation",
        "<w:rsidRoot w:val=\"00AB12CD\"/>",
        "<w:rsid w:val=\"0012AB78\"/>",
        "<m:mathPr>",
        "<w:themeFontLang w:val=\"en-US\" w:eastAsia=\"zh-TW\"/>",
        "<w:clrSchemeMapping w:bg1=\"light1\" w:t1=\"dark1\"/>",
        "<w:shapeDefaults>",
        "<w:decimalSymbol w:val=\".\"/>",
        "<w:listSeparator w:val=\",\"/>",
    ]

    /// Builds a minimal .docx whose settings.xml carries the rich child set.
    /// - Parameter extraSettingsChildren: raw OOXML spliced *before* the rich
    ///   children (e.g. `<w:trackChanges/>` for the disable-path fixtures).
    private func buildFixture(extraSettingsChildren: String = "") throws -> URL {
        let staging = FileManager.default.temporaryDirectory
            .appendingPathComponent("settings-tree-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: staging, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: staging) }

        func write(_ content: String, to relativePath: String) throws {
            let url = staging.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(), withIntermediateDirectories: true)
            try content.write(to: url, atomically: true, encoding: .utf8)
        }

        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="xml" ContentType="application/xml"/>
                <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
            </Types>
            """, to: "[Content_Types].xml")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """, to: "_rels/.rels")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
            </Relationships>
            """, to: "word/_rels/document.xml.rels")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Settings preservation fixture</w:t></w:r></w:p></w:body></w:document>
            """, to: "word/document.xml")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml">\(extraSettingsChildren)\(Self.richSettingsChildren)</w:settings>
            """, to: "word/settings.xml")

        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("settings-tree-\(UUID().uuidString).docx")
        let archive = try Archive(url: docxURL, accessMode: .create)
        let base = staging.resolvingSymlinksInPath().path
        let enumerator = FileManager.default.enumerator(
            at: staging, includingPropertiesForKeys: [.isDirectoryKey])!
        for case let fileURL as URL in enumerator {
            let isDir = (try? fileURL.resourceValues(forKeys: [.isDirectoryKey]))?.isDirectory ?? false
            if isDir { continue }
            let entry = String(fileURL.resolvingSymlinksInPath().path.dropFirst(base.count + 1))
            try archive.addEntry(with: entry, fileURL: fileURL, compressionMethod: .deflate)
        }
        return docxURL
    }

    /// Saves the document and returns the serialized `word/settings.xml`.
    private func savedSettingsXML(_ document: WordDocument) throws -> String {
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("settings-tree-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(document, to: outURL)
        let archive = try Archive(url: outURL, accessMode: .read)
        guard let entry = archive["word/settings.xml"] else {
            XCTFail("output docx has no word/settings.xml")
            return ""
        }
        var data = Data()
        _ = try archive.extract(entry) { data.append($0) }
        return String(decoding: data, as: UTF8.self)
    }

    private func assertRichChildrenPreserved(
        in xml: String, file: StaticString = #filePath, line: UInt = #line
    ) {
        for probe in Self.preservationProbes {
            XCTAssertTrue(
                xml.contains(probe),
                "settings.xml lost \(probe) after settings-dirtying mutation (#69 template overwrite)",
                file: file, line: line)
        }
    }

    // MARK: - Reader populates typed flags from tree

    func testReaderPopulatesTypedFlagsFromSettingsTree() throws {
        let fixture = try buildFixture(
            extraSettingsChildren: "<w:trackChanges/><w:evenAndOddHeaders/>")
        defer { try? FileManager.default.removeItem(at: fixture) }
        let doc = try DocxReader.read(from: fixture)

        XCTAssertTrue(doc.isTrackChangesEnabled(),
                      "reader must populate trackChanges from settings tree")
        XCTAssertTrue(doc.evenAndOddHeaders,
                      "reader must populate evenAndOddHeaders from settings tree")
    }

    func testReaderDefaultsFlagsFalseWhenAbsent() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        let doc = try DocxReader.read(from: fixture)

        XCTAssertFalse(doc.isTrackChangesEnabled())
        XCTAssertFalse(doc.evenAndOddHeaders)
    }

    // MARK: - enableTrackChanges / disableTrackChanges

    func testEnableTrackChangesPreservesAllSettingsChildren() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        doc.enableTrackChanges(author: "Tester")

        let xml = try savedSettingsXML(doc)
        XCTAssertTrue(xml.contains("<w:trackChanges/>"),
                      "enableTrackChanges must persist <w:trackChanges/> to settings.xml")
        assertRichChildrenPreserved(in: xml)
    }

    func testTrackChangesInsertedBeforeDefaultTabStop() throws {
        // CT_Settings is an xsd:sequence: trackChanges precedes defaultTabStop.
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        doc.enableTrackChanges(author: "Tester")

        let xml = try savedSettingsXML(doc)
        guard let flagRange = xml.range(of: "<w:trackChanges/>"),
              let anchorRange = xml.range(of: "<w:defaultTabStop") else {
            return XCTFail("expected both trackChanges and defaultTabStop in output")
        }
        XCTAssertLessThan(flagRange.lowerBound, anchorRange.lowerBound,
                          "trackChanges must be inserted before defaultTabStop (CT_Settings order)")
    }

    func testDisableTrackChangesRemovesFlagPreservesChildren() throws {
        let fixture = try buildFixture(extraSettingsChildren: "<w:trackChanges/>")
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        XCTAssertTrue(doc.isTrackChangesEnabled(), "fixture precondition")
        doc.disableTrackChanges()

        let xml = try savedSettingsXML(doc)
        XCTAssertFalse(xml.contains("<w:trackChanges"),
                       "disableTrackChanges must remove the flag from settings.xml")
        assertRichChildrenPreserved(in: xml)
    }

    // MARK: - setEvenAndOddHeaders

    func testSetEvenAndOddHeadersOnPreservesChildren() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        doc.setEvenAndOddHeaders(true)

        let xml = try savedSettingsXML(doc)
        XCTAssertTrue(xml.contains("<w:evenAndOddHeaders/>"),
                      "setEvenAndOddHeaders(true) must persist the flag")
        assertRichChildrenPreserved(in: xml)
    }

    func testSetEvenAndOddHeadersOffRemovesFlagPreservesChildren() throws {
        let fixture = try buildFixture(extraSettingsChildren: "<w:evenAndOddHeaders/>")
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        XCTAssertTrue(doc.evenAndOddHeaders, "fixture precondition")
        doc.setEvenAndOddHeaders(false)

        let xml = try savedSettingsXML(doc)
        XCTAssertFalse(xml.contains("<w:evenAndOddHeaders"),
                       "setEvenAndOddHeaders(false) must remove the flag")
        assertRichChildrenPreserved(in: xml)
    }

    // MARK: - Independent flags do not clobber each other

    func testTogglingEvenAndOddHeadersDoesNotDropWordAuthoredTrackChanges() throws {
        // Word (not Swift) enabled track changes; Swift only toggles headers.
        // Flag sync must not silently disable Word's revision tracking.
        let fixture = try buildFixture(extraSettingsChildren: "<w:trackChanges/>")
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        doc.setEvenAndOddHeaders(true)

        let xml = try savedSettingsXML(doc)
        XCTAssertTrue(xml.contains("<w:trackChanges/>"),
                      "toggling evenAndOddHeaders must not drop Word-authored trackChanges")
        XCTAssertTrue(xml.contains("<w:evenAndOddHeaders/>"))
        assertRichChildrenPreserved(in: xml)
    }

    // MARK: - Scratch mode fallback (no source archive → no settings tree)

    func testScratchModeSettingsIncludeFlags() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "scratch"))
        doc.enableTrackChanges(author: "Tester")
        doc.setEvenAndOddHeaders(true)

        let xml = try savedSettingsXML(doc)
        XCTAssertTrue(xml.contains("<w:trackChanges/>"),
                      "scratch-mode settings must include enabled trackChanges")
        XCTAssertTrue(xml.contains("<w:evenAndOddHeaders/>"),
                      "scratch-mode settings must include enabled evenAndOddHeaders")
        // Template essentials still present.
        XCTAssertTrue(xml.contains("<w:defaultTabStop"))
    }

    // MARK: - No-mutation overlay path stays byte-preserving (regression guard)

    func testUntouchedSettingsRoundTripByteEqual() throws {
        let fixture = try buildFixture(extraSettingsChildren: "<w:trackChanges/>")
        defer { try? FileManager.default.removeItem(at: fixture) }

        let inputArchive = try Archive(url: fixture, accessMode: .read)
        var inputData = Data()
        _ = try inputArchive.extract(inputArchive["word/settings.xml"]!) { inputData.append($0) }

        let doc = try DocxReader.read(from: fixture)
        let xml = try savedSettingsXML(doc)
        XCTAssertEqual(xml, String(decoding: inputData, as: UTF8.self),
                       "untouched settings.xml must round-trip byte-equal (overlay path)")
    }
}
