import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// word-aligned-state-sync Phase 1 task 2.6 — `docx-container-parsing`
/// Requirement: "All parts preserved via XmlNode tree alongside typed views".
///
/// `DocxReader.read(from:)` SHALL load every XML part of the package into
/// `document.xmlTrees`, not just the nine part classes the typed model
/// consumes. Parts the typed model does not understand (customXml/*,
/// word/theme/*, word/fontTable.xml, word/webSettings.xml, docProps/*) must
/// be addressable by the Phase 2 op log and — per the spec scenario "Unknown
/// part survives round-trip" — serialize byte-equal when untouched.
final class AllPartsTreeLoadingTests: XCTestCase {

    // MARK: - Fixture

    /// Fake PNG: magic header + arbitrary payload. Enough to verify binary
    /// parts are (a) not parsed into xmlTrees and (b) preserved byte-equal.
    private static let fakePNG = Data([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])
        + Data((0..<64).map { UInt8($0) })

    private struct Fixture {
        let url: URL
        let customXmlItem: String
        let themeXml: String
    }

    private func buildFixture(includeMalformedPart: Bool = false) throws -> Fixture {
        let staging = FileManager.default.temporaryDirectory
            .appendingPathComponent("all-parts-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: staging, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: staging) }

        func write(_ content: String, to relativePath: String) throws {
            let url = staging.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(), withIntermediateDirectories: true)
            try content.write(to: url, atomically: true, encoding: .utf8)
        }
        func writeData(_ data: Data, to relativePath: String) throws {
            let url = staging.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(), withIntermediateDirectories: true)
            try data.write(to: url)
        }

        let malformedOverride = includeMalformedPart
            ? #"<Override PartName="/word/broken.xml" ContentType="application/xml"/>"# : ""
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="xml" ContentType="application/xml"/>
                <Default Extension="png" ContentType="image/png"/>
                <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
                <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
                <Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>
                <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
                <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>\(malformedOverride)
            </Types>
            """, to: "[Content_Types].xml")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
                <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
            </Relationships>
            """, to: "_rels/.rels")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
                <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
                <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/>
                <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>
                <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
            </Relationships>
            """, to: "word/_rels/document.xml.rels")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>All-parts fixture</w:t></w:r></w:p></w:body></w:document>
            """, to: "word/document.xml")

        // Parts outside the typed model's nine classes:
        let customXmlItem = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <thesisMeta xmlns="urn:ntpu:thesis"><advisor>Dr. Example</advisor><committee size="3"/></thesisMeta>
            """
        try write(customXmlItem, to: "customXml/item1.xml")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <ds:datastoreItem ds:itemID="{D5C1A9B2-0000-4000-8000-000000000001}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>
            """, to: "customXml/itemProps1.xml")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/>
            </Relationships>
            """, to: "customXml/_rels/item1.xml.rels")
        let themeXml = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1></a:clrScheme></a:themeElements></a:theme>
            """
        try write(themeXml, to: "word/theme/theme1.xml")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font w:name="DFKai-SB"><w:charset w:val="88"/><w:family w:val="script"/></w:font></w:fonts>
            """, to: "word/fontTable.xml")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:optimizeForBrowser/><w:allowPNG/></w:webSettings>
            """, to: "word/webSettings.xml")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>All parts</dc:title><dc:creator>fixture</dc:creator></cp:coreProperties>
            """, to: "docProps/core.xml")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>fixture</Application></Properties>
            """, to: "docProps/app.xml")
        try writeData(Self.fakePNG, to: "word/media/image1.png")

        if includeMalformedPart {
            try write("<?xml version=\"1.0\"?><broken><unclosed>", to: "word/broken.xml")
        }

        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("all-parts-\(UUID().uuidString).docx")
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
        return Fixture(url: docxURL, customXmlItem: customXmlItem, themeXml: themeXml)
    }

    private func extractPart(_ partPath: String, from docxURL: URL) throws -> Data {
        let archive = try Archive(url: docxURL, accessMode: .read)
        guard let entry = archive[partPath] else {
            XCTFail("\(partPath) missing from \(docxURL.lastPathComponent)")
            return Data()
        }
        var data = Data()
        _ = try archive.extract(entry) { data.append($0) }
        return data
    }

    // MARK: - Every XML part lands in xmlTrees

    func testReaderLoadsEveryXmlPartIntoTree() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture.url) }
        let doc = try DocxReader.read(from: fixture.url)

        let expectedParts = [
            "customXml/item1.xml",
            "customXml/itemProps1.xml",
            "word/theme/theme1.xml",
            "word/fontTable.xml",
            "word/webSettings.xml",
            "docProps/core.xml",
            "docProps/app.xml",
        ]
        for part in expectedParts {
            XCTAssertNotNil(
                doc.xmlTrees[part],
                "xmlTrees must contain \(part) (task 2.6: every Content_Types part loads into the tree)")
        }
        XCTAssertEqual(doc.xmlTrees["customXml/item1.xml"]?.root.localName, "thesisMeta")
        XCTAssertEqual(doc.xmlTrees["word/theme/theme1.xml"]?.root.localName, "theme")
    }

    func testBinaryPartsAreNotParsedIntoTree() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture.url) }
        let doc = try DocxReader.read(from: fixture.url)

        XCTAssertNil(doc.xmlTrees["word/media/image1.png"],
                     "binary parts must not be parsed as XML trees")
    }

    // MARK: - Spec scenario: unknown part survives round-trip

    func testUnknownPartsSurviveRoundTripByteEqual() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture.url) }

        var doc = try DocxReader.read(from: fixture.url)
        // Touch the body so this is a real save, not a no-op copy.
        doc.appendParagraph(Paragraph(text: "mutation"))

        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("all-parts-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(doc, to: outURL)

        for part in ["customXml/item1.xml", "customXml/itemProps1.xml",
                     "word/theme/theme1.xml", "word/webSettings.xml"] {
            let input = try extractPart(part, from: fixture.url)
            let output = try extractPart(part, from: outURL)
            XCTAssertEqual(input, output, "\(part) must round-trip byte-equal when untouched")
        }
        let inputPNG = try extractPart("word/media/image1.png", from: fixture.url)
        XCTAssertEqual(inputPNG, try extractPart("word/media/image1.png", from: outURL),
                       "binary media must round-trip byte-equal")
    }

    // MARK: - Malformed part: loud diagnostic, no hard failure, bytes preserved

    func testMalformedXmlPartRecordsDiagnosticAndPreservesBytes() throws {
        let fixture = try buildFixture(includeMalformedPart: true)
        defer { try? FileManager.default.removeItem(at: fixture.url) }

        // Read must not throw: the malformed part is outside the typed model,
        // and its bytes are preserved by the overlay path regardless.
        var doc = try DocxReader.read(from: fixture.url)

        XCTAssertNil(doc.xmlTrees["word/broken.xml"])
        XCTAssertNotNil(doc.xmlTreeLoadFailures["word/broken.xml"],
                        "parse failure must be recorded loudly, not swallowed")

        doc.appendParagraph(Paragraph(text: "mutation"))
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("all-parts-broken-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(doc, to: outURL)

        let input = try extractPart("word/broken.xml", from: fixture.url)
        XCTAssertEqual(input, try extractPart("word/broken.xml", from: outURL),
                       "malformed part bytes must still round-trip via overlay copy")
    }
}
