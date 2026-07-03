import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// word-aligned-state-sync Phase 1 tasks 2.8–2.10 — `docx-revision-parsing`
/// Requirements:
/// - "Revision elements preserved via tree on round-trip" (2.8)
/// - "Unknown revision children are preserved verbatim" +
///   "Nested property-change revisions round-trip via tree" (2.9)
/// - "Revision-source typed view backed by tree" (2.10)
///
/// Same contract shape as task 2.7: typed accessors keep working, the part
/// tree retains every revision element, and no-mutation round-trip is
/// byte-equal — including `<w:ins>` metadata, nested `<w:rPrChange>` /
/// `<w:pPrChange>` snapshots, and future `w16cid:*` revision-extension
/// elements the typed model does not understand.
final class RevisionTreePreservationTests: XCTestCase {

    // MARK: - Fixture

    /// Body: tracked insert + tracked delete + rPrChange + pPrChange +
    /// w16cid extension element inside a paragraph.
    private static let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:ins w:id="42" w:author="Alice" w:date="2026-05-04T10:00:00Z"><w:r><w:t>inserted text</w:t></w:r></w:ins><w:del w:id="43" w:author="Bob" w:date="2026-05-04T11:00:00Z"><w:r><w:delText>deleted text</w:delText></w:r></w:del></w:p><w:p><w:pPr><w:pPrChange w:id="44" w:author="Alice" w:date="2026-05-04T12:00:00Z"><w:pPr><w:jc w:val="center"/></w:pPr></w:pPrChange></w:pPr><w:r><w:rPr><w:b/><w:rPrChange w:id="7" w:author="Alice" w:date="2026-05-04T12:30:00Z"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>formatted run</w:t></w:r><w16cid:commentsExtensible w16cid:durableId="123ABC"/></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHdr1"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>
        """

    /// Header with its own tracked insert (for the source-disambiguation
    /// scenario of task 2.10).
    private static let headerXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:ins w:id="90" w:author="Carol" w:date="2026-05-05T09:00:00Z"><w:r><w:t>header insert</w:t></w:r></w:ins></w:p></w:hdr>
        """

    /// Footnote with a tracked insert.
    private static let footnotesXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:ins w:id="91" w:author="Dave" w:date="2026-05-05T10:00:00Z"><w:r><w:t>footnote insert</w:t></w:r></w:ins></w:p></w:footnote></w:footnotes>
        """

    private func buildFixture() throws -> URL {
        let staging = FileManager.default.temporaryDirectory
            .appendingPathComponent("revision-tree-staging-\(UUID().uuidString)")
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
                <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
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
                <Relationship Id="rIdHdr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                <Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
            </Relationships>
            """, to: "word/_rels/document.xml.rels")
        try write(Self.documentXML, to: "word/document.xml")
        try write(Self.headerXML, to: "word/header1.xml")
        try write(Self.footnotesXML, to: "word/footnotes.xml")

        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("revision-tree-\(UUID().uuidString).docx")
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

    // MARK: - Task 2.8: revision elements — typed accessors AND tree retention

    func testTypedRevisionAccessorsStillWork() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        let doc = try DocxReader.read(from: fixture)

        let revisions = doc.getRevisionsFull()
        XCTAssertTrue(revisions.contains { $0.author == "Alice" },
                      "typed parse must surface Alice's tracked insert")
        XCTAssertTrue(revisions.contains { $0.author == "Bob" },
                      "typed parse must surface Bob's tracked delete")
    }

    func testRevisionElementsRetainedInTree() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        let doc = try DocxReader.read(from: fixture)

        func collect(_ node: XmlNode, localName: String, into out: inout [XmlNode]) {
            if node.kind == .element && node.localName == localName { out.append(node) }
            for child in node.children { collect(child, localName: localName, into: &out) }
        }
        guard let docRoot = doc.xmlTrees["word/document.xml"]?.root else {
            return XCTFail("document tree missing")
        }
        var insNodes: [XmlNode] = []
        collect(docRoot, localName: "ins", into: &insNodes)
        XCTAssertEqual(insNodes.count, 1, "tree must retain the body <w:ins>")
        let attrs = Dictionary(uniqueKeysWithValues: insNodes[0].attributes.map {
            ($0.localName, $0.value)
        })
        XCTAssertEqual(attrs["id"], "42")
        XCTAssertEqual(attrs["author"], "Alice")
        XCTAssertEqual(attrs["date"], "2026-05-04T10:00:00Z")
    }

    func testRevisionMarkupRoundTripsByteEqual() throws {
        // Spec scenario "Revision metadata survives round-trip" — no-mutation
        // save must reproduce <w:ins>/<w:del> byte-equal.
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        let doc = try DocxReader.read(from: fixture)

        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("revision-tree-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(doc, to: outURL)

        XCTAssertEqual(try extractPart("word/document.xml", from: fixture),
                       try extractPart("word/document.xml", from: outURL),
                       "document.xml with revision markup must round-trip byte-equal")
    }

    // MARK: - Task 2.9: nested property-change revisions + unknown revision children

    func testNestedPropertyChangeRevisionsSurviveByteEqual() throws {
        // Covers spec scenarios "rPrChange survives round-trip" (nested
        // <w:rPr> snapshot) and the pPrChange analogue. Byte-equality of the
        // whole part (previous test) implies these, but assert the specific
        // elements so a future writer regression names the culprit.
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        let doc = try DocxReader.read(from: fixture)

        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("revision-tree-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(doc, to: outURL)

        let out = String(decoding: try extractPart("word/document.xml", from: outURL), as: UTF8.self)
        XCTAssertTrue(out.contains(
            #"<w:rPrChange w:id="7" w:author="Alice" w:date="2026-05-04T12:30:00Z"><w:rPr><w:i/></w:rPr></w:rPrChange>"#),
            "nested rPrChange incl. its <w:rPr> snapshot must survive verbatim")
        XCTAssertTrue(out.contains(
            #"<w:pPrChange w:id="44" w:author="Alice" w:date="2026-05-04T12:00:00Z"><w:pPr><w:jc w:val="center"/></w:pPr></w:pPrChange>"#),
            "nested pPrChange incl. its <w:pPr> snapshot must survive verbatim")
        XCTAssertTrue(out.contains(#"<w16cid:commentsExtensible w16cid:durableId="123ABC"/>"#),
            "w16cid revision-extension element must survive verbatim")
    }

    func testUnknownRevisionChildrenDoNotSurfaceAsUnknownSentinel() throws {
        // Spec scenario "w16 revision extension is preserved": no typed model
        // surfaces an opaque `unknown` sentinel for the extension element.
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        let doc = try DocxReader.read(from: fixture)

        for revision in doc.getRevisionsFull() {
            XCTAssertFalse(String(describing: revision.type).lowercased().contains("unknown"),
                           "no revision may carry an opaque unknown sentinel type")
        }
    }

    // MARK: - Task 2.10: revision-source disambiguation unchanged

    func testRevisionSourceDisambiguationAcrossContainers() throws {
        // Spec scenario "Revisions report container source after the change":
        // body, header1, footnote1 revisions carry matching source enums.
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        let doc = try DocxReader.read(from: fixture)

        let revisions = doc.getRevisionsFull()
        XCTAssertTrue(revisions.contains { $0.author == "Alice" && $0.source == .body },
                      "body revision must report source .body")
        XCTAssertTrue(revisions.contains {
            if $0.author == "Carol", case .header = $0.source { return true }
            return false
        }, "header revision must report source .header")
        XCTAssertTrue(revisions.contains { $0.author == "Dave" && $0.source == .footnote(id: 1) },
                      "footnote revision must report source .footnote(id: 1)")
    }
}
