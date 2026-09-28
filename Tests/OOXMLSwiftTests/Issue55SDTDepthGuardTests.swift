import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Regression tests for [PsychQuant/ooxml-swift#55](https://github.com/PsychQuant/ooxml-swift/issues/55).
///
/// Surfaced during verify of #25 (Security Reviewer F1, pre-existing):
/// `Table` has an explicit nesting depth guard
/// (`DocxReader.MAX_TABLE_NEST_DEPTH = 5`, and `Document.insertNestedTable`'s
/// own API-level `WordError.nestedTooDeep` check), but `<w:sdt>` (Structured
/// Document Tag / content control) nesting had NONE — a crafted `.docx`
/// with a deep SDT-in-SDT-in-SDT chain could recurse arbitrarily and
/// exhaust the call stack, with no bound at all (not even a truncate-and-
/// continue fallback).
///
/// Three independent, previously-unguarded recursion paths all handle
/// `<w:sdt>` nesting and all needed the same fix:
///
/// 1. `SDTParser.parseSDT` — paragraph-level (inline) SDT nesting, e.g. a
///    `<w:sdt>` inside a `<w:p>`, whose `<w:sdtContent>` may itself contain
///    sibling `<w:sdt>` elements.
/// 2. `DocxReader.parseBodyChildren`'s `case "sdt":` branch — block-level
///    (body-level) SDT nesting, where the SDT is a direct child of
///    `<w:body>` rather than living inside a paragraph.
/// 3. `DocxReader.parseContainerChildBodyChildren`'s `case "sdt":` branch —
///    the same block-level shape, but for header/footer/footnote/endnote
///    container parts (a structurally separate function from #2, per the
///    existing "mirrors parseBodyChildren's case sdt branch" comment at
///    that call site).
///
/// All three now share `SDTParser.MAX_SDT_NEST_DEPTH = 5` (matching
/// `DocxReader.MAX_TABLE_NEST_DEPTH`'s value and reasoning) and throw
/// `WordError.invalidDocx` cleanly — not crash, not hang, not silently
/// truncate — the moment nesting would exceed it.
final class Issue55SDTDepthGuardTests: XCTestCase {

    // MARK: - Fixture construction

    private func buildFixture(documentXML: String, headerXML: String? = nil) throws -> URL {
        let staging = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue55-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: staging, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: staging) }

        func write(_ content: String, to relativePath: String) throws {
            let url = staging.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(), withIntermediateDirectories: true)
            try content.write(to: url, atomically: true, encoding: .utf8)
        }

        var contentTypesOverrides = """
                <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            """
        var docRelsExtra = ""
        if headerXML != nil {
            contentTypesOverrides += """

                <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
            """
            docRelsExtra = """
                <Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
            """
        }

        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="xml" ContentType="application/xml"/>
            \(contentTypesOverrides)
            </Types>
            """, to: "[Content_Types].xml")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """, to: "_rels/.rels")
        try write(documentXML, to: "word/document.xml")
        if let headerXML {
            try write(headerXML, to: "word/header1.xml")
            try write("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                \(docRelsExtra)
                </Relationships>
                """, to: "word/_rels/document.xml.rels")
        }

        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue55-\(UUID().uuidString).docx")
        do {
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
        } catch {
            try? FileManager.default.removeItem(at: docxURL)
            throw error
        }
        return docxURL
    }

    /// `levels` nested `<w:sdt>` elements, paragraph-level (all inside one
    /// `<w:p>`) — exercises `SDTParser.parseSDT`'s recursion. `levels == 0`
    /// means no SDT at all (a plain paragraph).
    private func makeParagraphLevelNestedSDTDocumentXML(levels: Int) -> String {
        var open = ""
        var close = ""
        for _ in 0..<levels {
            open += "<w:sdt><w:sdtContent>"
            close = "</w:sdtContent></w:sdt>" + close
        }
        return """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
            <w:p>\(open)<w:r><w:t>leaf</w:t></w:r>\(close)</w:p>\
            <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>\
            </w:body></w:document>
            """
    }

    /// `levels` nested `<w:sdt>` elements, block-level (direct children of
    /// `<w:body>`, not inside any `<w:p>`) — exercises
    /// `DocxReader.parseBodyChildren`'s `case "sdt":` recursion.
    private func makeBlockLevelNestedSDTDocumentXML(levels: Int) -> String {
        var open = ""
        var close = ""
        for _ in 0..<levels {
            open += "<w:sdt><w:sdtContent>"
            close = "</w:sdtContent></w:sdt>" + close
        }
        return """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
            \(open)<w:p><w:r><w:t>leaf</w:t></w:r></w:p>\(close)\
            <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>\
            </w:body></w:document>
            """
    }

    /// Same block-level nesting shape, but rooted at `<w:hdr>` — exercises
    /// `DocxReader.parseContainerChildBodyChildren`'s `case "sdt":`
    /// recursion (a structurally separate function from
    /// `parseBodyChildren`, used for header/footer/footnote/endnote parts).
    private func makeHeaderNestedSDTXML(levels: Int) -> String {
        var open = ""
        var close = ""
        for _ in 0..<levels {
            open += "<w:sdt><w:sdtContent>"
            close = "</w:sdtContent></w:sdt>" + close
        }
        return """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\
            \(open)<w:p><w:r><w:t>leaf</w:t></w:r></w:p>\(close)\
            </w:hdr>
            """
    }

    // MARK: - Paragraph-level (SDTParser.parseSDT)

    func testParagraphLevelSDTAtLimitReadsSuccessfully() throws {
        let url = try buildFixture(documentXML: makeParagraphLevelNestedSDTDocumentXML(levels: SDTParser.MAX_SDT_NEST_DEPTH + 1))
        defer { try? FileManager.default.removeItem(at: url) }

        var doc = try DocxReader.read(from: url)
        defer { doc.close() }
        guard case .paragraph(let para) = doc.body.children[0] else {
            XCTFail("expected a paragraph at body.children[0]")
            return
        }
        XCTAssertEqual(para.contentControls.count, 1, "one top-level SDT")
        XCTAssertEqual(nestingDepth(of: para.contentControls[0]), SDTParser.MAX_SDT_NEST_DEPTH)
    }

    func testParagraphLevelSDTOneOverLimitRefusesCleanlyNotCrash() throws {
        let url = try buildFixture(documentXML: makeParagraphLevelNestedSDTDocumentXML(levels: SDTParser.MAX_SDT_NEST_DEPTH + 2))
        defer { try? FileManager.default.removeItem(at: url) }

        XCTAssertThrowsError(try DocxReader.read(from: url)) { error in
            guard case WordError.invalidDocx(let message) = error else {
                XCTFail("expected WordError.invalidDocx, got \(error)")
                return
            }
            XCTAssertTrue(message.contains("SDT"), "expected message to name SDT nesting, got: \(message)")
            XCTAssertTrue(message.contains("\(SDTParser.MAX_SDT_NEST_DEPTH)"), "expected message to name the limit, got: \(message)")
        }
    }

    // MARK: - Block-level (DocxReader.parseBodyChildren)

    func testBlockLevelSDTAtLimitReadsSuccessfully() throws {
        let url = try buildFixture(documentXML: makeBlockLevelNestedSDTDocumentXML(levels: SDTParser.MAX_SDT_NEST_DEPTH))
        defer { try? FileManager.default.removeItem(at: url) }

        var doc = try DocxReader.read(from: url)
        defer { doc.close() }
        XCTAssertEqual(blockSDTNestingDepth(doc.body.children), SDTParser.MAX_SDT_NEST_DEPTH)
    }

    func testBlockLevelSDTOneOverLimitRefusesCleanlyNotCrash() throws {
        let url = try buildFixture(documentXML: makeBlockLevelNestedSDTDocumentXML(levels: SDTParser.MAX_SDT_NEST_DEPTH + 1))
        defer { try? FileManager.default.removeItem(at: url) }

        XCTAssertThrowsError(try DocxReader.read(from: url)) { error in
            guard case WordError.invalidDocx(let message) = error else {
                XCTFail("expected WordError.invalidDocx, got \(error)")
                return
            }
            XCTAssertTrue(message.contains("SDT"), "expected message to name SDT nesting, got: \(message)")
        }
    }

    /// A pathological chain well past the limit (20 levels — 4x
    /// `MAX_SDT_NEST_DEPTH`) must still refuse cleanly (not hang, not
    /// crash) — proves the guard actually bounds recursion at a fixed
    /// depth rather than merely happening to survive one borderline case.
    /// (Deliberately NOT hundreds of levels deep: that starts probing an
    /// unrelated, pre-existing raw-XML-depth ceiling in `XmlTreeReader`'s
    /// own recursive-descent parser — which runs before `parseBodyChildren`
    /// even begins and is out of #55's scope — rather than this guard.)
    func testBlockLevelSDTFarOverLimitRefusesCleanlyNotCrashOrHang() throws {
        let url = try buildFixture(documentXML: makeBlockLevelNestedSDTDocumentXML(levels: 20))
        defer { try? FileManager.default.removeItem(at: url) }

        XCTAssertThrowsError(try DocxReader.read(from: url)) { error in
            guard case WordError.invalidDocx = error else {
                XCTFail("expected WordError.invalidDocx, got \(error)")
                return
            }
        }
    }

    // MARK: - Container-level (DocxReader.parseContainerChildBodyChildren, via header)

    func testHeaderContainerSDTAtLimitReadsSuccessfully() throws {
        let url = try buildFixture(
            documentXML: makeBlockLevelNestedSDTDocumentXML(levels: 0),
            headerXML: makeHeaderNestedSDTXML(levels: SDTParser.MAX_SDT_NEST_DEPTH))
        defer { try? FileManager.default.removeItem(at: url) }

        var doc = try DocxReader.read(from: url)
        defer { doc.close() }
        XCTAssertEqual(doc.headers.count, 1)
        XCTAssertEqual(blockSDTNestingDepth(doc.headers[0].bodyChildren), SDTParser.MAX_SDT_NEST_DEPTH)
    }

    func testHeaderContainerSDTOneOverLimitRefusesCleanlyNotCrash() throws {
        let url = try buildFixture(
            documentXML: makeBlockLevelNestedSDTDocumentXML(levels: 0),
            headerXML: makeHeaderNestedSDTXML(levels: SDTParser.MAX_SDT_NEST_DEPTH + 1))
        defer { try? FileManager.default.removeItem(at: url) }

        XCTAssertThrowsError(try DocxReader.read(from: url)) { error in
            guard case WordError.invalidDocx(let message) = error else {
                XCTFail("expected WordError.invalidDocx, got \(error)")
                return
            }
            XCTAssertTrue(message.contains("SDT"), "expected message to name SDT nesting, got: \(message)")
        }
    }

    // MARK: - Helpers

    private func nestingDepth(of control: ContentControl) -> Int {
        guard let deepest = control.children.map(nestingDepth(of:)).max() else { return 0 }
        return 1 + deepest
    }

    private func blockSDTNestingDepth(_ children: [BodyChild]) -> Int {
        var maxDepth = 0
        for child in children {
            if case .contentControl(_, let inner) = child {
                maxDepth = max(maxDepth, 1 + blockSDTNestingDepth(inner))
            }
        }
        return maxDepth
    }
}
