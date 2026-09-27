import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Regression tests for [PsychQuant/ooxml-swift#164](https://github.com/PsychQuant/ooxml-swift/issues/164).
///
/// Before this fix: `<w:rPr>` children `RunProperties` does not model
/// (`<w:bCs>`, `<w:iCs>`, `<w:vanish>`, …) were dropped whenever
/// `word/document.xml` was regenerated from the typed model — which happens
/// for the WHOLE document as soon as ANY paragraph is edited through the
/// legacy direct-typed surface (`updateParagraph`, `replaceText`,
/// che-word-mcp's `format_text`). `<w:vanish/>` is hidden text, so an
/// unrelated edit made hidden text visible, silently.
///
/// After this fix: `DocxReader.parseRunProperties` captures any `<w:rPr>`
/// child not in its recognized set into `RunProperties.rawChildren` (same
/// "if not typed, preserve as raw" pattern as `Run.rawElements` and
/// `ParagraphProperties.rawChildren`, #168), and `RunProperties.toXML()`
/// re-emits each at its ECMA-376 canonical `CT_RPr` position.
final class Issue164RunPropertiesRawChildrenTests: XCTestCase {

    // MARK: - Section A: unit-level parse/emit round trip (no docx I/O)

    private func parseRunFragment(_ xml: String) throws -> Run {
        let wrapper = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          \(xml)
        </w:document>
        """
        let doc = try XMLDocument(xmlString: wrapper)
        guard let root = doc.rootElement(),
              let runElement = root.elements(forName: "w:r").first else {
            throw NSError(domain: "test", code: 0, userInfo: [NSLocalizedDescriptionKey: "no <w:r> found"])
        }
        return try DocxReader.parseRun(from: runElement, relationships: RelationshipsCollection())
    }

    /// The issue's exact minimal repro run: bold (typed) plus three
    /// unmodeled rPr children.
    private static let watchedRunXML =
        "<w:r><w:rPr><w:b/><w:bCs w:val=\"0\"/><w:iCs/><w:vanish/></w:rPr><w:t>alpha</w:t></w:r>"

    func testParseCapturesUnmodeledRPrChildren() throws {
        let run = try parseRunFragment(Self.watchedRunXML)
        let names = run.properties.rawChildren?.map(\.name)
        XCTAssertEqual(names, ["bCs", "iCs", "vanish"],
                       "unmodeled rPr children must be captured, in source order")
        // Modeled fields are unaffected — still populated the old way.
        XCTAssertTrue(run.properties.bold)
    }

    func testToXMLReemitsRawChildrenAtCanonicalPositions() throws {
        let run = try parseRunFragment(Self.watchedRunXML)
        let xml = run.toXML()
        XCTAssertTrue(xml.contains("<w:bCs w:val=\"0\"/>"), "got: \(xml)")
        XCTAssertTrue(xml.contains("<w:iCs/>"), "got: \(xml)")
        XCTAssertTrue(xml.contains("<w:vanish/>"), "got: \(xml)")
        // CT_RPr canonical order: b(3) bCs(4) i(5) iCs(6) … vanish(17).
        let bRange = try XCTUnwrap(xml.range(of: "<w:b/>"))
        let bCsRange = try XCTUnwrap(xml.range(of: "<w:bCs"))
        let iCsRange = try XCTUnwrap(xml.range(of: "<w:iCs/>"))
        let vanishRange = try XCTUnwrap(xml.range(of: "<w:vanish/>"))
        XCTAssertLessThan(bRange.lowerBound, bCsRange.lowerBound)
        XCTAssertLessThan(bCsRange.lowerBound, iCsRange.lowerBound)
        XCTAssertLessThan(iCsRange.lowerBound, vanishRange.lowerBound)
    }

    func testParseThenEmitRoundTripsUnmodeledChildrenVerbatim() throws {
        let run = try parseRunFragment(Self.watchedRunXML)
        let reparsed = try parseRunFragment(run.toXML())
        XCTAssertEqual(reparsed.properties.rawChildren?.map(\.name), ["bCs", "iCs", "vanish"])
    }

    // MARK: - Section A2: the remaining 4/7 elements the issue's Expected names
    // (independent review finding F1)
    //
    // Issue #164's Expected: "A regression test should cover at least
    // `vanish`, `bCs`, `iCs`, `caps`, `smallCaps`, `spacing`, `shd`" — the
    // three tests above only exercise the first three. `caps`/`smallCaps`
    // are empty ST_OnOff elements like `bCs`/`iCs`; `spacing`/`shd` carry
    // attributes (character spacing in twentieths of a point; run shading),
    // exercising a different shape than the boolean-flag elements above.

    private static let watchedRunXML2 =
        "<w:r><w:rPr><w:caps/><w:smallCaps/><w:spacing w:val=\"20\"/>"
        + "<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"FFFF00\"/></w:rPr><w:t>beta</w:t></w:r>"

    func testParseCapturesCapsSmallCapsSpacingShd() throws {
        let run = try parseRunFragment(Self.watchedRunXML2)
        let names = run.properties.rawChildren?.map(\.name)
        XCTAssertEqual(names, ["caps", "smallCaps", "spacing", "shd"],
                       "unmodeled rPr children must be captured, in source order")
    }

    func testToXMLReemitsCapsSmallCapsSpacingShdAtCanonicalPositions() throws {
        let run = try parseRunFragment(Self.watchedRunXML2)
        let xml = run.toXML()
        XCTAssertTrue(xml.contains("<w:caps/>"), "got: \(xml)")
        XCTAssertTrue(xml.contains("<w:smallCaps/>"), "got: \(xml)")
        XCTAssertTrue(xml.contains("<w:spacing w:val=\"20\"/>"), "got: \(xml)")
        XCTAssertTrue(xml.contains("w:shd") && xml.contains("w:fill=\"FFFF00\""), "got: \(xml)")
        // CT_RPr canonical order: caps(7) smallCaps(8) … spacing(20) … shd(30).
        let capsRange = try XCTUnwrap(xml.range(of: "<w:caps/>"))
        let smallCapsRange = try XCTUnwrap(xml.range(of: "<w:smallCaps/>"))
        let spacingRange = try XCTUnwrap(xml.range(of: "<w:spacing"))
        let shdRange = try XCTUnwrap(xml.range(of: "<w:shd"))
        XCTAssertLessThan(capsRange.lowerBound, smallCapsRange.lowerBound)
        XCTAssertLessThan(smallCapsRange.lowerBound, spacingRange.lowerBound)
        XCTAssertLessThan(spacingRange.lowerBound, shdRange.lowerBound)
    }

    func testParseThenEmitRoundTripsCapsSmallCapsSpacingShdVerbatim() throws {
        let run = try parseRunFragment(Self.watchedRunXML2)
        let reparsed = try parseRunFragment(run.toXML())
        XCTAssertEqual(reparsed.properties.rawChildren?.map(\.name), ["caps", "smallCaps", "spacing", "shd"])
        // Attribute-bearing elements must keep their attributes, not just
        // their element name (same "name-only checks would still pass if
        // attributes were dropped" concern the existing tests guard against).
        let shd = reparsed.properties.rawChildren?.first { $0.name == "shd" }
        XCTAssertTrue(shd?.xml.contains("w:fill=\"FFFF00\"") ?? false, "got: \(String(describing: shd?.xml))")
        let spacing = reparsed.properties.rawChildren?.first { $0.name == "spacing" }
        XCTAssertTrue(spacing?.xml.contains("w:val=\"20\"") ?? false, "got: \(String(describing: spacing?.xml))")
    }

    /// Full-docx regression, mirroring Section B below but for the second
    /// element set — an unrelated typed edit elsewhere must not drop
    /// caps/smallCaps/spacing/shd from this untouched run either.
    func testUpdateParagraphElsewherePreservesCapsSmallCapsSpacingShdOnUntouchedRun() throws {
        let documentXML = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
            <w:p>\(Self.watchedRunXML2)</w:p>\
            <w:p><w:r><w:t>gamma</w:t></w:r></w:p>\
            <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>\
            </w:body></w:document>
            """
        let fixture = try buildFixture(documentXML: documentXML)
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        try doc.updateParagraph(at: 1, text: "GAMMA")

        let outURL = try save(doc)
        defer { try? FileManager.default.removeItem(at: outURL) }
        let outputXML = try extractDocumentXML(from: outURL)
        let tree = try XmlTreeReader.parse(outputXML)

        func containsMarker(_ node: XmlNode) -> Bool {
            if node.kind == .text, node.textContent.contains("beta") { return true }
            return node.children.contains(where: containsMarker)
        }
        func findParagraph(_ node: XmlNode) -> XmlNode? {
            if node.kind == .element, node.localName == "p", containsMarker(node) { return node }
            for child in node.children { if let found = findParagraph(child) { return found } }
            return nil
        }
        let paragraph = try XCTUnwrap(findParagraph(tree.root), "no <w:p> containing 'beta' found")
        let run = try XCTUnwrap(paragraph.children.first { $0.kind == .element && $0.localName == "r" })
        let rPr = try XCTUnwrap(run.children.first { $0.kind == .element && $0.localName == "rPr" })
        let names = rPr.children.filter { $0.kind == .element }.map(\.localName)
        for expected in ["caps", "smallCaps", "spacing", "shd"] {
            XCTAssertEqual(names.filter { $0 == expected }.count, 1,
                           "\(expected) must survive exactly once")
        }
    }

    // MARK: - Section B: full-docx regression via the issue's exact repro

    /// `document.xml` for the fixture — mirrors the issue body exactly:
    /// paragraph 0 carries the watched run (bold + bCs/iCs/vanish), paragraph
    /// 1 ("gamma") is the one the repro edits.
    private static let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
        <w:p>\(watchedRunXML)</w:p>\
        <w:p><w:r><w:t>gamma</w:t></w:r></w:p>\
        <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>\
        </w:body></w:document>
        """

    private func buildFixture(documentXML: String? = nil) throws -> URL {
        let staging = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue164-staging-\(UUID().uuidString)")
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
            </Types>
            """, to: "[Content_Types].xml")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """, to: "_rels/.rels")
        try write(documentXML ?? Self.documentXML, to: "word/document.xml")

        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue164-\(UUID().uuidString).docx")
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

    private func extractDocumentXML(from docxURL: URL) throws -> Data {
        let archive = try Archive(url: docxURL, accessMode: .read)
        guard let entry = archive["word/document.xml"] else {
            XCTFail("word/document.xml missing from \(docxURL.lastPathComponent)")
            return Data()
        }
        var data = Data()
        _ = try archive.extract(entry) { data.append($0) }
        return data
    }

    /// Locates the first `<w:r>` child of the paragraph containing `marker`
    /// and returns its `<w:rPr>` node (nil if the run has none).
    private func watchedRunRPrNode(in documentXML: Data) throws -> XmlNode? {
        let tree = try XmlTreeReader.parse(documentXML)

        func containsMarker(_ node: XmlNode) -> Bool {
            if node.kind == .text, node.textContent.contains("alpha") { return true }
            return node.children.contains(where: containsMarker)
        }
        func findParagraph(_ node: XmlNode) -> XmlNode? {
            if node.kind == .element, node.localName == "p", containsMarker(node) {
                return node
            }
            for child in node.children {
                if let found = findParagraph(child) { return found }
            }
            return nil
        }
        guard let paragraph = findParagraph(tree.root) else {
            XCTFail("no <w:p> containing 'alpha' found")
            return nil
        }
        guard let run = paragraph.children.first(where: { $0.kind == .element && $0.localName == "r" }) else {
            XCTFail("watched paragraph has no <w:r>")
            return nil
        }
        return run.children.first { $0.kind == .element && $0.localName == "rPr" }
    }

    private func assertWatchedRunSurvived(
        in outputDocumentXML: Data,
        file: StaticString = #filePath, line: UInt = #line
    ) throws {
        let rPr = try XCTUnwrap(
            watchedRunRPrNode(in: outputDocumentXML),
            "the untouched run must still have a <w:rPr>", file: file, line: line)
        let names = rPr.children.filter { $0.kind == .element }.map(\.localName)
        XCTAssertTrue(names.contains("b"), "typed bold must survive", file: file, line: line)
        for expected in ["bCs", "iCs", "vanish"] {
            XCTAssertEqual(names.filter { $0 == expected }.count, 1,
                           "\(expected) must survive exactly once — vanish missing means hidden text became visible",
                           file: file, line: line)
        }
    }

    private func save(_ doc: WordDocument) throws -> URL {
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue164-out-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: outURL)
        return outURL
    }

    /// The issue's exact repro: `updateParagraph(at: 1, text:)` touches a
    /// DIFFERENT paragraph than the one carrying bCs/iCs/vanish.
    func testUpdateParagraphElsewherePreservesUnmodeledRPrChildrenOnUntouchedRun() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        try doc.updateParagraph(at: 1, text: "GAMMA")
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))

        let outURL = try save(doc)
        defer { try? FileManager.default.removeItem(at: outURL) }
        let outputXML = try extractDocumentXML(from: outURL)
        try assertWatchedRunSurvived(in: outputXML)
    }

    /// Same shape, using `replaceText` instead of `updateParagraph` — the
    /// issue names both as reproducing identically.
    func testReplaceTextElsewherePreservesUnmodeledRPrChildrenOnUntouchedRun() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        _ = try doc.replaceText(find: "gamma", with: "GAMMA")
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))

        let outURL = try save(doc)
        defer { try? FileManager.default.removeItem(at: outURL) }
        let outputXML = try extractDocumentXML(from: outURL)
        try assertWatchedRunSurvived(in: outputXML)
    }
}
