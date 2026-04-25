import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Regression tests for the 4 blocking findings (F1–F4) from the
/// PsychQuant/che-word-mcp#56 verification report
/// (https://github.com/PsychQuant/che-word-mcp/issues/56#issuecomment-4319691177).
/// All four were silent corruption bugs that pre-v0.19.2 round-tests didn't
/// catch because the smoke tests asserted concat-text SHA256 + element-count
/// parity — neither covers run-property loss, marker desync, revision wrapper
/// drop, or container namespace strip.
final class Issue56FollowupTests: XCTestCase {

    // MARK: - F1: Hyperlink.toXML iterates runs + emits raw fields

    /// Source-loaded hyperlink with multi-run inner content (different
    /// formatting per run + an unmodeled vendor attribute) must round-trip
    /// every inner `RunProperties` and the `rawAttributes` entry.
    /// Pre-v0.19.2: `Hyperlink.toXML()` collapsed runs to a single hardcoded
    /// styled `<w:r>` and dropped `rawAttributes` entirely.
    func testHyperlinkInnerRunPropertiesAndRawAttributesRoundTrip() throws {
        let bold = Run(text: "Bold ", properties: RunProperties(bold: true))
        let plain = Run(text: "and plain", properties: RunProperties())
        let hyperlink = Hyperlink(
            id: "h1",
            runs: [bold, plain],
            relationshipId: "rId99",
            tooltip: "Click me",
            rawAttributes: ["w:tgtFrame": "_blank", "w:docLocation": "section1"],
            rawChildren: [],
            position: 0
        )

        let xml = hyperlink.toXML()

        // Inner run properties survive — both <w:b/> from the bold run AND a
        // plain `<w:r>` for the second run must be present.
        XCTAssertTrue(
            xml.contains("<w:b/>") || xml.contains("<w:b "),
            "Hyperlink.toXML must preserve <w:b/> from inner run RunProperties. Output:\n\(xml)"
        )
        // Both runs must appear, not collapsed.
        let runCount = xml.components(separatedBy: "<w:r>").count - 1
        XCTAssertGreaterThanOrEqual(runCount, 2, "Multi-run hyperlink must emit >=2 <w:r>. Got \(runCount). Output:\n\(xml)")

        // rawAttributes must surface on the open tag.
        XCTAssertTrue(
            xml.contains("w:tgtFrame=\"_blank\""),
            "Hyperlink.toXML must emit rawAttributes[w:tgtFrame]. Output:\n\(xml)"
        )
        XCTAssertTrue(
            xml.contains("w:docLocation=\"section1\""),
            "Hyperlink.toXML must emit rawAttributes[w:docLocation]. Output:\n\(xml)"
        )

        // Concatenated text must match (the v0.19.0 contract).
        XCTAssertTrue(
            xml.contains("Bold "), "Bold run text lost. Output:\n\(xml)"
        )
        XCTAssertTrue(
            xml.contains("and plain"), "Plain run text lost. Output:\n\(xml)"
        )
    }

    /// Empty-runs (API-built path before populating runs) falls back to the
    /// hardcoded styled-run template so output stays valid OOXML.
    func testHyperlinkEmptyRunsFallsBackToStyledRunTemplate() {
        var hl = Hyperlink(id: "h2", text: "ignored", anchor: "bookmark1")
        hl.runs = []  // simulate caller blanking runs
        let xml = hl.toXML()
        XCTAssertTrue(xml.contains("w:rStyle"), "Empty-runs fallback must emit hardcoded Hyperlink rStyle. Output:\n\(xml)")
        XCTAssertTrue(xml.contains("w:anchor=\"bookmark1\""), "anchor attribute lost. Output:\n\(xml)")
    }

    // MARK: - F2: addBookmark / deleteBookmark sync bookmarkMarkers

    /// `addBookmark` on a paragraph that already has source-loaded markers
    /// (i.e., `bookmarkMarkers` non-empty so it goes through the
    /// sort-by-position emit path) must also append matching
    /// `BookmarkRangeMarker` entries — otherwise the new bookmark is
    /// silently dropped on save (the typed `bookmarks` entry is created
    /// but the writer only emits markers).
    func testAddBookmarkSyncsBookmarkMarkers() throws {
        var doc = WordDocument()
        // Stage a paragraph with a pre-existing source-loaded bookmark
        // (id=10) so `bookmarkMarkers` is non-empty before we mutate.
        var para = Paragraph(text: "Hello")
        para.bookmarks.append(Bookmark(id: 10, name: "existing"))
        para.bookmarkMarkers.append(BookmarkRangeMarker(kind: .start, id: 10, position: 0))
        para.bookmarkMarkers.append(BookmarkRangeMarker(kind: .end, id: 10, position: 1))
        doc.appendParagraph(para)

        let newId = try doc.insertBookmark(name: "newbookmark", at: 0)

        // Read back the paragraph and verify the markers were synced.
        let paragraphs = doc.getParagraphs()
        XCTAssertEqual(paragraphs.count, 1)
        let updated = paragraphs[0]

        // bookmarks list has both entries.
        XCTAssertEqual(updated.bookmarks.count, 2, "bookmarks list must have both old + new entries")
        XCTAssertTrue(updated.bookmarks.contains { $0.id == newId && $0.name == "newbookmark" })

        // bookmarkMarkers has 4 entries (2 pairs: existing + new).
        XCTAssertEqual(
            updated.bookmarkMarkers.count, 4,
            "bookmarkMarkers must contain both old pair AND new pair (2 + 2 = 4)"
        )
        XCTAssertEqual(
            updated.bookmarkMarkers.filter { $0.id == newId }.count, 2,
            "New bookmark must have BOTH start AND end markers"
        )
        XCTAssertTrue(
            updated.bookmarkMarkers.contains { $0.id == newId && $0.kind == .start },
            "Missing start marker for new bookmark"
        )
        XCTAssertTrue(
            updated.bookmarkMarkers.contains { $0.id == newId && $0.kind == .end },
            "Missing end marker for new bookmark"
        )
    }

    /// `deleteBookmark` must remove BOTH the typed `bookmarks` entry AND the
    /// matching `bookmarkMarkers` entries. Pre-v0.19.2: only `bookmarks`
    /// was cleared, leaving zombie `<w:bookmarkStart w:id="N" w:name=""/>`
    /// emit because the marker survived but the name lookup fell through to
    /// `?? ""`.
    func testDeleteBookmarkSyncsBookmarkMarkersRemoval() throws {
        var doc = WordDocument()
        var para = Paragraph(text: "Hello")
        para.bookmarks.append(Bookmark(id: 42, name: "tobedeleted"))
        para.bookmarkMarkers.append(BookmarkRangeMarker(kind: .start, id: 42, position: 0))
        para.bookmarkMarkers.append(BookmarkRangeMarker(kind: .end, id: 42, position: 1))
        doc.appendParagraph(para)

        try doc.deleteBookmark(name: "tobedeleted")

        let paragraphs = doc.getParagraphs()
        XCTAssertEqual(paragraphs.count, 1)
        let updated = paragraphs[0]
        XCTAssertTrue(updated.bookmarks.isEmpty, "bookmarks must be cleared")
        XCTAssertTrue(
            updated.bookmarkMarkers.isEmpty,
            "bookmarkMarkers for deleted bookmark must also be cleared. Found: \(updated.bookmarkMarkers)"
        )
    }

    // MARK: - F3: <w:ins> / <w:moveFrom> / <w:moveTo> position + revisionId on inner runs

    /// Reader must assign `position = childPosition` AND `revisionId = revId`
    /// to every run extracted from `<w:ins>`. Pre-v0.19.2: position defaulted
    /// to 0 (sort path put inserted runs at paragraph front), revisionId was
    /// never set (sort path emitted bare `<w:r>` with no `<w:ins>` wrapper).
    func testReaderAssignsPositionAndRevisionIdToInsRuns() throws {
        // Build a minimal docx with: <w:r>before</w:r><w:ins ...><w:r>inserted</w:r></w:ins><w:r>after</w:r>
        let url = try buildRevisionFixture(insertedText: "inserted-text-marker")
        defer { try? FileManager.default.removeItem(at: url) }

        let doc = try DocxReader.read(from: url)
        defer { var d = doc; d.close() }
        let paragraphs = doc.getParagraphs()
        XCTAssertEqual(paragraphs.count, 1)
        let para = paragraphs[0]

        // Find the run carrying our inserted marker text.
        let insertedRun = para.runs.first { $0.text == "inserted-text-marker" }
        XCTAssertNotNil(insertedRun, "Inserted run must be parsed into paragraph.runs. All runs: \(para.runs.map { $0.text })")
        guard let ir = insertedRun else { return }

        // F3a: position must be set (non-zero, since this is the 2nd <w:p> child).
        XCTAssertNotEqual(
            ir.position, 0,
            "Inserted run position must be set to source childPosition (>0 since it's not the first child). Got 0."
        )

        // F3b: revisionId must be set so sort-by-position emit can re-wrap.
        XCTAssertNotNil(ir.revisionId, "Inserted run revisionId must be set so sort path can re-wrap with <w:ins>")
    }

    /// End-to-end F3: source-loaded paragraph with a `<w:ins>` wrapper must
    /// round-trip the wrapper (Writer's sort-by-position path must regenerate
    /// `<w:ins>...</w:ins>` from grouped consecutive same-revisionId runs).
    func testInsWrapperRoundTripsThroughSortByPositionEmit() throws {
        let url = try buildRevisionFixture(insertedText: "inserted-text-marker")
        defer { try? FileManager.default.removeItem(at: url) }

        var doc = try DocxReader.read(from: url)
        defer { doc.close() }
        // Force document.xml regeneration through Writer.
        doc.markPartDirty("word/document.xml")
        let savedURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("rev-saved-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: savedURL) }
        try DocxWriter.write(doc, to: savedURL)

        // Read the saved document.xml as raw string.
        let savedDocXML = try readDocumentXMLString(from: savedURL)
        XCTAssertTrue(
            savedDocXML.contains("<w:ins"),
            "Saved document.xml must contain <w:ins> wrapper after round-trip. F3 failed.\nSaved XML:\n\(savedDocXML)"
        )
        XCTAssertTrue(
            savedDocXML.contains("inserted-text-marker"),
            "Inserted text must survive. Saved XML:\n\(savedDocXML)"
        )
    }

    // MARK: - F4: Container root namespace preservation

    /// Header containing extension namespaces (`mc`, `wp`, `w14`) must round-
    /// trip those declarations on `<w:hdr>`. Pre-v0.19.2: Header.toXML used a
    /// hardcoded 5-namespace template (`w`/`r`/`v`/`o`/`w10`) that silently
    /// dropped any additional declarations from source.
    func testHeaderRootAttributesRoundTrip() {
        var header = Header(
            id: "rId99",
            paragraphs: [Paragraph(text: "header text")],
            type: .default,
            originalFileName: "header1.xml",
            rootAttributes: [
                "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                "xmlns:v": "urn:schemas-microsoft-com:vml",
                "xmlns:o": "urn:schemas-microsoft-com:office:office",
                "xmlns:w10": "urn:schemas-microsoft-com:office:word",
                "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
                "xmlns:wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
                "mc:Ignorable": "w14",
            ]
        )
        // suppress unused mutation warning
        _ = header

        let xml = header.toXML()

        // Every captured attribute must appear on the open tag.
        for prefix in ["w", "r", "v", "o", "w10", "mc", "wp", "w14"] {
            XCTAssertTrue(
                xml.contains("xmlns:\(prefix)="),
                "Header.toXML must emit xmlns:\(prefix). Output:\n\(xml)"
            )
        }
        XCTAssertTrue(
            xml.contains("mc:Ignorable=\"w14\""),
            "Header.toXML must emit mc:Ignorable. Output:\n\(xml)"
        )
    }

    /// API-built header (rootAttributes empty) falls back to the hardcoded
    /// 5-namespace template. This is the no-regression backstop for F4.
    func testHeaderEmptyRootAttributesFallsBackToFiveNamespaceTemplate() {
        let header = Header(id: "rId99", paragraphs: [], type: .default)
        let xml = header.toXML()
        for prefix in ["w", "r", "v", "o", "w10"] {
            XCTAssertTrue(xml.contains("xmlns:\(prefix)="), "Default template missing xmlns:\(prefix). Output:\n\(xml)")
        }
    }

    /// FootnotesCollection rootAttributes round-trip. Default template (when
    /// rootAttributes empty) is `xmlns:w` + `xmlns:r` only.
    func testFootnotesRootAttributesRoundTrip() {
        let collection = FootnotesCollection(
            footnotes: [],
            rootAttributes: [
                "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
            ]
        )
        let xml = collection.toXML()
        XCTAssertTrue(xml.contains("xmlns:w14="), "FootnotesCollection.toXML must emit xmlns:w14 from rootAttributes. Output:\n\(xml)")
    }

    // MARK: - Test fixture helpers

    /// Build a minimal docx with one paragraph containing
    /// `<w:r>before</w:r><w:ins ...><w:r>inserted</w:r></w:ins><w:r>after</w:r>`
    /// for F3 round-trip tests. Only `[Content_Types].xml` + 2 rels + the
    /// document — minimum viable for DocxReader.
    private func buildRevisionFixture(insertedText: String) throws -> URL {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("rev-fixture-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: stagingDir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: stagingDir) }

        func write(_ s: String, to rel: String) throws {
            let url = stagingDir.appendingPathComponent(rel)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(),
                withIntermediateDirectories: true
            )
            try s.write(to: url, atomically: true, encoding: .utf8)
        }

        try write(#"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>
        """#, to: "[Content_Types].xml")

        try write(#"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """#, to: "_rels/.rels")

        try write(#"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        </Relationships>
        """#, to: "word/_rels/document.xml.rels")

        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:r><w:t xml:space="preserve">before </w:t></w:r>
              <w:ins w:id="1" w:author="reviewer" w:date="2026-04-26T00:00:00Z">
                <w:r><w:t xml:space="preserve">\(insertedText)</w:t></w:r>
              </w:ins>
              <w:r><w:t xml:space="preserve"> after</w:t></w:r>
            </w:p>
          </w:body>
        </w:document>
        """
        try write(documentXML, to: "word/document.xml")

        let outputURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("rev-fixture-\(UUID().uuidString).docx")
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

    private func readDocumentXMLString(from url: URL) throws -> String {
        let archive = try Archive(url: url, accessMode: .read)
        guard let entry = archive["word/document.xml"] else {
            throw NSError(domain: "Issue56FollowupTests", code: 1, userInfo: [NSLocalizedDescriptionKey: "no word/document.xml"])
        }
        var data = Data()
        _ = try archive.extract(entry) { data.append($0) }
        return String(data: data, encoding: .utf8) ?? ""
    }
}
