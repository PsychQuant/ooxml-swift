import XCTest
@testable import OOXMLSwift

/// v0.19.3+ regression tests for the round 2 verify findings on
/// PsychQuant/che-word-mcp#56 (https://github.com/PsychQuant/che-word-mcp/issues/56#issuecomment-4320157395).
/// This batch covers the **Hyperlink suite** (P0-1, P0-2, P0-3, P1-7) — the
/// other batches (sort path / revision / bookmark) ship in follow-up commits
/// inside the same v0.19.3 release.
final class Issue56RoundtripCompletenessTests: XCTestCase {

    // MARK: - P0-1: API path Hyperlink visual styling preserved

    /// `Hyperlink.external` populates a Run via `Run(text:)`. Pre-v0.19.3 the
    /// new `toXML()` walked `runs` directly and emitted a plain `<w:r><w:t>`
    /// without `<w:rStyle w:val="Hyperlink"/>`, `<w:color w:val="0563C1"/>`,
    /// or `<w:u w:val="single"/>` — so every hyperlink built via the 5 MCP
    /// `insert_*hyperlink` tools rendered without blue-underline styling.
    func testExternalHyperlinkAppliesHyperlinkVisualStyling() {
        let hl = Hyperlink.external(
            id: "h1",
            text: "click",
            url: "https://example.com",
            relationshipId: "rId1"
        )
        let xml = hl.toXML()

        XCTAssertTrue(
            xml.contains("w:rStyle w:val=\"Hyperlink\""),
            "API-built external hyperlink must include <w:rStyle w:val=\"Hyperlink\"/>. Output:\n\(xml)"
        )
        XCTAssertTrue(
            xml.contains("w:color w:val=\"0563C1\""),
            "API-built external hyperlink must include <w:color w:val=\"0563C1\"/>. Output:\n\(xml)"
        )
        XCTAssertTrue(
            xml.contains("w:u w:val=\"single\""),
            "API-built external hyperlink must include <w:u w:val=\"single\"/>. Output:\n\(xml)"
        )
    }

    /// Same contract for internal (anchor-based) hyperlinks built via
    /// `Hyperlink.internal(...)`.
    func testInternalHyperlinkAppliesHyperlinkVisualStyling() {
        let hl = Hyperlink.internal(
            id: "h2",
            text: "see section",
            bookmarkName: "section1"
        )
        let xml = hl.toXML()

        XCTAssertTrue(
            xml.contains("w:rStyle w:val=\"Hyperlink\""),
            "API-built internal hyperlink must include <w:rStyle w:val=\"Hyperlink\"/>. Output:\n\(xml)"
        )
        XCTAssertTrue(
            xml.contains("w:color w:val=\"0563C1\""),
            "API-built internal hyperlink must include <w:color w:val=\"0563C1\"/>. Output:\n\(xml)"
        )
        XCTAssertTrue(
            xml.contains("w:u w:val=\"single\""),
            "API-built internal hyperlink must include <w:u w:val=\"single\"/>. Output:\n\(xml)"
        )
    }

    // MARK: - P0-2: Reader preserves w:tgtFrame / w:docLocation via rawAttributes

    /// `parseHyperlink` listed `w:tgtFrame` and `w:docLocation` in
    /// `recognizedAttrs` (so they were skipped from `rawAttributes`), but the
    /// Hyperlink model has no typed field and `toXML()` never emits them.
    /// Net effect: source attributes silently dropped on round-trip. Fix:
    /// remove them from `recognizedAttrs` so they fall into rawAttributes,
    /// where the writer already emits them.
    func testHyperlinkTgtFrameRoundTripsThroughReaderAndWriter() throws {
        let xmlSrc = """
        <w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId5" w:tgtFrame="_blank" w:docLocation="frag1">
            <w:r><w:t>click</w:t></w:r>
        </w:hyperlink>
        """
        let element = try XMLElement(xmlString: xmlSrc)

        let hl = try DocxReader.parseHyperlink(
            from: element,
            relationships: RelationshipsCollection(),
            position: 0
        )

        XCTAssertEqual(
            hl.rawAttributes["w:tgtFrame"],
            "_blank",
            "w:tgtFrame must land in rawAttributes for round-trip. Got rawAttributes=\(hl.rawAttributes)"
        )
        XCTAssertEqual(
            hl.rawAttributes["w:docLocation"],
            "frag1",
            "w:docLocation must land in rawAttributes for round-trip. Got rawAttributes=\(hl.rawAttributes)"
        )

        let outXml = hl.toXML()
        XCTAssertTrue(
            outXml.contains("w:tgtFrame=\"_blank\""),
            "w:tgtFrame must round-trip through Writer. Output:\n\(outXml)"
        )
        XCTAssertTrue(
            outXml.contains("w:docLocation=\"frag1\""),
            "w:docLocation must round-trip through Writer. Output:\n\(outXml)"
        )
    }

    // MARK: - P0-3: Hyperlink internal child order preserved

    /// Source `<w:hyperlink><w:r>A</w:r><w:sdt>X</w:sdt><w:r>B</w:r></w:hyperlink>`
    /// must round-trip with the same A → SDT → B order. Pre-v0.19.3 Reader
    /// split into `runs=[A,B]` and `rawChildren=[<w:sdt>X</w:sdt>]`, then
    /// Writer emitted `<w:r>A</w:r><w:r>B</w:r><w:sdt>X</w:sdt>` — visible
    /// text order changed.
    func testHyperlinkChildOrderPreservedAcrossRunAndNonRunChildren() throws {
        let xmlSrc = """
        <w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId7"><w:r><w:t>A</w:t></w:r><w:sdt><w:sdtContent><w:r><w:t>X</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t>B</w:t></w:r></w:hyperlink>
        """
        let element = try XMLElement(xmlString: xmlSrc)

        let hl = try DocxReader.parseHyperlink(
            from: element,
            relationships: RelationshipsCollection(),
            position: 0
        )

        let outXml = hl.toXML()

        // The output must have A before SDT before B (positions in source order).
        guard let aPos = outXml.range(of: ">A<")?.lowerBound,
              let sdtPos = outXml.range(of: "<w:sdt")?.lowerBound,
              let bPos = outXml.range(of: ">B<")?.lowerBound else {
            XCTFail("Output missing one of A / SDT / B markers. Output:\n\(outXml)")
            return
        }
        XCTAssertLessThan(aPos, sdtPos, "A must precede SDT in round-trip. Output:\n\(outXml)")
        XCTAssertLessThan(sdtPos, bPos, "SDT must precede B in round-trip. Output:\n\(outXml)")
    }

    // MARK: - P0-6: Revision wrapper preserved for non-text content

    /// Source `<w:ins><w:r><w:tab/></w:r></w:ins>` (revision inserts a tab —
    /// not text) must round-trip with the `<w:ins>` wrapper intact. Pre-v0.19.3
    /// Reader only created a `Revision` entry when the inner concatenated text
    /// was non-empty; tab/break/image/fieldChar insertions yielded empty text,
    /// no Revision was added to `paragraph.revisions`, and the sort path's
    /// run-grouping fell back to emitting a naked `<w:r>` — wrapper silently
    /// dropped, the insertion looked accepted in Word post-save.
    func testRevisionWrapperPreservedForNonTextContent() throws {
        let xmlSrc = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:r><w:t>before</w:t></w:r>
            <w:ins w:id="5" w:author="alice" w:date="2026-04-25T00:00:00Z"><w:r><w:tab/></w:r></w:ins>
            <w:r><w:t>after</w:t></w:r>
        </w:p>
        """
        let element = try XMLElement(xmlString: xmlSrc)
        let paragraph = try DocxReader.parseParagraph(
            from: element,
            relationships: RelationshipsCollection(),
            styles: [],
            numbering: Numbering()
        )
        let xml = paragraph.toXML()

        XCTAssertTrue(
            xml.contains("<w:ins"),
            "Revision wrapper must round-trip even when inserted content is non-text. Output:\n\(xml)"
        )
        XCTAssertTrue(
            xml.contains("</w:ins>"),
            "Revision wrapper must close properly. Output:\n\(xml)"
        )
        XCTAssertTrue(
            xml.contains("<w:tab"),
            "Inserted <w:tab/> must round-trip inside the wrapper. Output:\n\(xml)"
        )
    }

    // MARK: - P0-7: Revision wrapper preserves nested non-run children

    /// Source `<w:ins><w:hyperlink r:id="rId1"><w:r>foo</w:r></w:hyperlink></w:ins>`
    /// — Track Changes mode where the user inserted a hyperlink — must
    /// round-trip with the hyperlink intact inside the `<w:ins>` wrapper.
    /// Pre-v0.19.3 Reader's `for insRun in childElement.elements(forName: "w:r")`
    /// only descended into direct `<w:r>` children, so the hyperlink (and its
    /// inner runs) was silently dropped on parse.
    func testRevisionWrapperPreservesNestedHyperlink() throws {
        let xmlSrc = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <w:r><w:t>before</w:t></w:r>
            <w:ins w:id="9" w:author="alice" w:date="2026-04-25T00:00:00Z"><w:hyperlink r:id="rId7"><w:r><w:t>linked-text</w:t></w:r></w:hyperlink></w:ins>
            <w:r><w:t>after</w:t></w:r>
        </w:p>
        """
        let element = try XMLElement(xmlString: xmlSrc)
        let paragraph = try DocxReader.parseParagraph(
            from: element,
            relationships: RelationshipsCollection(),
            styles: [],
            numbering: Numbering()
        )
        let xml = paragraph.toXML()

        XCTAssertTrue(
            xml.contains("<w:ins"),
            "Revision wrapper must round-trip. Output:\n\(xml)"
        )
        XCTAssertTrue(
            xml.contains("<w:hyperlink"),
            "Nested hyperlink inside <w:ins> must round-trip. Output:\n\(xml)"
        )
        XCTAssertTrue(
            xml.contains("linked-text"),
            "Nested hyperlink text must survive round-trip. Output:\n\(xml)"
        )
    }

    // MARK: - P1-7: Hyperlink.id unique across duplicate r:id

    /// Two `<w:hyperlink>` elements with the same `r:id` (legitimate when two
    /// links share a relationship target — e.g., two "click here" anchors for
    /// the same URL) must parse into Hyperlink instances with **distinct**
    /// `id` fields. The legacy `id = rId ?? anchor ?? "hl-\(position)"`
    /// returned the same id for both, breaking MCP tools that find/edit/
    /// delete hyperlinks by `id`.
    // MARK: - P0-4: Sort path emits contentControls

    /// Source-loaded paragraph carrying any positioned-child (so it routes to
    /// `toXMLSortedByPosition`) plus a `ContentControl` on `paragraph.contentControls`
    /// must emit the SDT in the output. Pre-v0.19.3 the sort path simply
    /// dropped `contentControls` (the doc-comment's "emit AFTER" claim was
    /// not implemented), silently losing every paragraph-level SDT on save.
    func testSortPathEmitsContentControls() {
        var para = Paragraph(text: "before")
        para.bookmarkMarkers = [
            BookmarkRangeMarker(kind: .start, id: 1, position: 0)
        ]
        para.contentControls = [
            ContentControl.richText(tag: "myTag", alias: "Field A", content: "Hello"),
        ]

        let xml = para.toXML()

        XCTAssertTrue(
            xml.contains("<w:sdt>"),
            "Sort path must emit <w:sdt>. Output:\n\(xml)"
        )
        XCTAssertTrue(
            xml.contains("Hello"),
            "Sort path must preserve SDT content text. Output:\n\(xml)"
        )
    }

    // MARK: - P0-5: Sort path emits commentIds / footnoteIds / endnoteIds / hasPageBreak

    /// Source-loaded paragraph (sort path) with `insert_comment` (legacy
    /// `commentIds` collection) must emit `<w:commentRangeStart>` /
    /// `<w:commentRangeEnd>` / `<w:commentReference>`. Pre-v0.19.3 these
    /// silently dropped on save when the paragraph already had any
    /// positioned-child.
    func testSortPathEmitsLegacyCommentIds() {
        var para = Paragraph(text: "anchor")
        para.bookmarkMarkers = [
            BookmarkRangeMarker(kind: .start, id: 1, position: 0)
        ]
        para.commentIds = [42]

        let xml = para.toXML()

        XCTAssertTrue(
            xml.contains("<w:commentRangeStart w:id=\"42\""),
            "Sort path must emit commentRangeStart. Output:\n\(xml)"
        )
        XCTAssertTrue(
            xml.contains("<w:commentRangeEnd w:id=\"42\""),
            "Sort path must emit commentRangeEnd. Output:\n\(xml)"
        )
        XCTAssertTrue(
            xml.contains("<w:commentReference w:id=\"42\""),
            "Sort path must emit commentReference. Output:\n\(xml)"
        )
    }

    /// Same contract for footnote refs, endnote refs, and hasPageBreak.
    func testSortPathEmitsLegacyFootnoteEndnoteAndPageBreak() {
        var para = Paragraph(text: "anchor")
        para.bookmarkMarkers = [
            BookmarkRangeMarker(kind: .start, id: 1, position: 0)
        ]
        para.footnoteIds = [7]
        para.endnoteIds = [9]
        para.hasPageBreak = true

        let xml = para.toXML()

        XCTAssertTrue(
            xml.contains("<w:footnoteReference w:id=\"7\""),
            "Sort path must emit footnoteReference. Output:\n\(xml)"
        )
        XCTAssertTrue(
            xml.contains("<w:endnoteReference w:id=\"9\""),
            "Sort path must emit endnoteReference. Output:\n\(xml)"
        )
        XCTAssertTrue(
            xml.contains("<w:br w:type=\"page\""),
            "Sort path must emit page break. Output:\n\(xml)"
        )
    }

    // MARK: - P0-8: hasSourcePositionedChildren routes hyperlink-only paragraphs

    /// Source paragraph `<w:r>A</w:r><w:hyperlink>L</w:hyperlink><w:r>B</w:r>`
    /// — runs at positions 1 and 3, hyperlink at position 2 — must round-trip
    /// as A → L → B. Pre-v0.19.3 the `hasSourcePositionedChildren` predicate
    /// excluded hyperlinks, so this paragraph went to the legacy path which
    /// emits all runs first then all hyperlinks → A B L (visible text order
    /// changed).
    /// v0.19.5+ (#56 R5-CONT P0 #6): updated to use realistic source positions
    /// (1, 2, 3) instead of (0, 1, 2). Per R5 P0 #2, `position == 0` is now
    /// reserved for the API-built sentinel — runs/hyperlinks with position 0
    /// emit at end-of-paragraph, not at "logical first". Reader-loaded
    /// children always start at 1 (DocxReader.parseParagraph childPosition).
    func testParagraphWithInterleavedRunsAndHyperlinkRoutesToSortPath() {
        var runA = Run(text: "A")
        runA.position = 1
        var runB = Run(text: "B")
        runB.position = 3
        var hl = Hyperlink.external(
            id: "h1",
            text: "L",
            url: "https://example.com",
            relationshipId: "rId1"
        )
        hl.position = 2

        var para = Paragraph(runs: [runA, runB])
        para.hyperlinks = [hl]

        let xml = para.toXML()

        guard let aPos = xml.range(of: ">A<")?.lowerBound,
              let lPos = xml.range(of: ">L<")?.lowerBound,
              let bPos = xml.range(of: ">B<")?.lowerBound else {
            XCTFail("Output missing one of A / L / B markers. Output:\n\(xml)")
            return
        }
        XCTAssertLessThan(aPos, lPos, "A must precede hyperlink L. Output:\n\(xml)")
        XCTAssertLessThan(lPos, bPos, "Hyperlink L must precede B. Output:\n\(xml)")
    }

    // MARK: - P1-1: nextBookmarkId calibrated from source max

    /// After Reader loads a document containing a bookmark with `w:id="10"`,
    /// the next `addBookmark` API call must allocate id 11 (or higher) — not
    /// id 1 (which would collide with the source bookmark and either crash
    /// Word's schema validation or produce undefined behavior). Pre-v0.19.3
    /// `nextBookmarkId` initialized to 1 and Reader never bumped it; F2 made
    /// the silent-drop bug into a silent-overwrite bug.
    func testReaderCalibratesNextBookmarkIdFromSourceMax() throws {
        // Build a minimal valid docx with a single source bookmark id=10.
        let tmpDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("ooxml-test-\(UUID().uuidString)")
        let extracted = tmpDir.appendingPathComponent("docx")
        try FileManager.default.createDirectory(at: extracted.appendingPathComponent("_rels"), withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: extracted.appendingPathComponent("word/_rels"), withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: tmpDir) }

        let contentTypes = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>
        """
        try contentTypes.write(to: extracted.appendingPathComponent("[Content_Types].xml"), atomically: true, encoding: .utf8)

        let rootRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """
        try rootRels.write(to: extracted.appendingPathComponent("_rels/.rels"), atomically: true, encoding: .utf8)

        let documentXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:bookmarkStart w:id="10" w:name="src"/>
                    <w:r><w:t>anchor</w:t></w:r>
                    <w:bookmarkEnd w:id="10"/>
                </w:p>
            </w:body>
        </w:document>
        """
        try documentXml.write(to: extracted.appendingPathComponent("word/document.xml"), atomically: true, encoding: .utf8)

        // Zip into a .docx via the package writer's ZipHelper (file system-based).
        let docxURL = tmpDir.appendingPathComponent("test.docx")
        try ZipHelper.zip(extracted, to: docxURL)

        var doc = try DocxReader.read(from: docxURL)
        defer { doc.close() }

        let newId = try doc.insertBookmark(name: "added")
        XCTAssertGreaterThan(
            newId, 10,
            "nextBookmarkId must be calibrated past source max (10). Got \(newId)."
        )
    }

    // MARK: - P1-4: API-built paragraph addBookmark preserves run-wrap semantics

    /// `var doc = WordDocument(); addParagraph("Hello"); addBookmark("foo")` —
    /// the bookmark must wrap "Hello" via `<w:bookmarkStart/><w:r>Hello</w:r><w:bookmarkEnd/>`,
    /// matching v3.12.0 behavior. F2's `appendBookmarkSyncingMarkers` blindly
    /// added markers to every paragraph, forcing API-built paragraphs onto the
    /// sort path which emits a zero-width point bookmark at the paragraph end —
    /// silent semantic regression for callers expecting the bookmark to span
    /// the run text.
    func testAddBookmarkOnApiBuiltParagraphPreservesRunWrapSemantics() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "Hello"))
        _ = try doc.insertBookmark(name: "foo")

        // Locate the rendered paragraph XML.
        guard let lastPara = doc.body.children.last,
              case .paragraph(let para) = lastPara else {
            XCTFail("Expected a paragraph at body tail")
            return
        }
        let xml = para.toXML()

        guard let startPos = xml.range(of: "<w:bookmarkStart")?.lowerBound,
              let runPos = xml.range(of: ">Hello<")?.lowerBound,
              let endPos = xml.range(of: "<w:bookmarkEnd")?.lowerBound else {
            XCTFail("Output missing bookmarkStart/run/bookmarkEnd. Output:\n\(xml)")
            return
        }
        XCTAssertLessThan(startPos, runPos, "bookmarkStart must precede the wrapped run. Output:\n\(xml)")
        XCTAssertLessThan(runPos, endPos, "Wrapped run must precede bookmarkEnd. Output:\n\(xml)")
    }

    func testParsedHyperlinksHaveUniqueIdEvenWhenSharingRelationshipId() throws {
        let xml1 = """
        <w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId5"><w:r><w:t>A</w:t></w:r></w:hyperlink>
        """
        let xml2 = """
        <w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId5"><w:r><w:t>B</w:t></w:r></w:hyperlink>
        """
        let el1 = try XMLElement(xmlString: xml1)
        let el2 = try XMLElement(xmlString: xml2)

        let hl1 = try DocxReader.parseHyperlink(
            from: el1,
            relationships: RelationshipsCollection(),
            position: 3
        )
        let hl2 = try DocxReader.parseHyperlink(
            from: el2,
            relationships: RelationshipsCollection(),
            position: 7
        )

        XCTAssertNotEqual(
            hl1.id, hl2.id,
            "Two hyperlinks sharing r:id but at different positions must get distinct ids. Got both = \"\(hl1.id)\""
        )
    }
}
