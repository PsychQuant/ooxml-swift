import XCTest
@testable import OOXMLSwift

/// v0.19.4+ regression tests for the round 3 verify findings on
/// PsychQuant/che-word-mcp#56 (https://github.com/PsychQuant/che-word-mcp/issues/56#issuecomment-4321007538).
///
/// R3 found 6 new P0 regressions caused by the R2 (v0.19.3) fixes themselves.
/// Each test in this file targets one R3-NEW finding and follows the per-task
/// verify discipline established in the spectra change
/// `che-word-mcp-issue-56-r3-stack-completion`:
///
/// - R3-NEW-1: Hyperlink mutation API SHALL round-trip on source-loaded hyperlinks
/// - R3-NEW-2: ContentControl SHALL expose a position: Int field and emit in source order
/// - R3-NEW-3: insertComment SHALL emit anchor markers on source paragraphs with existing comment markers
/// - R3-NEW-4: Mixed-content revision wrappers SHALL populate both raw and typed representations
/// - R3-NEW-5: nextBookmarkId calibration SHALL scan all bookmark-bearing document parts
/// - R3-NEW-6: Direct-emit XML attribute values SHALL be escaped to prevent injection
final class Issue56R3StackTests: XCTestCase {

    // MARK: - R3-NEW-1: Hyperlink mutation round-trip on source-loaded hyperlinks

    /// Spec scenario "replaceText on source-loaded hyperlink emits new text":
    /// load doc with `<w:hyperlink r:id="rId1"><w:r><w:t>old</w:t></w:r></w:hyperlink>`,
    /// call `replaceText("old", with: "new")`, re-emit, expect "new" in the
    /// hyperlink and no "old".
    ///
    /// Pre-v0.19.4: Reader populates BOTH `runs` AND `children`. Writer's
    /// "if !children.isEmpty use children else runs" priority makes `children`
    /// the source of truth, but `replaceText` only mutates `runs`. The saved
    /// XML re-emits the original "old" — silent edit-failure regression.
    func testReplaceText_OnSourceLoadedHyperlink_EmitsNewText() throws {
        let docxURL = try Self.buildHyperlinkSourceFixture(text: "old")
        defer { try? FileManager.default.removeItem(at: docxURL) }

        var doc = try DocxReader.read(from: docxURL)
        defer { doc.close() }

        let count = try doc.replaceText(find: "old", with: "new")
        XCTAssertEqual(count, 1, "replaceText should report 1 substitution")

        // Locate the paragraph carrying the hyperlink and re-emit.
        guard case .paragraph(let para) = doc.body.children[0] else {
            XCTFail("Expected paragraph at body[0]")
            return
        }
        let xml = para.toXML()

        // Bug: writer prefers `children` (still says "old"); fix: writer prefers `runs`.
        XCTAssertTrue(
            xml.contains(">new<"),
            "Re-emitted hyperlink XML must contain 'new' (the replacement). Output:\n\(xml)"
        )
        XCTAssertFalse(
            xml.contains(">old<"),
            "Re-emitted hyperlink XML must NOT contain 'old' (the original). Output:\n\(xml)"
        )
    }

    /// Spec scenario "updateHyperlink text setter on source-loaded hyperlink emits new text":
    /// same setup, but mutate via `updateHyperlink(hyperlinkId:text:)` instead of
    /// `replaceText`. Both APIs feed through `Hyperlink.text` setter / `runs`
    /// mutation, so both fail under the v0.19.3 priority order.
    func testUpdateHyperlinkText_OnSourceLoadedHyperlink_EmitsNewText() throws {
        let docxURL = try Self.buildHyperlinkSourceFixture(text: "old")
        defer { try? FileManager.default.removeItem(at: docxURL) }

        var doc = try DocxReader.read(from: docxURL)
        defer { doc.close() }

        // Reader assigns id from r:id + position; for our single hyperlink it's "rId1@0".
        guard case .paragraph(let para) = doc.body.children[0],
              let firstHyperlink = para.hyperlinks.first else {
            XCTFail("Expected source-loaded hyperlink in body[0]")
            return
        }
        try doc.updateHyperlink(hyperlinkId: firstHyperlink.id, text: "Updated")

        guard case .paragraph(let updatedPara) = doc.body.children[0] else {
            XCTFail("Expected paragraph at body[0] after update")
            return
        }
        let xml = updatedPara.toXML()

        XCTAssertTrue(
            xml.contains(">Updated<"),
            "Re-emitted hyperlink XML must contain 'Updated' after updateHyperlink. Output:\n\(xml)"
        )
        XCTAssertFalse(
            xml.contains(">old<"),
            "Re-emitted hyperlink XML must NOT contain 'old' after updateHyperlink. Output:\n\(xml)"
        )
    }

    // MARK: - R3-NEW-2: ContentControl SHALL expose a position: Int field and emit in source order

    /// Spec scenario "SDT between runs round-trips at its source position":
    /// load doc with `<w:p><w:r><w:t>A</w:t></w:r><w:sdt>X</w:sdt><w:r><w:t>B</w:t></w:r></w:p>`,
    /// re-emit, expect output order A, X (inside sdt), B.
    ///
    /// Pre-v0.19.4: ContentControl has no `position` field. DocxReader.parseParagraph
    /// appends source-loaded SDTs to `paragraph.contentControls` without
    /// passing `childPosition`, and `Paragraph.toXMLSortedByPosition` emits
    /// `contentControls` unconditionally in the post-content section after
    /// the sorted positioned-entry list. Result: SDT moves to end (A→B→SDT).
    func testParagraphLevelSDT_BetweenRuns_RoundTripsAtSourcePosition() throws {
        let docxURL = try Self.buildSDTBetweenRunsFixture()
        defer { try? FileManager.default.removeItem(at: docxURL) }

        var doc = try DocxReader.read(from: docxURL)
        defer { doc.close() }

        guard case .paragraph(let para) = doc.body.children[0] else {
            XCTFail("Expected paragraph at body[0]")
            return
        }
        let xml = para.toXML()

        guard let aPos = xml.range(of: ">A<")?.lowerBound,
              let sdtPos = xml.range(of: "<w:sdt")?.lowerBound,
              let bPos = xml.range(of: ">B<")?.lowerBound else {
            XCTFail("Output missing A / sdt / B markers. Output:\n\(xml)")
            return
        }
        XCTAssertLessThan(aPos, sdtPos,
                          "Run 'A' must precede <w:sdt> in source order. Output:\n\(xml)")
        XCTAssertLessThan(sdtPos, bPos,
                          "<w:sdt> must precede run 'B' in source order. Pre-fix the SDT moved to end (A→B→SDT). Output:\n\(xml)")
    }

    // MARK: - R3-NEW-3: insertComment SHALL emit anchor markers on source paragraphs with existing comment markers

    /// Spec scenario "insertComment on source paragraph with existing comments emits new anchors":
    /// load doc with `<w:commentRangeStart w:id="3"/>` markers, call insertComment,
    /// expect both id=3 markers and id=4 markers in re-emitted XML.
    ///
    /// Pre-v0.19.4: Paragraph.toXMLSortedByPosition gates the legacy commentIds
    /// emit (start, end, reference) on `commentRangeMarkers.isEmpty`. When source
    /// has any commentRangeMarkers, the new commentId from insertComment lands
    /// in `commentIds` but no `<w:commentRangeStart>` / `<w:commentRangeEnd>` /
    /// `<w:commentReference>` are emitted for it — comment side-bar shows the
    /// comment but no scope highlight.
    func testInsertComment_OnSourceParagraphWithExistingComments_EmitsNewAnchors() throws {
        let docxURL = try Self.buildExistingCommentsFixture()
        defer { try? FileManager.default.removeItem(at: docxURL) }

        var doc = try DocxReader.read(from: docxURL)
        defer { doc.close() }

        let newCommentId = try doc.insertComment(text: "second", author: "Tester", paragraphIndex: 0)
        XCTAssertNotEqual(newCommentId, 3, "new comment id must differ from source id 3")

        guard case .paragraph(let para) = doc.body.children[0] else {
            XCTFail("Expected paragraph at body[0]")
            return
        }
        let xml = para.toXML()

        // Original id=3 markers must survive.
        XCTAssertTrue(
            xml.contains("<w:commentRangeStart w:id=\"3\""),
            "Original commentRangeStart id=3 must survive. Output:\n\(xml)"
        )
        XCTAssertTrue(
            xml.contains("<w:commentRangeEnd w:id=\"3\""),
            "Original commentRangeEnd id=3 must survive. Output:\n\(xml)"
        )

        // New commentId must produce matching start + end + reference markers.
        let newIdStr = String(newCommentId)
        XCTAssertTrue(
            xml.contains("<w:commentRangeStart w:id=\"\(newIdStr)\""),
            "New commentRangeStart id=\(newIdStr) must be emitted alongside source markers. Output:\n\(xml)"
        )
        XCTAssertTrue(
            xml.contains("<w:commentRangeEnd w:id=\"\(newIdStr)\""),
            "New commentRangeEnd id=\(newIdStr) must be emitted. Output:\n\(xml)"
        )
        XCTAssertTrue(
            xml.contains("<w:commentReference w:id=\"\(newIdStr)\""),
            "New commentReference id=\(newIdStr) must be emitted. Output:\n\(xml)"
        )
    }

    // MARK: - R3-NEW-4: Mixed-content revision wrappers SHALL populate both raw and typed representations

    /// Spec scenario "w:ins wrapping a hyperlink populates both representations":
    /// load doc with `<w:ins w:id="5" w:author="Alice"...><w:hyperlink ...><w:r><w:t>linked</w:t></w:r></w:hyperlink></w:ins>`,
    /// expect both `paragraph.unrecognizedChildren` (with raw `<w:ins` XML) AND
    /// `paragraph.revisions` (with id=5, type=.insertion, author="Alice").
    ///
    /// Pre-v0.19.4: P0-7 hasNonRunChild branch appended to unrecognizedChildren
    /// then `break` — skipped the typed Revision append. MCP get_revisions /
    /// accept_revision / reject_revision tools see no Revision but the raw XML
    /// still emits. User clicks "accept all revisions" → wrappers persist →
    /// silent UX bug.
    func testMixedContentInsRevision_PopulatesBothRawAndTypedRevision() throws {
        let docxURL = try Self.buildMixedContentInsRevisionFixture()
        defer { try? FileManager.default.removeItem(at: docxURL) }

        let doc = try DocxReader.read(from: docxURL)
        defer { var d = doc; d.close() }

        guard case .paragraph(let para) = doc.body.children[0] else {
            XCTFail("Expected paragraph at body[0]")
            return
        }

        // Raw representation must exist (P0-7 byte-fidelity contract).
        let hasRawIns = para.unrecognizedChildren.contains {
            $0.name == "ins" && $0.rawXML.contains("<w:ins") && $0.rawXML.contains("w:id=\"5\"")
        }
        XCTAssertTrue(
            hasRawIns,
            "paragraph.unrecognizedChildren must contain the verbatim <w:ins> wrapper. Got: \(para.unrecognizedChildren.map { $0.name })"
        )

        // Typed representation must ALSO exist (R3-NEW-4 fix).
        let typedRev = para.revisions.first { $0.id == 5 }
        XCTAssertNotNil(
            typedRev,
            "paragraph.revisions must contain a typed Revision with id=5 so MCP tooling sees it. Got revisions: \(para.revisions.map { ($0.id, $0.type) })"
        )
        if let r = typedRev {
            XCTAssertEqual(r.type, .insertion, "typed Revision type must be .insertion")
            XCTAssertEqual(r.author, "Alice", "typed Revision author must be 'Alice'")
        }
    }

    /// Spec scenario "acceptRevision on mixed-content wrapper unwraps inner content":
    /// after accept, the inner `<w:hyperlink>` survives without the `<w:ins>` wrapper,
    /// the typed Revision and the raw entry both disappear.
    func testAcceptRevision_OnMixedContentWrapper_UnwrapsInnerContent() throws {
        let docxURL = try Self.buildMixedContentInsRevisionFixture()
        defer { try? FileManager.default.removeItem(at: docxURL) }

        var doc = try DocxReader.read(from: docxURL)
        defer { doc.close() }

        try doc.acceptRevision(revisionId: 5)

        guard case .paragraph(let para) = doc.body.children[0] else {
            XCTFail("Expected paragraph at body[0]")
            return
        }
        let xml = para.toXML()

        XCTAssertFalse(
            xml.contains("<w:ins"),
            "After accept, paragraph XML must NOT contain the <w:ins> wrapper. Output:\n\(xml)"
        )
        XCTAssertTrue(
            xml.contains("<w:hyperlink"),
            "After accept, the inner <w:hyperlink> must survive. Output:\n\(xml)"
        )
        XCTAssertNil(
            para.revisions.first { $0.id == 5 },
            "After accept, paragraph.revisions must NOT contain Revision id=5"
        )
        XCTAssertFalse(
            para.unrecognizedChildren.contains { $0.rawXML.contains("<w:ins") && $0.rawXML.contains("w:id=\"5\"") },
            "After accept, paragraph.unrecognizedChildren must NOT contain the <w:ins> raw entry"
        )
    }

    /// Spec scenario "rejectRevision on mixed-content insertion removes the entire wrapper and content":
    /// after reject of an insertion, both the wrapper AND the inner hyperlink are gone.
    func testRejectRevision_OnMixedContentWrapper_RemovesEntireWrapper() throws {
        let docxURL = try Self.buildMixedContentInsRevisionFixture()
        defer { try? FileManager.default.removeItem(at: docxURL) }

        var doc = try DocxReader.read(from: docxURL)
        defer { doc.close() }

        try doc.rejectRevision(revisionId: 5)

        guard case .paragraph(let para) = doc.body.children[0] else {
            XCTFail("Expected paragraph at body[0]")
            return
        }
        let xml = para.toXML()

        XCTAssertFalse(
            xml.contains("<w:ins"),
            "After reject, paragraph XML must NOT contain the <w:ins> wrapper. Output:\n\(xml)"
        )
        XCTAssertFalse(
            xml.contains("<w:hyperlink"),
            "After reject of an insertion, the inner <w:hyperlink> must be removed too. Output:\n\(xml)"
        )
        XCTAssertNil(
            para.revisions.first { $0.id == 5 },
            "After reject, paragraph.revisions must NOT contain Revision id=5"
        )
        XCTAssertFalse(
            para.unrecognizedChildren.contains { $0.rawXML.contains("<w:ins") && $0.rawXML.contains("w:id=\"5\"") },
            "After reject, paragraph.unrecognizedChildren must NOT contain the <w:ins> raw entry"
        )
    }

    // MARK: - R3-NEW-5: nextBookmarkId calibration SHALL scan all bookmark-bearing document parts

    /// Spec scenario "Bookmark in table cell calibrates nextBookmarkId":
    /// max id 99 inside a table cell, expect nextBookmarkId == 100.
    ///
    /// Pre-v0.19.4: calibration only iterated `body.children` matching `.paragraph`,
    /// missed `.table` cells. Subsequent `insertBookmark` would allocate id 1
    /// (no calibration applied) → collision with source id 99 → silent
    /// overwrite or schema rejection.
    func testBookmarkInTableCell_CalibratesNextBookmarkId() throws {
        let docxURL = try Self.buildBookmarkInTableFixture(maxId: 99)
        defer { try? FileManager.default.removeItem(at: docxURL) }

        var doc = try DocxReader.read(from: docxURL)
        defer { doc.close() }

        let newId = try doc.insertBookmark(name: "added")
        XCTAssertGreaterThan(
            newId, 99,
            "nextBookmarkId must be calibrated past table-cell source max (99). Got \(newId)."
        )
    }

    /// Spec scenario "Bookmark in header calibrates nextBookmarkId":
    /// body max 5, header max 50, expect nextBookmarkId == 51.
    func testBookmarkInHeader_CalibratesNextBookmarkId() throws {
        let docxURL = try Self.buildBookmarkInHeaderFixture(bodyId: 5, headerId: 50)
        defer { try? FileManager.default.removeItem(at: docxURL) }

        var doc = try DocxReader.read(from: docxURL)
        defer { doc.close() }

        let newId = try doc.insertBookmark(name: "added")
        XCTAssertGreaterThan(
            newId, 50,
            "nextBookmarkId must be calibrated past header source max (50). Got \(newId)."
        )
    }

    // MARK: - R3-NEW-6: Direct-emit XML attribute values SHALL be escaped to prevent injection

    /// Spec scenario "rStyle value with embedded quote round-trips without injection":
    /// rStyle value `x"/><injected/><w:dummy w:val="y` must round-trip through
    /// toXML() and re-parse as a single attribute value, NOT as additional
    /// sibling elements.
    func testRStyle_WithEmbeddedQuoteAndAngleBracket_DoesNotInjectSiblings() throws {
        let maliciousRStyle = "x\"/><injected/><w:dummy w:val=\"y"
        var props = RunProperties()
        props.rStyle = maliciousRStyle

        let xml = props.toXML()

        // Re-parse the emitted XML. Wrap in a synthetic root so XMLDocument
        // accepts the fragment.
        let wrapped = "<w:rPr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\(xml)</w:rPr>"
        let parsed = try XMLDocument(xmlString: wrapped, options: [])

        // Bug: pre-fix the unescaped quote+angle-bracket payload becomes
        // additional sibling elements <injected/> and <w:dummy>.
        let injectedNodes = try parsed.nodes(forXPath: "//*[local-name()='injected']")
        XCTAssertTrue(
            injectedNodes.isEmpty,
            "Re-parsed XML must NOT contain <injected/> sibling. rStyle escape failed. Output:\n\(xml)"
        )
        let dummyNodes = try parsed.nodes(forXPath: "//*[local-name()='dummy']")
        XCTAssertTrue(
            dummyNodes.isEmpty,
            "Re-parsed XML must NOT contain <w:dummy> sibling. rStyle escape failed. Output:\n\(xml)"
        )

        // Positive check: the rStyle attribute must contain the original
        // (decoded) malicious string verbatim.
        let rStyleNodes = try parsed.nodes(forXPath: "//*[local-name()='rStyle']")
        XCTAssertEqual(rStyleNodes.count, 1, "Expected exactly one <w:rStyle> element")
        if let rStyleEl = rStyleNodes.first as? XMLElement,
           let valAttr = rStyleEl.attribute(forName: "w:val") {
            XCTAssertEqual(
                valAttr.stringValue, maliciousRStyle,
                "Round-tripped w:val must equal original rStyle string verbatim."
            )
        } else {
            XCTFail("rStyle node has no w:val attribute. Output:\n\(xml)")
        }
    }

    /// Audit corner cases for color / fontName per R3-NEW-6 task 7.5.
    /// Each row of the spec scenario's escape table is exercised: ampersand,
    /// apostrophe, less-than, greater-than. Validates round-trip equality.
    func testRunPropertyAttributes_EscapeAcrossAuditedSites() throws {
        let cases: [(rStyle: String?, color: String?, fontName: String?, label: String)] = [
            ("Heading1", nil, nil, "normal rStyle, no escape needed"),
            ("A & B", nil, nil, "ampersand in rStyle"),
            ("O'Connor", nil, nil, "apostrophe in rStyle"),
            (nil, "FF0000", nil, "normal hex color"),
            (nil, "x\"/><evil/>", nil, "quote-injection in color"),
            (nil, nil, "Arial & Black", "ampersand in fontName"),
            (nil, nil, "Bad\"<font", "quote+angle in fontName"),
        ]
        for c in cases {
            var props = RunProperties()
            props.rStyle = c.rStyle
            props.color = c.color
            props.fontName = c.fontName
            let xml = props.toXML()
            // Re-parse and verify exact round-trip on each populated attribute.
            let wrapped = "<w:rPr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\(xml)</w:rPr>"
            let parsed = try XMLDocument(xmlString: wrapped, options: [])
            if let rs = c.rStyle {
                let val = (try parsed.nodes(forXPath: "//*[local-name()='rStyle']").first as? XMLElement)?
                    .attribute(forName: "w:val")?.stringValue
                XCTAssertEqual(val, rs, "Case '\(c.label)' rStyle round-trip failed. xml=\(xml)")
            }
            if let col = c.color {
                let val = (try parsed.nodes(forXPath: "//*[local-name()='color']").first as? XMLElement)?
                    .attribute(forName: "w:val")?.stringValue
                XCTAssertEqual(val, col, "Case '\(c.label)' color round-trip failed. xml=\(xml)")
            }
            if let fn = c.fontName {
                let val = (try parsed.nodes(forXPath: "//*[local-name()='rFonts']").first as? XMLElement)?
                    .attribute(forName: "w:ascii")?.stringValue
                XCTAssertEqual(val, fn, "Case '\(c.label)' fontName round-trip failed. xml=\(xml)")
            }
            // No injected/evil sibling elements regardless of input.
            XCTAssertTrue(
                try parsed.nodes(forXPath: "//*[local-name()='evil']").isEmpty,
                "Case '\(c.label)' produced <evil/> sibling — escape failed. xml=\(xml)"
            )
        }
    }

    // MARK: - Fixture builders

    /// Builds a minimal valid `.docx` whose body contains exactly one paragraph
    /// with one source-loaded `<w:hyperlink>` wrapping the given text. Returns
    /// the path to the assembled `.docx` file (caller cleans up).
    private static func buildHyperlinkSourceFixture(text: String) throws -> URL {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("r3-hl-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(
            at: stagingDir.appendingPathComponent("_rels"),
            withIntermediateDirectories: true
        )
        try FileManager.default.createDirectory(
            at: stagingDir.appendingPathComponent("word/_rels"),
            withIntermediateDirectories: true
        )
        defer { try? FileManager.default.removeItem(at: stagingDir) }

        let contentTypes = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>
        """
        try contentTypes.write(
            to: stagingDir.appendingPathComponent("[Content_Types].xml"),
            atomically: true, encoding: .utf8
        )

        let rootRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """
        try rootRels.write(
            to: stagingDir.appendingPathComponent("_rels/.rels"),
            atomically: true, encoding: .utf8
        )

        let documentRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
        </Relationships>
        """
        try documentRels.write(
            to: stagingDir.appendingPathComponent("word/_rels/document.xml.rels"),
            atomically: true, encoding: .utf8
        )

        let documentXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <w:body>
                <w:p><w:hyperlink r:id="rId1"><w:r><w:t>\(text)</w:t></w:r></w:hyperlink></w:p>
            </w:body>
        </w:document>
        """
        try documentXml.write(
            to: stagingDir.appendingPathComponent("word/document.xml"),
            atomically: true, encoding: .utf8
        )

        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("r3-hl-fixture-\(UUID().uuidString).docx")
        try ZipHelper.zip(stagingDir, to: docxURL)
        return docxURL
    }

    /// Builds a minimal valid `.docx` whose body contains exactly one paragraph
    /// `<w:p><w:r><w:t>A</w:t></w:r><w:sdt><w:sdtContent><w:r><w:t>X</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t>B</w:t></w:r></w:p>`.
    /// Used by R3-NEW-2 source-position round-trip tests.
    private static func buildSDTBetweenRunsFixture() throws -> URL {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("r3-sdt-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(
            at: stagingDir.appendingPathComponent("_rels"),
            withIntermediateDirectories: true
        )
        try FileManager.default.createDirectory(
            at: stagingDir.appendingPathComponent("word/_rels"),
            withIntermediateDirectories: true
        )
        defer { try? FileManager.default.removeItem(at: stagingDir) }

        let contentTypes = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>
        """
        try contentTypes.write(
            to: stagingDir.appendingPathComponent("[Content_Types].xml"),
            atomically: true, encoding: .utf8
        )

        let rootRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """
        try rootRels.write(
            to: stagingDir.appendingPathComponent("_rels/.rels"),
            atomically: true, encoding: .utf8
        )

        let emptyDocRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        </Relationships>
        """
        try emptyDocRels.write(
            to: stagingDir.appendingPathComponent("word/_rels/document.xml.rels"),
            atomically: true, encoding: .utf8
        )

        let documentXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p><w:r><w:t>A</w:t></w:r><w:sdt><w:sdtPr><w:tag w:val="t1"/></w:sdtPr><w:sdtContent><w:r><w:t>X</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t>B</w:t></w:r></w:p>
            </w:body>
        </w:document>
        """
        try documentXml.write(
            to: stagingDir.appendingPathComponent("word/document.xml"),
            atomically: true, encoding: .utf8
        )

        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("r3-sdt-fixture-\(UUID().uuidString).docx")
        try ZipHelper.zip(stagingDir, to: docxURL)
        return docxURL
    }

    /// Builds a minimal valid `.docx` with a paragraph carrying source-loaded
    /// `<w:commentRangeStart w:id="3"/>` markers wrapping a run, plus a comments.xml
    /// part declaring comment id=3. Used by R3-NEW-3 tests.
    private static func buildExistingCommentsFixture() throws -> URL {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("r3-cmt-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(
            at: stagingDir.appendingPathComponent("_rels"),
            withIntermediateDirectories: true
        )
        try FileManager.default.createDirectory(
            at: stagingDir.appendingPathComponent("word/_rels"),
            withIntermediateDirectories: true
        )
        defer { try? FileManager.default.removeItem(at: stagingDir) }

        let contentTypes = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
        </Types>
        """
        try contentTypes.write(
            to: stagingDir.appendingPathComponent("[Content_Types].xml"),
            atomically: true, encoding: .utf8
        )

        let rootRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """
        try rootRels.write(
            to: stagingDir.appendingPathComponent("_rels/.rels"),
            atomically: true, encoding: .utf8
        )

        let docRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
        </Relationships>
        """
        try docRels.write(
            to: stagingDir.appendingPathComponent("word/_rels/document.xml.rels"),
            atomically: true, encoding: .utf8
        )

        let documentXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p><w:commentRangeStart w:id="3"/><w:r><w:t>anchored text</w:t></w:r><w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r></w:p>
            </w:body>
        </w:document>
        """
        try documentXml.write(
            to: stagingDir.appendingPathComponent("word/document.xml"),
            atomically: true, encoding: .utf8
        )

        let commentsXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:comment w:id="3" w:author="Original" w:date="2026-04-26T00:00:00Z" w:initials="O">
                <w:p><w:r><w:t>existing note</w:t></w:r></w:p>
            </w:comment>
        </w:comments>
        """
        try commentsXml.write(
            to: stagingDir.appendingPathComponent("word/comments.xml"),
            atomically: true, encoding: .utf8
        )

        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("r3-cmt-fixture-\(UUID().uuidString).docx")
        try ZipHelper.zip(stagingDir, to: docxURL)
        return docxURL
    }

    /// Builds a minimal valid `.docx` with a paragraph containing
    /// `<w:ins w:id="5" w:author="Alice"...><w:hyperlink ...><w:r><w:t>linked</w:t></w:r></w:hyperlink></w:ins>`.
    /// Triggers the P0-7 hasNonRunChild raw-capture path. Used by R3-NEW-4 tests.
    private static func buildMixedContentInsRevisionFixture() throws -> URL {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("r3-rev-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(
            at: stagingDir.appendingPathComponent("_rels"),
            withIntermediateDirectories: true
        )
        try FileManager.default.createDirectory(
            at: stagingDir.appendingPathComponent("word/_rels"),
            withIntermediateDirectories: true
        )
        defer { try? FileManager.default.removeItem(at: stagingDir) }

        let contentTypes = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>
        """
        try contentTypes.write(
            to: stagingDir.appendingPathComponent("[Content_Types].xml"),
            atomically: true, encoding: .utf8
        )

        let rootRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """
        try rootRels.write(
            to: stagingDir.appendingPathComponent("_rels/.rels"),
            atomically: true, encoding: .utf8
        )

        let docRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
        </Relationships>
        """
        try docRels.write(
            to: stagingDir.appendingPathComponent("word/_rels/document.xml.rels"),
            atomically: true, encoding: .utf8
        )

        let documentXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <w:body>
                <w:p><w:ins w:id="5" w:author="Alice" w:date="2026-04-25T10:00:00Z"><w:hyperlink r:id="rId10"><w:r><w:t>linked</w:t></w:r></w:hyperlink></w:ins></w:p>
            </w:body>
        </w:document>
        """
        try documentXml.write(
            to: stagingDir.appendingPathComponent("word/document.xml"),
            atomically: true, encoding: .utf8
        )

        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("r3-rev-fixture-\(UUID().uuidString).docx")
        try ZipHelper.zip(stagingDir, to: docxURL)
        return docxURL
    }

    /// Builds a `.docx` whose ONLY bookmark lives inside a table cell paragraph,
    /// using the supplied numeric id. Used by R3-NEW-5 table-cell calibration test.
    private static func buildBookmarkInTableFixture(maxId: Int) throws -> URL {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("r3-tbl-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(
            at: stagingDir.appendingPathComponent("_rels"),
            withIntermediateDirectories: true
        )
        try FileManager.default.createDirectory(
            at: stagingDir.appendingPathComponent("word/_rels"),
            withIntermediateDirectories: true
        )
        defer { try? FileManager.default.removeItem(at: stagingDir) }

        let contentTypes = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>
        """
        try contentTypes.write(
            to: stagingDir.appendingPathComponent("[Content_Types].xml"),
            atomically: true, encoding: .utf8
        )

        let rootRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """
        try rootRels.write(
            to: stagingDir.appendingPathComponent("_rels/.rels"),
            atomically: true, encoding: .utf8
        )

        let docRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        </Relationships>
        """
        try docRels.write(
            to: stagingDir.appendingPathComponent("word/_rels/document.xml.rels"),
            atomically: true, encoding: .utf8
        )

        let documentXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:tbl>
                    <w:tr>
                        <w:tc>
                            <w:p><w:bookmarkStart w:id="\(maxId)" w:name="cell-bm"/><w:r><w:t>cell</w:t></w:r><w:bookmarkEnd w:id="\(maxId)"/></w:p>
                        </w:tc>
                    </w:tr>
                </w:tbl>
                <w:p><w:r><w:t>after table</w:t></w:r></w:p>
            </w:body>
        </w:document>
        """
        try documentXml.write(
            to: stagingDir.appendingPathComponent("word/document.xml"),
            atomically: true, encoding: .utf8
        )

        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("r3-tbl-fixture-\(UUID().uuidString).docx")
        try ZipHelper.zip(stagingDir, to: docxURL)
        return docxURL
    }

    /// Builds a `.docx` with a body bookmark (id `bodyId`) AND a header bookmark
    /// (id `headerId`). Used by R3-NEW-5 header calibration test.
    private static func buildBookmarkInHeaderFixture(bodyId: Int, headerId: Int) throws -> URL {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("r3-hdr-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(
            at: stagingDir.appendingPathComponent("_rels"),
            withIntermediateDirectories: true
        )
        try FileManager.default.createDirectory(
            at: stagingDir.appendingPathComponent("word/_rels"),
            withIntermediateDirectories: true
        )
        defer { try? FileManager.default.removeItem(at: stagingDir) }

        let contentTypes = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
        </Types>
        """
        try contentTypes.write(
            to: stagingDir.appendingPathComponent("[Content_Types].xml"),
            atomically: true, encoding: .utf8
        )

        let rootRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """
        try rootRels.write(
            to: stagingDir.appendingPathComponent("_rels/.rels"),
            atomically: true, encoding: .utf8
        )

        let docRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
        </Relationships>
        """
        try docRels.write(
            to: stagingDir.appendingPathComponent("word/_rels/document.xml.rels"),
            atomically: true, encoding: .utf8
        )

        let documentXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <w:body>
                <w:p><w:bookmarkStart w:id="\(bodyId)" w:name="body-bm"/><w:r><w:t>body</w:t></w:r><w:bookmarkEnd w:id="\(bodyId)"/></w:p>
                <w:sectPr><w:headerReference r:id="rId10" w:type="default"/></w:sectPr>
            </w:body>
        </w:document>
        """
        try documentXml.write(
            to: stagingDir.appendingPathComponent("word/document.xml"),
            atomically: true, encoding: .utf8
        )

        let headerXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:p><w:bookmarkStart w:id="\(headerId)" w:name="header-bm"/><w:r><w:t>header</w:t></w:r><w:bookmarkEnd w:id="\(headerId)"/></w:p>
        </w:hdr>
        """
        try headerXml.write(
            to: stagingDir.appendingPathComponent("word/header1.xml"),
            atomically: true, encoding: .utf8
        )

        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("r3-hdr-fixture-\(UUID().uuidString).docx")
        try ZipHelper.zip(stagingDir, to: docxURL)
        return docxURL
    }
}

// MARK: - R3-NEW-6 Audit Table (task 7.4)
//
// Audit of every direct-emit attribute site that interpolates an unbounded
// `String` value into XML attribute via Swift `\"\(value)\"` interpolation.
// Each row is either ESCAPED (now routed through the local
// `escapeXMLAttribute` helper) or SAFE-BY-CONSTRUCTION (the value comes from
// a Swift enum / Int / known-finite pattern).
//
// | File / Line                                            | Attribute             | Status               | Notes |
// |--------------------------------------------------------|-----------------------|----------------------|-------|
// | Run.swift RunProperties.toXML  rStyle                  | w:rStyle/@w:val       | ESCAPED              | R3-NEW-6 fix |
// | Run.swift RunProperties.toXML  fontName                | w:rFonts/@w:ascii etc | ESCAPED              | R3-NEW-6 audit |
// | Run.swift RunProperties.toXML  color                   | w:color/@w:val        | ESCAPED              | R3-NEW-6 audit |
// | Run.swift RunProperties.toXML  underline.rawValue      | w:u/@w:val            | SAFE-BY-CONSTRUCTION | enum |
// | Run.swift RunProperties.toXML  fontSize (Int)          | w:sz/@w:val           | SAFE-BY-CONSTRUCTION | Int |
// | Run.swift RunProperties.toXML  highlight.rawValue      | w:highlight/@w:val    | SAFE-BY-CONSTRUCTION | enum |
// | Run.swift RunProperties.toXML  verticalAlign.rawValue  | w:vertAlign/@w:val    | SAFE-BY-CONSTRUCTION | enum |
// | Hyperlink.swift toXML  rId / anchor / tooltip          | r:id / w:anchor / w:tooltip | ESCAPED        | pre-existing local escapeXML helper |
// | Hyperlink.swift toXML  rawAttributes loop              | various               | ESCAPED              | pre-existing local escapeXML helper |
// | Paragraph.swift bookmarkStart name                     | w:bookmarkStart/@w:name | ESCAPED            | pre-existing escapeXMLAttribute |
// | Paragraph.swift permissionRangeMarker emit             | w:permStart/@*        | ESCAPED              | pre-existing escapeXMLAttribute |
// | Paragraph.swift fldSimple instr                        | w:fldSimple/@w:instr  | ESCAPED              | pre-existing escapeXMLAttribute |
// | Footer.swift / DocxWriter.swift escape                 | various               | ESCAPED              | each file has its own local helper |
// | Revision.swift toChangeXML  color                      | w:color/@w:val        | ESCAPED              | R3-NEW-6 codex P1 fix (parallel emit path inside rPrChange) |
// | Revision.swift toChangeXML  fontName                   | w:rFonts/@w:ascii etc | ESCAPED              | pre-existing escapeXML in toChangeXML |
//
// Future cleanup (out of R3 scope): consolidate the four parallel
// `escapeXML` / `escapeXMLAttribute` helpers (Run.swift, Paragraph.swift,
// Hyperlink.swift, Revision.swift, Footer.swift, DocxWriter.swift) into a
// shared `XMLEscape.swift` once a follow-up change scopes the cross-cutting
// refactor. For now keeping local copies preserves zero-coupling between
// Models/ and IO/.
