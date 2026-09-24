import XCTest
@testable import OOXMLSwift

/// PsychQuant/ooxml-swift#173: `DocxReader` resolves styles/theme/fontTable/
/// numbering, and validates the main part, via relationships the same way
/// `DocumentFormattingProfile.importOfficial` does (PsychQuant/macdoc#213) —
/// not fixed default paths.
final class DocxReaderFormattingPartRelationshipTests: XCTestCase {
    static let relNS = "http://schemas.openxmlformats.org/package/2006/relationships"
    static let officeRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    static let a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    static let w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    // MARK: - Helpers

    func directory() throws -> URL {
        let url = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString)
        try FileManager.default.createDirectory(at: url, withIntermediateDirectories: true)
        addTeardownBlock { try? FileManager.default.removeItem(at: url) }
        return url
    }

    /// A real, writer-produced `.docx` as raw parts, so every fixture in
    /// this file starts from something `DocxReader.read` already parses
    /// end-to-end (unlike `DocumentFormattingProfileTests.template()`'s
    /// hand-written minimal XML, which is shaped for `importOfficial`
    /// only). A freshly-constructed `WordDocument()` has `styles`,
    /// `settings` and `fontTable` relationships but no `theme` — Word's
    /// default template ships no theme part either.
    func baseParts() throws -> [String: Data] {
        let dir = try directory()
        let source = dir.appendingPathComponent("source.docx")
        var seed = WordDocument()
        seed.appendParagraph(Paragraph(text: "BODY"))
        try DocxWriter.write(seed, to: source)
        return try RawPartChannel.readAllParts(from: source)
    }

    func package(_ parts: [String: Data]) throws -> URL {
        let root = try directory()
        let source = root.appendingPathComponent("source")
        for (path, bytes) in parts {
            let file = source.appendingPathComponent(path)
            try FileManager.default.createDirectory(at: file.deletingLastPathComponent(), withIntermediateDirectories: true)
            try bytes.write(to: file)
        }
        let url = root.appendingPathComponent("test.docx")
        try ZipHelper.zipToData(source).write(to: url)
        return url
    }

    func replacing(_ data: Data, _ old: String, with new: String) -> Data {
        let text = String(decoding: data, as: UTF8.self)
        XCTAssertTrue(text.contains(old), "fixture no longer contains \(old)")
        return Data(text.replacingOccurrences(of: old, with: new).utf8)
    }

    // MARK: - #173: styles/theme/fontTable resolved by relationship

    /// The `styles` relationship names a non-default part; a STALE part
    /// still sits at the conventional `word/styles.xml` path (as it would
    /// after a tool moved styles without cleaning up the old file, or in
    /// the exact scenario `writeFormattingParts`'s existing `old-styles.xml`
    /// repair test exercises from the other direction). The reader must
    /// follow the relationship, not the stale default-path file.
    func testDocxReaderResolvesStylesByRelationshipNotStaleDefaultPath() throws {
        var parts = try baseParts()
        let stylesXML = try XCTUnwrap(parts["word/styles.xml"])
        // Move styles content (marked) to a non-default part; leave a
        // DIFFERENT (stale, unmarked) styles.xml at the default path.
        parts["word/customStyles.xml"] = replacing(stylesXML, "w:styleId=\"Normal\"", with: "w:styleId=\"Normal\" w:customMarker=\"1\"")
        parts["word/styles.xml"] = stylesXML
        var rels = String(decoding: try XCTUnwrap(parts["word/_rels/document.xml.rels"]), as: UTF8.self)
        XCTAssertTrue(rels.contains("Target=\"styles.xml\""), "fixture rels no longer name styles.xml directly")
        rels = rels.replacingOccurrences(of: "Target=\"styles.xml\"", with: "Target=\"customStyles.xml\"")
        parts["word/_rels/document.xml.rels"] = Data(rels.utf8)

        var doc = try DocxReader.read(from: try package(parts))
        defer { doc.close() }
        // Content still lands under the canonical key even though it was
        // read from a non-default on-disk part.
        let stylesTree = try XCTUnwrap(doc.xmlTrees["word/styles.xml"])
        let hasMarker = ProfileXML.walk(stylesTree.root).contains {
            $0.attributes.contains { $0.localName == "customMarker" }
        }
        XCTAssertTrue(hasMarker, "styles should come from the relationship target (customStyles.xml), not the stale default path")
    }

    /// Codex R1 MEDIUM-3: reading via the relationship isn't enough on its
    /// own — a typed edit on a document whose styles live at a non-default,
    /// EXISTING part must still save correctly (not just the earlier
    /// `old-styles.xml`-style test's case, where the named target doesn't
    /// physically exist and the write path never even touches the custom
    /// part). The writer's existing "always publish to the canonical
    /// default path, repoint every registration" decision (#173, unchanged
    /// here) means the OUTPUT is expected at `word/styles.xml`, and
    /// `word/customStyles.xml` becomes an orphan (known limitation,
    /// documented in the #173 commit) — this test pins down that the EDIT
    /// itself lands correctly rather than being silently lost or reverted
    /// to the stale default-path content.
    func testDocxReaderEditAfterReadingStylesFromNonDefaultPathSavesCorrectly() throws {
        var parts = try baseParts()
        let stylesXML = try XCTUnwrap(parts["word/styles.xml"])
        parts["word/customStyles.xml"] = stylesXML
        // Leave a STALE, DIFFERENT default-path file so a silent
        // read-from-default regression would be caught by content, not
        // just by success/failure.
        parts["word/styles.xml"] = replacing(stylesXML, "w:styleId=\"Normal\"", with: "w:styleId=\"Normal\" w:staleMarker=\"1\"")
        var rels = String(decoding: try XCTUnwrap(parts["word/_rels/document.xml.rels"]), as: UTF8.self)
        rels = rels.replacingOccurrences(of: "Target=\"styles.xml\"", with: "Target=\"customStyles.xml\"")
        parts["word/_rels/document.xml.rels"] = Data(rels.utf8)

        var doc = try DocxReader.read(from: try package(parts))
        defer { doc.close() }
        try doc.updateStyle(id: "Normal", with: StyleUpdate(name: "Edited Normal"))
        let output = try directory().appendingPathComponent("edited.docx")
        try DocxWriter.write(doc, to: output)

        var reopened = try DocxReader.read(from: output)
        defer { reopened.close() }
        XCTAssertEqual(reopened.styles.first { $0.id == "Normal" }?.name, "Edited Normal", "the edit made against the relationship-resolved content must be the one that gets saved")
        let hasStaleMarker = ProfileXML.walk(try XCTUnwrap(reopened.xmlTrees["word/styles.xml"]).root).contains {
            $0.attributes.contains { $0.localName == "staleMarker" }
        }
        XCTAssertFalse(hasStaleMarker, "the saved styles must not silently be the stale default-path content")
    }

    /// Codex R1 MEDIUM-6: an External styles relationship (`TargetMode="External"`,
    /// disallowed for an in-package formatting part per `ProfileXML.implicitPart`)
    /// falls back to the default path rather than failing the whole read —
    /// the same tolerance a duplicate or missing relationship already gets
    /// (`testDocxReaderEditAfterReadingStylesFromNonDefaultPathSavesCorrectly`
    /// and the RelationshipsTests suite cover those; this is the third
    /// `implicitPart`-throwing case `resolvedFormattingPart` catches).
    func testDocxReaderFallsBackToDefaultStylesPathWhenRelationshipIsExternal() throws {
        var parts = try baseParts()
        var rels = String(decoding: try XCTUnwrap(parts["word/_rels/document.xml.rels"]), as: UTF8.self)
        rels = rels.replacingOccurrences(
            of: "Target=\"styles.xml\"",
            with: "Target=\"https://example.com/styles.xml\" TargetMode=\"External\"")
        parts["word/_rels/document.xml.rels"] = Data(rels.utf8)
        var doc = try DocxReader.read(from: try package(parts))
        defer { doc.close() }
        XCTAssertEqual(doc.styles.first { $0.id == "Normal" }?.name, "Normal", "falls back to the default-path styles.xml instead of failing the read")
    }

    /// `word/fontTable.xml`'s relationship (present by default) is
    /// repointed to a non-default part, and a `theme` relationship —
    /// entirely absent from a freshly-constructed document — is added
    /// pointing at a non-default part too. Both a stale default-path file
    /// AND the relationship-targeted file exist for each; the resolved
    /// content must be the relationship-targeted one.
    func testDocxReaderResolvesThemeFontTableByRelationship() throws {
        var parts = try baseParts()
        let staleTheme = Data("<a:theme xmlns:a=\"\(Self.a)\" name=\"Stale\"/>".utf8)
        let realTheme = Data("<a:theme xmlns:a=\"\(Self.a)\" name=\"Real\"/>".utf8)
        let staleFonts = try replacing(XCTUnwrap(parts["word/fontTable.xml"]), "<w:fonts", with: "<w:fonts w:staleMarker=\"1\"")
        let realFonts = try replacing(XCTUnwrap(parts["word/fontTable.xml"]), "<w:fonts", with: "<w:fonts w:realMarker=\"1\"")
        parts["word/theme/theme1.xml"] = staleTheme
        parts["word/custom/realTheme.xml"] = realTheme
        parts["word/fontTable.xml"] = staleFonts
        parts["word/custom/realFonts.xml"] = realFonts

        var rels = String(decoding: try XCTUnwrap(parts["word/_rels/document.xml.rels"]), as: UTF8.self)
        XCTAssertTrue(rels.contains("Target=\"fontTable.xml\""), "fixture rels no longer name fontTable.xml directly")
        rels = rels.replacingOccurrences(of: "Target=\"fontTable.xml\"", with: "Target=\"custom/realFonts.xml\"")
        rels = rels.replacingOccurrences(
            of: "</Relationships>",
            with: "<Relationship Id=\"rIdTheme\" Type=\"\(Self.officeRel)/theme\" Target=\"custom/realTheme.xml\"/></Relationships>")
        parts["word/_rels/document.xml.rels"] = Data(rels.utf8)

        var doc = try DocxReader.read(from: try package(parts))
        defer { doc.close() }
        let state = try XCTUnwrap(doc.formattingState)
        XCTAssertEqual(state.themeData, realTheme, "theme should come from the relationship target, not the stale default path")
        XCTAssertEqual(state.fontsData, realFonts, "fontTable should come from the relationship target, not the stale default path")
    }

    // MARK: - #173: main part resolved via _rels/.rels, fail-loud on mismatch

    /// `_rels/.rels`'s officeDocument relationship names a part other than
    /// `word/document.xml`, while a STALE `word/document.xml` still exists.
    /// Before #173 this stale file would be silently read; after, the
    /// mismatch is refused loudly instead.
    func testDocxReaderThrowsOnNonDefaultMainPartInsteadOfSilentlyReadingStaleDocument() throws {
        var parts = try baseParts()
        let packageRels = Data("""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="\(Self.relNS)">
        <Relationship Id="rId1" Type="\(Self.officeRel)/officeDocument" Target="word/main2.xml"/>
        </Relationships>
        """.utf8)
        parts["_rels/.rels"] = packageRels
        // word/document.xml (the stale file) stays exactly as baseParts left it.

        XCTAssertThrowsError(try DocxReader.read(from: try package(parts))) { error in
            let message = String(describing: error)
            XCTAssertTrue(message.contains("word/main2.xml") || message.contains("officeDocument"), message)
        }
    }

    /// No `_rels/.rels` at all — real-world incomplete packages. Falls back
    /// to `word/document.xml` exactly as before #173 (regression guard, not
    /// a new behavior).
    func testDocxReaderFallsBackToDefaultMainPartWhenPackageRelsAbsent() throws {
        var parts = try baseParts()
        parts["_rels/.rels"] = nil
        var doc = try DocxReader.read(from: try package(parts))
        defer { doc.close() }
        XCTAssertEqual(doc.getText(), "BODY")
    }

    /// The common case: `_rels/.rels` names `word/document.xml` explicitly
    /// (whatever spelling `DocxWriter` uses). Must keep working exactly as
    /// before #173.
    func testDocxReaderReadsNormallyWhenMainPartRelationshipMatchesDefault() throws {
        let parts = try baseParts()
        XCTAssertNotNil(parts["_rels/.rels"], "writer-produced package should carry a package-level rels part")
        var doc = try DocxReader.read(from: try package(parts))
        defer { doc.close() }
        XCTAssertEqual(doc.getText(), "BODY")
    }
}

