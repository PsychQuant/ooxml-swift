import XCTest
@testable import OOXMLSwift

/// The official-template import contract after PsychQuant/macdoc#212, #213
/// and #214: which parts `importOfficial` reads (the ones the main part's
/// relationships name), which bytes it accepts (UTF-8 only, agreeing with
/// the XML declaration), and how an incomplete template is reported.
final class DocumentFormattingProfileImportTests: XCTestCase {
    static let w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    static let relNS = "http://schemas.openxmlformats.org/package/2006/relationships"
    static let officeRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    // MARK: - Helpers

    func directory() throws -> URL {
        let url = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString)
        try FileManager.default.createDirectory(at: url, withIntermediateDirectories: true)
        addTeardownBlock { try? FileManager.default.removeItem(at: url) }
        return url
    }

    /// The parts of the shared well-formed template fixture. Its rels name
    /// styles, theme and fontTable at their default paths.
    func baseParts() throws -> [String: Data] {
        let url = try DocumentFormattingProfileTests().template()
        defer { try? FileManager.default.removeItem(at: url.deletingLastPathComponent()) }
        return try RawPartChannel.readAllParts(from: url)
    }

    func package(_ parts: [String: Data]) throws -> URL {
        let root = try directory()
        let source = root.appendingPathComponent("source")
        for (path, bytes) in parts {
            let file = source.appendingPathComponent(path)
            try FileManager.default.createDirectory(at: file.deletingLastPathComponent(), withIntermediateDirectories: true)
            try bytes.write(to: file)
        }
        let url = root.appendingPathComponent("template.dotx")
        try ZipHelper.zipToData(source).write(to: url)
        return url
    }

    func importing(_ parts: [String: Data]) throws -> DocumentFormattingProfile {
        try DocumentFormattingProfile.importOfficial(from: package(parts))
    }

    func relationships(_ entries: [(type: String, target: String)], extra: String = "") -> Data {
        let body = entries.enumerated().map { index, entry in
            "<Relationship Id=\"rId\(index + 1)\" Type=\"\(Self.officeRel)/\(entry.type)\" Target=\"\(entry.target)\"/>"
        }.joined()
        return Data("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"\(Self.relNS)\">\(body)\(extra)</Relationships>".utf8)
    }

    static let defaultTargets: [(type: String, target: String)] = [
        ("styles", "styles.xml"), ("theme", "theme/theme1.xml"), ("fontTable", "fontTable.xml")
    ]

    func replacing(_ data: Data?, _ old: String, with new: String) throws -> Data {
        let text = String(decoding: try XCTUnwrap(data), as: UTF8.self)
        XCTAssertTrue(text.contains(old), "fixture no longer contains \(old)")
        return Data(text.replacingOccurrences(of: old, with: new).utf8)
    }

    func assertImport(_ parts: [String: Data], throws expected: DocumentFormattingProfileError,
                      _ message: String = "", file: StaticString = #filePath, line: UInt = #line) {
        XCTAssertThrowsError(try importing(parts), message, file: file, line: line) { error in
            XCTAssertEqual(error as? DocumentFormattingProfileError, expected, message, file: file, line: line)
        }
    }

    /// Styles without theme references, so a template may omit its theme.
    static let themeFreeStyles = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:styles xmlns:w=\"\(w)\"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val=\"21\"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"a\"><w:name w:val=\"Normal\"/></w:style></w:styles>"

    // MARK: - PsychQuant/macdoc#213 relationship-resolved formatting parts

    /// Each formatting part is read from the part its relationship names,
    /// in every Target spelling that resolves there, never from the default
    /// path. Stale parts at the default paths are ignored.
    func testImportReadsFormattingPartsNamedByRelationships() throws {
        let base = try baseParts()
        for (stylesTarget, stylesPart) in [("customStyles.xml", "word/customStyles.xml"),
                                           ("/word/customStyles.xml", "word/customStyles.xml"),
                                           ("./custom%53tyles.xml", "word/customStyles.xml"),
                                           ("../formatting/styles.xml", "formatting/styles.xml")] {
            var parts = base
            parts[stylesPart] = try replacing(base["word/styles.xml"], "x:val=\"Title\"", with: "x:val=\"Relationship Title\"")
            parts["word/theme/customTheme.xml"] = try replacing(base["word/theme/theme1.xml"], "typeface=\"Aptos\"", with: "typeface=\"Relationship Theme\"")
            parts["word/customFonts.xml"] = try replacing(base["word/fontTable.xml"], "w:name=\"Aptos\"", with: "w:name=\"Relationship Font\"")
            parts["word/styles.xml"] = try replacing(base["word/styles.xml"], "x:val=\"Title\"", with: "x:val=\"Stale Title\"")
            parts["word/theme/theme1.xml"] = try replacing(base["word/theme/theme1.xml"], "typeface=\"Aptos\"", with: "typeface=\"Stale Theme\"")
            parts["word/fontTable.xml"] = try replacing(base["word/fontTable.xml"], "w:name=\"Aptos\"", with: "w:name=\"Stale Font\"")
            parts["word/_rels/document.xml.rels"] = relationships([("styles", stylesTarget), ("theme", "theme/customTheme.xml"), ("fontTable", "customFonts.xml")])
            let profile = try importing(parts)
            let styles = try XCTUnwrap(profile.stylesXML), theme = try XCTUnwrap(profile.themeXML), fonts = try XCTUnwrap(profile.fontsXML)
            XCTAssertTrue(styles.contains("Relationship Title"), stylesTarget)
            XCTAssertFalse(styles.contains("Stale Title"), stylesTarget)
            XCTAssertTrue(theme.contains("Relationship Theme"), stylesTarget)
            XCTAssertFalse(theme.contains("Stale Theme"), stylesTarget)
            XCTAssertTrue(fonts.contains("Relationship Font"), stylesTarget)
            XCTAssertFalse(fonts.contains("Stale Font"), stylesTarget)
        }
    }

    /// A template whose formatting parts exist only at the non-default
    /// paths its relationships name imports; 3.10.0 reported the fixed
    /// `word/styles.xml` as missing.
    func testImportSucceedsWhenFormattingPartsExistOnlyAtRelationshipTargets() throws {
        var parts = try baseParts()
        parts["word/customStyles.xml"] = parts.removeValue(forKey: "word/styles.xml")
        parts["word/theme/customTheme.xml"] = parts.removeValue(forKey: "word/theme/theme1.xml")
        parts["word/customFonts.xml"] = parts.removeValue(forKey: "word/fontTable.xml")
        parts["word/_rels/document.xml.rels"] = relationships([("styles", "customStyles.xml"), ("theme", "theme/customTheme.xml"), ("fontTable", "customFonts.xml")])
        let profile = try importing(parts)
        XCTAssertNotNil(profile.stylesXML)
        XCTAssertNotNil(profile.themeXML)
        XCTAssertNotNil(profile.fontsXML)
    }

    /// A relationship whose Target names a part the package does not
    /// contain is refused, even when a part exists at the default path.
    func testImportRejectsRelationshipWhoseTargetPartIsMissing() throws {
        let base = try baseParts()
        for (type, target, resolved) in [("styles", "customStyles.xml", "word/customStyles.xml"),
                                         ("theme", "theme/theme9.xml", "word/theme/theme9.xml"),
                                         ("fontTable", "../customFonts.xml", "customFonts.xml"),
                                         ("numbering", "customNumbering.xml", "word/customNumbering.xml")] {
            var parts = base
            var targets = Self.defaultTargets.filter { $0.type != type }
            targets.append((type, target))
            parts["word/_rels/document.xml.rels"] = relationships(targets)
            parts["word/numbering.xml"] = Data("<w:numbering xmlns:w=\"\(Self.w)\"/>".utf8)
            assertImport(parts, throws: .missingRequiredFormatting("\(type) relationship 的 Target「\(target)」指向不存在的 part \(resolved)"), type)
        }
    }

    /// Without a styles relationship the package has no styles part in the
    /// OPC sense; the default-path part is not read in its place.
    func testImportRejectsTemplateWithoutStylesRelationship() throws {
        var parts = try baseParts()
        parts["word/_rels/document.xml.rels"] = relationships(Self.defaultTargets.filter { $0.type != "styles" })
        assertImport(parts, throws: .missingRequiredFormatting("word/_rels/document.xml.rels 沒有 styles relationship"))
        parts.removeValue(forKey: "word/_rels/document.xml.rels")
        assertImport(parts, throws: .missingRequiredFormatting("word/_rels/document.xml.rels 沒有 styles relationship"), "no rels part")
    }

    /// Theme and font table are optional: without a relationship an
    /// orphan part at the default path is not part of the template, so it is
    /// not captured. Styles that still need the theme are then refused by
    /// the existing completeness rule.
    func testImportIgnoresOrphanThemeAndFontTable() throws {
        var parts = try baseParts()
        parts["word/styles.xml"] = Data(Self.themeFreeStyles.utf8)
        parts["word/_rels/document.xml.rels"] = relationships([("styles", "styles.xml")])
        XCTAssertNotNil(parts["word/theme/theme1.xml"])
        XCTAssertNotNil(parts["word/fontTable.xml"])
        let profile = try importing(parts)
        XCTAssertNil(profile.themeXML)
        XCTAssertNil(profile.fontsXML)

        var themed = try baseParts()
        themed["word/_rels/document.xml.rels"] = relationships([("styles", "styles.xml"), ("fontTable", "fontTable.xml")])
        assertImport(themed, throws: .missingRequiredFormatting("theme used by styles"))
    }

    /// Numbering is only a refusal check (nothing of it is persisted), so it
    /// is checked conservatively: the part the numbering relationship names
    /// AND the default `word/numbering.xml`. A numbering part at a
    /// non-default path can no longer slip past the first-version refusal.
    func testImportChecksNumberingAtTheRelationshipTargetAndTheDefaultPath() throws {
        let numbering = Data("<w:numbering xmlns:w=\"\(Self.w)\"><w:abstractNum w:abstractNumId=\"0\"/><w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num></w:numbering>".utf8)
        let empty = Data("<w:numbering xmlns:w=\"\(Self.w)\"/>".utf8)
        var parts = try baseParts()
        parts["word/customNumbering.xml"] = numbering
        parts["word/_rels/document.xml.rels"] = relationships(Self.defaultTargets + [("numbering", "customNumbering.xml")])
        assertImport(parts, throws: .unsupportedNumbering, "relationship target")

        parts["word/customNumbering.xml"] = empty
        XCTAssertNoThrow(try importing(parts), "empty numbering at the relationship target")
        parts["word/numbering.xml"] = numbering
        assertImport(parts, throws: .unsupportedNumbering, "default path is still checked")
    }

    /// A formatting relationship pointing outside the package is refused;
    /// an external part is never fetched.
    func testImportRejectsExternalFormattingRelationship() throws {
        for type in ["styles", "theme", "fontTable", "numbering"] {
            var parts = try baseParts()
            let external = "<Relationship Id=\"rIdExternal\" Type=\"\(Self.officeRel)/\(type)\" Target=\"https://example.invalid/\(type).xml\" TargetMode=\"External\"/>"
            parts["word/_rels/document.xml.rels"] = relationships(Self.defaultTargets.filter { $0.type != type }, extra: external)
            assertImport(parts, throws: .invalidSnapshot("\(type) relationship 的 TargetMode 為 External，格式 part 必須在套件內"), type)
        }
    }

    /// The numbering relationship gets the same OPC checks as the other
    /// implicit relationships on import.
    func testImportRejectsDuplicateOrDirectoryNumberingRelationship() throws {
        var parts = try baseParts()
        parts["word/_rels/document.xml.rels"] = relationships(Self.defaultTargets + [("numbering", "numbering.xml"), ("numbering", "numbering2.xml")])
        assertImport(parts, throws: .duplicateRelationship("numbering"))
        parts["word/_rels/document.xml.rels"] = relationships(Self.defaultTargets + [("numbering", "numbering.xml/.")])
        assertImport(parts, throws: .invalidRelationshipTarget("numbering.xml/."))
    }
}
