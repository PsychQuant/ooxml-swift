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

    // MARK: - PsychQuant/macdoc#214 encoding gate on the profile path

    /// Every part import reads must be UTF-8 and must not declare another
    /// encoding: Word decodes by the declaration while this library decodes
    /// UTF-8, so the stored snapshot could differ from what Word shows.
    func testImportRejectsPartsDeclaringNonUTF8Encodings() throws {
        let base = try baseParts()
        for part in ["word/styles.xml", "word/document.xml", "word/theme/theme1.xml", "word/fontTable.xml", "word/_rels/document.xml.rels"] {
            for declared in ["ISO-8859-1", "Shift_JIS", "UTF-16", "windows-1252"] {
                var parts = base
                let body = String(decoding: try XCTUnwrap(base[part], part), as: UTF8.self)
                let stripped = body.hasPrefix("<?xml") ? String(body[body.range(of: "?>")!.upperBound...]) : body
                parts[part] = Data("<?xml version=\"1.0\" encoding=\"\(declared)\" standalone=\"yes\"?>\(stripped)".utf8)
                assertImport(parts, throws: .invalidSnapshot("\(part) 的 XML 宣告編碼為「\(declared)」，格式 profile 只接受 UTF-8"), "\(part) \(declared)")
            }
        }
        var numbered = base
        numbered["word/_rels/document.xml.rels"] = relationships(Self.defaultTargets + [("numbering", "numbering.xml")])
        numbered["word/numbering.xml"] = Data("<?xml version='1.0' encoding='Shift_JIS'?><w:numbering xmlns:w=\"\(Self.w)\"/>".utf8)
        assertImport(numbered, throws: .invalidSnapshot("word/numbering.xml 的 XML 宣告編碼為「Shift_JIS」，格式 profile 只接受 UTF-8"), "numbering")
    }

    /// UTF-8 in any spelling Word or other producers use is accepted: no
    /// declaration, `UTF-8` in any case, single quotes, a UTF-8 BOM, and
    /// whitespace before the declaration (which the tree reader tolerates).
    func testImportAcceptsUTF8DeclarationSpellings() throws {
        let base = try baseParts()
        let styles = String(decoding: try XCTUnwrap(base["word/styles.xml"]), as: UTF8.self)
        for (label, prefix) in [("none", Data()), ("upper", Data("<?xml version=\"1.0\" encoding=\"UTF-8\"?>".utf8)),
                                ("lower", Data("<?xml version=\"1.0\" encoding=\"utf-8\"?>".utf8)),
                                ("single", Data("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\r\n".utf8)),
                                ("no-encoding", Data("<?xml version=\"1.0\"?>".utf8)),
                                ("bom", Data([0xEF, 0xBB, 0xBF]) + Data("<?xml version=\"1.0\" encoding=\"UTF-8\"?>".utf8)),
                                ("leading-space", Data("\n <?xml version=\"1.0\" encoding=\"UTF-8\"?>".utf8))] {
            var parts = base
            parts["word/styles.xml"] = prefix + Data(styles.utf8)
            XCTAssertNoThrow(try importing(parts), label)
        }
    }

    /// Bytes that are not valid UTF-8 are refused instead of being stored
    /// with U+FFFD replacement characters; so are UTF-16/UTF-32 byte orders
    /// and NUL bytes. UTF-16 was already refused in 3.10.0 (no charset
    /// expansion); the message now names the encoding.
    func testImportRejectsBytesThatAreNotUTF8() throws {
        let base = try baseParts()
        let styles = String(decoding: try XCTUnwrap(base["word/styles.xml"]), as: UTF8.self)
        var shiftJIS = base
        shiftJIS["word/styles.xml"] = try XCTUnwrap(styles.replacingOccurrences(of: "x:val=\"Title\"", with: "x:val=\"見出し\"").data(using: .shiftJIS))
        assertImport(shiftJIS, throws: .invalidSnapshot("word/styles.xml 不是合法的 UTF-8（含無法解碼的位元組或 NUL），格式 profile 只接受 UTF-8"), "undeclared Shift_JIS bytes")
        var nul = base
        nul["word/styles.xml"] = Data(styles.utf8) + Data([0])
        assertImport(nul, throws: .invalidSnapshot("word/styles.xml 不是合法的 UTF-8（含無法解碼的位元組或 NUL），格式 profile 只接受 UTF-8"), "NUL")
        for (label, encoding, bom) in [("be-bom", String.Encoding.utf16BigEndian, [UInt8]([0xFE, 0xFF])), ("le-bom", .utf16LittleEndian, [0xFF, 0xFE]),
                                       ("be", .utf16BigEndian, []), ("le", .utf16LittleEndian, []), ("utf32le-bom", .utf32LittleEndian, [0xFF, 0xFE, 0, 0])] {
            var parts = base
            parts["word/styles.xml"] = Data(bom) + styles.replacingOccurrences(of: "UTF-8", with: "UTF-16").data(using: encoding)!
            assertImport(parts, throws: .invalidSnapshot("word/styles.xml 是 UTF-16／UTF-32 編碼，格式 profile 只接受 UTF-8"), label)
        }
    }

    /// A declaration the gate cannot read is refused rather than guessed.
    func testImportRejectsUnreadableXMLDeclaration() throws {
        let base = try baseParts()
        let styles = String(decoding: try XCTUnwrap(base["word/styles.xml"]), as: UTF8.self)
        for declaration in ["<?xml version=\"1.0\" encoding=Shift_JIS?>", "<?xml version=\"1.0\" encoding=\"UTF-8\" encoding=\"Shift_JIS\"?>",
                            "<?xml version=\"1.0\" encoding=\"UTF-8\"", "<?xml version=\"1.0\" encoding=\"UTF-8'?>"] {
            var parts = base
            parts["word/styles.xml"] = Data((declaration + styles).utf8)
            assertImport(parts, throws: .invalidSnapshot("word/styles.xml 的 XML 宣告無法解析"), declaration)
        }
    }

    /// A stored snapshot whose payload declares another encoding is refused
    /// on decode, like the import it claims to come from.
    func testDecodeRejectsSnapshotPayloadDeclaringNonUTF8Encoding() throws {
        let profile = try importing(try baseParts())
        let encoded = try JSONEncoder().encode(profile)
        XCTAssertNoThrow(try JSONDecoder().decode(DocumentFormattingProfile.self, from: encoded))
        for key in ["stylesXML", "sectionXML", "themeXML", "fontsXML"] {
            var json = try XCTUnwrap(JSONSerialization.jsonObject(with: encoded) as? [String: Any])
            let payload = try XCTUnwrap(json[key] as? String, key)
            let body = payload.hasPrefix("<?xml") ? String(payload[payload.range(of: "?>")!.upperBound...]) : payload
            json[key] = "<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?>" + body
            XCTAssertThrowsError(try JSONDecoder().decode(DocumentFormattingProfile.self, from: JSONSerialization.data(withJSONObject: json)), key) { error in
                XCTAssertEqual(error as? DocumentFormattingProfileError, .invalidSnapshot("格式快照 \(key) 的 XML 宣告編碼為「ISO-8859-1」，格式 profile 只接受 UTF-8"), key)
            }
        }
    }

    // MARK: - PsychQuant/macdoc#214 DocxReader keeps its behavior (locked)

    /// Decision: the general reader is unchanged. For `word/styles.xml`,
    /// bytes that are valid UTF-8 are decoded as UTF-8 whatever the XML
    /// declaration says (the declaration is not consulted); other bytes go
    /// through the UTF-16 path or are refused (see
    /// testReaderBOMAndDeclarationMismatchCounterexamples). Only the
    /// profile import path fails closed.
    func testReaderDecodesUTF8ValidStylesAsUTF8DespiteNonUTF8Declaration() throws {
        let source = try directory().appendingPathComponent("source.docx")
        try DocxWriter.write(WordDocument(), to: source)
        var parts = try RawPartChannel.readAllParts(from: source)
        for declared in ["ISO-8859-1", "Shift_JIS"] {
            parts["word/styles.xml"] = Data("<?xml version=\"1.0\" encoding=\"\(declared)\"?><w:styles xmlns:w=\"\(Self.w)\"><w:style w:type=\"paragraph\" w:styleId=\"X\"><w:name w:val=\"名稱\"/></w:style></w:styles>".utf8)
            let input = try package(parts)
            var doc = try DocxReader.read(from: input)
            defer { doc.close() }
            XCTAssertEqual(doc.styles.first { $0.id == "X" }?.name, "名稱", declared)
        }
    }

    /// Characterization, not endorsement (a follow-up is suggested in
    /// PsychQuant/macdoc#214): for `word/document.xml` the typed model is
    /// built by libxml2 (`XMLDocument(data:)`), which honours the
    /// declaration, while the lossless tree decodes UTF-8. An untouched save
    /// keeps the bytes; a save after a body edit writes a UTF-8 declaration
    /// over the original bytes, so what Word shows for untouched text changes
    /// from the declared decoding to UTF-8.
    func testReaderDocumentPartTypedTextFollowsDeclarationWhileTreeStaysUTF8() throws {
        let source = try directory().appendingPathComponent("source.docx")
        var seed = WordDocument()
        seed.appendParagraph(Paragraph(text: "PLACEHOLDER"))
        try DocxWriter.write(seed, to: source)
        var parts = try RawPartChannel.readAllParts(from: source)
        let original = String(decoding: try XCTUnwrap(parts["word/document.xml"]), as: UTF8.self)
        XCTAssertTrue(original.contains("encoding=\"UTF-8\""))
        let declared = Data(original.replacingOccurrences(of: "encoding=\"UTF-8\"", with: "encoding=\"ISO-8859-1\"")
            .replacingOccurrences(of: "PLACEHOLDER", with: "名稱").utf8)
        parts["word/document.xml"] = declared
        var doc = try DocxReader.read(from: try package(parts))
        defer { doc.close() }
        XCTAssertEqual(doc.getText(), String(data: Data("名稱".utf8), encoding: .isoLatin1), "typed text follows the declaration")
        let treeText = ProfileXML.walk(try XCTUnwrap(doc.xmlTrees["word/document.xml"]).root).filter { $0.kind == .text }.map(\.textContent).joined()
        XCTAssertEqual(treeText, "名稱", "the tree decodes UTF-8")
        let untouched = try directory().appendingPathComponent("untouched.docx")
        try DocxWriter.write(doc, to: untouched)
        XCTAssertEqual(try RawPartChannel.readAllParts(from: untouched)["word/document.xml"], declared)
        doc.appendParagraph(Paragraph(text: "edit"))
        let edited = try directory().appendingPathComponent("edited.docx")
        try DocxWriter.write(doc, to: edited)
        let saved = String(decoding: try XCTUnwrap(RawPartChannel.readAllParts(from: edited)["word/document.xml"]), as: UTF8.self)
        XCTAssertTrue(saved.hasPrefix("<?xml version=\"1.0\" encoding=\"UTF-8\""), String(saved.prefix(80)))
        XCTAssertTrue(saved.contains("名稱"))
    }

    // MARK: - PsychQuant/macdoc#212 actionable completeness errors

    func styles(docDefaults: String) -> Data {
        Data("<w:styles xmlns:w=\"\(Self.w)\">\(docDefaults)<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"a\"><w:name w:val=\"Normal\"/></w:style></w:styles>".utf8)
    }

    func themeFreeParts(docDefaults: String) throws -> [String: Data] {
        var parts = try baseParts()
        parts["word/styles.xml"] = styles(docDefaults: docDefaults)
        return parts
    }

    /// The error names exactly the missing field, and its description tells
    /// the user how to make the default explicit. Word's implicit defaults
    /// are never filled in on the user's behalf.
    func testMissingDocDefaultsFieldsAreNamedPrecisely() throws {
        let sz = "docDefaults/rPrDefault/rPr/sz", pPr = "docDefaults/pPrDefault"
        for (docDefaults, expected) in [
            // 90_template_ja's shape: fonts and language but no size, and an
            // empty pPrDefault (which is accepted).
            ("<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"Century\"/><w:lang w:val=\"en-US\"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>", sz),
            ("<w:docDefaults><w:rPrDefault/><w:pPrDefault/></w:docDefaults>", sz),
            ("<w:docDefaults><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>", sz),
            ("<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val=\"21\"/></w:rPr></w:rPrDefault></w:docDefaults>", pPr),
            ("<w:docDefaults/>", "\(sz), \(pPr)"),
            ("", "\(sz), \(pPr)")
        ] {
            assertImport(try themeFreeParts(docDefaults: docDefaults), throws: .missingRequiredFormatting(expected), docDefaults)
        }
        XCTAssertNoThrow(try importing(try themeFreeParts(docDefaults: "<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val=\"21\"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>")),
                         "an explicit size with an empty pPrDefault is complete")
    }

    /// A size that is present but not a positive integer is invalid data,
    /// not missing data.
    func testInvalidDocDefaultsSizeIsReportedAsInvalid() throws {
        for value in ["0", "-2", "abc", ""] {
            let parts = try themeFreeParts(docDefaults: "<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val=\"\(value)\"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>")
            assertImport(parts, throws: .invalidSnapshot("docDefaults/rPrDefault/rPr/sz 必須是正整數（半點），實際為「\(value)」"), value)
        }
        let parts = try themeFreeParts(docDefaults: "<w:docDefaults><w:rPrDefault><w:rPr><w:sz/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>")
        assertImport(parts, throws: .invalidSnapshot("docDefaults/rPrDefault/rPr/sz 必須是正整數（半點），實際為「」"), "no val")
    }

    /// The descriptions carry the fix: which dialog, which button, and the
    /// equivalent XML, for each missing default.
    func testCompletenessErrorDescriptionsAreActionable() {
        let size = DocumentFormattingProfileError.missingRequiredFormatting("docDefaults/rPrDefault/rPr/sz").errorDescription ?? ""
        for fragment in ["docDefaults/rPrDefault/rPr/sz", "預設字級", "隱含", "字型", "設為預設值", "<w:sz w:val="] {
            XCTAssertTrue(size.contains(fragment), "size description lacks \(fragment): \(size)")
        }
        let paragraph = DocumentFormattingProfileError.missingRequiredFormatting("docDefaults/pPrDefault").errorDescription ?? ""
        for fragment in ["docDefaults/pPrDefault", "段落", "設為預設值", "<w:pPrDefault>"] {
            XCTAssertTrue(paragraph.contains(fragment), "paragraph description lacks \(fragment): \(paragraph)")
        }
        let both = DocumentFormattingProfileError.missingRequiredFormatting("docDefaults/rPrDefault/rPr/sz, docDefaults/pPrDefault").errorDescription ?? ""
        XCTAssertTrue(both.contains("<w:sz w:val=") && both.contains("<w:pPrDefault>"), both)
        let numbering = DocumentFormattingProfileError.unsupportedNumbering.errorDescription ?? ""
        for fragment in ["編號", "abstractNum", "numId", "第一版", "inherit"] {
            XCTAssertTrue(numbering.contains(fragment), "numbering description lacks \(fragment): \(numbering)")
        }
        let other = DocumentFormattingProfileError.missingRequiredFormatting("final body sectPr").errorDescription
        XCTAssertEqual(other, "格式快照缺少必要資料：final body sectPr", "fields without a known fix keep the plain message")
    }

    // MARK: - Real templates (MACDOC_TEMPLATE_DIR)

    /// 90_template_ja relies on Word's implicit default font size; its
    /// empty pPrDefault is accepted. A temp-dir copy that adds only the
    /// explicit size imports, so the size is the one missing field. The
    /// source is read in place and never modified.
    func testRealBaselineTemplateIsMissingOnlyTheExplicitDefaultSize() throws {
        let url = try TemplateFixtureGate.requireTemplate(TemplateFixtureGate.baselineTemplateName)
        let before = try Data(contentsOf: url)
        XCTAssertThrowsError(try DocumentFormattingProfile.importOfficial(from: url)) { error in
            XCTAssertEqual(error as? DocumentFormattingProfileError, .missingRequiredFormatting("docDefaults/rPrDefault/rPr/sz"))
            XCTAssertTrue((error as? LocalizedError)?.errorDescription?.contains("設為預設值") == true)
        }
        var parts = try RawPartChannel.readAllParts(from: url)
        var styles = String(decoding: try XCTUnwrap(parts["word/styles.xml"]), as: UTF8.self)
        let defaults = try XCTUnwrap(styles.range(of: "<w:docDefaults>")?.lowerBound)..<XCTUnwrap(styles.range(of: "</w:docDefaults>")?.upperBound)
        let explicit = String(styles[defaults]).replacingOccurrences(of: "<w:lang ", with: "<w:sz w:val=\"21\"/><w:lang ")
        XCTAssertNotEqual(explicit, String(styles[defaults]), "fixture docDefaults shape changed; update this derivation")
        XCTAssertTrue(explicit.contains("<w:pPrDefault/>"), "the derivation keeps the empty pPrDefault")
        styles.replaceSubrange(defaults, with: explicit)
        parts["word/styles.xml"] = Data(styles.utf8)
        XCTAssertNoThrow(try importing(parts))
        XCTAssertEqual(try Data(contentsOf: url), before)
    }
}
