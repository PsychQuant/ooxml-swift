import XCTest
@testable import OOXMLSwift

final class DocumentFormattingProfileTests: XCTestCase {
    let w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    let a = "http://schemas.openxmlformats.org/drawingml/2006/main"

    func testFreshReducerFormattingSurvivesBothWritersAndLaterTypedEdit() throws {
        for useProfile in [false, true] {
            let url = try directory().appendingPathComponent("original.docx")
            try DocxWriter.writeData(WordDocument()).write(to: url)
            var doc = try DocxReader.read(from: url)
            defer { doc.close() }
            if useProfile { try doc.applyFormattingProfile(DocumentFormattingProfile.importOfficial(from: template()), context: .existingDocument) }
            let size = try XCTUnwrap(ProfileXML.walk(doc.xmlTrees["word/styles.xml"]!.root).first { $0.localName == "sz" })
            size.libraryUUID = UUID()
            try doc.apply(operations: [.updateAttribute(target: ElementID(node: size)!, prefix: "w", localName: "val", value: "48")])
            for output in [try parts(doc), try parts(doc, authoring: true)] {
                let root = try ProfileXML.parse(output["word/styles.xml"]!)
                let defaults = try XCTUnwrap(ProfileXML.child(root, "docDefaults"))
                XCTAssertEqual(ProfileXML.walk(defaults).first { $0.localName == "sz" }.flatMap { ProfileXML.value($0, "val") }, "48")
            }
            try doc.updateStyle(id: "Normal", with: StyleUpdate(name: "Later typed edit"))
            for output in [try parts(doc), try parts(doc, authoring: true)] {
                let defaults = try XCTUnwrap(ProfileXML.child(ProfileXML.parse(output["word/styles.xml"]!), "docDefaults"))
                XCTAssertEqual(ProfileXML.walk(defaults).first { $0.localName == "sz" }.flatMap { ProfileXML.value($0, "val") }, "48")
                XCTAssertTrue(output["word/styles.xml"]!.contains("Later typed edit"))
            }
        }
    }

    func testLaterAuthoritativeThemeFontsAndCarriedStylesRemainDurable() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.applyFormattingProfile(DocumentFormattingProfile.importOfficial(from: template()), context: .newDocument)
        let theme = try XCTUnwrap(ProfileXML.walk(doc.xmlTrees["word/theme/theme1.xml"]!.root).first { $0.localName == "latin" })
        theme.libraryUUID = UUID()
        try doc.apply(operations: [.updateAttribute(target: ElementID(node: theme)!, prefix: nil, localName: "typeface", value: "Later Theme")])
        let font = try XCTUnwrap(doc.xmlTrees["word/fontTable.xml"]?.root.children.first { $0.localName == "font" })
        font.libraryUUID = UUID()
        try doc.apply(operations: [.updateAttribute(target: ElementID(node: font)!, prefix: "w", localName: "name", value: "Reducer Font")])
        for output in [try parts(doc), try parts(doc, authoring: true)] {
            XCTAssertTrue(output["word/fontTable.xml"]!.contains("Reducer Font"))
        }
        let changedStyles = try ProfileXML.string(doc.xmlTrees["word/styles.xml"]!.root).replacingOccurrences(of: "w:val=\"24\"", with: "w:val=\"48\"")
        try doc.apply(operations: [.carryPart(partPath: "word/styles.xml", xml: changedStyles), .carryPart(partPath: "word/fontTable.xml", xml: "<w:fonts xmlns:w=\"\(w)\"><w:font w:name=\"Later Font\"/></w:fonts>")])
        try doc.updateStyle(id: "Normal", with: StyleUpdate(name: "After XML edits"))
        for output in [try parts(doc), try parts(doc, authoring: true)] {
            XCTAssertTrue(output["word/styles.xml"]!.contains("w:val=\"48\""))
            XCTAssertTrue(output["word/styles.xml"]!.contains("After XML edits"))
            XCTAssertTrue(output["word/theme/theme1.xml"]!.contains("Later Theme"))
            XCTAssertTrue(output["word/fontTable.xml"]!.contains("Later Font"))
        }
    }

    func testUndoRestoresFormattingBaselineBeforeLaterTypedEdit() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.applyFormattingProfile(DocumentFormattingProfile.importOfficial(from: template()), context: .newDocument)
        let size = try XCTUnwrap(ProfileXML.walk(doc.xmlTrees["word/styles.xml"]!.root).first { $0.localName == "sz" })
        size.libraryUUID = UUID()
        try doc.apply(operations: [.batchBegin(label: "Update defaults"), .updateAttribute(target: ElementID(node: size)!, prefix: "w", localName: "val", value: "48"), .batchEnd])
        let op = try XCTUnwrap(doc.operationLog.entries.first?.opID)
        try doc.apply(operations: [.undo(targetOpID: op)])
        try doc.updateStyle(id: "Normal", with: StyleUpdate(name: "After undo"))
        for output in [try parts(doc), try parts(doc, authoring: true)] {
            let defaults = try XCTUnwrap(ProfileXML.child(ProfileXML.parse(output["word/styles.xml"]!), "docDefaults"))
            XCTAssertEqual(ProfileXML.walk(defaults).first { $0.localName == "sz" }.flatMap { ProfileXML.value($0, "val") }, "24")
        }
    }

    func testRemovedDefaultsRemainAbsentAfterLaterTypedEdit() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.applyFormattingProfile(DocumentFormattingProfile.importOfficial(from: template()), context: .newDocument)
        let defaults = try XCTUnwrap(ProfileXML.child(doc.xmlTrees["word/styles.xml"]!.root, "docDefaults"))
        defaults.libraryUUID = UUID()
        try doc.apply(operations: [.removeNode(target: ElementID(node: defaults)!)])
        for output in [try parts(doc), try parts(doc, authoring: true)] {
            XCTAssertNil(ProfileXML.child(try ProfileXML.parse(output["word/styles.xml"]!), "docDefaults"))
        }
        try doc.updateStyle(id: "Normal", with: StyleUpdate(name: "No defaults"))
        for output in [try parts(doc), try parts(doc, authoring: true)] {
            XCTAssertNil(ProfileXML.child(try ProfileXML.parse(output["word/styles.xml"]!), "docDefaults"))
        }
    }

    func testMixedCarryAndReducerUsesCommittedFormattingAfterTypedEdit() throws {
        for carryFirst in [true, false] {
            var doc = WordDocument.emptyAuthoringDocument()
            try doc.applyFormattingProfile(DocumentFormattingProfile.importOfficial(from: template()), context: .newDocument)
            let size = try XCTUnwrap(ProfileXML.walk(doc.xmlTrees["word/styles.xml"]!.root).first { $0.localName == "sz" })
            size.libraryUUID = UUID()
            let carried = try ProfileXML.string(doc.xmlTrees["word/styles.xml"]!.root)
                .replacingOccurrences(of: "w:val=\"24\"", with: "w:val=\"48\"")
                .replacingOccurrences(of: "<w:name w:val=\"Normal\"/>", with: "<w:name w:val=\"Carried name\"/>")
            let carry = Operation.carryPart(partPath: "word/styles.xml", xml: carried)
            let reducer = Operation.updateAttribute(target: ElementID(node: size)!, prefix: "w", localName: "val", value: "64")
            try doc.apply(operations: carryFirst ? [carry, reducer] : [reducer, carry])
            XCTAssertNil(doc.carriedParts["word/styles.xml"])
            XCTAssertEqual(doc.styles.first { $0.id == "Normal" }?.name, "Normal")
            for output in [try parts(doc), try parts(doc, authoring: true)] {
                let defaults = try XCTUnwrap(ProfileXML.child(ProfileXML.parse(output["word/styles.xml"]!), "docDefaults"))
                XCTAssertEqual(ProfileXML.walk(defaults).first { $0.localName == "sz" }.flatMap { ProfileXML.value($0, "val") }, "64")
            }
            try doc.updateStyle(id: "Normal", with: StyleUpdate(name: "After mixed"))
            for output in [try parts(doc), try parts(doc, authoring: true)] {
                let defaults = try XCTUnwrap(ProfileXML.child(ProfileXML.parse(output["word/styles.xml"]!), "docDefaults"))
                XCTAssertEqual(ProfileXML.walk(defaults).first { $0.localName == "sz" }.flatMap { ProfileXML.value($0, "val") }, "64")
                XCTAssertTrue(output["word/styles.xml"]!.contains("After mixed"))
            }
        }
    }

    func testAliasedTargetStylesKeepReferencesAndUnknownMarkup() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        let styles = "<x:styles xmlns:x=\"\(w)\" xmlns:custom=\"urn:target-owned\"><x:style x:type=\"paragraph\" x:default=\"1\" x:styleId=\"TargetDefault\"><x:name x:val=\"Default\"/></x:style><x:style x:type=\"paragraph\" x:styleId=\"AliasedTarget\"><x:name x:val=\"Keep target\"/><x:basedOn x:val=\"TargetDefault\"/><custom:preserve custom:setting=\"keep\"/><w:extension xmlns:w=\"urn:shadow\" x:flag=\"kept\"/></x:style></x:styles>"
        try doc.apply(operations: [.carryPart(partPath: "word/styles.xml", xml: styles), .appendParagraph(in: nil, paragraph: ParagraphPayload(text: "Keep reference", styleId: "AliasedTarget", paraId: "ABCD1234"))])
        try doc.applyFormattingProfile(DocumentFormattingProfile.importOfficial(from: template()), context: .existingDocument)
        let output = try parts(doc)
        XCTAssertTrue(output["word/styles.xml"]!.contains("AliasedTarget"))
        XCTAssertTrue(output["word/styles.xml"]!.contains("TargetDefault"))
        XCTAssertTrue(output["word/styles.xml"]!.contains("custom:preserve"))
        XCTAssertTrue(output["word/document.xml"]!.contains("AliasedTarget"))
        XCTAssertNotNil(ProfileXML.walk(try ProfileXML.parse(output["word/styles.xml"]!)).first { $0.namespaceURI == "urn:shadow" && $0.localName == "extension" })
    }

    func testDuplicateSingletonFormattingIsRejectedDuringImportAndDecode() throws {
        let profile = try DocumentFormattingProfile.importOfficial(from: template())
        for (key, xml) in [
            ("stylesXML", profile.stylesXML!.replacingOccurrences(of: "<w:name w:val=\"Normal\"/>", with: "<w:name w:val=\"Normal\"/><w:next w:val=\"a\"/><w:next w:val=\"Missing\"/>")),
            ("stylesXML", profile.stylesXML!.replacingOccurrences(of: "<w:sz w:val=\"24\"/>", with: "<w:sz w:val=\"24\"/><w:sz w:val=\"48\"/>")),
            ("stylesXML", profile.stylesXML!.replacingOccurrences(of: "</w:docDefaults>", with: "<w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>")),
            ("sectionXML", profile.sectionXML!.replacingOccurrences(of: "</w:sectPr>", with: "<w:pgSz w:w=\"1\" w:h=\"2\"/></w:sectPr>"))
        ] {
            var json = try XCTUnwrap(JSONSerialization.jsonObject(with: JSONEncoder().encode(profile)) as? [String: Any])
            json[key] = xml
            XCTAssertThrowsError(try JSONDecoder().decode(DocumentFormattingProfile.self, from: JSONSerialization.data(withJSONObject: json)))
            if key == "stylesXML" { XCTAssertThrowsError(try DocumentFormattingProfile.importOfficial(from: template(styles: xml))) }
        }
    }

    func testGenericReaderPreservesUTF16AncillaryBytesWithoutProfile() throws {
        let root = try directory()
        let package = root.appendingPathComponent("package")
        let original = root.appendingPathComponent("original.docx")
        try DocxWriter.writeData(WordDocument()).write(to: original)
        var all = try RawPartChannel.readAllParts(from: original)
        let fonts = "<?xml version=\"1.0\" encoding=\"UTF-16\"?><w:fonts xmlns:w=\"\(w)\"><w:font w:name=\"UTF16 Font\"/></w:fonts>".data(using: .utf16)!
        all["word/fontTable.xml"] = fonts
        for (path, bytes) in all {
            let url = package.appendingPathComponent(path)
            try FileManager.default.createDirectory(at: url.deletingLastPathComponent(), withIntermediateDirectories: true)
            try bytes.write(to: url)
        }
        try ZipHelper.zipToData(package).write(to: original)
        var doc = try DocxReader.read(from: original)
        defer { doc.close() }
        let saved = root.appendingPathComponent("saved.docx")
        try DocxWriter.write(doc, to: saved)
        XCTAssertEqual(try RawPartChannel.readAllParts(from: saved)["word/fontTable.xml"], fonts)
        try doc.updateStyle(id: "Normal", with: StyleUpdate(name: "Still UTF16"))
        try DocxWriter.write(doc, to: saved)
        XCTAssertEqual(try RawPartChannel.readAllParts(from: saved)["word/fontTable.xml"], fonts)
    }

    // MARK: - #195 canonicalWordTree namespace-declaration dedup

    /// `xmlns:w` must appear only on the root of a `canonicalWordTree`
    /// output; descendants inherit the root's binding through ordinary XML
    /// scoping instead of each re-declaring it. Regression for #195.
    func testCanonicalWordTreeDeclaresNamespaceOnlyOnce() throws {
        let nested = """
        <x:styles xmlns:x="\(w)">
          <x:style x:type="paragraph" x:styleId="a">
            <x:name x:val="Normal"/>
            <x:pPr>
              <x:rPr><x:rFonts x:ascii="Aptos"/><x:sz x:val="24"/></x:rPr>
            </x:pPr>
          </x:style>
        </x:styles>
        """
        let canonical = ProfileXML.canonicalWordTree(try ProfileXML.parse(nested))
        let serialized = try ProfileXML.string(canonical)
        XCTAssertEqual(serialized.components(separatedBy: "xmlns:w=\"\(w)\"").count - 1, 1)
    }

    /// Same fixture, decoded from UTF-16 the way `DocxReader` decodes
    /// `word/styles.xml` when it carries a non-UTF-8 encoding declaration.
    func testCanonicalWordTreeDeclaresNamespaceOnlyOnceForUTF16StylesInput() throws {
        let root = try directory()
        let package = root.appendingPathComponent("package")
        let original = root.appendingPathComponent("original.docx")
        try DocxWriter.writeData(WordDocument()).write(to: original)
        var all = try RawPartChannel.readAllParts(from: original)
        let styles = """
        <?xml version="1.0" encoding="UTF-16"?>
        <x:styles xmlns:x="\(w)">
          <x:docDefaults><x:rPrDefault><x:rPr><x:sz x:val="24"/></x:rPr></x:rPrDefault></x:docDefaults>
          <x:style x:type="paragraph" x:default="1" x:styleId="Normal">
            <x:name x:val="Normal"/>
            <x:pPr><x:rPr><x:rFonts x:ascii="Aptos"/></x:rPr></x:pPr>
          </x:style>
        </x:styles>
        """
        all["word/styles.xml"] = styles.data(using: .utf16)!
        for (path, bytes) in all {
            let url = package.appendingPathComponent(path)
            try FileManager.default.createDirectory(at: url.deletingLastPathComponent(), withIntermediateDirectories: true)
            try bytes.write(to: url)
        }
        try ZipHelper.zipToData(package).write(to: original)
        var doc = try DocxReader.read(from: original)
        defer { doc.close() }
        let originalStylesXML = try XCTUnwrap(doc.formattingState?.originalStylesXML)
        XCTAssertEqual(originalStylesXML.components(separatedBy: "xmlns:w=\"\(w)\"").count - 1, 1)
    }

    /// Regression lock for `DocxReader`'s manual copy of the canonical
    /// root's namespace declarations onto the detached `docDefaults` clone
    /// (`DocxReader.swift` around the `defaultsXML` assignment). Before
    /// #195's fix this copy was largely redundant, because every node —
    /// `docDefaults` included — already carried its own `xmlns:w`. After the
    /// fix, non-root nodes no longer self-declare, so this manual copy is
    /// the *only* source of `xmlns:w` on the detached, independently
    /// re-parsed `defaultsXML` string; if it were ever deleted as
    /// "dead code", `defaultsXML` would become invalid standalone XML.
    func testDetachedDocDefaultsCloneStillCarriesRootNamespaceDeclarations() throws {
        let root = try directory()
        let package = root.appendingPathComponent("package")
        let original = root.appendingPathComponent("original.docx")
        try DocxWriter.writeData(WordDocument()).write(to: original)
        var all = try RawPartChannel.readAllParts(from: original)
        all["word/styles.xml"] = Data("""
        <x:styles xmlns:x="\(w)">
          <x:docDefaults><x:rPrDefault><x:rPr><x:rFonts x:ascii="Aptos"/><x:sz x:val="24"/></x:rPr></x:rPrDefault></x:docDefaults>
          <x:style x:type="paragraph" x:default="1" x:styleId="Normal"><x:name x:val="Normal"/></x:style>
        </x:styles>
        """.utf8)
        for (path, bytes) in all {
            let url = package.appendingPathComponent(path)
            try FileManager.default.createDirectory(at: url.deletingLastPathComponent(), withIntermediateDirectories: true)
            try bytes.write(to: url)
        }
        try ZipHelper.zipToData(package).write(to: original)
        var doc = try DocxReader.read(from: original)
        defer { doc.close() }
        let defaultsXML = try XCTUnwrap(doc.formattingState?.defaultsXML)
        // Must be independently parseable — it is later re-parsed on its own
        // (e.g. in `writeFormattingParts`) and spliced back in as a child.
        let parsed = try ProfileXML.parse(defaultsXML)
        XCTAssertEqual(parsed.localName, "docDefaults")
        XCTAssertEqual(defaultsXML.components(separatedBy: "xmlns:w=\"\(w)\"").count - 1, 1)
    }

    /// Aliased target prefixes and unknown foreign namespaces must not be
    /// affected by the dedup — the `w` binding still collapses to one
    /// declaration while the foreign `custom:`/`urn:shadow` bindings, which
    /// are distinct URIs, still declare normally.
    func testCanonicalWordTreeDedupPreservesAliasedAndUnknownNamespaces() throws {
        let styles = "<x:styles xmlns:x=\"\(w)\" xmlns:custom=\"urn:target-owned\"><x:style x:type=\"paragraph\" x:default=\"1\" x:styleId=\"TargetDefault\"><x:name x:val=\"Default\"/></x:style><x:style x:type=\"paragraph\" x:styleId=\"AliasedTarget\"><x:name x:val=\"Keep target\"/><x:basedOn x:val=\"TargetDefault\"/><custom:preserve custom:setting=\"keep\"/><w:extension xmlns:w=\"urn:shadow\" x:flag=\"kept\"/></x:style></x:styles>"
        let canonical = ProfileXML.canonicalWordTree(try ProfileXML.parse(styles))
        let serialized = try ProfileXML.string(canonical)
        XCTAssertEqual(serialized.components(separatedBy: "xmlns:w=\"\(w)\"").count - 1, 1)
        XCTAssertTrue(serialized.contains("AliasedTarget"))
        XCTAssertTrue(serialized.contains("TargetDefault"))
        XCTAssertTrue(serialized.contains("custom:preserve"))
        XCTAssertNotNil(ProfileXML.walk(canonical).first { $0.namespaceURI == "urn:shadow" && $0.localName == "extension" })
    }

    /// A second typed style edit must not grow the number of `xmlns:w`
    /// declarations in the re-serialized `word/styles.xml` — the durable
    /// baseline the first edit persisted must already be deduped, and
    /// `mergeFormattingChanges`'s `deepClone()` of untouched styles must not
    /// reintroduce per-node declarations. (The count itself need not be 1:
    /// `state.defaultsXML` is deliberately kept as an independently
    /// self-declaring standalone document — see `withWordNamespace()` — so
    /// it still carries its own `xmlns:w` once it is spliced back in as a
    /// `docDefaults` child; that is a fixed, non-growing part of the count,
    /// not the per-node bug #195 targets.)
    func testSecondTypedEditDoesNotGrowNamespaceDeclarationCount() throws {
        // The growth only compounds on reopen: each `DocxReader.read` reruns
        // `canonicalWordTree` on whatever `word/styles.xml` currently holds,
        // so a first typed edit's (pre-fix, per-node) declarations become
        // the input to the second edit's canonicalization, doubling up.
        let profile = try DocumentFormattingProfile.importOfficial(from: template())
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.applyFormattingProfile(profile, context: .newDocument)
        try doc.updateStyle(id: "Normal", with: StyleUpdate(name: "First edit"))
        let firstURL = try directory().appendingPathComponent("first.docx")
        try DocxWriter.write(doc, to: firstURL)
        doc.close()
        let firstOutput = String(decoding: try XCTUnwrap(try RawPartChannel.readAllParts(from: firstURL)["word/styles.xml"]), as: UTF8.self)
        let firstCount = firstOutput.components(separatedBy: "xmlns:w=\"\(w)\"").count - 1

        var reopened = try DocxReader.read(from: firstURL)
        defer { reopened.close() }
        try reopened.updateStyle(id: "TitleLocal", with: StyleUpdate(name: "Second edit"))
        let secondURL = try directory().appendingPathComponent("second.docx")
        try DocxWriter.write(reopened, to: secondURL)
        let secondOutput = String(decoding: try XCTUnwrap(try RawPartChannel.readAllParts(from: secondURL)["word/styles.xml"]), as: UTF8.self)
        let secondCount = secondOutput.components(separatedBy: "xmlns:w=\"\(w)\"").count - 1

        XCTAssertEqual(secondCount, firstCount)
        XCTAssertTrue(firstOutput.contains("First edit"))
        XCTAssertTrue(secondOutput.contains("First edit"))
        XCTAssertTrue(secondOutput.contains("Second edit"))
    }

    func directory() throws -> URL {
        let url = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString)
        try FileManager.default.createDirectory(at: url, withIntermediateDirectories: true)
        addTeardownBlock { try? FileManager.default.removeItem(at: url) }
        return url
    }

    func testEffectiveThemeMatchesProfileThenCallerEditsAcrossWritersAndReopen() throws {
        let profile = try DocumentFormattingProfile.importOfficial(from: template())
        let sourceTheme = try XCTUnwrap(profile.themeXML).replacingOccurrences(of: "Aptos", with: "Source Font")
        let callerTheme = sourceTheme.replacingOccurrences(of: "Source Font", with: "Caller Font")
            .replacingOccurrences(of: "</a:theme>", with: "<x:custom xmlns:x=\"urn:caller-theme\" value=\"keep\"/></a:theme>")
        for sourceHasTheme in [false, true] {
            let root = try directory(), source = root.appendingPathComponent("source.docx")
            var initial = WordDocument.emptyAuthoringDocument()
            if sourceHasTheme { try initial.apply(operations: [.carryPart(partPath: "word/theme/theme1.xml", xml: sourceTheme)]) }
            try initial.writeAuthoringPackage(to: source)
            var doc = try DocxReader.read(from: source)
            defer { doc.close() }
            try doc.applyFormattingProfile(profile, context: .existingDocument)
            let effective = String(decoding: try XCTUnwrap(doc.effectiveThemeData()), as: UTF8.self)
            XCTAssertTrue(effective.contains("Aptos"))
            XCTAssertTrue(effective.contains("DFKai-SB"))
            XCTAssertFalse(effective.contains("Source Font"))
            for output in [try parts(doc), try parts(doc, authoring: true)] {
                XCTAssertEqual(output["word/theme/theme1.xml"], effective)
            }
            try doc.apply(operations: [.carryPart(partPath: "word/theme/theme1.xml", xml: callerTheme)])
            XCTAssertEqual(try doc.effectiveThemeData(), Data(callerTheme.utf8))
            try doc.updateStyle(id: "Normal", with: StyleUpdate(name: "After caller theme"))
            for authoring in [false, true] {
                let output = root.appendingPathComponent("edited-\(authoring).docx")
                if authoring { try doc.writeAuthoringPackage(to: output) }
                else { try DocxWriter.write(doc, to: output) }
                XCTAssertEqual(try RawPartChannel.readAllParts(from: output)["word/theme/theme1.xml"], Data(callerTheme.utf8))
                var reopened = try DocxReader.read(from: output)
                defer { reopened.close() }
                XCTAssertEqual(try reopened.effectiveThemeData(), Data(callerTheme.utf8))
            }
        }
    }

    func testCarriedThemeWithoutProfilePersistsInBothWritersAndRegistersMetadata() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        let xml = "<a:theme xmlns:a=\"\(a)\" name=\"Caller\"><x:custom xmlns:x=\"urn:caller-theme\"/></a:theme>"
        XCTAssertNil(try doc.effectiveThemeData())
        doc.markPartDirty("word/theme/theme1.xml")
        try doc.apply(operations: [.carryPart(partPath: "word/theme/theme1.xml", xml: xml)])
        XCTAssertEqual(try doc.effectiveThemeData(), Data(xml.utf8))
        for output in [try parts(doc), try parts(doc, authoring: true)] {
            XCTAssertEqual(output["word/theme/theme1.xml"], xml)
            XCTAssertTrue(output["[Content_Types].xml"]!.contains("/word/theme/theme1.xml"))
            XCTAssertTrue(output["word/_rels/document.xml.rels"]!.contains("theme/theme1.xml"))
        }
    }

    func testUnmodifiedCarriedThemeKeepsReplayMetadataByteExact() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        let original = try parts(doc, authoring: true)
        let types = original["[Content_Types].xml"]! + "\n<!-- caller content types -->"
        let rels = original["word/_rels/document.xml.rels"]! + "\n<!-- caller relationships -->"
        let theme = "<a:theme xmlns:a=\"\(a)\" name=\"Replay\"/>"
        try doc.apply(operations: [
            .carryPart(partPath: "word/theme/theme1.xml", xml: theme),
            .carryPart(partPath: "[Content_Types].xml", xml: types),
            .carryPart(partPath: "word/_rels/document.xml.rels", xml: rels)
        ])
        let replayed = try parts(doc, authoring: true)
        XCTAssertEqual(replayed["word/theme/theme1.xml"], theme)
        XCTAssertEqual(replayed["[Content_Types].xml"], types)
        XCTAssertEqual(replayed["word/_rels/document.xml.rels"], rels)
    }

    func testTypedLatentStylesSetAndClearSurviveProfileFinalizationInBothWriters() throws {
        for profile in [DocumentFormattingProfile.inherit, try DocumentFormattingProfile.importOfficial(from: template())] {
            var doc = WordDocument.emptyAuthoringDocument()
            try doc.applyFormattingProfile(profile, context: .newDocument)
            doc.setLatentStyles([LatentStyle(name: "Heading 9", uiPriority: 9, semiHidden: true, unhideWhenUsed: false, qFormat: false)])
            for output in [try parts(doc), try parts(doc, authoring: true)] {
                let styles = try ProfileXML.parse(output["word/styles.xml"]!)
                let latent = try XCTUnwrap(ProfileXML.child(styles, "latentStyles"))
                let entry = try XCTUnwrap(ProfileXML.child(latent, "lsdException"))
                XCTAssertEqual(ProfileXML.value(entry, "name"), "Heading 9")
                XCTAssertEqual(ProfileXML.value(entry, "semiHidden"), "1")
            }
            doc.setLatentStyles([])
            for output in [try parts(doc), try parts(doc, authoring: true)] {
                XCTAssertNil(ProfileXML.child(try ProfileXML.parse(output["word/styles.xml"]!), "latentStyles"))
            }
        }
    }

    func template(styles: String? = nil, section: String? = nil, extras: [String: String] = [:]) throws -> URL {
        let dir = try directory()
        let source = dir.appendingPathComponent("source")
        var parts = [
            "word/styles.xml": styles ?? """
            <x:styles xmlns:x="\(w)" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <x:docDefaults><x:rPrDefault><x:rPr><x:rFonts x:asciiTheme="minorHAnsi" x:eastAsiaTheme="minorEastAsia"/><x:sz x:val="24"/></x:rPr></x:rPrDefault><x:pPrDefault><x:pPr><x:spacing x:after="160" x:line="278" x:lineRule="auto"/></x:pPr></x:pPrDefault></x:docDefaults>
              <x:style x:type="paragraph" x:default="1" x:styleId="a"><x:name x:val="Normal"/><x:pPr><x:spacing x:after="160"/></x:pPr></x:style>
              <x:style x:type="paragraph" x:styleId="TitleLocal"><x:name x:val="Title"/><x:basedOn x:val="a"/><x:rPr><x:b/></x:rPr></x:style>
            </x:styles>
            """,
            "word/document.xml": "<x:document xmlns:x=\"\(w)\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><x:body><x:p><x:r><x:t>PRIVATE BODY</x:t></x:r></x:p>\(section ?? "<x:sectPr x:rsidSect=\"PRIVATE-REVISION\"><x:headerReference x:type=\"default\" r:id=\"rId9\"/><x:pgSz x:w=\"11906\" x:h=\"16838\"/><x:pgMar x:top=\"1440\" x:right=\"1800\" x:bottom=\"1440\" x:left=\"1800\" x:header=\"720\" x:footer=\"720\" x:gutter=\"0\"/></x:sectPr>")</x:body></x:document>",
            "word/theme/theme1.xml": "<a:theme xmlns:a=\"\(a)\" name=\"Office\"><a:themeElements><a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"Aptos Display\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Hant\" typeface=\"新細明體\"/></a:majorFont><a:minorFont><a:latin typeface=\"Aptos\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Hant\" typeface=\"新細明體\"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>",
            "word/fontTable.xml": "<w:fonts xmlns:w=\"\(w)\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:font w:name=\"Aptos\"><w:embedRegular r:id=\"rIdSecret\"/></w:font></w:fonts>",
            "word/vbaProject.bin": "PRIVATE MACRO", "docProps/core.xml": "PRIVATE AUTHOR",
            "word/_rels/document.xml.rels": "PRIVATE EXTERNAL LINK"
        ]
        parts.merge(extras) { _, new in new }
        for (path, text) in parts {
            let file = source.appendingPathComponent(path)
            try FileManager.default.createDirectory(at: file.deletingLastPathComponent(), withIntermediateDirectories: true)
            try Data(text.utf8).write(to: file)
        }
        let url = dir.appendingPathComponent("Normal.dotm")
        try ZipHelper.zipToData(source).write(to: url)
        return url
    }

    func parts(_ doc: WordDocument, authoring: Bool = false) throws -> [String: String] {
        let url = try directory().appendingPathComponent("out.docx")
        if authoring { try doc.writeAuthoringPackage(to: url) }
        else { try DocxWriter.writeData(doc).write(to: url) }
        return try RawPartChannel.readAllParts(from: url).mapValues { String(decoding: $0, as: UTF8.self) }
    }

    func testSafeSnapshotNormalizesAliasesExcludesPrivateDataAndSurvivesSourceRemoval() throws {
        let url = try template()
        let original = try Data(contentsOf: url)
        let profile = try DocumentFormattingProfile.importOfficial(from: url)
        let encoded = try JSONEncoder().encode(profile)
        let json = String(decoding: encoded, as: UTF8.self)
        for secret in ["PRIVATE", "rId9", "rIdSecret", "embedRegular", url.path, "vbaProject"] { XCTAssertFalse(json.contains(secret), secret) }
        XCTAssertEqual(try Data(contentsOf: url), original)
        try FileManager.default.removeItem(at: url)
        let decoded = try JSONDecoder().decode(DocumentFormattingProfile.self, from: encoded)
        var doc = WordDocument()
        try doc.applyFormattingProfile(decoded, context: .newDocument)
        XCTAssertEqual(doc.sectionProperties.pageSize, PageSize(width: 11906, height: 16838))
        XCTAssertEqual(doc.sectionProperties.pageMargins.left, 1800)
        let output = try parts(doc)
        XCTAssertTrue(output["word/styles.xml"]!.contains("w:val=\"24\""))
        XCTAssertTrue(output["word/styles.xml"]!.contains("w:eastAsia=\"DFKai-SB\""))
        XCTAssertFalse(output["word/styles.xml"]!.contains("eastAsiaTheme"))
        XCTAssertTrue(output["word/theme/theme1.xml"]!.contains("Aptos"))
        XCTAssertFalse(output["word/theme/theme1.xml"]!.contains("新細明體"))
        XCTAssertTrue(output["word/theme/theme1.xml"]!.contains("typeface=\"DFKai-SB\""))
        XCTAssertTrue(output["word/fontTable.xml"]!.contains("w:name=\"DFKai-SB\""))
        XCTAssertFalse(output["word/styles.xml"]!.contains("w:eastAsia=\"標楷體\""))
    }

    func testBothWritersPersistDefaultsAfterTypedStyleMutationAndRetainReferencedStyles() throws {
        let profile = try DocumentFormattingProfile.importOfficial(from: template())
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: [.appendParagraph(in: nil, paragraph: ParagraphPayload(text: "hello", styleId: "Heading1", paraId: "ABCD1234"))])
        try doc.applyFormattingProfile(profile, context: .existingDocument)
        var changedRun = RunProperties()
        changedRun.rFonts = RFontsProperties(ascii: "Caller Western", hAnsi: "Caller Western")
        try doc.updateStyle(id: "Heading1", with: StyleUpdate(name: "Updated heading", runProperties: changedRun))
        for output in [try parts(doc), try parts(doc, authoring: true), try parts(doc)] {
            let styles = output["word/styles.xml"]!
            XCTAssertTrue(styles.contains("w:val=\"24\""))
            XCTAssertTrue(styles.contains("Updated heading"))
            XCTAssertTrue(styles.contains("Caller Western"))
            XCTAssertTrue(styles.contains("w:styleId=\"Heading1\""))
            XCTAssertEqual(styles.components(separatedBy: "w:default=\"1\"").count - 1, 1)
            XCTAssertTrue(output["word/document.xml"]!.contains("11906"))
            XCTAssertTrue(output["word/_rels/document.xml.rels"]!.contains("theme/theme1.xml"))
            XCTAssertTrue(output["[Content_Types].xml"]!.contains("/word/theme/theme1.xml"))
            XCTAssertNotNil(output["word/fontTable.xml"])
        }
    }

    func testInheritOnlyRemovesGeneratorFontsForNewDocuments() throws {
        var existing = WordDocument()
        existing.properties.created = Date(timeIntervalSince1970: 1_700_000_000)
        existing.properties.modified = existing.properties.created
        let before = try parts(existing)
        try existing.applyFormattingProfile(.inherit, context: .existingDocument)
        XCTAssertEqual(try parts(existing), before)
        var fresh = WordDocument()
        fresh.styles.append(Style(id: "Explicit", name: "Explicit", type: .paragraph, runProperties: RunProperties(fontName: "Caller Font")))
        try fresh.applyFormattingProfile(.inherit, context: .newDocument)
        let output = try parts(fresh)
        let fonts = output["word/styles.xml"]! + output["word/fontTable.xml"]!
        XCTAssertFalse(fonts.contains("Calibri"))
        XCTAssertFalse(fonts.contains("Times New Roman"))
        XCTAssertTrue(fonts.contains("Caller Font"))
    }

    func testInheritPreservesExplicitSameValueFontsAndIgnoresNonFontEdits() throws {
        for explicitMode in ["setter", "axes", "initializer"] {
            var doc = WordDocument.emptyAuthoringDocument()
            let normal = try XCTUnwrap(doc.styles.firstIndex { $0.id == "Normal" })
            let before = doc.styles[normal]
            if explicitMode == "axes" {
                doc.styles[normal].runProperties?.rFonts = RFontsProperties(ascii: "Calibri", hAnsi: "Calibri", eastAsia: "Calibri", cs: "Calibri")
            } else if explicitMode == "initializer" {
                doc.styles[normal].runProperties = RunProperties(fontName: "Calibri")
            } else {
                doc.styles[normal].runProperties?.fontName = "Calibri"
                XCTAssertEqual(doc.styles[normal], before, "font origin must not affect content equality")
            }
            doc.styles[1].name = "Renamed generated heading"
            doc.styles[1].runProperties?.bold = false
            try doc.applyFormattingProfile(.inherit, context: .newDocument)
            for output in [try parts(doc), try parts(doc, authoring: true)] {
                let styles = try ProfileXML.parse(output["word/styles.xml"]!)
                let normal = try XCTUnwrap(styles.children.first { ProfileXML.value($0, "styleId") == "Normal" })
                XCTAssertEqual(ProfileXML.walk(normal).first { $0.localName == "rFonts" }.flatMap { ProfileXML.value($0, "ascii") }, "Calibri")
                let heading = try XCTUnwrap(styles.children.first { ProfileXML.value($0, "styleId") == "Heading1" })
                XCTAssertFalse(ProfileXML.walk(heading).contains { $0.localName == "rFonts" })
            }
        }
    }

    func testReaderRejectsStylesDTDInEveryUnicodeByteOrder() throws {
        let sourceURL = try directory().appendingPathComponent("source.docx")
        try DocxWriter.write(WordDocument(), to: sourceURL)
        let original = try RawPartChannel.readAllParts(from: sourceURL)
        let xml = "<?xml version=\"1.0\"?><!DOCTYPE w:styles [<!ENTITY probe \"EXPANDED\">]><w:styles xmlns:w=\"\(w)\"><w:style w:type=\"paragraph\" w:styleId=\"X\"><w:name w:val=\"&probe;\"/></w:style></w:styles>"
        let encodings: [(String, String.Encoding, [UInt8])] = [
            ("utf8", .utf8, []),
            ("utf16be", .utf16BigEndian, []), ("utf16be-bom", .utf16BigEndian, [0xFE, 0xFF]),
            ("utf16le", .utf16LittleEndian, []), ("utf16le-bom", .utf16LittleEndian, [0xFF, 0xFE]),
            ("utf32be", .utf32BigEndian, []), ("utf32be-bom", .utf32BigEndian, [0, 0, 0xFE, 0xFF]),
            ("utf32le", .utf32LittleEndian, []), ("utf32le-bom", .utf32LittleEndian, [0xFF, 0xFE, 0, 0])
        ]
        for (name, encoding, bom) in encodings {
            var parts = original
            parts["word/styles.xml"] = Data(bom) + xml.data(using: encoding)!
            let input = try directory().appendingPathComponent("\(name).docx")
            try writePackage(parts, to: input)
            XCTAssertThrowsError(try { () -> [Style] in
                var document = try DocxReader.read(from: input)
                defer { document.close() }
                return document.styles
            }(), name) { error in
                if name.hasPrefix("utf32") {
                    guard case WordError.invalidDocx = error else { return XCTFail("unsupported encoding reached the parser: \(error)") }
                } else {
                    XCTAssertEqual(error as? XMLHardeningError, .dtdNotAllowed(part: "word/styles.xml"))
                }
            }
        }
    }

    func testReaderRejectsUnsupportedUTF32StylesWithoutDTD() throws {
        let sourceURL = try directory().appendingPathComponent("source.docx")
        try DocxWriter.write(WordDocument(), to: sourceURL)
        let original = try RawPartChannel.readAllParts(from: sourceURL)
        let xml = "<?xml version=\"1.0\"?><w:styles xmlns:w=\"\(w)\"><w:style w:type=\"paragraph\" w:styleId=\"X\"><w:name w:val=\"Unexpanded\"/></w:style></w:styles>"
        for (encoding, bom) in [(String.Encoding.utf32BigEndian, [UInt8]()), (.utf32BigEndian, [0, 0, 0xFE, 0xFF]), (.utf32LittleEndian, []), (.utf32LittleEndian, [0xFF, 0xFE, 0, 0])] {
            var parts = original
            parts["word/styles.xml"] = Data(bom) + xml.data(using: encoding)!
            let input = try directory().appendingPathComponent("unsupported.docx")
            try writePackage(parts, to: input)
            XCTAssertThrowsError(try { () -> [Style] in
                var document = try DocxReader.read(from: input)
                defer { document.close() }
                return document.styles
            }()) { error in
                guard case WordError.invalidDocx = error else { return XCTFail("unsupported encoding reached the parser: \(error)") }
            }
        }
    }

    func testTypedUTF16StylesEditPreservesUnknownMarkupAndMetadata() throws {
        for (encoding, aliased, hasDefaults) in [(String.Encoding.utf16, false, true), (.utf16BigEndian, true, true), (.utf16LittleEndian, false, false)] {
        let url = try directory().appendingPathComponent("utf16.docx")
        try DocxWriter.writeData(WordDocument()).write(to: url)
        var source = try RawPartChannel.readAllParts(from: url)
        var styles = String(decoding: source["word/styles.xml"]!, as: UTF8.self)
            .replacingOccurrences(of: "UTF-8", with: "UTF-16")
            .replacingOccurrences(of: "</w:styles>", with: "<w:style w:type=\"paragraph\" w:styleId=\"Target\"><w:name w:val=\"保留樣式\"/><w:aliases w:val=\"目標別名\"/><x:keep xmlns:x=\"urn:target\" x:value=\"保真\"/></w:style></w:styles>")
        if !hasDefaults { styles = styles.replacingOccurrences(of: "<w:docDefaults>[\\s\\S]*?</w:docDefaults>", with: "", options: .regularExpression) }
        let encoded = aliased ? styles.replacingOccurrences(of: "w:", with: "z:").replacingOccurrences(of: "xmlns:w=", with: "xmlns:z=") : styles
        source["word/styles.xml"] = encoded.data(using: encoding)!
        try writePackage(source, to: url)
        var doc = try DocxReader.read(from: url)
        defer { doc.close() }
        try doc.updateStyle(id: "Target", with: StyleUpdate(name: "重新命名"))
        let output = try directory().appendingPathComponent("edited.docx")
        try DocxWriter.write(doc, to: output)
        let saved = try RawPartChannel.readAllParts(from: output)
        let parsed = try ProfileXML.parse(saved["word/styles.xml"]!)
        XCTAssertNotNil(ProfileXML.walk(parsed).first { $0.namespaceURI == "urn:target" && $0.localName == "keep" })
        XCTAssertTrue(String(decoding: saved["word/styles.xml"]!, as: UTF8.self).contains("目標別名"))
        for path in ["[Content_Types].xml", "word/_rels/document.xml.rels"] { XCTAssertEqual(saved[path], source[path], path) }
        var reopened = try DocxReader.read(from: output)
        defer { reopened.close() }
        XCTAssertEqual(reopened.styles.first { $0.id == "Target" }?.name, "重新命名")
        // Authoring's explicit carry contract supplies the unedited package
        // parts, while the reader's styles baseline drives the typed merge.
        try doc.apply(operations: source.keys.sorted().filter { $0 != "word/styles.xml" }.map {
            .carryPart(partPath: $0, xml: String(decoding: source[$0]!, as: UTF8.self))
        })
        let authoring = try directory().appendingPathComponent("authoring.docx")
        try doc.writeAuthoringPackage(to: authoring)
        let authored = try RawPartChannel.readAllParts(from: authoring)
        XCTAssertEqual(authored["word/styles.xml"], saved["word/styles.xml"])
        for path in ["[Content_Types].xml", "word/_rels/document.xml.rels"] { XCTAssertEqual(authored[path], source[path], path) }
        }
    }

    func testMetadataRepairPreservesUnknownAttributesOnUpdatedRegistrations() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        let relNS = "http://schemas.openxmlformats.org/package/2006/relationships"
        let typeNS = "http://schemas.openxmlformats.org/package/2006/content-types"
        try doc.apply(operations: [
            .carryPart(partPath: "[Content_Types].xml", xml: "<Types xmlns=\"\(typeNS)\" xmlns:x=\"urn:metadata\"><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/styles.xml\" ContentType=\"wrong/type\" x:keep=\"type-owner\"/></Types>"),
            .carryPart(partPath: "word/_rels/document.xml.rels", xml: "<Relationships xmlns=\"\(relNS)\" xmlns:x=\"urn:metadata\"><Relationship Id=\"ownedStyleID\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"old-styles.xml\" x:keep=\"rel-owner\"/></Relationships>")
        ])
        try doc.applyFormattingProfile(DocumentFormattingProfile.importOfficial(from: template()), context: .existingDocument)
        let result = try parts(doc, authoring: true)
        let types = try ProfileXML.parse(result["[Content_Types].xml"]!)
        let registration = try XCTUnwrap(types.children.first { $0.attributeValue(prefix: nil, localName: "PartName") == "/word/styles.xml" })
        XCTAssertEqual(registration.attributeValue(prefix: "x", localName: "keep"), "type-owner")
        XCTAssertEqual(registration.attributeValue(prefix: nil, localName: "ContentType"), "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")
        let rels = try ProfileXML.parse(result["word/_rels/document.xml.rels"]!)
        let relationship = try XCTUnwrap(rels.children.first { $0.attributeValue(prefix: nil, localName: "Id") == "ownedStyleID" })
        XCTAssertEqual(relationship.attributeValue(prefix: "x", localName: "keep"), "rel-owner")
        XCTAssertEqual(relationship.attributeValue(prefix: nil, localName: "Target"), "styles.xml")
    }

    func testOfficialWithoutThemePreservesCarriedTargetTheme() throws {
        let imported = try DocumentFormattingProfile.importOfficial(from: template())
        var json = try XCTUnwrap(JSONSerialization.jsonObject(with: JSONEncoder().encode(imported)) as? [String: Any])
        json.removeValue(forKey: "themeXML")
        json["stylesXML"] = imported.stylesXML?.replacingOccurrences(of: " w:asciiTheme=\"minorHAnsi\"", with: "").replacingOccurrences(of: " w:eastAsiaTheme=\"minorEastAsia\"", with: "")
        let profile = try JSONDecoder().decode(DocumentFormattingProfile.self, from: JSONSerialization.data(withJSONObject: json))
        var doc = WordDocument.emptyAuthoringDocument()
        let theme = "<a:theme xmlns:a=\"\(a)\" name=\"Target owned\"/>"
        try doc.apply(operations: [.carryPart(partPath: "word/theme/theme1.xml", xml: theme)])
        try doc.applyFormattingProfile(profile, context: .existingDocument)
        for output in [try parts(doc), try parts(doc, authoring: true)] {
            XCTAssertEqual(output["word/theme/theme1.xml"], theme)
            XCTAssertTrue(output["word/_rels/document.xml.rels"]!.contains("theme/theme1.xml"))
        }
    }

    func testUnprofiledStyleEditDoesNotRewriteRegisteredMetadata() throws {
        for target in ["styles.xml", "/word/styles.xml"] {
        let url = try directory().appendingPathComponent("plain.docx")
        try DocxWriter.writeData(WordDocument()).write(to: url)
        var original = try RawPartChannel.readAllParts(from: url)
        original["word/_rels/document.xml.rels"] = Data(String(decoding: original["word/_rels/document.xml.rels"]!, as: UTF8.self).replacingOccurrences(of: "Target=\"styles.xml\"", with: "Target=\"\(target)\"").utf8)
        try writePackage(original, to: url)
        var doc = try DocxReader.read(from: url)
        defer { doc.close() }
        try doc.updateStyle(id: "Normal", with: StyleUpdate(name: "Ordinary edit"))
        try DocxWriter.write(doc, to: url)
        let saved = try RawPartChannel.readAllParts(from: url)
        for path in ["[Content_Types].xml", "word/_rels/document.xml.rels"] { XCTAssertEqual(saved[path], original[path], path) }
        }
    }

    func writePackage(_ parts: [String: Data], to url: URL) throws {
        let root = try directory()
        for (path, bytes) in parts {
            let target = root.appendingPathComponent(path)
            try FileManager.default.createDirectory(at: target.deletingLastPathComponent(), withIntermediateDirectories: true)
            try bytes.write(to: target)
        }
        try ZipHelper.zipToData(root).write(to: url)
    }

    func testOfficialPreservesCompletePackageRelationshipsAcrossWritersAndReplay() throws {
        let root = try directory(), sourceURL = root.appendingPathComponent("source.docx")
        try DocxWriter.writeData(WordDocument()).write(to: sourceURL)
        var source = try RawPartChannel.readAllParts(from: sourceURL)
        let r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        let registrations = [("image", "rIdImage", "media/pixel.png"), ("header", "rIdHeader", "header1.xml"), ("footer", "rIdFooter", "footer1.xml"), ("hyperlink", "rIdLink", "https://example.com/target")]
        source["word/document.xml"] = Data("""
        <w:document xmlns:w="\(w)" xmlns:r="\(r)" xmlns:a="\(a)" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body><w:p><w:hyperlink r:id="rIdLink"><w:r><w:t>保留超連結</w:t></w:r></w:hyperlink><w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/><wp:docPr id="1" name="Pixel"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="pixel.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9525" cy="9525"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr></w:body></w:document>
        """.utf8)
        source["word/media/pixel.png"] = Data(base64Encoded: "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=")!
        source["word/header1.xml"] = Data("<w:hdr xmlns:w=\"\(w)\"><w:p><w:r><w:t>頁首</w:t></w:r></w:p></w:hdr>".utf8)
        source["word/footer1.xml"] = Data("<w:ftr xmlns:w=\"\(w)\"><w:p><w:r><w:t>頁尾</w:t></w:r></w:p></w:ftr>".utf8)
        var rels = String(decoding: source["word/_rels/document.xml.rels"]!, as: UTF8.self)
        for (kind, id, target) in registrations {
            rels = rels.replacingOccurrences(of: "</Relationships>", with: "<Relationship Id=\"\(id)\" Type=\"\(r)/\(kind)\" Target=\"\(target)\"\(kind == "hyperlink" ? " TargetMode=\"External\"" : "")/></Relationships>")
        }
        source["word/_rels/document.xml.rels"] = Data(rels.utf8)
        var types = String(decoding: source["[Content_Types].xml"]!, as: UTF8.self)
        types = types.replacingOccurrences(of: "</Types>", with: "<Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/><Override PartName=\"/word/footer1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml\"/></Types>")
        source["[Content_Types].xml"] = Data(types.utf8)
        try writePackage(source, to: sourceURL)
        let profile = try DocumentFormattingProfile.importOfficial(from: template())
        var log = OperationLog()
        // Binary data uses the actual binary operation, never lossy UTF-8.
        for path in source.keys.sorted() {
            if path.hasSuffix(".png") { log.append(.carryBinaryPart(partPath: path, base64: source[path]!.base64EncodedString()), source: .swift) }
            else { log.append(.carryPart(partPath: path, xml: String(decoding: source[path]!, as: UTF8.self)), source: .swift) }
        }
        let script = root.appendingPathComponent("replay.mdocx.swift")
        try ScriptExporter.exportSwift(log: log).write(to: script, atomically: true, encoding: .utf8)
        for mode in ["ordinary", "authoring", "script"] {
            let output = root.appendingPathComponent("\(mode).docx")
            if mode == "script" { _ = try scriptPipelineExecute(scriptPath: script.path, outputPath: output.path, formattingProfile: profile) }
            else {
                var doc = mode == "ordinary" ? try DocxReader.read(from: sourceURL) : WordDocument.emptyAuthoringDocument()
                defer { doc.close() }
                if mode == "authoring" { try doc.apply(log: log) }
                try doc.applyFormattingProfile(profile, context: .existingDocument)
                if mode == "ordinary" { try DocxWriter.write(doc, to: output) }
                else { try doc.writeAuthoringPackage(to: output) }
            }
            func check(_ bytes: [String: Data]) throws {
                let formattingParts: Set<String> = ["word/document.xml", "word/styles.xml", "word/fontTable.xml", "word/theme/theme1.xml", "[Content_Types].xml", "word/_rels/document.xml.rels"]
                for path in source.keys where !formattingParts.contains(path) { XCTAssertEqual(bytes[path], source[path], "\(mode): \(path)") }
                let rels = try ProfileXML.parse(bytes["word/_rels/document.xml.rels"]!)
                let types = try ProfileXML.parse(bytes["[Content_Types].xml"]!)
                let body = try ProfileXML.parse(bytes["word/document.xml"]!)
                for (kind, id, target) in registrations {
                    let matches = rels.children.filter { $0.attributeValue(prefix: nil, localName: "Id") == id }
                    XCTAssertEqual(matches.count, 1, "\(mode): \(id)")
                    XCTAssertEqual(matches.first?.attributeValue(prefix: nil, localName: "Target"), target)
                    XCTAssertEqual(matches.first?.attributeValue(prefix: nil, localName: "Type"), "\(r)/\(kind)")
                    XCTAssertTrue(ProfileXML.walk(body).contains { $0.attributes.contains { $0.prefix == "r" && $0.value == id } })
                    if kind == "hyperlink" { XCTAssertEqual(matches.first?.attributeValue(prefix: nil, localName: "TargetMode"), "External") }
                    else { XCTAssertEqual(bytes["word/" + target], source["word/" + target], target) }
                }
                for part in ["header1.xml", "footer1.xml", "styles.xml", "fontTable.xml", "theme/theme1.xml"] {
                    XCTAssertEqual(types.children.filter { $0.attributeValue(prefix: nil, localName: "PartName") == "/word/" + part }.count, 1, part)
                }
                XCTAssertEqual(types.children.filter { $0.attributeValue(prefix: nil, localName: "Extension") == "png" }.count, 1)
            }
            try check(RawPartChannel.readAllParts(from: output))
            var reopened = try DocxReader.read(from: output)
            defer { reopened.close() }
            try reopened.updateStyle(id: "Normal", with: StyleUpdate(name: "再儲存"))
            let again = root.appendingPathComponent("\(mode)-again.docx")
            try DocxWriter.write(reopened, to: again)
            try check(RawPartChannel.readAllParts(from: again))
        }
    }

    func testProfileBeforeStagingOmitsGeneratorFontsAndPreservesDirectFonts() throws {
        var generated = WordDocument()
        generated.body.children.append(.paragraph(Paragraph(runs: [Run(text: "code", properties: RunProperties(fontName: "Menlo"))])))
        try generated.applyFormattingProfile(.inherit, context: .newDocument)
        let url = try directory().appendingPathComponent("staging.docx")
        try DocxWriter.writeData(generated).write(to: url)
        var staged = try DocxReader.read(from: url)
        defer { staged.close() }
        try staged.applyFormattingProfile(.inherit, context: .newDocument)
        let output = try parts(staged)
        XCTAssertFalse(output["word/styles.xml"]!.contains("Calibri"))
        XCTAssertTrue(output["word/document.xml"]!.contains("Menlo"))
    }

    func testReadbackNewContextDoesNotGuessFontOrigin() throws {
        let url = try directory().appendingPathComponent("staging.docx")
        try DocxWriter.writeData(WordDocument()).write(to: url)
        var staged = try DocxReader.read(from: url)
        defer { staged.close() }
        try staged.applyFormattingProfile(.inherit, context: .newDocument)
        for output in [try parts(staged), try parts(staged, authoring: true)] {
            XCTAssertTrue(output["word/styles.xml"]!.contains("Calibri"))
            let defaults = try XCTUnwrap(ProfileXML.child(ProfileXML.parse(output["word/styles.xml"]!), "docDefaults"))
            let fonts = try XCTUnwrap(ProfileXML.walk(defaults).first { $0.localName == "rFonts" })
            XCTAssertEqual(ProfileXML.value(fonts, "ascii"), "Calibri")
            XCTAssertEqual(ProfileXML.value(fonts, "cs"), "Times New Roman")
        }
    }

    func testOfficialOverridesEastAsiaOnRetainedStylesAndPreservesTargetNumbering() throws {
        var doc = WordDocument()
        var run = RunProperties()
        run.rFonts = RFontsProperties(ascii: "Caller Western", eastAsia: "Conflicting Font")
        doc.styles.append(Style(id: "Referenced", name: "Referenced", type: .paragraph, basedOn: "Normal", runProperties: run))
        let number = try doc.createNumberingDefinition(levels: [Level(ilvl: 0, numFmt: .decimal, lvlText: "%1.", indent: 720)])
        let originalNumbering = doc.numbering
        try doc.applyFormattingProfile(DocumentFormattingProfile.importOfficial(from: template()), context: .existingDocument)
        XCTAssertEqual(doc.numbering, originalNumbering)
        XCTAssertGreaterThan(number, 0)
        let output = try parts(doc)
        XCTAssertFalse(output["word/styles.xml"]!.contains("Conflicting Font"))
        XCTAssertTrue(output["word/styles.xml"]!.contains("Caller Western"))
    }

    func testSnapshotRejectsDuplicateAttributesAndInvalidGeometryBeforeMutation() throws {
        let profile = try DocumentFormattingProfile.importOfficial(from: template())
        let encoded = try JSONEncoder().encode(profile)
        for value in [profile.sectionXML!.replacingOccurrences(of: "w:w=\"11906\"", with: "w:w=\"11906\" w:w=\"1\""),
                      profile.sectionXML!.replacingOccurrences(of: "11906", with: "-5")] {
            var json = try XCTUnwrap(JSONSerialization.jsonObject(with: encoded) as? [String: Any])
            json["sectionXML"] = value
            XCTAssertThrowsError(try JSONDecoder().decode(DocumentFormattingProfile.self, from: JSONSerialization.data(withJSONObject: json)))
        }
    }

    func testStyleRenamePreservesImportedTableFormattingAndLigatures() throws {
        let source = try RawPartChannel.readAllParts(from: template())
        let styles = String(decoding: source["word/styles.xml"]!, as: UTF8.self)
            .replacingOccurrences(of: "</x:styles>", with: "<x:style x:type=\"table\" x:styleId=\"TableBase\"><x:name x:val=\"Table base\"/><x:tblPr><x:tblInd x:w=\"720\" x:type=\"dxa\"/><x:tblCellMar><x:top x:w=\"72\" x:type=\"dxa\"/></x:tblCellMar></x:tblPr></x:style></x:styles>")
            .replacingOccurrences(of: "<x:sz x:val=\"24\"/>", with: "<x:sz x:val=\"24\"/><v:ligatures xmlns:v=\"http://schemas.microsoft.com/office/word/2010/wordml\" v:val=\"standard\"/>")
        var doc = WordDocument()
        try doc.applyFormattingProfile(DocumentFormattingProfile.importOfficial(from: template(styles: styles)), context: .newDocument)
        try doc.updateStyle(id: "TableBase", with: StyleUpdate(name: "Renamed table"))
        for output in [try parts(doc), try parts(doc, authoring: true)] {
            let xml = output["word/styles.xml"]!
            XCTAssertTrue(xml.contains("Renamed table"))
            XCTAssertTrue(xml.contains("w:tblInd"))
            XCTAssertTrue(xml.contains("w:w=\"720\""))
            XCTAssertTrue(xml.contains("w14:ligatures"))
            XCTAssertTrue(xml.contains("w:w=\"72\""))
        }
    }

    func testImportRejectsNumberingMissingRequiredDataAndMalformedXML() throws {
        for extras in [
            ["word/numbering.xml": "<w:numbering xmlns:w=\"\(w)\"><w:abstractNum w:abstractNumId=\"0\"/></w:numbering>"],
            ["word/styles.xml": "<broken>"],
            ["word/styles.xml": "<w:styles xmlns:w=\"\(w)\"/>"],
            ["word/document.xml": "<w:document xmlns:w=\"\(w)\"><w:body/></w:document>"]
        ] { XCTAssertThrowsError(try DocumentFormattingProfile.importOfficial(from: template(extras: extras))) }
    }

    func testDecodeValidatesVersionAndUnsafePayload() throws {
        let profile = try DocumentFormattingProfile.importOfficial(from: template())
        let encoded = try JSONEncoder().encode(profile)
        var json = try XCTUnwrap(JSONSerialization.jsonObject(with: encoded) as? [String: Any])
        json["schemaVersion"] = 999
        XCTAssertThrowsError(try JSONDecoder().decode(DocumentFormattingProfile.self, from: JSONSerialization.data(withJSONObject: json)))
        json["schemaVersion"] = 1
        json["stylesXML"] = "<w:styles xmlns:w=\"\(w)\"><w:p>PRIVATE</w:p></w:styles>"
        XCTAssertThrowsError(try JSONDecoder().decode(DocumentFormattingProfile.self, from: JSONSerialization.data(withJSONObject: json)))
    }

    func testProfileRunsBeforeVerificationAndLeavesExistingOutputUntouched() throws {
        let dir = try directory()
        let script = dir.appendingPathComponent("source.mdocx.swift")
        let output = dir.appendingPathComponent("output.docx")
        let log = OperationLog()
        try ScriptExporter.exportSwift(log: log).write(to: script, atomically: true, encoding: .utf8)
        _ = try scriptPipelineExecute(scriptPath: script.path, outputPath: output.path)
        let original = try Data(contentsOf: output)
        let profile = try DocumentFormattingProfile.importOfficial(from: template())
        let result = try scriptPipelineExecute(scriptPath: script.path, outputPath: output.path, verifyAgainst: output.path, overwrite: true, formattingProfile: profile)
        XCTAssertEqual(result.verified, false)
        XCTAssertNil(result.written)
        XCTAssertEqual(try Data(contentsOf: output), original)
    }

    func testReopenThenTypedMutationRetainsImportedDefaultsAndTheme() throws {
        let profile = try DocumentFormattingProfile.importOfficial(from: template())
        var doc = WordDocument()
        try doc.applyFormattingProfile(profile, context: .newDocument)
        let url = try directory().appendingPathComponent("reopen.docx")
        try DocxWriter.writeData(doc).write(to: url)
        var reopened = try DocxReader.read(from: url)
        defer { reopened.close() }
        try reopened.updateStyle(id: "Normal", with: StyleUpdate(name: "Changed after reopening"))
        let output = try parts(reopened)
        XCTAssertTrue(output["word/styles.xml"]!.contains("w:val=\"24\""))
        XCTAssertTrue(output["word/styles.xml"]!.contains("Changed after reopening"))
        XCTAssertTrue(output["word/styles.xml"]!.contains("w:line=\"278\""))
        XCTAssertNotNil(output["word/theme/theme1.xml"])
    }

    func testFinalSectionPreservesTargetReferencesOtherSectionsAndValueCopy() throws {
        let profile = try DocumentFormattingProfile.importOfficial(from: template())
        let xml = "<w:document xmlns:w=\"\(w)\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body><w:p><w:pPr><w:sectPr><w:pgSz w:w=\"9000\" w:h=\"10000\"/></w:sectPr></w:pPr><w:r><w:t>Keep me</w:t></w:r></w:p><w:sectPr><w:headerReference w:type=\"default\" r:id=\"rId22\"/><w:footerReference w:type=\"first\" r:id=\"rId23\"/><w:pgSz w:w=\"10000\" w:h=\"11000\"/></w:sectPr></w:body></w:document>"
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: [.carryPart(partPath: "word/document.xml", xml: xml.replacingOccurrences(of: "</w:sectPr></w:body>", with: "<w:titlePg/></w:sectPr></w:body>"))])
        let old = doc
        let before = try XmlTreeWriter.serialize(old.xmlTrees["word/document.xml"]!)
        try doc.applyFormattingProfile(profile, context: .existingDocument)
        let output = try parts(doc, authoring: true)["word/document.xml"]!
        XCTAssertTrue(output.contains("11906"))
        XCTAssertTrue(output.contains("9000"))
        XCTAssertTrue(output.contains("rId22"))
        XCTAssertTrue(output.contains("rId23"))
        XCTAssertTrue(output.contains("Keep me"))
        let tree = try XmlTreeReader.parse(Data(output.utf8))
        let final = try XCTUnwrap(ProfileXML.child(tree.root, "body")?.children.last { $0.localName == "sectPr" })
        XCTAssertLessThan(try XCTUnwrap(final.children.firstIndex { $0.localName == "pgSz" }), try XCTUnwrap(final.children.firstIndex { $0.localName == "titlePg" }))
        XCTAssertEqual(try XmlTreeWriter.serialize(old.xmlTrees["word/document.xml"]!), before)
    }

    func testAliasedNonzeroNumberingAndDanglingStyleChainsAreRejected() throws {
        let base = "<x:styles xmlns:x=\"\(w)\"><x:docDefaults><x:rPrDefault><x:rPr><x:sz x:val=\"24\"/></x:rPr></x:rPrDefault><x:pPrDefault><x:pPr/></x:pPrDefault></x:docDefaults><x:style x:type=\"paragraph\" x:default=\"1\" x:styleId=\"a\">%@</x:style></x:styles>"
        for inserted in ["<x:pPr><x:numPr><x:numId x:val=\"9\"/></x:numPr></x:pPr>", "<x:basedOn x:val=\"missing\"/>", "<x:basedOn x:val=\"a\"/>"] {
            XCTAssertThrowsError(try DocumentFormattingProfile.importOfficial(from: template(styles: String(format: base, inserted))))
        }
    }

    func testDecodeRejectsUnsafeNamespacesAndInvalidNestedFormatting() throws {
        let profile = try DocumentFormattingProfile.importOfficial(from: template())
        let encoded = try JSONEncoder().encode(profile)
        for payload in [
            "<x:headerReference xmlns:x=\"\(w)\"/>",
            "<x:rFonts xmlns:x=\"urn:hostile\" x:ascii=\"Hidden\"/>",
            "<w:font w:name=\"Not a style child\"/>",
            "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"other\"/>"
        ] {
            var json = try XCTUnwrap(JSONSerialization.jsonObject(with: encoded) as? [String: Any])
            json["stylesXML"] = profile.stylesXML!.replacingOccurrences(of: "</w:styles>", with: payload + "</w:styles>")
            XCTAssertThrowsError(try JSONDecoder().decode(DocumentFormattingProfile.self, from: JSONSerialization.data(withJSONObject: json)), payload)
        }
    }

    func testOptionalReadOnlyTemplateCompatibility() throws {
        guard let path = ProcessInfo.processInfo.environment["OOXML_PROFILE_TEMPLATE"] else { throw XCTSkip("optional local template not selected") }
        let url = URL(fileURLWithPath: path)
        let before = try Data(contentsOf: url)
        let profile = try DocumentFormattingProfile.importOfficial(from: url)
        var doc = WordDocument()
        try doc.applyFormattingProfile(profile, context: .newDocument)
        XCTAssertEqual(doc.sectionProperties.pageSize, PageSize(width: 11906, height: 16838))
        XCTAssertEqual(doc.sectionProperties.pageMargins.left, 1800)
        XCTAssertTrue(try parts(doc)["word/styles.xml"]!.contains("w:eastAsia=\"DFKai-SB\""))
        XCTAssertEqual(try Data(contentsOf: url), before)
    }
}
