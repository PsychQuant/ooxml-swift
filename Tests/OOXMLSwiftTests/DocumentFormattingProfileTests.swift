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

    func directory() throws -> URL {
        let url = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString)
        try FileManager.default.createDirectory(at: url, withIntermediateDirectories: true)
        addTeardownBlock { try? FileManager.default.removeItem(at: url) }
        return url
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
        XCTAssertTrue(output["word/styles.xml"]!.contains("標楷體"))
        XCTAssertFalse(output["word/styles.xml"]!.contains("eastAsiaTheme"))
        XCTAssertTrue(output["word/theme/theme1.xml"]!.contains("Aptos"))
        XCTAssertFalse(output["word/theme/theme1.xml"]!.contains("新細明體"))
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

    func testNewDocumentContextAfterStagingStillOmitsGeneratorFontsAndPreservesDirectFonts() throws {
        var generated = WordDocument()
        generated.body.children.append(.paragraph(Paragraph(runs: [Run(text: "code", properties: RunProperties(fontName: "Menlo"))])))
        let url = try directory().appendingPathComponent("staging.docx")
        try DocxWriter.writeData(generated).write(to: url)
        var staged = try DocxReader.read(from: url)
        defer { staged.close() }
        try staged.applyFormattingProfile(.inherit, context: .newDocument)
        let output = try parts(staged)
        XCTAssertFalse(output["word/styles.xml"]!.contains("Calibri"))
        XCTAssertTrue(output["word/document.xml"]!.contains("Menlo"))
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
        XCTAssertTrue(try parts(doc)["word/styles.xml"]!.contains("標楷體"))
        XCTAssertEqual(try Data(contentsOf: url), before)
    }
}
