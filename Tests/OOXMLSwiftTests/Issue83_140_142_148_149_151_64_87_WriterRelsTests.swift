import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Regression coverage for the `DocxWriter` relationships / packaging cluster:
/// PsychQuant/ooxml-swift #83, #140, #142, #148, #149, #216, #87, #151.
///
/// #83 / #140 / #149 share one root cause (documented in the #83/#140 issue
/// bodies and the #140 first comment, which explicitly says the cluster PR
/// #141 only turned the failure from a crash into a *named refusal* — "根治
/// 在這裡", the real fix belongs in THIS issue): `DocxWriter` hardcoded
/// `rId1`/`rId2`/`rId3`(/`rId4` for numbering) for styles/settings/fontTable/
/// numbering instead of discovering them from the package. The #140 second
/// comment explicitly asks that its exact repro (insert an image, THEN create
/// a numbering definition, via the public API, in that order) be added as a
/// regression test — `testCreatingNumberingAfterAnImageAlreadyHoldsItsSlotIdSavesWithoutCollision`
/// below is that literal repro.
final class Issue83_140_142_148_149_151_64_87_WriterRelsTests: XCTestCase {

    // MARK: - Shared fixture helpers (synthetic packages only — rule 6: no
    // third-party document content is ever committed; every byte below is
    // authored by this test file).

    private let pkgNS = "http://schemas.openxmlformats.org/package/2006/relationships"
    private let rNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    private let wNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    private let aNS = "http://schemas.openxmlformats.org/drawingml/2006/main"
    private let stylesType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    private let settingsType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    private let fontTableType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
    private let numberingType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
    private let imageType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    // Not a type DocxWriter manages — stands in for a real unmanaged part
    // (webSettings) the way #149's repro table uses it.
    private let webSettingsType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"

    private func zip(_ parts: [String: String]) throws -> Data {
        let archive = try Archive(accessMode: .create)
        for name in parts.keys.sorted() {
            let data = Data(parts[name]!.utf8)
            try archive.addEntry(with: name, type: .file, uncompressedSize: Int64(data.count),
                                 compressionMethod: .deflate) { position, size in
                let start = data.startIndex.advanced(by: Int(position))
                return data.subdata(in: start..<start.advanced(by: size))
            }
        }
        return try XCTUnwrap(archive.data)
    }

    private func rels(_ body: String) -> String {
        #"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="\#(pkgNS)">"# + body + "</Relationships>"
    }

    private func minimalDocumentXML(embedding imageId: String? = nil) -> String {
        let drawing = imageId.map { #"<w:p><w:r><w:drawing><a:blip r:embed="\#($0)" xmlns:a="\#(aNS)"/></w:drawing></w:r></w:p>"# } ?? "<w:p><w:r><w:t>hi</w:t></w:r></w:p>"
        return #"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="\#(wNS)" xmlns:r="\#(rNS)"><w:body>"#
            + drawing + "<w:sectPr/></w:body></w:document>"
    }

    private let minimalStylesXML = #"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#
    private let minimalSettingsXML = #"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#
    private let minimalFontTableXML = #"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#
    private let minimalWebSettingsXML = #"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#
    private let minimalNumberingXML = #"""
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>
    <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
    </w:numbering>
    """#
    private let onePxPNG = Data([
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4, 0x89,
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54,
        0x78, 0x9C, 0x62, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01,
        0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00,
        0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82
    ]) // used as raw bytes, not a base64 string

    private let contentTypesXML = #"""
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Default Extension="png" ContentType="image/png"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
    <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
    <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
    <Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>
    <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
    </Types>
    """#
    private let packageRelsXML = #"""
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>
    """#
    private let coreXML = #"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"/>"#
    private let appXML = #"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"/>"#

    /// A synthetic .docx built entirely by this test file: styles/settings/
    /// fontTable/numbering/webSettings/image relationship ids are all
    /// parameters, so a test can place them at ids other than DocxWriter's
    /// historical rId1..rId4 hardcode.
    private func syntheticPackage(
        stylesId: String, settingsId: String, fontTableId: String,
        webSettingsId: String? = nil, numberingId: String? = nil,
        imageId: String? = nil, extraRelsXML: String = ""
    ) throws -> Data {
        var docRels = """
        <Relationship Id="\(stylesId)" Type="\(stylesType)" Target="styles.xml"/>
        <Relationship Id="\(settingsId)" Type="\(settingsType)" Target="settings.xml"/>
        <Relationship Id="\(fontTableId)" Type="\(fontTableType)" Target="fontTable.xml"/>
        """
        if let webSettingsId {
            docRels += #"<Relationship Id="\#(webSettingsId)" Type="\#(webSettingsType)" Target="webSettings.xml"/>"#
        }
        if let numberingId {
            docRels += #"<Relationship Id="\#(numberingId)" Type="\#(numberingType)" Target="numbering.xml"/>"#
        }
        if let imageId {
            docRels += #"<Relationship Id="\#(imageId)" Type="\#(imageType)" Target="media/image1.png"/>"#
        }
        docRels += extraRelsXML

        var parts: [String: String] = [
            "[Content_Types].xml": contentTypesXML,
            "_rels/.rels": packageRelsXML,
            "word/_rels/document.xml.rels": rels(docRels),
            "word/document.xml": minimalDocumentXML(embedding: imageId),
            "word/styles.xml": minimalStylesXML,
            "word/settings.xml": minimalSettingsXML,
            "word/fontTable.xml": minimalFontTableXML,
            "word/webSettings.xml": minimalWebSettingsXML,
            "docProps/core.xml": coreXML,
            "docProps/app.xml": appXML,
        ]
        if numberingId != nil { parts["word/numbering.xml"] = minimalNumberingXML }

        let archive = try Archive(accessMode: .create)
        for name in parts.keys.sorted() {
            let data = Data(parts[name]!.utf8)
            try archive.addEntry(with: name, type: .file, uncompressedSize: Int64(data.count),
                                 compressionMethod: .deflate) { position, size in
                let start = data.startIndex.advanced(by: Int(position))
                return data.subdata(in: start..<start.advanced(by: size))
            }
        }
        if imageId != nil {
            try archive.addEntry(with: "word/media/image1.png", type: .file, uncompressedSize: Int64(onePxPNG.count),
                                 compressionMethod: .deflate) { position, size in
                let start = self.onePxPNG.startIndex.advanced(by: Int(position))
                return self.onePxPNG.subdata(in: start..<start.advanced(by: size))
            }
        }
        return try XCTUnwrap(archive.data)
    }

    /// Read a synthetic package built above into a `WordDocument` (overlay
    /// mode: `archiveTempDir` is set).
    private func readSyntheticDocument(_ data: Data) throws -> WordDocument {
        let url = FileManager.default.temporaryDirectory.appendingPathComponent("i83149-\(UUID().uuidString).docx")
        try data.write(to: url)
        defer { try? FileManager.default.removeItem(at: url) }
        return try DocxReader.read(from: url)
    }

    /// Save `document` (overlay mode requires a real file target — `writeData`
    /// always uses scratch mode per its own doc comment) and return the
    /// produced package's `word/_rels/document.xml.rels` scan.
    private func writeAndScanDocumentRels(_ document: WordDocument) throws -> (data: Data, relsXML: String, scan: (imageIds: [String], duplicateIds: [String], parsed: Bool, allIds: [String], structure: [String])) {
        let dest = FileManager.default.temporaryDirectory.appendingPathComponent("i83149-out-\(UUID().uuidString).docx")
        try DocxWriter.write(document, to: dest)
        defer { try? FileManager.default.removeItem(at: dest) }
        let data = try Data(contentsOf: dest)
        let dir = try ZipHelper.unzip(dest)
        defer { ZipHelper.cleanup(dir) }
        let relsData = try Data(contentsOf: dir.appendingPathComponent("word/_rels/document.xml.rels"))
        let relsXML = String(data: relsData, encoding: .utf8) ?? ""
        return (data, relsXML, PackageInspector.scanRels(relsData))
    }

    private func markAllDirtyForRelsRewrite(_ document: inout WordDocument) {
        // A minimal edit that is guaranteed to force `needsDocumentRels` —
        // used by tests whose point is the SLOT resolution, not the dirty
        // gating itself (#142's tests below exercise the gating directly).
        // `markTypedDirty` is `internal`, reachable here via `@testable import`.
        document.markTypedDirty("word/_rels/document.xml.rels")
    }

    // MARK: - #83 — managed rel ids discovered from the source, not hardcoded

    /// #83 body: "Discover the managed-rel rIds from the source
    /// `document.xml.rels` (match by `Type`, reuse the source's rId +
    /// target) instead of hardcoding rId1–3/rId4." This is the literal shape
    /// from the issue: styles/settings/fontTable/numbering all live at ids
    /// other than rId1..rId4 (as a non-Word writer would number them).
    func testStylesSettingsFontTableNumberingSurviveAtNonStandardOriginalIds() throws {
        let data = try syntheticPackage(
            stylesId: "rId9", settingsId: "rId10", fontTableId: "rId11", numberingId: "rId12")
        var document = try readSyntheticDocument(data)
        defer { document.close() }
        XCTAssertFalse(document.numbering.abstractNums.isEmpty, "fixture must actually carry a numbering definition")

        markAllDirtyForRelsRewrite(&document)
        let result = try writeAndScanDocumentRels(document)

        XCTAssertEqual(result.scan.duplicateIds, [], "no id should be declared twice")
        func typeCount(_ type: String) -> Int {
            result.relsXML.components(separatedBy: "Type=\"\(type)\"").count - 1
        }
        XCTAssertEqual(typeCount(stylesType), 1, "styles relationship must survive exactly once:\n\(result.relsXML)")
        XCTAssertEqual(typeCount(settingsType), 1, "settings relationship must survive exactly once:\n\(result.relsXML)")
        XCTAssertEqual(typeCount(fontTableType), 1, "fontTable relationship must survive exactly once:\n\(result.relsXML)")
        XCTAssertEqual(typeCount(numberingType), 1, "numbering relationship must survive exactly once (#83's reported drop):\n\(result.relsXML)")
    }

    /// #83's suggested fix explicitly: "reuse the source's rId + target".
    /// A managed type that is genuinely present in the original package
    /// keeps its ORIGINAL id — round-trip stays maximally byte-stable, and
    /// nothing in `word/document.xml` could dangle since none of these four
    /// types are ever referenced by id from the body.
    func testManagedTypesReuseTheirOriginalIdVerbatim() throws {
        let data = try syntheticPackage(
            stylesId: "rId9", settingsId: "rId10", fontTableId: "rId11", numberingId: "rId12")
        var document = try readSyntheticDocument(data)
        defer { document.close() }

        markAllDirtyForRelsRewrite(&document)
        let result = try writeAndScanDocumentRels(document)

        XCTAssertTrue(result.relsXML.contains(#"Id="rId9" Type="\#(stylesType)" Target="styles.xml""#),
                      "styles must keep its original id rId9 and target:\n\(result.relsXML)")
        XCTAssertTrue(result.relsXML.contains(#"Id="rId12" Type="\#(numberingType)" Target="numbering.xml""#),
                      "numbering must keep its original id rId12 and target:\n\(result.relsXML)")
    }

    // MARK: - #140 — fixed slots must not collide with a legitimate package

    /// #140 body's exact scenario: "一份把 rId1 用在圖片的合法 .docx 讀進來再存"
    /// — a real image relationship legitimately occupies rId1 (the writer's
    /// historical styles slot). Per the first comment, PR #141 only turned
    /// this into a NAMED refusal (`WordError.invalidDocx`); the comment says
    /// the real fix — the package must actually SAVE — belongs here.
    func testImageLegitimatelyAtRId1Saves() throws {
        let data = try syntheticPackage(
            stylesId: "rId9", settingsId: "rId10", fontTableId: "rId11",
            imageId: "rId1")
        var document = try readSyntheticDocument(data)
        defer { document.close() }
        XCTAssertEqual(document.images.map(\.id), ["rId1"])

        markAllDirtyForRelsRewrite(&document)
        let result = try writeAndScanDocumentRels(document)

        XCTAssertEqual(result.scan.duplicateIds, [], "must not refuse a legitimate package (#140 first comment: refusing is not the fix)")
        XCTAssertTrue(result.scan.imageIds.contains("rId1"), "the image relationship must still be declared at rId1:\n\(result.relsXML)")
        func typeCount(_ type: String) -> Int { result.relsXML.components(separatedBy: "Type=\"\(type)\"").count - 1 }
        XCTAssertEqual(typeCount(stylesType), 1)
    }

    /// #140 second comment — VERBATIM repro requested for the regression
    /// suite: "建立新文件，插入一張圖片（圖片拿到 rId4），再建立編號定義
    /// （numbering 在寫出時固定佔用 rId4），存檔 → 拒絕". Uses ONLY the
    /// public typed API, in exactly this order (the comment notes the
    /// REVERSE order already worked before this fix).
    func testCreatingNumberingAfterAnImageAlreadyHoldsItsSlotIdSavesWithoutCollision() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "body"))

        let imageId = try doc.insertImage(
            base64: onePxPNG.base64EncodedString(), fileName: "image1.png",
            widthPx: 1, heightPx: 1)
        XCTAssertTrue(document_hasNoNumberingYet(doc), "numbering must not exist yet — this is the order #140's repro depends on")

        _ = try doc.createNumberingDefinition(levels: [
            Level(ilvl: 0, numFmt: .decimal, lvlText: "%1.", indent: 720)
        ])

        // Pre-fix: this throws WordError.invalidDocx naming #140 ("would
        // declare 1 relationship id twice: rId4").
        let data = try DocxWriter.writeData(doc)

        let dir = try ZipHelper.unzip(data: data)
        defer { ZipHelper.cleanup(dir) }
        let relsData = try Data(contentsOf: dir.appendingPathComponent("word/_rels/document.xml.rels"))
        let scan = PackageInspector.scanRels(relsData)
        XCTAssertEqual(scan.duplicateIds, [], "no relationship id may be declared twice")
        XCTAssertTrue(scan.imageIds.contains(imageId), "the image must keep its allocated id \(imageId)")
        let relsXML = String(data: relsData, encoding: .utf8) ?? ""
        XCTAssertEqual(relsXML.components(separatedBy: "Type=\"\(numberingType)\"").count - 1, 1,
                       "numbering relationship must be declared exactly once, at an id other than the image's:\n\(relsXML)")
    }

    private func document_hasNoNumberingYet(_ doc: WordDocument) -> Bool {
        doc.numbering.abstractNums.isEmpty
    }

    // MARK: - #149 — fixed-slot collision with an UNMANAGED original relationship

    /// #149's repro table: fontTable in the original sits at a non-standard
    /// id (here `rId7`) while an UNMANAGED type (`webSettings`) legitimately
    /// occupies the writer's historical fontTable slot (`rId3`). Pre-fix,
    /// pass 1 preserves webSettings at rId3 first, and the (hardcoded rId3)
    /// typed fontTable descriptor is then silently skipped in pass 2 because
    /// rId3 is already `emittedIds` — fontTable disappears, `word/fontTable.xml`
    /// becomes an orphan part.
    func testFontTableSurvivesWhenAnUnmanagedRelAlreadyHoldsItsHistoricalSlot() throws {
        let data = try syntheticPackage(
            stylesId: "rId1", settingsId: "rId2", fontTableId: "rId7",
            webSettingsId: "rId3")
        var document = try readSyntheticDocument(data)
        defer { document.close() }

        markAllDirtyForRelsRewrite(&document)
        let result = try writeAndScanDocumentRels(document)

        XCTAssertEqual(result.scan.duplicateIds, [])
        func typeCount(_ type: String) -> Int { result.relsXML.components(separatedBy: "Type=\"\(type)\"").count - 1 }
        XCTAssertEqual(typeCount(fontTableType), 1, "fontTable relationship must not be dropped (#149):\n\(result.relsXML)")
        XCTAssertEqual(typeCount(webSettingsType), 1, "the unmanaged webSettings relationship must survive untouched:\n\(result.relsXML)")
    }

    // MARK: - #142 — decoded comparison, not raw text scanning

    /// #142 body, mechanism 1: a decoy comment containing what LOOKS like a
    /// declaration for the new image's id fools `hasNewTypedRelationships`'s
    /// raw `.contains("Id=\"...\"")` check into believing the relationship
    /// already exists, so the rels part is never rewritten and the new
    /// image's relationship is never declared — "存檔成功、圖片靜默消失"
    /// (macdoc#175's failure shape, per the issue). The new image is added
    /// directly to `document.images` (bypassing `insertImage`'s own
    /// unconditional `markTypedDirty`) so this test exercises exactly the
    /// function the issue names, not the dirty-tracking safety net around it.
    /// **Outcome note**: fixing `hasNewTypedRelationships` makes it correctly
    /// judge "yes, something new" for this shape (the bug this test targets).
    /// What happens next is decided by a SEPARATE, already-existing gate in
    /// `documentRelationshipsXML` that refuses to merge into ANY rels file
    /// containing a comment/CDATA/processing-instruction at all (a comment is
    /// structure a text-based merge cannot trust, regardless of what it says)
    /// — so the observable result here is a NAMED REFUSAL, not the image
    /// relationship being declared. That is still the fix: pre-fix this
    /// scenario was "存檔成功、圖片靜默消失" (exit 0, image silently gone,
    /// #142's own words); post-fix it is a loud, named error. This test pins
    /// exactly that outcome so it cannot regress back to silent success.
    func testDecoyCommentDoesNotSuppressANewlyAddedImageRelationship() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "body"))
        let seed = FileManager.default.temporaryDirectory.appendingPathComponent("i142-seed-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: seed)
        defer { try? FileManager.default.removeItem(at: seed) }

        let dir = try ZipHelper.unzip(seed)
        defer { ZipHelper.cleanup(dir) }
        let relsURL = dir.appendingPathComponent("word/_rels/document.xml.rels")
        var relsXML = try String(contentsOf: relsURL, encoding: .utf8)
        // The decoy: a commented-out declaration for the id the test is
        // about to give the new image.
        relsXML = relsXML.replacingOccurrences(of: "</Relationships>", with: #"<!-- Id="rId9" --></Relationships>"#)
        try relsXML.write(to: relsURL, atomically: true, encoding: .utf8)
        let damaged = FileManager.default.temporaryDirectory.appendingPathComponent("i142-damaged-\(UUID().uuidString).docx")
        try ZipHelper.zip(dir, to: damaged)
        defer { try? FileManager.default.removeItem(at: damaged) }

        var document = try DocxReader.read(from: damaged)
        defer { document.close() }
        XCTAssertTrue(document.modifiedPartsView.isEmpty, "no-op load must not mark anything dirty")

        // Directly append — NOT via `insertImage` — so nothing marks
        // "word/_rels/document.xml.rels" or "[Content_Types].xml" dirty.
        // `hasNewTypedRelationships` is then the ONLY gate deciding whether
        // the rels part gets rewritten.
        document.images.append(ImageReference(id: "rId9", fileName: "image1.png", contentType: "image/png", data: onePxPNG))

        let out = FileManager.default.temporaryDirectory.appendingPathComponent("i142-out-\(UUID().uuidString).docx")
        var thrown: Error?
        XCTAssertThrowsError(try DocxWriter.write(document, to: out)) { thrown = $0 }
        let message = (thrown as? WordError)?.errorDescription ?? String(describing: thrown)
        XCTAssertTrue(message.contains("comment"), "must name the comment as the reason it refuses to merge, got: \(message)")
        XCTAssertFalse(FileManager.default.fileExists(atPath: out.path),
                        "a refused save must not leave a package on disk with the image silently missing")
    }

    /// #142 body, mechanism 2 (entity-written id): `Id="rId&#211;"` decodes to
    /// `rId9`, so a raw substring search for `Id="rId9"` never matches it —
    /// the shape `hasNewTypedRelationships` itself would misjudge.
    ///
    /// **Finding, not a fix here**: a package that genuinely declares a
    /// relationship id with a character/entity reference ALREADY trips a
    /// separate, earlier-added gate (`documentRelationshipsXML`'s raw-vs-
    /// parsed id view comparison, `PsychQuant/ooxml-swift#142`'s own later
    /// hardening) the moment anything forces a rels rewrite — it refuses
    /// with a named error before `RelationshipsOverlay.merge` ever runs, so
    /// `hasNewTypedRelationships`'s answer for this specific shape can no
    /// longer matter in practice. This test pins that refusal instead of a
    /// silent-drop, so a future change to that gate cannot silently
    /// reintroduce the mechanism #142 named.
    func testEntityEncodedOriginalIdIsRefusedRatherThanSilentlyMismatched() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "body"))
        let seed = FileManager.default.temporaryDirectory.appendingPathComponent("i142b-seed-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: seed)
        defer { try? FileManager.default.removeItem(at: seed) }

        let dir = try ZipHelper.unzip(seed)
        defer { ZipHelper.cleanup(dir) }
        let relsURL = dir.appendingPathComponent("word/_rels/document.xml.rels")
        var relsXML = try String(contentsOf: relsURL, encoding: .utf8)
        // A real (not commented) relationship declared with an entity-escaped
        // id spelling: rId&#211; decodes to rId9.
        relsXML = relsXML.replacingOccurrences(
            of: "</Relationships>",
            with: #"<Relationship Id="rId&#211;" Type="\#(webSettingsType)" Target="webSettings.xml"/></Relationships>"#)
        try relsXML.write(to: relsURL, atomically: true, encoding: .utf8)
        try #"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:webSettings xmlns:w="\#(wNS)"/>"#
            .write(to: dir.appendingPathComponent("word/webSettings.xml"), atomically: true, encoding: .utf8)
        let damaged = FileManager.default.temporaryDirectory.appendingPathComponent("i142b-damaged-\(UUID().uuidString).docx")
        try ZipHelper.zip(dir, to: damaged)
        defer { try? FileManager.default.removeItem(at: damaged) }

        var document = try DocxReader.read(from: damaged)
        defer { document.close() }
        // Forces the rels rewrite so the entity-escaped id is actually
        // looked at (a no-op save would leave the untouched rels alone).
        _ = document.addHeader(text: "trigger merge")

        let out = FileManager.default.temporaryDirectory.appendingPathComponent("i142b-out-\(UUID().uuidString).docx")
        var thrown: Error?
        XCTAssertThrowsError(try DocxWriter.write(document, to: out)) { thrown = $0 }
        let message = (thrown as? WordError)?.errorDescription ?? String(describing: thrown)
        XCTAssertTrue(message.contains("#142"), "must name #142 as the refusal reason, got: \(message)")
    }

    // MARK: - #148 — rels merge must not double-escape character/entity references

    /// #148's own reproduction: a `Target` containing a character reference
    /// survives one merge unchanged, then is corrupted a little more on every
    /// further save (`&#47;&#214;` → the parser's `/>` reading gets re-escaped
    /// as literal `&amp;#47;&amp;#214;`, etc.). This test saves the SAME
    /// unmanaged relationship across three saves (each forced to go through
    /// the merge by touching an unrelated header) and asserts the Target text
    /// is byte-identical every time.
    func testUnmanagedRelationshipWithEntityReferencedTargetIsNotProgressivelyDoubleEscaped() throws {
        let originalTarget = "a&#47;&#214;b.xml"   // parser reads this as "a/>b.xml"
        let extra = #"<Relationship Id="rId50" Type="http://example.test/unmanaged" Target="\#(originalTarget)"/>"#
        let data = try syntheticPackage(
            stylesId: "rId1", settingsId: "rId2", fontTableId: "rId3",
            extraRelsXML: extra)
        var document = try readSyntheticDocument(data)
        defer { document.close() }

        for round in 1...3 {
            _ = document.addHeader(text: "h\(round)")   // forces the merge every round
            let dest = FileManager.default.temporaryDirectory.appendingPathComponent("i148-\(round)-\(UUID().uuidString).docx")
            try DocxWriter.write(document, to: dest)
            defer { try? FileManager.default.removeItem(at: dest) }
            let dir = try ZipHelper.unzip(dest)
            defer { ZipHelper.cleanup(dir) }
            let relsXML = try String(contentsOf: dir.appendingPathComponent("word/_rels/document.xml.rels"), encoding: .utf8)
            XCTAssertTrue(relsXML.contains(#"Target="\#(originalTarget)""#),
                          "round \(round): Target must stay exactly \(originalTarget), got:\n\(relsXML)")
            document = try DocxReader.read(from: dest)
        }
    }

    /// The `TargetMode` variant from the R10 Devil's-Advocate comment on
    /// #148: `TargetMode="Ex&#116;ernal"` decodes to `External` before the
    /// save and must not become the literal `Ex&#116;ernal` after it.
    func testUnmanagedHyperlinkTargetModeWithEntityReferenceIsNotReescaped() throws {
        let extra = #"<Relationship Id="rId50" Type="http://example.test/unmanaged-hyperlink" Target="https://example.com" TargetMode="Ex&#116;ernal"/>"#
        let data = try syntheticPackage(
            stylesId: "rId1", settingsId: "rId2", fontTableId: "rId3",
            extraRelsXML: extra)
        var document = try readSyntheticDocument(data)
        defer { document.close() }
        _ = document.addHeader(text: "trigger merge")

        let dest = FileManager.default.temporaryDirectory.appendingPathComponent("i148b-\(UUID().uuidString).docx")
        try DocxWriter.write(document, to: dest)
        defer { try? FileManager.default.removeItem(at: dest) }
        let dir = try ZipHelper.unzip(dest)
        defer { ZipHelper.cleanup(dir) }
        let relsXML = try String(contentsOf: dir.appendingPathComponent("word/_rels/document.xml.rels"), encoding: .utf8)
        XCTAssertTrue(relsXML.contains(#"TargetMode="Ex&#116;ernal""#),
                      "TargetMode spelling must be preserved verbatim, not re-escaped:\n\(relsXML)")
    }

    // MARK: - #151 — reproducible writer output

    /// #151 body: `props.created ?? Date()` / `props.modified ?? Date()` make
    /// two saves of the SAME unset-timestamp document differ. This does not
    /// wait on a clock tick — it calls `writeData` twice in immediate
    /// succession and requires `docProps/core.xml` to be byte-identical both
    /// times regardless of timing.
    func testFreshDocumentWithNoExplicitTimestampsProducesReproducibleCoreProperties() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "x"))
        XCTAssertNil(doc.properties.created)
        XCTAssertNil(doc.properties.modified)

        let data1 = try DocxWriter.writeData(doc)
        let data2 = try DocxWriter.writeData(doc)
        let dir1 = try ZipHelper.unzip(data: data1)
        let dir2 = try ZipHelper.unzip(data: data2)
        defer { ZipHelper.cleanup(dir1); ZipHelper.cleanup(dir2) }
        let core1 = try Data(contentsOf: dir1.appendingPathComponent("docProps/core.xml"))
        let core2 = try Data(contentsOf: dir2.appendingPathComponent("docProps/core.xml"))
        XCTAssertEqual(core1, core2, "docProps/core.xml must not depend on wall-clock time when no timestamp was set")
    }

    /// #151 second comment (the zip mod-time cause): "`write(_:to:)` 與
    /// `writeData(_:)` straddle a 2-second tick" is exactly what a tight loop
    /// usually does NOT do (both calls land in the same wall-clock second by
    /// luck almost every time — that is precisely why the issue calls the
    /// existing byte-equality test flaky rather than reliably failing).
    /// Sleeping past both the zip's 2-second DOS-time resolution and
    /// `ISO8601DateFormatter`'s 1-second resolution turns the race into a
    /// deterministic repro instead of relying on scheduler luck.
    func testWriteDataIsByteEqualAcrossASlowTick() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "reproducibility"))
        let first = try DocxWriter.writeData(doc)
        Thread.sleep(forTimeInterval: 2.5)
        let second = try DocxWriter.writeData(doc)
        XCTAssertEqual(first, second, "writeData output must not depend on wall-clock time between calls (#151)")
    }

    // MARK: - #216 — Content_Types Overrides for commentsExtended / endnotes

    func testScratchModeDeclaresContentTypesOverridesForCommentsExtendedAndEndnotes() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "p0"))
        let cid = try doc.insertComment(text: "c", author: "a", paragraphIndex: 0)
        if let idx = doc.comments.comments.firstIndex(where: { $0.id == cid }) {
            doc.comments.comments[idx].done = true   // forces commentsExtended.xml
        }
        _ = try doc.insertEndnote(text: "e", paragraphIndex: 0)

        let data = try DocxWriter.writeData(doc)
        let dir = try ZipHelper.unzip(data: data)
        defer { ZipHelper.cleanup(dir) }
        let contentTypes = try String(contentsOf: dir.appendingPathComponent("[Content_Types].xml"), encoding: .utf8)
        XCTAssertTrue(contentTypes.contains("/word/commentsExtended.xml"), "Content_Types must declare commentsExtended.xml (#216):\n\(contentTypes)")
        XCTAssertTrue(contentTypes.contains("/word/endnotes.xml"), "Content_Types must declare endnotes.xml (#216):\n\(contentTypes)")

        let relsXML = try String(contentsOf: dir.appendingPathComponent("word/_rels/document.xml.rels"), encoding: .utf8)
        XCTAssertTrue(relsXML.contains("endnotes.xml"), "endnotes.xml must be referenced from a relationship, not left orphaned (#216):\n\(relsXML)")
    }

    func testOverlayModeDeclaresContentTypesOverridesForCommentsExtendedAndEndnotes() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "p0"))
        let seed = FileManager.default.temporaryDirectory.appendingPathComponent("i64-seed-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: seed)
        defer { try? FileManager.default.removeItem(at: seed) }

        var document = try DocxReader.read(from: seed)
        defer { document.close() }
        let cid = try document.insertComment(text: "c", author: "a", paragraphIndex: 0)
        if let idx = document.comments.comments.firstIndex(where: { $0.id == cid }) {
            document.comments.comments[idx].done = true
        }
        _ = try document.insertEndnote(text: "e", paragraphIndex: 0)
        // `insertComment`/`insertEndnote` mark their own parts dirty but not
        // "[Content_Types].xml" or the document rels (that gap, if any, is
        // outside #216's scope — see report). Mark them explicitly so this
        // test isolates #216's actual question: once Content_Types/rels DO
        // get rewritten, are the Overrides/relationship correct?
        document.markTypedDirty("[Content_Types].xml")
        document.markTypedDirty("word/_rels/document.xml.rels")

        let out = FileManager.default.temporaryDirectory.appendingPathComponent("i64-out-\(UUID().uuidString).docx")
        try DocxWriter.write(document, to: out)
        defer { try? FileManager.default.removeItem(at: out) }
        let dir = try ZipHelper.unzip(out)
        defer { ZipHelper.cleanup(dir) }
        let contentTypes = try String(contentsOf: dir.appendingPathComponent("[Content_Types].xml"), encoding: .utf8)
        XCTAssertTrue(contentTypes.contains("/word/commentsExtended.xml"), "overlay Content_Types must declare commentsExtended.xml (#216):\n\(contentTypes)")
        XCTAssertTrue(contentTypes.contains("/word/endnotes.xml"), "overlay Content_Types must declare endnotes.xml (#216):\n\(contentTypes)")
        let relsXML = try String(contentsOf: dir.appendingPathComponent("word/_rels/document.xml.rels"), encoding: .utf8)
        XCTAssertTrue(relsXML.contains("endnotes.xml"), "overlay: endnotes.xml must not be an orphan part (#216):\n\(relsXML)")
    }

    // MARK: - #87 — captured root attribute order stability

    /// #87: a freshly-created document's `<w:document>` root uses the full
    /// Word-canonical namespace cloud, in canonical order. Once read back
    /// (`documentRootAttributes` becomes a `[String: String]`) and saved
    /// again, `renderDocumentRootOpenTag` currently re-sorts to w/r-first +
    /// alphabetical — a THIRD open→edit→save must not keep permuting it
    /// further (idempotence is the part this test can assert without
    /// depending on a specific canonical order).
    func testDocumentRootAttributeOrderIsIdempotentAcrossRepeatedOpenEditSave() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "root order"))
        let first = FileManager.default.temporaryDirectory.appendingPathComponent("i87-1-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: first)
        defer { try? FileManager.default.removeItem(at: first) }

        func rootOpenTag(of url: URL) throws -> String {
            let dir = try ZipHelper.unzip(url)
            defer { ZipHelper.cleanup(dir) }
            let xml = try String(contentsOf: dir.appendingPathComponent("word/document.xml"), encoding: .utf8)
            guard let range = xml.range(of: "<w:document"), let close = xml.range(of: ">", range: range.upperBound..<xml.endIndex) else {
                XCTFail("no <w:document> open tag found"); return ""
            }
            return String(xml[range.lowerBound..<close.upperBound])
        }

        var document2 = try DocxReader.read(from: first)
        document2.appendParagraph(Paragraph(text: "edit 1"))
        let second = FileManager.default.temporaryDirectory.appendingPathComponent("i87-2-\(UUID().uuidString).docx")
        try DocxWriter.write(document2, to: second)
        defer { try? FileManager.default.removeItem(at: second) }
        document2.close()

        var document3 = try DocxReader.read(from: second)
        document3.appendParagraph(Paragraph(text: "edit 2"))
        let third = FileManager.default.temporaryDirectory.appendingPathComponent("i87-3-\(UUID().uuidString).docx")
        try DocxWriter.write(document3, to: third)
        defer { try? FileManager.default.removeItem(at: third) }
        document3.close()

        let tag2 = try rootOpenTag(of: second)
        let tag3 = try rootOpenTag(of: third)
        XCTAssertEqual(tag2, tag3, "the root open tag must stop changing once every attribute set is stable — repeated open/edit/save must be idempotent (#87)")
    }
}
