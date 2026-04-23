import Foundation

/// DOCX 檔案寫入器
public struct DocxWriter {

    /// 將 WordDocument 寫入 .docx 檔案
    ///
    /// **Overlay mode** (v0.12.0+): when `document.archiveTempDir != nil`,
    /// the writer overwrites typed-model parts directly into the preserved
    /// tempDir (rather than rebuilding from a scratch tempDir), then zips
    /// the merged result. This preserves all OOXML parts the typed model
    /// does not manage (theme/, webSettings.xml, people.xml, glossary/,
    /// etc.) byte-for-byte.
    ///
    /// **Scratch mode**: when `archiveTempDir == nil` (initializer-built
    /// documents), behavior is unchanged from prior releases — writer builds
    /// a fresh scratch tempDir from typed model only.
    public static func write(_ document: WordDocument, to url: URL) throws {
        if FileManager.default.fileExists(atPath: url.path) {
            try FileManager.default.removeItem(at: url)
        }

        if let archiveTempDir = document.archiveTempDir {
            // Overlay mode: write typed parts into preserved tempDir, then zip.
            try writeAllParts(document, to: archiveTempDir, overlayMode: true)
            let data = try ZipHelper.zipToData(archiveTempDir)
            try data.write(to: url)
        } else {
            // Scratch mode: existing behavior unchanged.
            let data = try writeData(document)
            try data.write(to: url)
        }
    }

    /// 將 WordDocument 壓縮成 in-memory .docx bytes（不落地）
    ///
    /// In-memory variant always uses scratch mode (no source archive to
    /// preserve from). Callers wanting overlay-mode preservation MUST use
    /// `write(_:to:)` instead.
    public static func writeData(_ document: WordDocument) throws -> Data {
        let tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("che-word-mcp")
            .appendingPathComponent(UUID().uuidString)

        try FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
        defer { ZipHelper.cleanup(tempDir) }

        try writeAllParts(document, to: tempDir, overlayMode: false)
        return try ZipHelper.zipToData(tempDir)
    }

    /// 將 WordDocument 的所有 OOXML parts 寫到指定目錄（共享 pipeline）
    ///
    /// - Parameters:
    ///   - overlayMode: when `true`, `[Content_Types].xml` is computed via
    ///     `ContentTypesOverlay` to preserve original Override entries for
    ///     unknown parts (theme, webSettings, etc.). When `false`, all
    ///     part XMLs are emitted from typed model only (scratch mode).
    private static func writeAllParts(_ document: WordDocument, to tempDir: URL, overlayMode: Bool) throws {
        try createDirectoryStructure(at: tempDir)

        let hasNumbering = !document.numbering.abstractNums.isEmpty
        let hasHeaders = !document.headers.isEmpty
        let hasFooters = !document.footers.isEmpty
        let dirty = document.modifiedParts

        // v0.13.0+: in overlay mode every typed-part writer is gated by the
        // corresponding part path appearing in `dirty`. Scratch mode (no
        // archiveTempDir) writes everything unconditionally to preserve prior
        // behavior. The helper computes new typed parts not declared in the
        // original Content_Types so writeContentTypes still runs when the typed
        // model added (e.g.) a fresh header/footer/image even if the dirty set
        // doesn't explicitly contain `[Content_Types].xml`.
        let needsContentTypes = !overlayMode
            || dirty.contains("[Content_Types].xml")
            || hasNewTypedParts(document)
        let needsDocumentRels = !overlayMode
            || dirty.contains("word/_rels/document.xml.rels")
            || hasNewTypedRelationships(document)

        if needsContentTypes {
            try writeContentTypes(to: tempDir, document: document, overlayMode: overlayMode)
        }
        if !overlayMode {
            // Top-level _rels/.rels is read-only in overlay mode (preserved
            // verbatim from the source archive). Scratch mode emits a fresh one.
            try writeRelationships(to: tempDir)
        }
        if needsDocumentRels {
            try writeDocumentRelationships(to: tempDir, document: document)
        }
        if !overlayMode || dirty.contains("word/document.xml") {
            try writeDocument(document, to: tempDir)
        }
        if !overlayMode || dirty.contains("word/styles.xml") {
            try writeStyles(document.styles, to: tempDir)
        }
        if !overlayMode || dirty.contains("word/settings.xml") {
            try writeSettings(to: tempDir)
        }
        if !overlayMode || dirty.contains("word/fontTable.xml") {
            try writeFontTable(to: tempDir)
        }
        if !overlayMode || dirty.contains("docProps/core.xml") {
            try writeCoreProperties(document.properties, to: tempDir)
        }
        if !overlayMode || dirty.contains("docProps/app.xml") {
            try writeAppProperties(to: tempDir)
        }

        if hasNumbering, !overlayMode || dirty.contains("word/numbering.xml") {
            try writeNumbering(document.numbering, to: tempDir)
        }
        if hasHeaders {
            for header in document.headers {
                if !overlayMode || dirty.contains("word/\(header.fileName)") {
                    try writeHeader(header, to: tempDir)
                }
            }
        }
        if hasFooters {
            for footer in document.footers {
                if !overlayMode || dirty.contains("word/\(footer.fileName)") {
                    try writeFooter(footer, to: tempDir)
                }
            }
        }
        if !document.images.isEmpty {
            // Image binary writing is per-image — only re-emit images whose
            // media path is dirty (covers new image insertion). Existing source
            // images are already in the preserved archive.
            if overlayMode {
                try writeNewImages(document.images, to: tempDir, dirty: dirty)
            } else {
                try writeImages(document.images, to: tempDir)
            }
        }
        if !document.comments.comments.isEmpty,
           !overlayMode || dirty.contains("word/comments.xml") {
            try writeComments(document.comments, to: tempDir)
        }
        if let extXML = document.comments.toExtendedXML(),
           !overlayMode || dirty.contains("word/commentsExtended.xml") {
            try writeCommentsExtended(extXML, to: tempDir)
        }
        if !document.footnotes.footnotes.isEmpty,
           !overlayMode || dirty.contains("word/footnotes.xml") {
            try writeFootnotes(document.footnotes, to: tempDir)
        }
        if !document.endnotes.endnotes.isEmpty,
           !overlayMode || dirty.contains("word/endnotes.xml") {
            try writeEndnotes(document.endnotes, to: tempDir)
        }
    }

    /// True when the typed model contains parts not declared in the source
    /// archive's `[Content_Types].xml` — for example, a freshly added header
    /// (`addHeader` produces a new `word/headerN.xml`) or a new media file
    /// from `insertImage`. Used to gate `writeContentTypes` in overlay mode.
    private static func hasNewTypedParts(_ document: WordDocument) -> Bool {
        guard let tempDir = document.archiveTempDir else { return false }
        let originalCT: String
        do {
            originalCT = try String(contentsOf: tempDir.appendingPathComponent("[Content_Types].xml"), encoding: .utf8)
        } catch {
            return true  // Original missing — must rewrite Content_Types
        }
        for header in document.headers
        where !originalCT.contains("/word/\(header.fileName)") {
            return true
        }
        for footer in document.footers
        where !originalCT.contains("/word/\(footer.fileName)") {
            return true
        }
        for image in document.images
        where !originalCT.contains("/word/media/\(image.fileName)") {
            // Media additions trigger via Default extension, but if the
            // extension is new (e.g., first .webp), Content_Types must update.
            let ext = (image.fileName as NSString).pathExtension.lowercased()
            if !originalCT.contains("Extension=\"\(ext)\"") { return true }
        }
        return false
    }

    /// True when the typed model has relationships not declared in the source
    /// archive's `word/_rels/document.xml.rels` — used to gate
    /// `writeDocumentRelationships` in overlay mode.
    private static func hasNewTypedRelationships(_ document: WordDocument) -> Bool {
        guard let tempDir = document.archiveTempDir else { return false }
        let originalRels: String
        do {
            originalRels = try String(contentsOf: tempDir.appendingPathComponent("word/_rels/document.xml.rels"), encoding: .utf8)
        } catch {
            return true  // No source rels — emit fresh
        }
        for header in document.headers where !originalRels.contains("Id=\"\(header.id)\"") {
            return true
        }
        for footer in document.footers where !originalRels.contains("Id=\"\(footer.id)\"") {
            return true
        }
        for image in document.images where !originalRels.contains("Id=\"\(image.id)\"") {
            return true
        }
        for hyperlinkRef in document.hyperlinkReferences
        where !originalRels.contains("Id=\"\(hyperlinkRef.relationshipId)\"") {
            return true
        }
        return false
    }

    /// Overlay-mode image writer — only emits media files that are in `dirty`
    /// (i.e., images added since `DocxReader.read()`). Existing source images
    /// remain in the preserved archive untouched.
    private static func writeNewImages(_ images: [ImageReference], to baseURL: URL, dirty: Set<String>) throws {
        for image in images where dirty.contains("word/media/\(image.fileName)") {
            let url = baseURL.appendingPathComponent("word/media/\(image.fileName)")
            try image.data.write(to: url)
        }
    }

    // MARK: - Directory Structure

    private static func createDirectoryStructure(at baseURL: URL) throws {
        let directories = [
            "_rels",
            "word",
            "word/_rels",
            "word/media",  // 圖片媒體目錄
            "docProps"
        ]

        for dir in directories {
            let dirURL = baseURL.appendingPathComponent(dir)
            try FileManager.default.createDirectory(at: dirURL, withIntermediateDirectories: true)
        }
    }

    // MARK: - Content Types

    private static func writeContentTypes(to baseURL: URL, document: WordDocument, overlayMode: Bool = false) throws {
        if overlayMode, document.archiveTempDir != nil {
            try writeContentTypesOverlay(to: baseURL, document: document)
            return
        }

        let hasNumbering = !document.numbering.abstractNums.isEmpty

        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Default Extension="png" ContentType="image/png"/>
            <Default Extension="jpeg" ContentType="image/jpeg"/>
            <Default Extension="jpg" ContentType="image/jpeg"/>
            <Default Extension="gif" ContentType="image/gif"/>
            <Default Extension="bmp" ContentType="image/bmp"/>
            <Default Extension="tiff" ContentType="image/tiff"/>
            <Default Extension="webp" ContentType="image/webp"/>
            <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
            <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
            <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
            <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
            <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
        """

        if hasNumbering {
            xml += """
                <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
            """
        }

        // 頁首
        for header in document.headers {
            xml += """
                <Override PartName="/word/\(header.fileName)" ContentType="\(Header.contentType)"/>
            """
        }

        // 頁尾
        for footer in document.footers {
            xml += """
                <Override PartName="/word/\(footer.fileName)" ContentType="\(Footer.contentType)"/>
            """
        }

        // 註解
        if !document.comments.comments.isEmpty {
            xml += """
                <Override PartName="/word/comments.xml" ContentType="\(CommentsCollection.contentType)"/>
            """
        }

        // commentsExtended（回覆和已解決狀態）
        if document.comments.hasExtendedComments {
            xml += """
                <Override PartName="/word/commentsExtended.xml" ContentType="\(CommentsCollection.extendedContentType)"/>
            """
        }

        // 腳註
        if !document.footnotes.footnotes.isEmpty {
            xml += """
                <Override PartName="/word/footnotes.xml" ContentType="\(FootnotesCollection.contentType)"/>
            """
        }

        // 尾註
        if !document.endnotes.endnotes.isEmpty {
            xml += """
                <Override PartName="/word/endnotes.xml" ContentType="\(EndnotesCollection.contentType)"/>
            """
        }

        xml += "</Types>"

        let url = baseURL.appendingPathComponent("[Content_Types].xml")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    /// Overlay-mode `[Content_Types].xml`: read original from preserved
    /// archive tempDir, compute typed-parts list, merge via
    /// `ContentTypesOverlay`, write merged result. Preserves Overrides for
    /// theme / webSettings / people / glossary / etc. that the typed model
    /// does not manage.
    private static func writeContentTypesOverlay(to baseURL: URL, document: WordDocument) throws {
        guard let archiveTempDir = document.archiveTempDir else {
            // Caller verified non-nil; defensive fallback to scratch mode.
            try writeContentTypes(to: baseURL, document: document, overlayMode: false)
            return
        }
        let originalCT = (try? String(
            contentsOf: archiveTempDir.appendingPathComponent("[Content_Types].xml"),
            encoding: .utf8
        )) ?? ""
        let overlay = ContentTypesOverlay(originalContentTypesXML: originalCT)
        let merged = overlay.merge(
            typedParts: typedPartDescriptors(for: document),
            typedManagedPatterns: typedManagedPatternsForOverlay(document)
        )
        let url = baseURL.appendingPathComponent("[Content_Types].xml")
        try merged.write(to: url, atomically: true, encoding: .utf8)
    }

    /// Compute the PartDescriptors the writer is about to emit (typed model
    /// state). Used by overlay-mode Content_Types merge.
    private static func typedPartDescriptors(for document: WordDocument) -> [PartDescriptor] {
        var parts: [PartDescriptor] = [
            PartDescriptor(partName: "/word/document.xml",
                           contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"),
            PartDescriptor(partName: "/word/styles.xml",
                           contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"),
            PartDescriptor(partName: "/word/settings.xml",
                           contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"),
            PartDescriptor(partName: "/word/fontTable.xml",
                           contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"),
            PartDescriptor(partName: "/docProps/core.xml",
                           contentType: "application/vnd.openxmlformats-package.core-properties+xml"),
            PartDescriptor(partName: "/docProps/app.xml",
                           contentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml")
        ]
        if !document.numbering.abstractNums.isEmpty {
            parts.append(PartDescriptor(
                partName: "/word/numbering.xml",
                contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
            ))
        }
        for header in document.headers {
            parts.append(PartDescriptor(partName: "/word/\(header.fileName)", contentType: Header.contentType))
        }
        for footer in document.footers {
            parts.append(PartDescriptor(partName: "/word/\(footer.fileName)", contentType: Footer.contentType))
        }
        if !document.comments.comments.isEmpty {
            parts.append(PartDescriptor(partName: "/word/comments.xml", contentType: CommentsCollection.contentType))
        }
        if document.comments.hasExtendedComments {
            parts.append(PartDescriptor(partName: "/word/commentsExtended.xml",
                                        contentType: CommentsCollection.extendedContentType))
        }
        if !document.footnotes.footnotes.isEmpty {
            parts.append(PartDescriptor(partName: "/word/footnotes.xml",
                                        contentType: FootnotesCollection.contentType))
        }
        if !document.endnotes.endnotes.isEmpty {
            parts.append(PartDescriptor(partName: "/word/endnotes.xml",
                                        contentType: EndnotesCollection.contentType))
        }
        return parts
    }

    /// PartName patterns the typed model owns. Overlay drops original Overrides
    /// matching these patterns when the typed parts list omits them (= deletion).
    private static func typedManagedPatternsForOverlay(_ document: WordDocument) -> [String] {
        return [
            "/word/document.xml",
            "/word/styles.xml",
            "/word/settings.xml",
            "/word/fontTable.xml",
            "/word/numbering.xml",
            "/word/header",   // prefix: header1.xml, header2.xml, ...
            "/word/footer",   // prefix: footer1.xml, footer2.xml, ...
            "/word/comments.xml",
            "/word/commentsExtended.xml",
            "/word/footnotes.xml",
            "/word/endnotes.xml",
            "/docProps/core.xml",
            "/docProps/app.xml"
        ]
    }

    // MARK: - Relationships

    private static func writeRelationships(to baseURL: URL) throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
            <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
        </Relationships>
        """

        let url = baseURL.appendingPathComponent("_rels/.rels")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    /// Type URLs the typed model owns. An original rel of one of these types
    /// will be re-emitted from the typed model's authoritative state; an
    /// original rel of any OTHER type is preserved verbatim by overlay merge
    /// (theme / webSettings / customXml / commentsExtensible / commentsIds /
    /// people / etc.). Added in v0.13.1 (closes che-word-mcp#35).
    private static let typedManagedRelationshipTypes: Set<String> = [
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        Header.relationshipType,
        Footer.relationshipType,
        Hyperlink.relationshipType,
        CommentsCollection.relationshipType,
        CommentsCollection.extendedRelationshipType,
        FootnotesCollection.relationshipType,
        EndnotesCollection.relationshipType,
    ]

    private static func writeDocumentRelationships(to baseURL: URL, document: WordDocument) throws {
        let originalRelsXML: String
        if let archiveTempDir = document.archiveTempDir {
            let url = archiveTempDir.appendingPathComponent("word/_rels/document.xml.rels")
            originalRelsXML = (try? String(contentsOf: url, encoding: .utf8)) ?? ""
        } else {
            originalRelsXML = ""
        }

        // Build allocator for newly-needed rIds (comments/footnotes/endnotes
        // when typed model has them but original rels doesn't).
        var typedReservedIds: [String] = ["rId1", "rId2", "rId3"]
        if !document.numbering.abstractNums.isEmpty { typedReservedIds.append("rId4") }
        for header in document.headers { typedReservedIds.append(header.id) }
        for footer in document.footers { typedReservedIds.append(footer.id) }
        for image in document.images { typedReservedIds.append(image.id) }
        for hyperlinkRef in document.hyperlinkReferences { typedReservedIds.append(hyperlinkRef.relationshipId) }
        let allocator = RelationshipIdAllocator(
            originalRelsXML: originalRelsXML,
            additionalReservedIds: typedReservedIds
        )

        let typedRels = buildTypedRelationships(document: document, allocator: allocator)

        let xml: String
        if document.archiveTempDir != nil && !originalRelsXML.isEmpty {
            // Overlay mode: merge typed rels into original to preserve unknown
            // types (theme / webSettings / people / customXml / etc.).
            let overlay = RelationshipsOverlay(originalRelsXML: originalRelsXML)
            xml = overlay.merge(
                typedRels: typedRels,
                typedManagedTypes: Self.typedManagedRelationshipTypes
            )
        } else {
            // Scratch mode (no source archive): emit fresh rels from typed model only.
            xml = serializeScratchRels(typedRels)
        }

        let url = baseURL.appendingPathComponent("word/_rels/document.xml.rels")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    /// Collect all rels the typed model wants to emit. Used by both overlay
    /// merge (where these go through `RelationshipsOverlay`) and scratch mode.
    private static func buildTypedRelationships(
        document: WordDocument,
        allocator: RelationshipIdAllocator
    ) -> [RelationshipDescriptor] {
        var rels: [RelationshipDescriptor] = [
            RelationshipDescriptor(
                id: "rId1",
                type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                target: "styles.xml", targetMode: nil
            ),
            RelationshipDescriptor(
                id: "rId2",
                type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
                target: "settings.xml", targetMode: nil
            ),
            RelationshipDescriptor(
                id: "rId3",
                type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
                target: "fontTable.xml", targetMode: nil
            ),
        ]
        if !document.numbering.abstractNums.isEmpty {
            rels.append(RelationshipDescriptor(
                id: "rId4",
                type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
                target: "numbering.xml", targetMode: nil
            ))
        }
        for header in document.headers {
            rels.append(RelationshipDescriptor(
                id: header.id, type: Header.relationshipType,
                target: header.fileName, targetMode: nil
            ))
        }
        for footer in document.footers {
            rels.append(RelationshipDescriptor(
                id: footer.id, type: Footer.relationshipType,
                target: footer.fileName, targetMode: nil
            ))
        }
        for image in document.images {
            rels.append(RelationshipDescriptor(
                id: image.id,
                type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                target: "media/\(image.fileName)", targetMode: nil
            ))
        }
        for hyperlinkRef in document.hyperlinkReferences {
            rels.append(RelationshipDescriptor(
                id: hyperlinkRef.relationshipId, type: Hyperlink.relationshipType,
                target: hyperlinkRef.url, targetMode: "External"
            ))
        }
        if !document.comments.comments.isEmpty {
            rels.append(RelationshipDescriptor(
                id: allocator.allocate(), type: CommentsCollection.relationshipType,
                target: "comments.xml", targetMode: nil
            ))
        }
        if document.comments.hasExtendedComments {
            rels.append(RelationshipDescriptor(
                id: allocator.allocate(), type: CommentsCollection.extendedRelationshipType,
                target: "commentsExtended.xml", targetMode: nil
            ))
        }
        if !document.footnotes.footnotes.isEmpty {
            rels.append(RelationshipDescriptor(
                id: allocator.allocate(), type: FootnotesCollection.relationshipType,
                target: "footnotes.xml", targetMode: nil
            ))
        }
        if !document.endnotes.endnotes.isEmpty {
            rels.append(RelationshipDescriptor(
                id: allocator.allocate(), type: EndnotesCollection.relationshipType,
                target: "endnotes.xml", targetMode: nil
            ))
        }
        return rels
    }

    /// Scratch-mode rels serializer (no source archive). Preserves the exact
    /// pre-v0.13.1 output for `create_document` callers.
    private static func serializeScratchRels(_ rels: [RelationshipDescriptor]) -> String {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        """
        for rel in rels {
            xml += "\n    <Relationship Id=\"\(rel.id)\" Type=\"\(rel.type)\" Target=\"\(escapeXML(rel.target))\""
            if let mode = rel.targetMode {
                xml += " TargetMode=\"\(mode)\""
            }
            xml += "/>"
        }
        xml += "\n</Relationships>"
        return xml
    }

    // MARK: - Document

    private static func writeDocument(_ document: WordDocument, to baseURL: URL) throws {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
        """

        // 段落和表格
        for child in document.body.children {
            switch child {
            case .paragraph(let para):
                xml += para.toXML()
            case .table(let table):
                xml += table.toXML()
            }
        }

        // 分節屬性（頁面設定）- 使用文件的 sectionProperties
        xml += document.sectionProperties.toXML()

        xml += "</w:body></w:document>"

        let url = baseURL.appendingPathComponent("word/document.xml")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Styles

    private static func writeStyles(_ styles: [Style], to baseURL: URL) throws {
        let xml = styles.toStylesXML()
        let url = baseURL.appendingPathComponent("word/styles.xml")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Numbering

    private static func writeNumbering(_ numbering: Numbering, to baseURL: URL) throws {
        let xml = numbering.toXML()
        let url = baseURL.appendingPathComponent("word/numbering.xml")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Header

    private static func writeHeader(_ header: Header, to baseURL: URL) throws {
        let xml = header.toXML()
        let url = baseURL.appendingPathComponent("word/\(header.fileName)")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Footer

    private static func writeFooter(_ footer: Footer, to baseURL: URL) throws {
        let xml: String

        // 如果有指定頁碼格式，使用頁碼格式生成 XML
        if let format = footer.pageNumberFormat {
            xml = footer.toXMLWithPageNumber(format: format, alignment: footer.pageNumberAlignment)
        } else if footer.paragraphs.isEmpty {
            // 沒有段落也沒有頁碼格式，使用預設簡單頁碼
            xml = footer.toXMLWithPageNumber(format: .simple)
        } else {
            // 有段落內容，使用一般 XML 輸出
            xml = footer.toXML()
        }

        let url = baseURL.appendingPathComponent("word/\(footer.fileName)")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Images

    private static func writeImages(_ images: [ImageReference], to baseURL: URL) throws {
        for image in images {
            let url = baseURL.appendingPathComponent("word/media/\(image.fileName)")
            try image.data.write(to: url)
        }
    }

    // MARK: - Settings

    private static func writeSettings(to baseURL: URL) throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:defaultTabStop w:val="720"/>
            <w:characterSpacingControl w:val="doNotCompress"/>
            <w:compat>
                <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
            </w:compat>
        </w:settings>
        """

        let url = baseURL.appendingPathComponent("word/settings.xml")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Font Table

    private static func writeFontTable(to baseURL: URL) throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:font w:name="Calibri">
                <w:panose1 w:val="020F0502020204030204"/>
                <w:charset w:val="00"/>
                <w:family w:val="swiss"/>
                <w:pitch w:val="variable"/>
            </w:font>
            <w:font w:name="Times New Roman">
                <w:panose1 w:val="02020603050405020304"/>
                <w:charset w:val="00"/>
                <w:family w:val="roman"/>
                <w:pitch w:val="variable"/>
            </w:font>
            <w:font w:name="Calibri Light">
                <w:panose1 w:val="020F0302020204030204"/>
                <w:charset w:val="00"/>
                <w:family w:val="swiss"/>
                <w:pitch w:val="variable"/>
            </w:font>
        </w:fonts>
        """

        let url = baseURL.appendingPathComponent("word/fontTable.xml")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Core Properties

    private static func writeCoreProperties(_ props: DocumentProperties, to baseURL: URL) throws {
        let dateFormatter = ISO8601DateFormatter()

        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                           xmlns:dc="http://purl.org/dc/elements/1.1/"
                           xmlns:dcterms="http://purl.org/dc/terms/"
                           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
        """

        if let title = props.title {
            xml += "<dc:title>\(escapeXML(title))</dc:title>"
        }
        if let subject = props.subject {
            xml += "<dc:subject>\(escapeXML(subject))</dc:subject>"
        }
        if let creator = props.creator {
            xml += "<dc:creator>\(escapeXML(creator))</dc:creator>"
        } else {
            xml += "<dc:creator>che-word-mcp</dc:creator>"
        }
        if let keywords = props.keywords {
            xml += "<cp:keywords>\(escapeXML(keywords))</cp:keywords>"
        }
        if let description = props.description {
            xml += "<dc:description>\(escapeXML(description))</dc:description>"
        }
        if let lastModifiedBy = props.lastModifiedBy {
            xml += "<cp:lastModifiedBy>\(escapeXML(lastModifiedBy))</cp:lastModifiedBy>"
        }
        if let revision = props.revision {
            xml += "<cp:revision>\(revision)</cp:revision>"
        } else {
            xml += "<cp:revision>1</cp:revision>"
        }

        let created = props.created ?? Date()
        xml += "<dcterms:created xsi:type=\"dcterms:W3CDTF\">\(dateFormatter.string(from: created))</dcterms:created>"

        let modified = props.modified ?? Date()
        xml += "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">\(dateFormatter.string(from: modified))</dcterms:modified>"

        xml += "</cp:coreProperties>"

        let url = baseURL.appendingPathComponent("docProps/core.xml")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - App Properties

    private static func writeAppProperties(to baseURL: URL) throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
            <Application>che-word-mcp</Application>
            <AppVersion>1.0.0</AppVersion>
        </Properties>
        """

        let url = baseURL.appendingPathComponent("docProps/app.xml")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Comments

    private static func writeComments(_ comments: CommentsCollection, to baseURL: URL) throws {
        let xml = comments.toXML()
        let url = baseURL.appendingPathComponent("word/comments.xml")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Comments Extended

    private static func writeCommentsExtended(_ xml: String, to baseURL: URL) throws {
        let url = baseURL.appendingPathComponent("word/commentsExtended.xml")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Footnotes

    private static func writeFootnotes(_ footnotes: FootnotesCollection, to baseURL: URL) throws {
        let xml = footnotes.toXML()
        let url = baseURL.appendingPathComponent("word/footnotes.xml")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Endnotes

    private static func writeEndnotes(_ endnotes: EndnotesCollection, to baseURL: URL) throws {
        let xml = endnotes.toXML()
        let url = baseURL.appendingPathComponent("word/endnotes.xml")
        try xml.write(to: url, atomically: true, encoding: .utf8)
    }

    // MARK: - Helpers

    private static func escapeXML(_ string: String) -> String {
        return string
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
            .replacingOccurrences(of: "'", with: "&apos;")
    }
}
