import Foundation

/// DOCX 檔案讀取器
public struct DocxReader {

    /// When `true`, the parser writes one line to stderr for each direct child
    /// of `<w:p>` whose local name is not recognized by `parseParagraph`. Use
    /// during development / tests to surface parser coverage gaps. Default
    /// `false` — production parses have zero overhead (the flag is checked
    /// before any string formatting).
    ///
    /// Not thread-safe for concurrent toggles during parallel parses. Toggle at
    /// test setup / teardown, not during document parsing.
    public static var debugLoggingEnabled: Bool = false

    /// 讀取 .docx 檔案並解析為 WordDocument
    public static func read(from url: URL) throws -> WordDocument {
        guard FileManager.default.fileExists(atPath: url.path) else {
            throw WordError.fileNotFound(url.path)
        }

        // 1. 解壓縮 ZIP
        let tempDir = try ZipHelper.unzip(url)

        // tempDir is retained on the returned WordDocument for preserve-by-default
        // round-trip fidelity (v0.12.0+). Only clean up on error paths — success
        // hands ownership to the document via `preservedArchive`. Caller releases
        // via `WordDocument.close()`.
        var transferOwnership = false
        defer {
            if !transferOwnership {
                ZipHelper.cleanup(tempDir)
            }
        }

        // 2. 讀取關係檔案 word/_rels/document.xml.rels
        let relationships = try parseRelationships(from: tempDir)

        // 3. 提取圖片資源
        let images = try extractImages(from: tempDir, relationships: relationships)

        // 4. 讀取 document.xml
        let documentURL = tempDir.appendingPathComponent("word/document.xml")
        guard FileManager.default.fileExists(atPath: documentURL.path) else {
            throw WordError.parseError("找不到 word/document.xml")
        }

        let documentData = try Data(contentsOf: documentURL)
        let documentXML = try XMLDocument(data: documentData)

        // 5. 讀取 styles.xml（先解析，用於語義標註）
        var document = WordDocument()

        // 5a. v0.19.0+ (PsychQuant/che-word-mcp#56): preserve every attribute
        // (xmlns:* declarations, mc:Ignorable, anything else) on the source
        // <w:document> root so DocxWriter can rebuild the open tag verbatim
        // instead of collapsing to the hardcoded xmlns:w + xmlns:r pair.
        // Without this, libxml2 reports "unbound prefix" on every body element
        // that references an undeclared namespace.
        //
        // Foundation's XMLDocument (libxml2 under the hood) splits namespace
        // declarations onto `XMLNode.namespaces` (separate from `attributes`)
        // and silently drops xmlns:* prefixes that are not referenced in the
        // tree it can see — so to capture the source set verbatim we re-parse
        // the root open tag from raw bytes instead of relying on the DOM.
        document.documentRootAttributes = Self.parseDocumentRootAttributes(from: documentData)

        let stylesURL = tempDir.appendingPathComponent("word/styles.xml")
        if FileManager.default.fileExists(atPath: stylesURL.path) {
            let stylesData = try Data(contentsOf: stylesURL)
            let stylesXML = try XMLDocument(data: stylesData)
            document.styles = try parseStyles(from: stylesXML)
            document.latentStyles = parseLatentStyles(from: stylesXML)
        }

        // 6. 讀取 numbering.xml（可選，用於清單語義標註）
        let numberingURL = tempDir.appendingPathComponent("word/numbering.xml")
        if FileManager.default.fileExists(atPath: numberingURL.path) {
            let numberingData = try Data(contentsOf: numberingURL)
            let numberingXML = try XMLDocument(data: numberingData)
            document.numbering = try parseNumbering(from: numberingXML)
        }

        // 7. 解析文件內容（傳入 styles 和 numbering 用於語義標註）
        document.body = try parseBody(
            from: documentXML,
            relationships: relationships,
            styles: document.styles,
            numbering: document.numbering
        )
        document.images = images

        // v0.19.3+ (#56 round 2 P1-1): calibrate nextBookmarkId past every
        // bookmark id present in the source so subsequent `insertBookmark`
        // calls cannot collide with existing source bookmarks. Pre-fix the
        // counter started at 1 regardless of source content; F2 turned the
        // resulting collision from a silent drop into a silent overwrite
        // (Word may schema-reject or treat both bookmarks as one).
        var maxBookmarkId = 0
        for child in document.body.children {
            if case .paragraph(let para) = child {
                for bookmark in para.bookmarks {
                    if bookmark.id > maxBookmarkId { maxBookmarkId = bookmark.id }
                }
                for marker in para.bookmarkMarkers {
                    if marker.id > maxBookmarkId { maxBookmarkId = marker.id }
                }
            }
        }
        if maxBookmarkId > 0 {
            document.nextBookmarkId = maxBookmarkId + 1
        }

        // 7b. 讀取 headers (Part C of ooxml-swift#1)
        // v0.13.0+: preserve `originalFileName` from rel.target so multi-instance
        // same-type headers (header1.xml..header6.xml) don't collapse to a single fileName.
        // v0.13.5+ (#55): validate rel.target BEFORE forming the read URL to
        // block path traversal at the read sink (the property setter on
        // Header.originalFileName is defense-in-depth for post-load mutation).
        for rel in relationships.relationships where rel.type == .header {
            guard isSafeRelativeOOXMLPath(rel.target) else {
                FileHandle.standardError.write(
                    Data("Warning: DocxReader skipped header rel '\(rel.id)' with unsafe target '\(rel.target)' (#55 security baseline)\n".utf8)
                )
                continue
            }
            let headerURL = tempDir.appendingPathComponent("word/\(rel.target)")
            guard FileManager.default.fileExists(atPath: headerURL.path) else { continue }
            let headerData = try Data(contentsOf: headerURL)
            let headerXML = try XMLDocument(data: headerData)
            let paragraphs = try parseContainerParagraphs(
                from: headerXML,
                relationships: relationships, styles: document.styles, numbering: document.numbering
            )
            // v0.19.2+ (#56 follow-up F4): preserve `<w:hdr>` root attributes
            // (xmlns:* + mc:Ignorable + vendor) so VML watermark prefixes
            // round-trip beyond the hardcoded 5-namespace template.
            let rootAttrs = Self.parseContainerRootAttributes(
                from: headerData, rootElementOpenPrefix: "<w:hdr"
            )
            document.headers.append(
                Header(id: rel.id, paragraphs: paragraphs, originalFileName: rel.target, rootAttributes: rootAttrs)
            )
        }

        // 7c. 讀取 footers
        // v0.13.0+: see headers comment above re: originalFileName preservation.
        // v0.13.5+ (#55): same path-traversal guard as headers.
        for rel in relationships.relationships where rel.type == .footer {
            guard isSafeRelativeOOXMLPath(rel.target) else {
                FileHandle.standardError.write(
                    Data("Warning: DocxReader skipped footer rel '\(rel.id)' with unsafe target '\(rel.target)' (#55 security baseline)\n".utf8)
                )
                continue
            }
            let footerURL = tempDir.appendingPathComponent("word/\(rel.target)")
            guard FileManager.default.fileExists(atPath: footerURL.path) else { continue }
            let footerData = try Data(contentsOf: footerURL)
            let footerXML = try XMLDocument(data: footerData)
            let paragraphs = try parseContainerParagraphs(
                from: footerXML,
                relationships: relationships, styles: document.styles, numbering: document.numbering
            )
            // v0.19.2+ (#56 follow-up F4): preserve `<w:ftr>` root attributes.
            let rootAttrs = Self.parseContainerRootAttributes(
                from: footerData, rootElementOpenPrefix: "<w:ftr"
            )
            document.footers.append(
                Footer(id: rel.id, paragraphs: paragraphs, originalFileName: rel.target, rootAttributes: rootAttrs)
            )
        }

        // 7d. 讀取 footnotes
        let footnotesURL = tempDir.appendingPathComponent("word/footnotes.xml")
        if FileManager.default.fileExists(atPath: footnotesURL.path) {
            let footnotesData = try Data(contentsOf: footnotesURL)
            let footnotesXML = try XMLDocument(data: footnotesData)
            // v0.19.2+ (#56 follow-up F4): preserve `<w:footnotes>` root attributes.
            document.footnotes.rootAttributes = Self.parseContainerRootAttributes(
                from: footnotesData, rootElementOpenPrefix: "<w:footnotes"
            )
            if let root = footnotesXML.rootElement() {
                for child in root.children ?? [] {
                    guard
                        let element = child as? XMLElement,
                        element.localName == "footnote",
                        let idStr = element.attribute(forName: "w:id")?.stringValue,
                        let id = Int(idStr)
                    else { continue }
                    // Skip structural footnotes (separator / continuationSeparator) by w:type attribute
                    let fnType = element.attribute(forName: "w:type")?.stringValue
                    if fnType == "separator" || fnType == "continuationSeparator" { continue }
                    let paragraphs = try parseContainerChildParagraphs(
                        in: element,
                        relationships: relationships, styles: document.styles, numbering: document.numbering
                    )
                    let text = paragraphs.map { $0.getText() }.joined(separator: " ")
                    var footnote = Footnote(id: id, text: text, paragraphIndex: 0)
                    footnote.paragraphs = paragraphs
                    document.footnotes.footnotes.append(footnote)
                }
            }
        }

        // 7e. 讀取 endnotes
        let endnotesURL = tempDir.appendingPathComponent("word/endnotes.xml")
        if FileManager.default.fileExists(atPath: endnotesURL.path) {
            let endnotesData = try Data(contentsOf: endnotesURL)
            let endnotesXML = try XMLDocument(data: endnotesData)
            // v0.19.2+ (#56 follow-up F4): preserve `<w:endnotes>` root attributes.
            document.endnotes.rootAttributes = Self.parseContainerRootAttributes(
                from: endnotesData, rootElementOpenPrefix: "<w:endnotes"
            )
            if let root = endnotesXML.rootElement() {
                for child in root.children ?? [] {
                    guard
                        let element = child as? XMLElement,
                        element.localName == "endnote",
                        let idStr = element.attribute(forName: "w:id")?.stringValue,
                        let id = Int(idStr)
                    else { continue }
                    // Skip structural endnotes (separator / continuationSeparator) by w:type attribute
                    let enType = element.attribute(forName: "w:type")?.stringValue
                    if enType == "separator" || enType == "continuationSeparator" { continue }
                    let paragraphs = try parseContainerChildParagraphs(
                        in: element,
                        relationships: relationships, styles: document.styles, numbering: document.numbering
                    )
                    let text = paragraphs.map { $0.getText() }.joined(separator: " ")
                    var endnote = Endnote(id: id, text: text, paragraphIndex: 0)
                    endnote.paragraphs = paragraphs
                    document.endnotes.endnotes.append(endnote)
                }
            }
        }

        // 8. 讀取 core.xml（可選）
        let coreURL = tempDir.appendingPathComponent("docProps/core.xml")
        if FileManager.default.fileExists(atPath: coreURL.path) {
            let coreData = try Data(contentsOf: coreURL)
            let coreXML = try XMLDocument(data: coreData)
            document.properties = try parseCoreProperties(from: coreXML)
        }

        // 8. 讀取 comments.xml（可選）
        let commentsURL = tempDir.appendingPathComponent("word/comments.xml")
        if FileManager.default.fileExists(atPath: commentsURL.path) {
            let commentsData = try Data(contentsOf: commentsURL)
            let commentsXML = try XMLDocument(data: commentsData)
            document.comments = try parseComments(from: commentsXML)
        }

        // 9. Link comment paragraphIndex from paragraph commentIds
        for (index, child) in document.body.children.enumerated() {
            if case .paragraph(let para) = child {
                for commentId in para.commentIds {
                    if let idx = document.comments.comments.firstIndex(where: { $0.id == commentId }) {
                        document.comments.comments[idx].paragraphIndex = index
                    }
                }
            }
        }

        // 10. 收集段落內的修訂記錄到 document.revisions
        // Body revisions (source = .body)
        for (index, child) in document.body.children.enumerated() {
            switch child {
            case .paragraph(let para):
                for var revision in para.revisions {
                    revision.paragraphIndex = index
                    revision.source = .body
                    document.revisions.revisions.append(revision)
                }
            case .table(let table):
                for (rowIdx, row) in table.rows.enumerated() {
                    for (colIdx, cell) in row.cells.enumerated() {
                        for (cellParaIdx, para) in cell.paragraphs.enumerated() {
                            for var revision in para.revisions {
                                revision.paragraphIndex = index
                                revision.source = .body
                                revision.tableRow = rowIdx
                                revision.tableColumn = colIdx
                                revision.cellParagraphIndex = cellParaIdx
                                document.revisions.revisions.append(revision)
                            }
                        }
                    }
                }
            case .contentControl:
                // #44 task 3.4: revision tracking inside block-level SDTs
                // is deferred — these wrappers don't carry their own
                // revisions, and child paragraph revisions are already
                // attached when each paragraph is parsed. A future task
                // would index revisions through the SDT children too.
                break
            }
        }

        // Header revisions (source = .header(id:))
        for header in document.headers {
            for (paraIdx, para) in header.paragraphs.enumerated() {
                for var revision in para.revisions {
                    revision.paragraphIndex = paraIdx
                    revision.source = .header(id: header.id)
                    document.revisions.revisions.append(revision)
                }
            }
        }

        // Footer revisions (source = .footer(id:))
        for footer in document.footers {
            for (paraIdx, para) in footer.paragraphs.enumerated() {
                for var revision in para.revisions {
                    revision.paragraphIndex = paraIdx
                    revision.source = .footer(id: footer.id)
                    document.revisions.revisions.append(revision)
                }
            }
        }

        // Footnote revisions (source = .footnote(id:))
        for footnote in document.footnotes.footnotes {
            for (paraIdx, para) in footnote.paragraphs.enumerated() {
                for var revision in para.revisions {
                    revision.paragraphIndex = paraIdx
                    revision.source = .footnote(id: footnote.id)
                    document.revisions.revisions.append(revision)
                }
            }
        }

        // Endnote revisions (source = .endnote(id:))
        for endnote in document.endnotes.endnotes {
            for (paraIdx, para) in endnote.paragraphs.enumerated() {
                for var revision in para.revisions {
                    revision.paragraphIndex = paraIdx
                    revision.source = .endnote(id: endnote.id)
                    document.revisions.revisions.append(revision)
                }
            }
        }

        // 11. 讀取 commentsExtended.xml（可選，Word 2012+ 回覆與已解決狀態）
        let commentsExtURL = tempDir.appendingPathComponent("word/commentsExtended.xml")
        if FileManager.default.fileExists(atPath: commentsExtURL.path) {
            let extData = try Data(contentsOf: commentsExtURL)
            let extXML = try XMLDocument(data: extData)
            try parseCommentsExtended(from: extXML, into: &document.comments)
        }

        // Hand tempDir ownership to the returned WordDocument. Caller MUST call
        // `doc.close()` when finished to release it (otherwise the tempDir leaks
        // until process exit; macOS reclaims `/tmp` on reboot).
        document.preservedArchive = PreservedArchive(tempDir: tempDir)
        transferOwnership = true

        // v0.13.0+: clear modifiedParts to empty as the final step before returning.
        // Guarantees freshly loaded documents start with `modifiedParts.isEmpty == true`,
        // so DocxWriter overlay mode skips every typed-part writer until the caller
        // mutates the typed model.
        document.modifiedParts.removeAll()

        return document
    }

    // MARK: - Relationships Parsing

    /// 解析關係檔案
    private static func parseRelationships(from tempDir: URL) throws -> RelationshipsCollection {
        var collection = RelationshipsCollection()

        let relsURL = tempDir.appendingPathComponent("word/_rels/document.xml.rels")
        guard FileManager.default.fileExists(atPath: relsURL.path) else {
            // 沒有關係檔案也是合法的
            return collection
        }

        let relsData = try Data(contentsOf: relsURL)
        let relsXML = try XMLDocument(data: relsData)

        // 取得所有 Relationship 節點
        let relNodes = try relsXML.nodes(forXPath: "//*[local-name()='Relationship']")

        for node in relNodes {
            guard let element = node as? XMLElement else { continue }

            guard let id = element.attribute(forName: "Id")?.stringValue,
                  let typeStr = element.attribute(forName: "Type")?.stringValue,
                  let target = element.attribute(forName: "Target")?.stringValue else {
                continue
            }

            let relationship = Relationship(
                id: id,
                type: RelationshipType(rawValue: typeStr),
                target: target
            )
            collection.relationships.append(relationship)
        }

        return collection
    }

    // MARK: - Image Extraction

    /// 從 word/media/ 提取圖片
    /// v0.13.1 (closes che-word-mcp#35 root cause A): rewritten to be
    /// **relationship-driven** instead of directory-driven. Pre-v0.13.1 the
    /// loop iterated `word/media/` and fell back to `"rId_\(fileName)"` when
    /// `targetToId` lookup missed — producing forged ids like `rId_image1.png`
    /// that violate the OOXML `rId[0-9]+` convention AND made
    /// `hasNewTypedRelationships` return true on no-op round-trip.
    ///
    /// Now: iterate `relationships.imageRelationships` (the authoritative
    /// source). For each rel, try multiple path normalizations to locate the
    /// file. If the file isn't found, skip the rel rather than forge an id.
    /// Orphan files in `word/media/` not referenced by any rel are dropped
    /// (they're unreferenced anyway).
    private static func extractImages(from tempDir: URL, relationships: RelationshipsCollection) throws -> [ImageReference] {
        var images: [ImageReference] = []

        let wordDir = tempDir.appendingPathComponent("word")
        guard FileManager.default.fileExists(atPath: wordDir.path) else {
            return images
        }

        for rel in relationships.imageRelationships {
            // Try multiple normalizations: media/X, ../media/X (rels file is
            // at word/_rels/, so ../media/X → word/media/X), word/media/X.
            // Also try the rel.target verbatim relative to word/.
            let candidates: [URL] = [
                wordDir.appendingPathComponent(rel.target),
                wordDir.appendingPathComponent(rel.target.replacingOccurrences(of: "../", with: "")),
                tempDir.appendingPathComponent(rel.target),
                tempDir.appendingPathComponent(rel.target.replacingOccurrences(of: "../", with: "")),
            ]
            guard let fileURL = candidates.first(where: { FileManager.default.fileExists(atPath: $0.path) }) else {
                // Skip orphan rel — can't materialize without the actual file.
                continue
            }
            let data = try Data(contentsOf: fileURL)
            let fileName = fileURL.lastPathComponent
            let ext = (fileName as NSString).pathExtension.lowercased()

            images.append(ImageReference(
                id: rel.id,
                fileName: fileName,
                contentType: mimeType(for: ext),
                data: data
            ))
        }

        return images
    }

    /// 取得副檔名對應的 MIME 類型
    private static func mimeType(for ext: String) -> String {
        switch ext {
        case "png": return "image/png"
        case "jpg", "jpeg": return "image/jpeg"
        case "gif": return "image/gif"
        case "bmp": return "image/bmp"
        case "tiff", "tif": return "image/tiff"
        case "webp": return "image/webp"
        case "emf": return "image/x-emf"
        case "wmf": return "image/x-wmf"
        default: return "application/octet-stream"
        }
    }

    // MARK: - Body Parsing

    private static func parseBody(
        from xml: XMLDocument,
        relationships: RelationshipsCollection,
        styles: [Style],
        numbering: Numbering
    ) throws -> Body {
        var body = Body()

        // v0.13.3+ (che-word-mcp#41 + save-durability prerequisite):
        // Serial parsing always (no parallel primitives in OOXML IO).
        // Parsing determinism is load-bearing for recover_from_autosave.
        guard let bodyEl = (try xml.nodes(forXPath: "//*[local-name()='body']").first) as? XMLElement
        else { return body }

        body.children = try parseBodyChildren(
            in: bodyEl,
            relationships: relationships,
            styles: styles,
            numbering: numbering,
            collectingTablesInto: &body.tables
        )
        return body
    }

    /// Recursively parse the body-level children of an XML element. Used for
    /// `<w:body>`, `<w:sdtContent>` (block-level SDT children), and `<w:tc>`
    /// (table cell contents that may include nested block-level SDTs).
    ///
    /// Walks direct children: `<w:p>` → paragraph, `<w:tbl>` → table,
    /// `<w:sdt>` → block-level ContentControl with recursively-parsed
    /// `<w:sdtContent>` children. Other elements are skipped.
    ///
    /// The `collectingTablesInto` parameter preserves the existing
    /// `body.tables` flat list for backwards compatibility with consumers
    /// that iterate it directly.
    internal static func parseBodyChildren(
        in element: XMLElement,
        relationships: RelationshipsCollection,
        styles: [Style],
        numbering: Numbering,
        collectingTablesInto tables: inout [Table]
    ) throws -> [BodyChild] {
        var children: [BodyChild] = []
        for node in element.children ?? [] {
            guard let el = node as? XMLElement else { continue }
            switch el.localName {
            case "p":
                let paragraph = try parseParagraph(
                    from: el,
                    relationships: relationships,
                    styles: styles,
                    numbering: numbering
                )
                children.append(.paragraph(paragraph))
            case "tbl":
                let table = try parseTable(
                    from: el,
                    relationships: relationships,
                    styles: styles,
                    numbering: numbering
                )
                children.append(.table(table))
                tables.append(table)
            case "sdt":
                // v0.15.0+ (#44 task 3.4): block-level SDT wrapping body
                // children. Parse SDT metadata via SDTParser (without
                // descending into <w:sdtContent>'s children), then recurse
                // here to parse the body children inside.
                let metadata = ContentControl(
                    sdt: SDTParser.parseSdtPr(from: el),
                    content: ""
                )
                var sdtChildren: [BodyChild] = []
                if let sdtContent = el.elements(forName: "w:sdtContent").first {
                    sdtChildren = try parseBodyChildren(
                        in: sdtContent,
                        relationships: relationships,
                        styles: styles,
                        numbering: numbering,
                        collectingTablesInto: &tables
                    )
                }
                children.append(.contentControl(metadata, children: sdtChildren))
            default:
                continue
            }
        }
        return children
    }

    // MARK: - Paragraph Parsing

    /// Parse a `<w:p>` element into a `Paragraph`.
    ///
    /// Internal (not private) so unit tests in `OOXMLSwiftTests` can exercise the parser
    /// directly via `@testable import OOXMLSwift` with hand-constructed `XMLElement`
    /// instances, bypassing the full .docx ZIP read path.
    internal static func parseParagraph(
        from element: XMLElement,
        relationships: RelationshipsCollection,
        styles: [Style],
        numbering: Numbering
    ) throws -> Paragraph {
        var paragraph = Paragraph()
        let isoFormatter = ISO8601DateFormatter()

        // 解析段落屬性
        if let pPr = element.elements(forName: "w:pPr").first {
            paragraph.properties = parseParagraphProperties(from: pPr)

            // Part B: detect <w:pPrChange> inside <w:pPr> for paragraph property change tracking
            if let pPrChange = pPr.elements(forName: "w:pPrChange").first {
                let revId = Int(pPrChange.attribute(forName: "w:id")?.stringValue ?? "0") ?? 0
                let author = pPrChange.attribute(forName: "w:author")?.stringValue ?? "Unknown"
                let date = pPrChange.attribute(forName: "w:date")?.stringValue
                    .flatMap { isoFormatter.date(from: $0) } ?? Date()
                let description = Self.summarizeParagraphPropertiesXML(pPrChange.elements(forName: "w:pPr").first)
                var rev = Revision(id: revId, type: .paragraphChange, author: author,
                                   paragraphIndex: 0, date: date)
                rev.previousFormatDescription = description
                paragraph.revisions.append(rev)
            }
        }

        // 解析 Runs（包含 w:ins/w:del 追蹤修訂）。
        //
        // v0.19.0+ (PsychQuant/che-word-mcp#56) Phase 2: track a `position`
        // counter that increments by 1 per direct child element of <w:p>, so
        // `BookmarkRangeMarker.position` (and the other position-indexed
        // raw-carriers Phase 4 adds) record source-document order. This is the
        // input the Phase 4 sort-by-position emit consumes — it has no effect
        // until Paragraph.toXML() refactor lands, so behavior is additive here.
        var childPosition = 0
        for child in element.children ?? [] {
            guard let childElement = child as? XMLElement else { continue }
            defer { childPosition += 1 }

            switch childElement.localName {
            case "pPr":
                // v0.19.1+ (#56 follow-up): pPr is consumed by the dedicated
                // `paragraph.properties = parseParagraphProperties(from: pPr)`
                // call ABOVE this loop. Letting it fall through to the
                // `default` branch in v0.19.0 silently captured pPr into
                // `unrecognizedChildren` AS WELL, causing the sort-by-position
                // emit to write `<w:pPr>` twice (once via the legacy pPr block
                // at the top of `Paragraph.toXMLSortedByPosition`, once
                // verbatim from `unrecognizedChildren`). xmllint accepts the
                // duplicate but file size grows by ~1 KB per round-trip per
                // paragraph. Skip explicitly here.
                break

            case "r":
                var parsedRun = try parseRun(from: childElement, relationships: relationships)
                // v0.19.0+ (#56) Phase 4: assign source-order position so the
                // sort-by-position emit interleaves runs with bookmarks /
                // hyperlinks / fldSimple / etc. correctly.
                parsedRun.position = childPosition
                paragraph.runs.append(parsedRun)
                // Part B: detect <w:rPrChange> inside <w:rPr> for run formatting change tracking
                if let rPrChangeRev = Self.detectRPrChangeRevision(in: childElement, isoFormatter: isoFormatter) {
                    paragraph.revisions.append(rPrChangeRev)
                }

            case "ins":
                // 插入修訂：提取 w:r 並建立 Revision
                let revId = Int(childElement.attribute(forName: "w:id")?.stringValue ?? "0") ?? 0
                let author = childElement.attribute(forName: "w:author")?.stringValue ?? "Unknown"
                let date = childElement.attribute(forName: "w:date")?.stringValue
                    .flatMap { isoFormatter.date(from: $0) } ?? Date()

                // v0.19.3+ (#56 round 2 P0-7): if the wrapper carries any
                // non-`<w:r>` child (`<w:hyperlink>`, `<w:sdt>`, `<w:fldSimple>`,
                // `<mc:AlternateContent>`, etc.), capture the whole wrapper
                // verbatim into `unrecognizedChildren` at this position so the
                // sort path emits it byte-for-byte. Pre-fix the per-run loop
                // skipped the non-run children entirely — Track Changes
                // hyperlink insertions silently lost the hyperlink on round-trip.
                if Self.hasNonRunChild(childElement) {
                    paragraph.unrecognizedChildren.append(
                        UnrecognizedChild(name: "ins", rawXML: childElement.xmlString, position: childPosition)
                    )
                    break
                }

                var insertedText = ""
                for insRun in childElement.elements(forName: "w:r") {
                    var parsedRun = try parseRun(from: insRun, relationships: relationships)
                    // v0.19.2+ (#56 follow-up F3): assign source-order position so
                    // sort-by-position emit doesn't collapse all wrapper-internal
                    // runs to position=0 (paragraph front). Set revisionId so the
                    // sort path can re-wrap consecutive same-revisionId runs in
                    // <w:ins>/<w:del>/<w:moveFrom>/<w:moveTo> on emit.
                    parsedRun.position = childPosition
                    parsedRun.revisionId = revId
                    paragraph.runs.append(parsedRun)
                    insertedText += parsedRun.text
                }

                // v0.19.3+ (#56 round 2 P0-6): always create the Revision
                // entry regardless of whether inserted text is empty. Pre-fix
                // the `!insertedText.isEmpty` guard meant insertions of pure
                // non-text content (`<w:tab/>`, `<w:br/>`, `<w:drawing>`,
                // `<w:fldChar>`) parsed runs with `revisionId` but produced
                // no Revision — the sort-path grouping then fell back to a
                // naked `<w:r>` and the wrapper silently disappeared.
                paragraph.revisions.append(Revision(
                    id: revId, type: .insertion, author: author,
                    paragraphIndex: 0, originalText: nil,
                    newText: insertedText, date: date
                ))

            case "del":
                // 刪除修訂：提取 w:delText 並建立 Revision
                let revId = Int(childElement.attribute(forName: "w:id")?.stringValue ?? "0") ?? 0
                let author = childElement.attribute(forName: "w:author")?.stringValue ?? "Unknown"
                let date = childElement.attribute(forName: "w:date")?.stringValue
                    .flatMap { isoFormatter.date(from: $0) } ?? Date()

                if Self.hasNonRunChild(childElement) {
                    paragraph.unrecognizedChildren.append(
                        UnrecognizedChild(name: "del", rawXML: childElement.xmlString, position: childPosition)
                    )
                    break
                }

                var deletedText = ""
                for delRun in childElement.elements(forName: "w:r") {
                    for delText in delRun.elements(forName: "w:delText") {
                        deletedText += delText.stringValue ?? ""
                    }
                    // v0.19.2+ (#56 F3): persist run with position + revisionId
                    // so sort-by-position emit can re-wrap with <w:del>. Without
                    // this, the deletion wrapper is silently dropped on round-trip
                    // and the deletion looks accepted in Word post-save.
                    var parsedRun = try parseRun(from: delRun, relationships: relationships)
                    parsedRun.position = childPosition
                    parsedRun.revisionId = revId
                    paragraph.runs.append(parsedRun)
                }

                paragraph.revisions.append(Revision(
                    id: revId, type: .deletion, author: author,
                    paragraphIndex: 0, originalText: deletedText,
                    newText: nil, date: date
                ))

            case "moveFrom":
                // 移動來源修訂（mirrors w:del）：抽取 w:r 並建立 Revision
                let revId = Int(childElement.attribute(forName: "w:id")?.stringValue ?? "0") ?? 0
                let author = childElement.attribute(forName: "w:author")?.stringValue ?? "Unknown"
                let date = childElement.attribute(forName: "w:date")?.stringValue
                    .flatMap { isoFormatter.date(from: $0) } ?? Date()

                if Self.hasNonRunChild(childElement) {
                    paragraph.unrecognizedChildren.append(
                        UnrecognizedChild(name: "moveFrom", rawXML: childElement.xmlString, position: childPosition)
                    )
                    break
                }

                var movedText = ""
                for moveRun in childElement.elements(forName: "w:r") {
                    var parsedRun = try parseRun(from: moveRun, relationships: relationships)
                    // v0.19.2+ (#56 F3): position + revisionId for sort path wrapping
                    parsedRun.position = childPosition
                    parsedRun.revisionId = revId
                    paragraph.runs.append(parsedRun)
                    movedText += parsedRun.text
                }

                paragraph.revisions.append(Revision(
                    id: revId, type: .moveFrom, author: author,
                    paragraphIndex: 0, originalText: movedText,
                    newText: nil, date: date
                ))

            case "moveTo":
                // 移動目標修訂（mirrors w:ins）：抽取 w:r 並建立 Revision
                let revId = Int(childElement.attribute(forName: "w:id")?.stringValue ?? "0") ?? 0
                let author = childElement.attribute(forName: "w:author")?.stringValue ?? "Unknown"
                let date = childElement.attribute(forName: "w:date")?.stringValue
                    .flatMap { isoFormatter.date(from: $0) } ?? Date()

                if Self.hasNonRunChild(childElement) {
                    paragraph.unrecognizedChildren.append(
                        UnrecognizedChild(name: "moveTo", rawXML: childElement.xmlString, position: childPosition)
                    )
                    break
                }

                var movedText = ""
                for moveRun in childElement.elements(forName: "w:r") {
                    var parsedRun = try parseRun(from: moveRun, relationships: relationships)
                    // v0.19.2+ (#56 F3): position + revisionId for sort path wrapping
                    parsedRun.position = childPosition
                    parsedRun.revisionId = revId
                    paragraph.runs.append(parsedRun)
                    movedText += parsedRun.text
                }

                paragraph.revisions.append(Revision(
                    id: revId, type: .moveTo, author: author,
                    paragraphIndex: 0, originalText: nil,
                    newText: movedText, date: date
                ))

            case "commentRangeStart":
                if let idStr = childElement.attribute(forName: "w:id")?.stringValue,
                   let id = Int(idStr) {
                    paragraph.commentIds.append(id)
                    // v0.19.0+ (#56) Phase 4: also populate the position-indexed
                    // marker so sort-by-position emit can re-emit at original offset.
                    paragraph.commentRangeMarkers.append(
                        CommentRangeMarker(kind: .start, id: id, position: childPosition)
                    )
                }

            case "commentRangeEnd":
                // v0.19.0+ (#56) Phase 4: previously dropped silently. Now
                // captured as a position-indexed marker for round-trip preservation.
                if let idStr = childElement.attribute(forName: "w:id")?.stringValue,
                   let id = Int(idStr) {
                    paragraph.commentRangeMarkers.append(
                        CommentRangeMarker(kind: .end, id: id, position: childPosition)
                    )
                }

            case "permStart":
                // v0.19.0+ (#56) Phase 4: editor-permission gate start.
                let id = childElement.attribute(forName: "w:id")?.stringValue ?? ""
                let editorGroup = childElement.attribute(forName: "w:edGrp")?.stringValue
                let editor = childElement.attribute(forName: "w:ed")?.stringValue
                paragraph.permissionRangeMarkers.append(
                    PermissionRangeMarker(
                        kind: .start, id: id,
                        editorGroup: editorGroup, editor: editor,
                        position: childPosition
                    )
                )

            case "permEnd":
                let id = childElement.attribute(forName: "w:id")?.stringValue ?? ""
                paragraph.permissionRangeMarkers.append(
                    PermissionRangeMarker(kind: .end, id: id, position: childPosition)
                )

            case "proofErr":
                // v0.19.0+ (#56) Phase 4: proof error markers (spelling /
                // grammar). Source attribute `w:type` distinguishes start / end
                // and spell / grammar.
                if let typeStr = childElement.attribute(forName: "w:type")?.stringValue,
                   let errType = ProofErrorMarker.ErrorType(rawValue: typeStr) {
                    paragraph.proofErrorMarkers.append(
                        ProofErrorMarker(type: errType, position: childPosition)
                    )
                }

            case "smartTag":
                // v0.19.0+ (#56) Phase 4: smart tag raw-carrier. Stored verbatim
                // because its <w:smartTagPr> child surface is vendor-specific.
                paragraph.smartTags.append(
                    SmartTagBlock(rawXML: childElement.xmlString, position: childPosition)
                )

            case "customXml":
                // v0.19.0+ (#56) Phase 4: custom XML wrapper raw-carrier.
                paragraph.customXmlBlocks.append(
                    CustomXmlBlock(rawXML: childElement.xmlString, position: childPosition)
                )

            case "dir":
                // v0.19.0+ (#56) Phase 4: directional override (RTL / LTR) wrapper.
                paragraph.bidiOverrides.append(
                    BidiOverrideBlock(element: .dir, rawXML: childElement.xmlString, position: childPosition)
                )

            case "bdo":
                // v0.19.0+ (#56) Phase 4: bidirectional override wrapper.
                paragraph.bidiOverrides.append(
                    BidiOverrideBlock(element: .bdo, rawXML: childElement.xmlString, position: childPosition)
                )

            case "bookmarkStart":
                // v0.19.0+ (PsychQuant/che-word-mcp#56) Phase 2: source-side
                // bookmark population. Each <w:bookmarkStart w:id w:name/>
                // produces both a typed Bookmark on `paragraph.bookmarks`
                // (legacy model used by 218 MCP tools) and a position-indexed
                // BookmarkRangeMarker on `paragraph.bookmarkMarkers` so Phase 4
                // sort-by-position emit can re-emit at original relative offset.
                if let idStr = childElement.attribute(forName: "w:id")?.stringValue,
                   let id = Int(idStr),
                   let name = childElement.attribute(forName: "w:name")?.stringValue {
                    paragraph.bookmarks.append(Bookmark(id: id, name: name))
                    paragraph.bookmarkMarkers.append(
                        BookmarkRangeMarker(kind: .start, id: id, position: childPosition)
                    )
                }

            case "bookmarkEnd":
                if let idStr = childElement.attribute(forName: "w:id")?.stringValue,
                   let id = Int(idStr) {
                    paragraph.bookmarkMarkers.append(
                        BookmarkRangeMarker(kind: .end, id: id, position: childPosition)
                    )
                }

            case "sdt":
                // v0.15.0+ (#44 task 3.1): paragraph-level SDT becomes a
                // first-class ContentControl on Paragraph.contentControls
                // (sibling of runs), not a Run.rawXML blob. See SDD
                // `che-word-mcp-content-controls-read-write`.
                //
                // v0.19.4+ (#56 R3-NEW-2): pass childPosition so the SDT
                // round-trips at its source position via
                // `Paragraph.toXMLSortedByPosition` (otherwise the post-content
                // legacy emit forces it to end-of-paragraph).
                paragraph.contentControls.append(
                    SDTParser.parseSDT(from: childElement, position: childPosition)
                )

            case "hyperlink":
                // v0.19.0+ (PsychQuant/che-word-mcp#56) Phase 3: parse
                // <w:hyperlink> wrappers as typed Hyperlink instances. Inner
                // <w:r> children populate `runs` (typed editable surface so
                // tool-mediated `replace_text` / `format_text` find content
                // inside hyperlinks); unrecognized attributes / non-Run
                // children flow into `rawAttributes` / `rawChildren` for
                // byte-level survival.
                let parsedHyperlink = try parseHyperlink(
                    from: childElement,
                    relationships: relationships,
                    position: childPosition
                )
                paragraph.hyperlinks.append(parsedHyperlink)

            case "fldSimple":
                // v0.19.0+ (#56) Phase 3: typed FieldSimple model so SEQ Table
                // captions / REF cross-references / TOC entries are addressable
                // by `replace_text` / `format_text`.
                let parsedField = try parseFieldSimple(
                    from: childElement,
                    relationships: relationships,
                    position: childPosition
                )
                paragraph.fieldSimples.append(parsedField)

            case "AlternateContent":
                // v0.19.0+ (#56) Phase 3: <mc:AlternateContent> hybrid model.
                // Captures verbatim raw XML for byte-equivalent emit and
                // extracts <mc:Fallback> runs as typed surface for tool reads.
                let parsedAC = try parseAlternateContent(
                    from: childElement,
                    relationships: relationships,
                    position: childPosition
                )
                paragraph.alternateContents.append(parsedAC)

            default:
                // v0.19.0+ (#56) Phase 4: every unrecognized <w:p> child is
                // captured verbatim with its source position so the round-trip
                // suite can XCTFail with the element name (per design decision
                // "ECMA-376 <w:p> schema as the completeness checklist").
                // Without this, dropped elements would silently disappear and
                // the test would have no way to surface the gap.
                let name = childElement.localName ?? "<nil>"
                paragraph.unrecognizedChildren.append(
                    UnrecognizedChild(
                        name: name,
                        rawXML: childElement.xmlString,
                        position: childPosition
                    )
                )
                if DocxReader.debugLoggingEnabled {
                    let line = "DocxReader.parseParagraph: captured unmodeled element \(name) at position \(childPosition)\n"
                    if let data = line.data(using: .utf8) {
                        FileHandle.standardError.write(data)
                    }
                }
            }
        }

        // 🆕 語義標註
        paragraph.semantic = detectParagraphSemantic(
            properties: paragraph.properties,
            runs: paragraph.runs,
            styles: styles,
            numbering: numbering
        )

        return paragraph
    }

    private static func parseParagraphProperties(from element: XMLElement) -> ParagraphProperties {
        var props = ParagraphProperties()

        // 樣式
        if let pStyle = element.elements(forName: "w:pStyle").first,
           let val = pStyle.attribute(forName: "w:val")?.stringValue {
            props.style = val
        }

        // 對齊
        if let jc = element.elements(forName: "w:jc").first,
           let val = jc.attribute(forName: "w:val")?.stringValue {
            props.alignment = Alignment(rawValue: val)
        }

        // 間距
        if let spacing = element.elements(forName: "w:spacing").first {
            var spacingProps = Spacing()
            if let before = spacing.attribute(forName: "w:before")?.stringValue {
                spacingProps.before = Int(before)
            }
            if let after = spacing.attribute(forName: "w:after")?.stringValue {
                spacingProps.after = Int(after)
            }
            if let line = spacing.attribute(forName: "w:line")?.stringValue {
                spacingProps.line = Int(line)
            }
            if let lineRule = spacing.attribute(forName: "w:lineRule")?.stringValue {
                spacingProps.lineRule = LineRule(rawValue: lineRule)
            }
            props.spacing = spacingProps
        }

        // 縮排
        if let ind = element.elements(forName: "w:ind").first {
            var indentation = Indentation()
            if let left = ind.attribute(forName: "w:left")?.stringValue {
                indentation.left = Int(left)
            }
            if let right = ind.attribute(forName: "w:right")?.stringValue {
                indentation.right = Int(right)
            }
            if let firstLine = ind.attribute(forName: "w:firstLine")?.stringValue {
                indentation.firstLine = Int(firstLine)
            }
            if let hanging = ind.attribute(forName: "w:hanging")?.stringValue {
                indentation.hanging = Int(hanging)
            }
            props.indentation = indentation
        }

        // 編號/項目符號 (w:numPr)
        if let numPr = element.elements(forName: "w:numPr").first {
            var numInfo: NumberingInfo?
            var numId: Int?
            var level: Int = 0

            if let ilvl = numPr.elements(forName: "w:ilvl").first,
               let val = ilvl.attribute(forName: "w:val")?.stringValue {
                level = Int(val) ?? 0
            }
            if let numIdEl = numPr.elements(forName: "w:numId").first,
               let val = numIdEl.attribute(forName: "w:val")?.stringValue {
                numId = Int(val)
            }

            if let id = numId {
                numInfo = NumberingInfo(numId: id, level: level)
            }
            props.numbering = numInfo
        }

        // 分頁控制
        if element.elements(forName: "w:keepNext").first != nil {
            props.keepNext = true
        }
        if element.elements(forName: "w:keepLines").first != nil {
            props.keepLines = true
        }
        if element.elements(forName: "w:pageBreakBefore").first != nil {
            props.pageBreakBefore = true
        }

        return props
    }

    // MARK: - Container Parsing Helpers (Part C of ooxml-swift#1)

    /// Parse `<w:p>` children directly under an XML document's root element
    /// (used for headers `<w:hdr>` and footers `<w:ftr>` which are paragraph containers).
    private static func parseContainerParagraphs(
        from xml: XMLDocument,
        relationships: RelationshipsCollection,
        styles: [Style],
        numbering: Numbering
    ) throws -> [Paragraph] {
        guard let root = xml.rootElement() else { return [] }
        return try parseContainerChildParagraphs(
            in: root,
            relationships: relationships, styles: styles, numbering: numbering
        )
    }

    /// Parse `<w:p>` children inside a given XML element
    /// (used for footnote/endnote individual entries that contain paragraphs).
    private static func parseContainerChildParagraphs(
        in element: XMLElement,
        relationships: RelationshipsCollection,
        styles: [Style],
        numbering: Numbering
    ) throws -> [Paragraph] {
        var paragraphs: [Paragraph] = []
        for child in element.children ?? [] {
            guard let childElement = child as? XMLElement, childElement.localName == "p" else { continue }
            let para = try parseParagraph(
                from: childElement,
                relationships: relationships,
                styles: styles,
                numbering: numbering
            )
            paragraphs.append(para)
        }
        return paragraphs
    }

    // MARK: - Nested Revision Helpers (Part B of ooxml-swift#1)

    /// Detect `<w:rPrChange>` inside a `<w:r>` element's `<w:rPr>` and return a `Revision`
    /// if found. Returns `nil` if no change-tracking element is present.
    private static func detectRPrChangeRevision(in runElement: XMLElement, isoFormatter: ISO8601DateFormatter) -> Revision? {
        guard
            let rPr = runElement.elements(forName: "w:rPr").first,
            let rPrChange = rPr.elements(forName: "w:rPrChange").first
        else { return nil }

        let revId = Int(rPrChange.attribute(forName: "w:id")?.stringValue ?? "0") ?? 0
        let author = rPrChange.attribute(forName: "w:author")?.stringValue ?? "Unknown"
        let date = rPrChange.attribute(forName: "w:date")?.stringValue
            .flatMap { isoFormatter.date(from: $0) } ?? Date()
        let description = summarizeRunPropertiesXML(rPrChange.elements(forName: "w:rPr").first)
        let priorProps = rPrChange.elements(forName: "w:rPr").first.map { parseRunProperties(from: $0) }

        var rev = Revision(id: revId, type: .formatChange, author: author,
                           paragraphIndex: 0, date: date)
        rev.previousFormatDescription = description
        rev.previousFormat = priorProps
        return rev
    }

    /// Human-readable summary of run formatting from a `<w:rPr>` XML element.
    private static func summarizeRunPropertiesXML(_ element: XMLElement?) -> String {
        guard let el = element else { return "no prior formatting" }
        var parts: [String] = []
        if el.elements(forName: "w:b").first != nil { parts.append("bold") }
        if el.elements(forName: "w:i").first != nil { parts.append("italic") }
        if el.elements(forName: "w:u").first != nil { parts.append("underline") }
        if el.elements(forName: "w:strike").first != nil { parts.append("strikethrough") }
        if let sz = el.elements(forName: "w:sz").first?.attribute(forName: "w:val")?.stringValue,
           let halfPt = Int(sz) {
            parts.append("\(halfPt / 2)pt")
        }
        if let font = el.elements(forName: "w:rFonts").first?.attribute(forName: "w:ascii")?.stringValue {
            parts.append(font)
        }
        if let color = el.elements(forName: "w:color").first?.attribute(forName: "w:val")?.stringValue {
            parts.append("color:\(color)")
        }
        return parts.isEmpty ? "no prior formatting" : parts.joined(separator: ", ")
    }

    /// Human-readable summary of paragraph formatting from a `<w:pPr>` XML element.
    private static func summarizeParagraphPropertiesXML(_ element: XMLElement?) -> String {
        guard let el = element else { return "no prior formatting" }
        var parts: [String] = []
        if let jc = el.elements(forName: "w:jc").first?.attribute(forName: "w:val")?.stringValue {
            parts.append("alignment: \(jc)")
        }
        if let spacing = el.elements(forName: "w:spacing").first {
            if let before = spacing.attribute(forName: "w:before")?.stringValue {
                parts.append("spacing-before: \(before)")
            }
            if let after = spacing.attribute(forName: "w:after")?.stringValue {
                parts.append("spacing-after: \(after)")
            }
        }
        if let ind = el.elements(forName: "w:ind").first {
            if let left = ind.attribute(forName: "w:left")?.stringValue {
                parts.append("indent-left: \(left)")
            }
        }
        if let pStyle = el.elements(forName: "w:pStyle").first?.attribute(forName: "w:val")?.stringValue {
            parts.append("style: \(pStyle)")
        }
        return parts.isEmpty ? "no prior formatting" : parts.joined(separator: ", ")
    }

    // MARK: - Run Parsing

    /// v0.19.0+ (PsychQuant/che-word-mcp#56) Phase 3: parse a `<w:hyperlink>`
    /// element into the hybrid `Hyperlink` model. Inner `<w:r>` children become
    /// typed Runs (so MCP tools can edit text inside hyperlinks). Recognized
    /// attributes (`r:id`, `w:anchor`, `w:tooltip`, `w:history`, `w:tgtFrame`,
    /// `w:docLocation`) populate typed fields. Anything else is captured by
    /// the raw passthrough fields so a no-op round-trip preserves the wrapper
    /// byte-equivalent.
    internal static func parseHyperlink(
        from element: XMLElement,
        relationships: RelationshipsCollection,
        position: Int
    ) throws -> Hyperlink {
        // v0.19.3+ (#56 round 2 P0-2): only attributes with a typed `Hyperlink`
        // field belong here. Removed `w:tgtFrame` and `w:docLocation` because
        // the model has no typed surface for them and `toXML()` doesn't emit
        // them — leaving them in `recognizedAttrs` silently dropped vendor /
        // browser-target attributes on round-trip. They now flow into
        // `rawAttributes` and the writer emits them via the alphabetical loop.
        let recognizedAttrs: Set<String> = [
            "r:id", "w:anchor", "w:tooltip", "w:history",
        ]

        let rId = element.attribute(forName: "r:id")?.stringValue
        let anchor = element.attribute(forName: "w:anchor")?.stringValue
        let tooltip = element.attribute(forName: "w:tooltip")?.stringValue
        // `w:history` defaults true; only `"0"` flips it false (per Hyperlink doc).
        let historyAttr = element.attribute(forName: "w:history")?.stringValue
        let history = (historyAttr != "0")

        var rawAttributes: [String: String] = [:]
        for attr in element.attributes ?? [] {
            guard let name = attr.name, !recognizedAttrs.contains(name) else { continue }
            rawAttributes[name] = attr.stringValue ?? ""
        }

        // v0.19.3+ (#56 round 2 P0-3): walk children once, building both the
        // ordered `children` list (source of truth for the writer) AND the
        // legacy `runs` / `rawChildren` projections (kept for backward-compat
        // reads from existing callers that still iterate the typed lists).
        var runs: [Run] = []
        var rawChildren: [String] = []
        var children: [HyperlinkChild] = []
        for child in element.children ?? [] {
            guard let childElement = child as? XMLElement else { continue }
            if childElement.localName == "r" {
                let run = try parseRun(from: childElement, relationships: relationships)
                runs.append(run)
                children.append(.run(run))
            } else {
                let raw = childElement.xmlString
                rawChildren.append(raw)
                children.append(.rawXML(raw))
            }
        }

        // For external hyperlinks, resolve the URL via the rels collection so
        // downstream consumers (e.g., visualization, link audit tools) can read
        // `Hyperlink.url` without separately joining against rels.
        var url: String? = nil
        if let rId = rId {
            url = relationships.relationships.first(where: { $0.id == rId })?.target
        }

        // v0.19.3+ (#56 round 2 P1-7): allocate a unique id by appending the
        // source position. Pre-fix `id = rId ?? anchor ?? "hl-\(position)"`
        // returned the same id when two hyperlinks shared the same `r:id`
        // (legitimate when two anchors target the same URL via one rels entry),
        // breaking MCP tools that find / edit / delete hyperlinks by id.
        // Format: `<rId-or-anchor-or-hl>@<position>` so the human-readable
        // prefix survives for debugging and the suffix guarantees uniqueness.
        let idPrefix = rId ?? anchor ?? "hl"
        let id = "\(idPrefix)@\(position)"

        return Hyperlink(
            id: id,
            runs: runs,
            relationshipId: rId,
            anchor: anchor,
            url: url,
            tooltip: tooltip,
            history: history,
            rawAttributes: rawAttributes,
            rawChildren: rawChildren,
            children: children,
            position: position
        )
    }

    /// v0.19.0+ (PsychQuant/che-word-mcp#56) Phase 3: parse a `<w:fldSimple>`
    /// element into the typed `FieldSimple` model. `w:instr` whitespace is
    /// preserved exactly so existing field-recalc tools that match on the raw
    /// instruction string continue to work. Inner `<w:r>` children populate
    /// `runs` so `replace_text` / `format_text` can edit the rendered field
    /// result. Unrecognized attributes flow into `rawAttributes`.
    internal static func parseFieldSimple(
        from element: XMLElement,
        relationships: RelationshipsCollection,
        position: Int
    ) throws -> FieldSimple {
        let recognizedAttrs: Set<String> = ["w:instr", "w:fldLock", "w:dirty"]

        let instr = element.attribute(forName: "w:instr")?.stringValue ?? ""

        var rawAttributes: [String: String] = [:]
        for attr in element.attributes ?? [] {
            guard let name = attr.name, !recognizedAttrs.contains(name) else { continue }
            rawAttributes[name] = attr.stringValue ?? ""
        }

        var runs: [Run] = []
        for child in element.children ?? [] {
            guard let childElement = child as? XMLElement else { continue }
            if childElement.localName == "r" {
                runs.append(try parseRun(from: childElement, relationships: relationships))
            }
        }

        return FieldSimple(
            instr: instr,
            runs: runs,
            rawAttributes: rawAttributes,
            position: position
        )
    }

    /// v0.19.0+ (PsychQuant/che-word-mcp#56) Phase 3: parse an
    /// `<mc:AlternateContent>` block into the hybrid `AlternateContent` model.
    /// `rawXML` captures the verbatim source XML (so Writer emits byte-equivalent
    /// output and `<mc:Choice>` content is preserved without typed modeling);
    /// `fallbackRuns` extracts `<w:r>` children inside `<mc:Fallback>` as
    /// typed Runs for tool-mediated read access (e.g., reporting math
    /// transliterations or letting a future SDD apply edits).
    internal static func parseAlternateContent(
        from element: XMLElement,
        relationships: RelationshipsCollection,
        position: Int
    ) throws -> AlternateContent {
        // Capture verbatim XML via XMLNode.xmlString (Foundation guarantees this
        // is well-formed XML even though it may differ from the source bytes by
        // canonicalization). For Phase 3 this is acceptable; if we ever need
        // strict byte equality, callers can mark the part dirty and the original
        // ZIP entry stays preserved-by-default via overlay mode.
        let rawXML = element.xmlString

        // Walk children to find <mc:Fallback> and extract its <w:r> children.
        var fallbackRuns: [Run] = []
        for child in element.children ?? [] {
            guard let childElement = child as? XMLElement else { continue }
            if childElement.localName == "Fallback" {
                for grandchild in childElement.children ?? [] {
                    guard let runElement = grandchild as? XMLElement else { continue }
                    if runElement.localName == "r" {
                        fallbackRuns.append(try parseRun(from: runElement, relationships: relationships))
                    }
                }
            }
        }

        return AlternateContent(rawXML: rawXML, fallbackRuns: fallbackRuns, position: position)
    }

    /// Parse a `<w:r>` element into a `Run`. Internal access (was private)
    /// so `@testable import OOXMLSwift` consumers can drive parser directly.
    internal static func parseRun(from element: XMLElement, relationships: RelationshipsCollection) throws -> Run {
        var run = Run(text: "")

        // 解析 Run 屬性
        if let rPr = element.elements(forName: "w:rPr").first {
            run.properties = parseRunProperties(from: rPr)
        }

        // 解析文字
        for t in element.elements(forName: "w:t") {
            run.text += t.stringValue ?? ""
        }

        // 解析圖片 (w:drawing)
        if let drawingElement = element.elements(forName: "w:drawing").first {
            run.drawing = try parseDrawing(from: drawingElement, relationships: relationships)
            // 🆕 圖片語義標註（標為 unknown，等後續分類）
            run.semantic = SemanticAnnotation.unknownImage
        }

        // 🆕 檢查是否為 OMML 公式 (m:oMath 或 m:oMathPara)
        // 使用 children 遍歷取代 XPath，避免 O(n²) 效能問題
        if let oMathElement = findFirstDescendant(of: element, localNames: ["oMath", "oMathPara"]) {
            run.rawXML = oMathElement.xmlString
            run.semantic = SemanticAnnotation.ommlFormula
        }

        // v0.14.0+ (che-word-mcp#52): preserve unknown direct children of <w:r>
        // (e.g., <w:pict> VML watermarks, <w:object> OLE embeds, <w:ruby>).
        // Recognized typed kinds are skipped because they're already captured
        // into typed fields above. Source-document order is preserved by
        // walking children sequentially.
        let recognizedRunChildren: Set<String> = ["rPr", "t", "drawing", "oMath", "oMathPara"]
        var collectedRawElements: [RawElement] = []
        for child in element.children ?? [] {
            guard let childElement = child as? XMLElement,
                  let localName = childElement.localName,
                  !recognizedRunChildren.contains(localName) else {
                continue
            }
            collectedRawElements.append(
                RawElement(name: localName, xml: childElement.xmlString)
            )
        }
        if !collectedRawElements.isEmpty {
            run.rawElements = collectedRawElements
        }

        return run
    }

    // MARK: - Drawing Parsing

    /// 解析 <w:drawing> 元素
    private static func parseDrawing(from element: XMLElement, relationships: RelationshipsCollection) throws -> Drawing? {
        // 尋找 inline 或 anchor 元素（使用 children 遍歷取代 XPath）
        if let inlineElement = findFirstDescendant(of: element, localNames: ["inline"]) {
            return try parseInlineDrawing(from: inlineElement, relationships: relationships)
        } else if let anchorElement = findFirstDescendant(of: element, localNames: ["anchor"]) {
            return try parseAnchorDrawing(from: anchorElement, relationships: relationships)
        }

        return nil
    }

    /// 解析 inline drawing
    private static func parseInlineDrawing(from element: XMLElement, relationships: RelationshipsCollection) throws -> Drawing? {
        // 取得尺寸 (wp:extent)
        guard let extentElement = findFirstDescendant(of: element, localNames: ["extent"]),
              let cxStr = extentElement.attribute(forName: "cx")?.stringValue,
              let cyStr = extentElement.attribute(forName: "cy")?.stringValue,
              let cx = Int(cxStr),
              let cy = Int(cyStr) else {
            return nil
        }

        // 取得圖片參照 (a:blip r:embed)
        guard let blipElement = findFirstDescendant(of: element, localNames: ["blip"]) else {
            return nil
        }

        // r:embed 屬性包含 relationship ID
        let embedId = blipElement.attribute(forName: "r:embed")?.stringValue
            ?? blipElement.attribute(forName: "embed")?.stringValue

        guard let imageId = embedId else {
            return nil
        }

        // 取得圖片名稱和描述 (wp:docPr)
        var name = "Picture"
        var description = ""

        if let docPrElement = findFirstDescendant(of: element, localNames: ["docPr"]) {
            if let nameAttr = docPrElement.attribute(forName: "name")?.stringValue {
                name = nameAttr
            }
            if let descrAttr = docPrElement.attribute(forName: "descr")?.stringValue {
                description = descrAttr
            }
        }

        let drawing = Drawing(
            type: .inline,
            width: cx,
            height: cy,
            imageId: imageId,
            name: name,
            description: description
        )

        return drawing
    }

    /// 解析 anchor drawing (浮動圖片)
    private static func parseAnchorDrawing(from element: XMLElement, relationships: RelationshipsCollection) throws -> Drawing? {
        // 取得尺寸
        guard let extentElement = findFirstDescendant(of: element, localNames: ["extent"]),
              let cxStr = extentElement.attribute(forName: "cx")?.stringValue,
              let cyStr = extentElement.attribute(forName: "cy")?.stringValue,
              let cx = Int(cxStr),
              let cy = Int(cyStr) else {
            return nil
        }

        // 取得圖片參照
        guard let blipElement = findFirstDescendant(of: element, localNames: ["blip"]) else {
            return nil
        }

        let embedId = blipElement.attribute(forName: "r:embed")?.stringValue
            ?? blipElement.attribute(forName: "embed")?.stringValue

        guard let imageId = embedId else {
            return nil
        }

        // 取得名稱和描述
        var name = "Picture"
        var description = ""

        if let docPrElement = findFirstDescendant(of: element, localNames: ["docPr"]) {
            if let nameAttr = docPrElement.attribute(forName: "name")?.stringValue {
                name = nameAttr
            }
            if let descrAttr = docPrElement.attribute(forName: "descr")?.stringValue {
                description = descrAttr
            }
        }

        var drawing = Drawing(
            type: .anchor,
            width: cx,
            height: cy,
            imageId: imageId,
            name: name,
            description: description
        )

        // 解析定位屬性
        var anchorPos = AnchorPosition()

        // 水平定位
        if let posHElement = findFirstDescendant(of: element, localNames: ["positionH"]) {
            if let relativeFrom = posHElement.attribute(forName: "relativeFrom")?.stringValue {
                anchorPos.horizontalRelativeFrom = HorizontalRelativeFrom(rawValue: relativeFrom) ?? .column
            }

            // posOffset 或 align
            let hOffsetElement = findFirstDescendant(of: posHElement, localNames: ["posOffset"])
            let hAlignElement = findFirstDescendant(of: posHElement, localNames: ["align"])

            if let offsetEl = hOffsetElement, let offsetStr = offsetEl.stringValue, let offset = Int(offsetStr) {
                anchorPos.horizontalOffset = offset
            } else if let alignEl = hAlignElement, let alignStr = alignEl.stringValue {
                anchorPos.horizontalAlignment = HorizontalAlignment(rawValue: alignStr)
            }
        }

        // 垂直定位
        if let posVElement = findFirstDescendant(of: element, localNames: ["positionV"]) {
            if let relativeFrom = posVElement.attribute(forName: "relativeFrom")?.stringValue {
                anchorPos.verticalRelativeFrom = VerticalRelativeFrom(rawValue: relativeFrom) ?? .paragraph
            }

            let offsetElement = findFirstDescendant(of: posVElement, localNames: ["posOffset"])
            let alignElement = findFirstDescendant(of: posVElement, localNames: ["align"])

            if let offsetEl = offsetElement, let offsetStr = offsetEl.stringValue, let offset = Int(offsetStr) {
                anchorPos.verticalOffset = offset
            } else if let alignEl = alignElement, let alignStr = alignEl.stringValue {
                anchorPos.verticalAlignment = VerticalAlignment(rawValue: alignStr)
            }
        }

        drawing.anchorPosition = anchorPos

        return drawing
    }

    /// v0.19.3+ (#56 round 2 P0-7): returns true when a revision wrapper
    /// (`<w:ins>`/`<w:del>`/`<w:moveFrom>`/`<w:moveTo>`) contains any direct
    /// child whose local name is NOT `w:r`. Such wrappers can hold nested
    /// `<w:hyperlink>`, `<w:sdt>`, `<w:fldSimple>`, `<mc:AlternateContent>`
    /// etc. (Word's Track Changes "user inserted a hyperlink" emit shape).
    /// The parseParagraph caller routes mixed-content wrappers to verbatim
    /// raw-carrier capture so the wrapper round-trips byte-equivalent
    /// instead of dropping the non-run children.
    private static func hasNonRunChild(_ element: XMLElement) -> Bool {
        for child in element.children ?? [] {
            guard let childElement = child as? XMLElement else { continue }
            if childElement.localName != "r" {
                return true
            }
        }
        return false
    }

    private static func parseRunProperties(from element: XMLElement) -> RunProperties {
        var props = RunProperties()

        // v0.19.3+ (#56 round 2 P0-1): rStyle reference (e.g., "Hyperlink"
        // for hyperlink-styled runs). Source-loaded runs preserve their style
        // name through round-trip; API-built hyperlinks set this so Word
        // applies the Hyperlink character style (blue + underline).
        if let rStyle = element.elements(forName: "w:rStyle").first,
           let val = rStyle.attribute(forName: "w:val")?.stringValue {
            props.rStyle = val
        }

        // 粗體
        if element.elements(forName: "w:b").first != nil {
            props.bold = true
        }

        // 斜體
        if element.elements(forName: "w:i").first != nil {
            props.italic = true
        }

        // 底線
        if let u = element.elements(forName: "w:u").first,
           let val = u.attribute(forName: "w:val")?.stringValue {
            props.underline = UnderlineType(rawValue: val)
        }

        // 刪除線
        if element.elements(forName: "w:strike").first != nil {
            props.strikethrough = true
        }

        // 字型大小
        if let sz = element.elements(forName: "w:sz").first,
           let val = sz.attribute(forName: "w:val")?.stringValue {
            props.fontSize = Int(val)
        }

        // 字型
        if let rFonts = element.elements(forName: "w:rFonts").first,
           let ascii = rFonts.attribute(forName: "w:ascii")?.stringValue {
            props.fontName = ascii
        }

        // 顏色
        if let color = element.elements(forName: "w:color").first,
           let val = color.attribute(forName: "w:val")?.stringValue {
            props.color = val
        }

        // 螢光標記
        if let highlight = element.elements(forName: "w:highlight").first,
           let val = highlight.attribute(forName: "w:val")?.stringValue {
            props.highlight = HighlightColor(rawValue: val)
        }

        // 垂直對齊
        if let vertAlign = element.elements(forName: "w:vertAlign").first,
           let val = vertAlign.attribute(forName: "w:val")?.stringValue {
            props.verticalAlign = VerticalAlign(rawValue: val)
        }

        return props
    }

    // MARK: - Table Parsing

    /// v0.17.0+ (#49): max table nesting depth — beyond this we throw rather
    /// than risk OOM on malformed input. Matches Word's own internal threshold.
    static let MAX_TABLE_NEST_DEPTH = 5

    private static func parseTable(
        from element: XMLElement,
        relationships: RelationshipsCollection,
        styles: [Style],
        numbering: Numbering,
        depth: Int = 0
    ) throws -> Table {
        guard depth <= MAX_TABLE_NEST_DEPTH else {
            throw WordError.invalidDocx("nested table depth exceeds \(MAX_TABLE_NEST_DEPTH)")
        }

        var table = Table()

        // 解析表格屬性 (extended in v0.17.0+ to also pick up tblInd / explicit
        // layout / conditional styles)
        if let tblPr = element.elements(forName: "w:tblPr").first {
            table.properties = parseTableProperties(from: tblPr)
            // tblInd
            if let tblInd = tblPr.elements(forName: "w:tblInd").first,
               let w = tblInd.attribute(forName: "w:w")?.stringValue,
               let value = Int(w) {
                table.tableIndent = value
            }
            // explicit layout (separate from properties.layout for round-trip clarity)
            if let layout = tblPr.elements(forName: "w:tblLayout").first,
               let val = layout.attribute(forName: "w:type")?.stringValue,
               let lay = TableLayout(rawValue: val) {
                table.explicitLayout = lay
            }
            // conditional styles
            for stylePr in tblPr.elements(forName: "w:tblStylePr") {
                guard let typeStr = stylePr.attribute(forName: "w:type")?.stringValue,
                      let type = TableConditionalStyleType(rawValue: typeStr)
                else { continue }
                var props = TableConditionalStyleProperties()
                if let rPr = stylePr.elements(forName: "w:rPr").first {
                    if rPr.elements(forName: "w:b").first != nil { props.bold = true }
                    if rPr.elements(forName: "w:i").first != nil { props.italic = true }
                    if let c = rPr.elements(forName: "w:color").first?.attribute(forName: "w:val")?.stringValue {
                        props.color = c
                    }
                    if let szStr = rPr.elements(forName: "w:sz").first?.attribute(forName: "w:val")?.stringValue,
                       let sz = Int(szStr) {
                        props.fontSize = sz
                    }
                }
                if let tcPr = stylePr.elements(forName: "w:tcPr").first,
                   let bg = tcPr.elements(forName: "w:shd").first?.attribute(forName: "w:fill")?.stringValue {
                    props.backgroundColor = bg
                }
                table.conditionalStyles.append(TableConditionalStyle(type: type, properties: props))
            }
        }

        // 解析表格行
        for tr in element.elements(forName: "w:tr") {
            let row = try parseTableRow(
                from: tr,
                relationships: relationships,
                styles: styles,
                numbering: numbering,
                depth: depth
            )
            table.rows.append(row)
        }

        return table
    }

    private static func parseTableProperties(from element: XMLElement) -> TableProperties {
        var props = TableProperties()

        // 寬度
        if let tblW = element.elements(forName: "w:tblW").first {
            if let w = tblW.attribute(forName: "w:w")?.stringValue {
                props.width = Int(w)
            }
            if let type = tblW.attribute(forName: "w:type")?.stringValue {
                props.widthType = WidthType(rawValue: type)
            }
        }

        // 對齊
        if let jc = element.elements(forName: "w:jc").first,
           let val = jc.attribute(forName: "w:val")?.stringValue {
            props.alignment = Alignment(rawValue: val)
        }

        // 版面配置
        if let layout = element.elements(forName: "w:tblLayout").first,
           let val = layout.attribute(forName: "w:type")?.stringValue {
            props.layout = TableLayout(rawValue: val)
        }

        return props
    }

    private static func parseTableRow(
        from element: XMLElement,
        relationships: RelationshipsCollection,
        styles: [Style],
        numbering: Numbering,
        depth: Int = 0
    ) throws -> TableRow {
        var row = TableRow()

        // 解析行屬性
        if let trPr = element.elements(forName: "w:trPr").first {
            row.properties = parseTableRowProperties(from: trPr)
        }

        // 解析儲存格
        for tc in element.elements(forName: "w:tc") {
            let cell = try parseTableCell(
                from: tc,
                relationships: relationships,
                styles: styles,
                numbering: numbering,
                depth: depth
            )
            row.cells.append(cell)
        }

        return row
    }

    private static func parseTableRowProperties(from element: XMLElement) -> TableRowProperties {
        var props = TableRowProperties()

        // 行高
        if let trHeight = element.elements(forName: "w:trHeight").first {
            if let val = trHeight.attribute(forName: "w:val")?.stringValue {
                props.height = Int(val)
            }
            if let hRule = trHeight.attribute(forName: "w:hRule")?.stringValue {
                props.heightRule = HeightRule(rawValue: hRule)
            }
        }

        // 表頭行
        if element.elements(forName: "w:tblHeader").first != nil {
            props.isHeader = true
        }

        // 禁止分割
        if element.elements(forName: "w:cantSplit").first != nil {
            props.cantSplit = true
        }

        return props
    }

    private static func parseTableCell(
        from element: XMLElement,
        relationships: RelationshipsCollection,
        styles: [Style],
        numbering: Numbering,
        depth: Int = 0
    ) throws -> TableCell {
        var cell = TableCell()
        cell.paragraphs = []

        // 解析儲存格屬性
        if let tcPr = element.elements(forName: "w:tcPr").first {
            cell.properties = parseTableCellProperties(from: tcPr)
        }

        // 解析段落（傳入 styles 和 numbering 用於語義標註）
        for p in element.elements(forName: "w:p") {
            let para = try parseParagraph(
                from: p,
                relationships: relationships,
                styles: styles,
                numbering: numbering
            )
            cell.paragraphs.append(para)
        }

        // v0.17.0+ (#49): nested tables — recurse into <w:tbl> children of <w:tc>
        for nestedEl in element.elements(forName: "w:tbl") {
            let nested = try parseTable(
                from: nestedEl,
                relationships: relationships,
                styles: styles,
                numbering: numbering,
                depth: depth + 1
            )
            cell.nestedTables.append(nested)
        }

        // 確保至少有一個段落
        if cell.paragraphs.isEmpty {
            cell.paragraphs.append(Paragraph())
        }

        return cell
    }

    private static func parseTableCellProperties(from element: XMLElement) -> TableCellProperties {
        var props = TableCellProperties()

        // 寬度
        if let tcW = element.elements(forName: "w:tcW").first {
            if let w = tcW.attribute(forName: "w:w")?.stringValue {
                props.width = Int(w)
            }
            if let type = tcW.attribute(forName: "w:type")?.stringValue {
                props.widthType = WidthType(rawValue: type)
            }
        }

        // 水平合併
        if let gridSpan = element.elements(forName: "w:gridSpan").first,
           let val = gridSpan.attribute(forName: "w:val")?.stringValue {
            props.gridSpan = Int(val)
        }

        // 垂直合併
        if let vMerge = element.elements(forName: "w:vMerge").first,
           let val = vMerge.attribute(forName: "w:val")?.stringValue {
            props.verticalMerge = VerticalMerge(rawValue: val)
        }

        // 垂直對齊
        if let vAlign = element.elements(forName: "w:vAlign").first,
           let val = vAlign.attribute(forName: "w:val")?.stringValue {
            props.verticalAlignment = CellVerticalAlignment(rawValue: val)
        }

        // 底色
        if let shd = element.elements(forName: "w:shd").first,
           let fill = shd.attribute(forName: "w:fill")?.stringValue {
            var shading = CellShading(fill: fill)
            if let color = shd.attribute(forName: "w:color")?.stringValue {
                shading.color = color
            }
            if let val = shd.attribute(forName: "w:val")?.stringValue {
                shading.pattern = ShadingPattern(rawValue: val)
            }
            props.shading = shading
        }

        // v0.17.0+ (#49): diagonal borders
        if let tcBorders = element.elements(forName: "w:tcBorders").first {
            let tl2br = parseBorder(tcBorders.elements(forName: "w:tl2br").first)
            let tr2bl = parseBorder(tcBorders.elements(forName: "w:tr2bl").first)
            if tl2br != nil || tr2bl != nil {
                if props.borders == nil { props.borders = CellBorders() }
                props.borders?.tl2br = tl2br
                props.borders?.tr2bl = tr2bl
            }
        }

        return props
    }

    /// v0.17.0+ helper: parse a single `<w:top|bottom|...>` border element.
    private static func parseBorder(_ element: XMLElement?) -> Border? {
        guard let el = element else { return nil }
        let style = BorderStyle(rawValue: el.attribute(forName: "w:val")?.stringValue ?? "single") ?? .single
        let size = Int(el.attribute(forName: "w:sz")?.stringValue ?? "4") ?? 4
        let color = el.attribute(forName: "w:color")?.stringValue ?? "000000"
        return Border(style: style, size: size, color: color)
    }

    // MARK: - Styles Parsing

    private static func parseStyles(from xml: XMLDocument) throws -> [Style] {
        var styles: [Style] = []

        let styleNodes = try xml.nodes(forXPath: "//*[local-name()='style']")

        for node in styleNodes {
            guard let element = node as? XMLElement else { continue }

            guard let styleId = element.attribute(forName: "w:styleId")?.stringValue else { continue }
            guard let typeStr = element.attribute(forName: "w:type")?.stringValue,
                  let type = StyleType(rawValue: typeStr) else { continue }

            var name = styleId
            if let nameElement = element.elements(forName: "w:name").first,
               let val = nameElement.attribute(forName: "w:val")?.stringValue {
                name = val
            }

            var style = Style(id: styleId, name: name, type: type)

            // 基於
            if let basedOn = element.elements(forName: "w:basedOn").first,
               let val = basedOn.attribute(forName: "w:val")?.stringValue {
                style.basedOn = val
            }

            // 下一樣式
            if let next = element.elements(forName: "w:next").first,
               let val = next.attribute(forName: "w:val")?.stringValue {
                style.nextStyle = val
            }

            // 預設
            if element.attribute(forName: "w:default")?.stringValue == "1" {
                style.isDefault = true
            }

            // 快速樣式
            style.isQuickStyle = element.elements(forName: "w:qFormat").first != nil

            // v0.16.0+ (#44 §8): linked paragraph↔character style
            if let link = element.elements(forName: "w:link").first,
               let val = link.attribute(forName: "w:val")?.stringValue {
                style.linkedStyleId = val
            }

            // v0.16.0+ (#44 §8): visibility flags
            style.hidden = element.elements(forName: "w:hidden").first != nil
            style.semiHidden = element.elements(forName: "w:semiHidden").first != nil

            // v0.16.0+ (#44 §8): localized name aliases — additional <w:name>
            // elements with xml:lang. The first <w:name> (already consumed
            // above) is the primary; subsequent ones with xml:lang are aliases.
            let nameElements = element.elements(forName: "w:name")
            for nameEl in nameElements.dropFirst() {
                guard let lang = nameEl.attribute(forName: "xml:lang")?.stringValue,
                      let val = nameEl.attribute(forName: "w:val")?.stringValue
                else { continue }
                style.aliases.append(StyleAlias(lang: lang, name: val))
            }

            // 段落屬性
            if let pPr = element.elements(forName: "w:pPr").first {
                style.paragraphProperties = parseParagraphProperties(from: pPr)
            }

            // Run 屬性
            if let rPr = element.elements(forName: "w:rPr").first {
                style.runProperties = parseRunProperties(from: rPr)
            }

            styles.append(style)
        }

        // 如果沒有讀到樣式，使用預設樣式
        if styles.isEmpty {
            styles = Style.defaultStyles
        }

        return styles
    }

    /// v0.16.0+ (#44 §8): parse `<w:latentStyles>` block into LatentStyle entries.
    /// Returns empty array when no block is present.
    private static func parseLatentStyles(from xml: XMLDocument) -> [LatentStyle] {
        guard let nodes = try? xml.nodes(forXPath: "//*[local-name()='lsdException']") else {
            return []
        }
        var entries: [LatentStyle] = []
        for node in nodes {
            guard let el = node as? XMLElement,
                  let name = el.attribute(forName: "w:name")?.stringValue
            else { continue }
            let priority = el.attribute(forName: "w:uiPriority")?.stringValue.flatMap(Int.init)
            let semiHidden = el.attribute(forName: "w:semiHidden")?.stringValue == "1"
            let unhideWhenUsed = el.attribute(forName: "w:unhideWhenUsed")?.stringValue == "1"
            let qFormat = el.attribute(forName: "w:qFormat")?.stringValue == "1"
            entries.append(LatentStyle(
                name: name,
                uiPriority: priority,
                semiHidden: semiHidden,
                unhideWhenUsed: unhideWhenUsed,
                qFormat: qFormat
            ))
        }
        return entries
    }

    // MARK: - Core Properties Parsing

    private static func parseCoreProperties(from xml: XMLDocument) throws -> DocumentProperties {
        var props = DocumentProperties()

        // 標題
        if let nodes = try? xml.nodes(forXPath: "//*[local-name()='title']"),
           let node = nodes.first {
            props.title = node.stringValue
        }

        // 主題
        if let nodes = try? xml.nodes(forXPath: "//*[local-name()='subject']"),
           let node = nodes.first {
            props.subject = node.stringValue
        }

        // 作者
        if let nodes = try? xml.nodes(forXPath: "//*[local-name()='creator']"),
           let node = nodes.first {
            props.creator = node.stringValue
        }

        // 關鍵字
        if let nodes = try? xml.nodes(forXPath: "//*[local-name()='keywords']"),
           let node = nodes.first {
            props.keywords = node.stringValue
        }

        // 描述
        if let nodes = try? xml.nodes(forXPath: "//*[local-name()='description']"),
           let node = nodes.first {
            props.description = node.stringValue
        }

        // 最後修改者
        if let nodes = try? xml.nodes(forXPath: "//*[local-name()='lastModifiedBy']"),
           let node = nodes.first {
            props.lastModifiedBy = node.stringValue
        }

        // 版本
        if let nodes = try? xml.nodes(forXPath: "//*[local-name()='revision']"),
           let node = nodes.first,
           let value = node.stringValue {
            props.revision = Int(value)
        }

        // 建立日期
        let dateFormatter = ISO8601DateFormatter()
        if let nodes = try? xml.nodes(forXPath: "//*[local-name()='created']"),
           let node = nodes.first,
           let value = node.stringValue {
            props.created = dateFormatter.date(from: value)
        }

        // 修改日期
        if let nodes = try? xml.nodes(forXPath: "//*[local-name()='modified']"),
           let node = nodes.first,
           let value = node.stringValue {
            props.modified = dateFormatter.date(from: value)
        }

        return props
    }

    // MARK: - Numbering Parsing

    /// 解析 numbering.xml
    private static func parseNumbering(from xml: XMLDocument) throws -> Numbering {
        var numbering = Numbering()

        // 解析抽象編號定義 (w:abstractNum)
        let abstractNumNodes = try xml.nodes(forXPath: "//*[local-name()='abstractNum']")
        for node in abstractNumNodes {
            guard let element = node as? XMLElement,
                  let abstractNumIdStr = element.attribute(forName: "w:abstractNumId")?.stringValue,
                  let abstractNumId = Int(abstractNumIdStr) else { continue }

            var levels: [Level] = []

            // 解析層級 (w:lvl)
            for lvlElement in element.elements(forName: "w:lvl") {
                guard let ilvlStr = lvlElement.attribute(forName: "w:ilvl")?.stringValue,
                      let ilvl = Int(ilvlStr) else { continue }

                var numFmt: NumberFormat = .decimal
                var lvlText = ""
                var start = 1
                var indent = 720  // 預設縮排
                var fontName: String?

                // 編號格式 (w:numFmt)
                if let numFmtEl = lvlElement.elements(forName: "w:numFmt").first,
                   let val = numFmtEl.attribute(forName: "w:val")?.stringValue {
                    numFmt = NumberFormat(rawValue: val) ?? .decimal
                }

                // 文字格式 (w:lvlText)
                if let lvlTextEl = lvlElement.elements(forName: "w:lvlText").first,
                   let val = lvlTextEl.attribute(forName: "w:val")?.stringValue {
                    lvlText = val
                }

                // 起始值 (w:start)
                if let startEl = lvlElement.elements(forName: "w:start").first,
                   let val = startEl.attribute(forName: "w:val")?.stringValue {
                    start = Int(val) ?? 1
                }

                // 縮排 (w:pPr/w:ind)
                if let pPr = lvlElement.elements(forName: "w:pPr").first,
                   let ind = pPr.elements(forName: "w:ind").first,
                   let left = ind.attribute(forName: "w:left")?.stringValue {
                    indent = Int(left) ?? 720
                }

                // 字型 (w:rPr/w:rFonts)
                if let rPr = lvlElement.elements(forName: "w:rPr").first,
                   let rFonts = rPr.elements(forName: "w:rFonts").first,
                   let ascii = rFonts.attribute(forName: "w:ascii")?.stringValue {
                    fontName = ascii
                }

                let level = Level(
                    ilvl: ilvl,
                    start: start,
                    numFmt: numFmt,
                    lvlText: lvlText,
                    indent: indent,
                    fontName: fontName
                )
                levels.append(level)
            }

            let abstractNum = AbstractNum(abstractNumId: abstractNumId, levels: levels)
            numbering.abstractNums.append(abstractNum)
        }

        // 解析編號實例 (w:num)
        let numNodes = try xml.nodes(forXPath: "//*[local-name()='num']")
        for node in numNodes {
            guard let element = node as? XMLElement,
                  let numIdStr = element.attribute(forName: "w:numId")?.stringValue,
                  let numId = Int(numIdStr) else { continue }

            // 取得對應的 abstractNumId
            guard let abstractNumIdRef = element.elements(forName: "w:abstractNumId").first,
                  let abstractNumIdStr = abstractNumIdRef.attribute(forName: "w:val")?.stringValue,
                  let abstractNumId = Int(abstractNumIdStr) else { continue }

            // v0.16.0+ (#44 §3): parse w:lvlOverride children
            var overrides: [LvlOverride] = []
            for ovEl in element.elements(forName: "w:lvlOverride") {
                guard let ilvlStr = ovEl.attribute(forName: "w:ilvl")?.stringValue,
                      let ilvl = Int(ilvlStr),
                      let startEl = ovEl.elements(forName: "w:startOverride").first,
                      let startStr = startEl.attribute(forName: "w:val")?.stringValue,
                      let start = Int(startStr)
                else { continue }
                overrides.append(LvlOverride(ilvl: ilvl, startOverride: start))
            }

            let num = Num(numId: numId, abstractNumId: abstractNumId, lvlOverrides: overrides)
            numbering.nums.append(num)
        }

        return numbering
    }

    // MARK: - Semantic Detection

    /// 偵測段落的語義類型
    private static func detectParagraphSemantic(
        properties: ParagraphProperties,
        runs: [Run],
        styles: [Style],
        numbering: Numbering
    ) -> SemanticAnnotation? {
        // 1. 檢查標題樣式
        if let styleName = properties.style {
            if let headingLevel = detectHeadingLevel(styleName: styleName, styles: styles) {
                return SemanticAnnotation.heading(headingLevel)
            }

            // 檢查 Title/Subtitle
            let lowerStyle = styleName.lowercased()
            if lowerStyle == "title" || lowerStyle.contains("title") {
                return SemanticAnnotation(type: .title)
            }
            if lowerStyle == "subtitle" || lowerStyle.contains("subtitle") {
                return SemanticAnnotation(type: .subtitle)
            }
        }

        // 2. 檢查編號/項目符號
        if let numInfo = properties.numbering {
            let isBullet = isBulletList(numId: numInfo.numId, numbering: numbering)
            if isBullet {
                return SemanticAnnotation.bulletItem(level: numInfo.level)
            } else {
                return SemanticAnnotation.numberedItem(level: numInfo.level)
            }
        }

        // 3. 檢查分頁符
        if properties.pageBreakBefore {
            return SemanticAnnotation.pageBreak
        }

        // 4. 檢查 runs 中是否有公式或圖片（段落級別標註）
        for run in runs {
            // 有 OMML 公式
            if let rawXML = run.rawXML, rawXML.contains("oMath") {
                return SemanticAnnotation.ommlFormula
            }
            // 有圖片
            if run.drawing != nil {
                return SemanticAnnotation.unknownImage
            }
        }

        // 5. 預設為一般段落
        return SemanticAnnotation.paragraph
    }

    /// 從樣式名稱偵測標題層級
    private static func detectHeadingLevel(styleName: String, styles: [Style]) -> Int? {
        let lowerName = styleName.lowercased()

        // 直接比對常見標題樣式 ID
        // Word 預設: Heading1, Heading2, ... Heading9
        // 或中文: 標題1, 標題2, ...
        if lowerName.hasPrefix("heading") {
            let numPart = lowerName.dropFirst("heading".count)
            if let level = Int(numPart), level >= 1, level <= 9 {
                return level
            }
        }

        // 檢查樣式定義中的 name
        if let style = styles.first(where: { $0.id == styleName }) {
            let displayName = style.name.lowercased()
            if displayName.hasPrefix("heading") {
                let numPart = displayName.dropFirst("heading".count).trimmingCharacters(in: .whitespaces)
                if let level = Int(numPart), level >= 1, level <= 9 {
                    return level
                }
            }
            // 檢查 basedOn 是否為標題樣式
            if let basedOn = style.basedOn {
                return detectHeadingLevel(styleName: basedOn, styles: styles)
            }
        }

        return nil
    }

    /// 判斷是否為項目符號清單
    private static func isBulletList(numId: Int, numbering: Numbering) -> Bool {
        // 找到對應的 numbering instance (Num)
        guard let num = numbering.nums.first(where: { $0.numId == numId }) else {
            return false
        }

        // 找到對應的 abstract numbering (AbstractNum)
        guard let abstractNum = numbering.abstractNums.first(where: { $0.abstractNumId == num.abstractNumId }) else {
            return false
        }

        // 檢查第一層的格式
        if let firstLevel = abstractNum.levels.first {
            // bullet 格式通常是 .bullet 或文字是符號
            if firstLevel.numFmt == .bullet {
                return true
            }
            // 檢查文字是否為符號（如 •、○、■ 等）
            let text = firstLevel.lvlText
            let bulletSymbols = ["•", "○", "■", "□", "◆", "◇", "▪", "▫", "●", "○", "\u{F0B7}", "\u{F0A7}"]
            for symbol in bulletSymbols {
                if text.contains(symbol) {
                    return true
                }
            }
        }

        return false
    }

    // MARK: - Comments Extended Parsing

    /// 解析 commentsExtended.xml（Word 2012+ 回覆與已解決狀態）
    private static func parseCommentsExtended(from xml: XMLDocument, into comments: inout CommentsCollection) throws {
        // commentsExtended.xml 的結構：
        // <w15:commentsEx>
        //   <w15:commentEx w15:paraId="..." w15:paraIdParent="..." w15:done="1"/>
        // </w15:commentsEx>
        let extNodes = try xml.nodes(forXPath: "//*[local-name()='commentEx']")

        for node in extNodes {
            guard let element = node as? XMLElement else { continue }

            // 取得 paraId
            let paraId = element.attribute(forName: "w15:paraId")?.stringValue
                ?? element.attribute(forName: "paraId")?.stringValue
            guard let paraId = paraId else { continue }

            // 找到對應的 comment（透過 paraId）
            guard let idx = comments.comments.firstIndex(where: { $0.paraId == paraId }) else { continue }

            // 解析 parentId（回覆）
            let parentParaId = element.attribute(forName: "w15:paraIdParent")?.stringValue
                ?? element.attribute(forName: "paraIdParent")?.stringValue
            if let parentParaId = parentParaId,
               let parentComment = comments.comments.first(where: { $0.paraId == parentParaId }) {
                comments.comments[idx].parentId = parentComment.id
            }

            // 解析 done（已解決）
            let doneStr = element.attribute(forName: "w15:done")?.stringValue
                ?? element.attribute(forName: "done")?.stringValue
            if doneStr == "1" {
                comments.comments[idx].done = true
            }
        }
    }

    // MARK: - Comments Parsing

    private static func parseComments(from xml: XMLDocument) throws -> CommentsCollection {
        var collection = CommentsCollection()

        // 取得所有註解節點
        let commentNodes = try xml.nodes(forXPath: "//*[local-name()='comment']")

        for node in commentNodes {
            guard let element = node as? XMLElement else { continue }

            // 解析註解 ID
            guard let idStr = element.attribute(forName: "w:id")?.stringValue,
                  let id = Int(idStr) else { continue }

            // 解析作者
            let author = element.attribute(forName: "w:author")?.stringValue ?? "Unknown"

            // 解析縮寫
            let initials = element.attribute(forName: "w:initials")?.stringValue

            // 解析日期
            let dateFormatter = ISO8601DateFormatter()
            var date = Date()
            if let dateStr = element.attribute(forName: "w:date")?.stringValue {
                date = dateFormatter.date(from: dateStr) ?? Date()
            }

            // 解析註解文字（從 w:p/w:r/w:t 取得）
            var text = ""
            let textNodes = try element.nodes(forXPath: ".//*[local-name()='t']")
            for textNode in textNodes {
                text += textNode.stringValue ?? ""
            }

            // 建立 Comment 物件
            // 注意：從 comments.xml 讀取時，paragraphIndex 需要從文件中的 commentRangeStart 來確定
            // 這裡先設為 -1，表示需要從文件內容對應
            var comment = Comment(
                id: id,
                author: author,
                text: text.trimmingCharacters(in: .whitespacesAndNewlines),
                paragraphIndex: -1,
                date: date,
                initials: initials
            )

            // 嘗試解析 w14:paraId（用於回覆連結）
            // 從段落屬性中取得
            if let pElement = element.elements(forName: "w:p").first {
                // w14:paraId 可能在段落屬性中
                if let paraIdAttr = pElement.attribute(forName: "w14:paraId")?.stringValue {
                    comment.paraId = paraIdAttr
                }
            }

            collection.comments.append(comment)
        }

        return collection
    }

    // MARK: - XPath-free Helpers

    /// 在子孫中搜尋第一個符合 localName 的元素（取代 XPath，避免 O(n²)）
    private static func findFirstDescendant(of node: XMLNode, localNames: [String]) -> XMLElement? {
        guard let children = node.children else { return nil }
        for child in children {
            guard let el = child as? XMLElement else { continue }
            if let name = el.localName, localNames.contains(name) {
                return el
            }
            if let found = findFirstDescendant(of: el, localNames: localNames) {
                return found
            }
        }
        return nil
    }

    // MARK: - Document root attribute extraction (PsychQuant/che-word-mcp#56)

    /// Parse every attribute (xmlns:* declarations, mc:Ignorable, etc.) from the
    /// `<w:document>` opening tag in raw `document.xml` bytes. Bypasses
    /// Foundation's `XMLDocument` because libxml2 silently drops xmlns:*
    /// declarations whose prefix it sees as unused — and OOXML routinely declares
    /// 30+ extension namespaces that are referenced only inside parts the typed
    /// model has not yet expanded (e.g., w16cex inside `<w:document>` body).
    ///
    /// Locates the substring between `<w:document` and the matching `>` (skipping
    /// `?>` from the XML prolog), splits attributes on whitespace while respecting
    /// quoted values, and returns them as a `[name: value]` map.
    static func parseDocumentRootAttributes(from data: Data) -> [String: String] {
        return parseContainerRootAttributes(from: data, rootElementOpenPrefix: "<w:document")
    }

    /// v0.19.2+ (#56 follow-up F4): generalized version of
    /// `parseDocumentRootAttributes` for any container's root element.
    /// Used by header (`<w:hdr`), footer (`<w:ftr`), footnotes
    /// (`<w:footnotes`) and endnotes (`<w:endnotes`) raw-byte ingestion so
    /// every part type — not just `word/document.xml` — preserves source
    /// `xmlns:*` declarations and other root-level attributes through a
    /// no-op round-trip.
    ///
    /// Pass the literal opening prefix including `<` and the element local
    /// name (e.g., `"<w:hdr"`). Behaviour matches the document.xml variant:
    /// returns `[:]` on UTF-8 decode failure, missing prefix, or unterminated
    /// open tag (Writer falls back to its hardcoded namespace template).
    static func parseContainerRootAttributes(
        from data: Data,
        rootElementOpenPrefix: String
    ) -> [String: String] {
        guard let raw = String(data: data, encoding: .utf8) else { return [:] }
        guard let openRange = raw.range(of: rootElementOpenPrefix) else { return [:] }
        // Find the closing `>` of the open tag (ignoring `>` inside attribute values).
        var idx = openRange.upperBound
        var inQuote: Character? = nil
        while idx < raw.endIndex {
            let c = raw[idx]
            if let q = inQuote {
                if c == q { inQuote = nil }
            } else if c == "\"" || c == "'" {
                inQuote = c
            } else if c == ">" {
                break
            }
            idx = raw.index(after: idx)
        }
        guard idx < raw.endIndex else { return [:] }
        // attrSlice is "<w:hdr ATTRS" — strip element name, keep ATTRS.
        let attrSlice = String(raw[raw.index(openRange.lowerBound, offsetBy: rootElementOpenPrefix.count)..<idx])
        return splitAttributes(attrSlice)
    }

    /// Split a whitespace-and-attribute string like ` xmlns:w="..." mc:Ignorable="w14"`
    /// into a name→value map. Handles either `"` or `'` quoting and ignores the
    /// trailing `/` self-closing slash.
    private static func splitAttributes(_ s: String) -> [String: String] {
        var result: [String: String] = [:]
        var i = s.startIndex
        let end = s.endIndex
        while i < end {
            // Skip whitespace.
            while i < end, s[i].isWhitespace { i = s.index(after: i) }
            if i >= end { break }
            if s[i] == "/" { break }   // self-closing slash before `>`.
            // Read attribute name up to `=`.
            let nameStart = i
            while i < end, s[i] != "=" { i = s.index(after: i) }
            if i >= end { break }
            let name = String(s[nameStart..<i]).trimmingCharacters(in: .whitespaces)
            i = s.index(after: i)   // skip `=`.
            // Skip whitespace before value.
            while i < end, s[i].isWhitespace { i = s.index(after: i) }
            if i >= end { break }
            // Read quoted value.
            let quote = s[i]
            guard quote == "\"" || quote == "'" else { break }
            i = s.index(after: i)
            let valueStart = i
            while i < end, s[i] != quote { i = s.index(after: i) }
            if i >= end { break }
            let value = String(s[valueStart..<i])
            i = s.index(after: i)
            if !name.isEmpty { result[name] = value }
        }
        return result
    }
}
