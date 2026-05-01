import Foundation

/// Word 文件結構
public struct WordDocument: Equatable {
    public var body: Body
    public var styles: [Style]
    public var properties: DocumentProperties
    public var numbering: Numbering
    public var sectionProperties: SectionProperties
    public var headers: [Header]
    public var footers: [Footer]
    public var images: [ImageReference]              // 圖片資源
    public var hyperlinkReferences: [HyperlinkReference] = []  // 超連結關係（用於 .rels）
    public var comments: CommentsCollection = CommentsCollection()   // 註解集合
    public var revisions: RevisionsCollection = RevisionsCollection() // 修訂集合
    public var footnotes: FootnotesCollection = FootnotesCollection() // 腳註集合
    public var endnotes: EndnotesCollection = EndnotesCollection()    // 尾註集合
    internal var nextBookmarkId: Int = 1       // 書籤 ID 計數器
    private var nextHyperlinkId: Int = 1      // 超連結 ID 計數器

    /// Source archive wrapper, set by `DocxReader.read(from:)` for
    /// preserve-by-default round-trip fidelity (v0.12.0+). `nil` for documents
    /// created via initializer without a source ZIP.
    ///
    /// **Lifecycle**: callers SHALL invoke `close()` when finished to release
    /// the underlying tempDir. Forgetting to call `close()` leaks the tempDir
    /// until process exit (macOS reclaims `/tmp` on reboot).
    ///
    /// **Excluded from Equatable** because two reads of the same source `.docx`
    /// produce different UUID-named tempDirs but should still be considered
    /// equal in document content.
    internal var preservedArchive: PreservedArchive?

    /// Convenience accessor exposing the unzip tempDir URL.
    /// `nil` for initializer-built documents without a source ZIP.
    ///
    /// **Public read-only**: callers (e.g., MCP servers implementing
    /// theme/header/footer CRUD) need to read original OOXML parts directly
    /// from the preserved archive. Mutation of the underlying archive is
    /// internal to `ooxml-swift`.
    public var archiveTempDir: URL? {
        return preservedArchive?.tempDir
    }

    /// OOXML part paths the typed model has mutated since `DocxReader.read()`
    /// returned (v0.13.0+). Used by `DocxWriter` overlay mode to skip writers
    /// for parts that have not changed, achieving true byte-for-byte preservation.
    ///
    /// Paths follow the OOXML archive convention (e.g., `"word/document.xml"`,
    /// `"word/header1.xml"`, `"word/theme/theme1.xml"`, `"docProps/core.xml"`).
    ///
    /// **Excluded from Equatable**: two reads of the same source `.docx`
    /// produce empty `modifiedParts` regardless, but that's not a sameness
    /// criterion either — it is per-instance mutation tracking state.
    internal var modifiedParts: Set<String> = []

    /// Read-only public accessor for `modifiedParts`. Used by tests and
    /// downstream consumers to verify which OOXML parts will be re-emitted on
    /// the next `DocxWriter.write()` call in overlay mode.
    public var modifiedPartsView: Set<String> {
        return modifiedParts
    }

    /// Mark an OOXML part path as dirty so `DocxWriter` overlay mode re-emits
    /// it on the next `write(_:to:)` call. Used by external consumers (e.g.,
    /// `che-word-mcp`) that write directly to `archiveTempDir` bypassing the
    /// typed mutation methods (e.g., editing `word/theme/theme1.xml` via raw
    /// XML manipulation).
    ///
    /// Idempotent — inserting the same path twice has no additional effect.
    public mutating func markPartDirty(_ partPath: String) {
        modifiedParts.insert(partPath)
    }

    /// v0.16.0+ (#44 §8): latentStyles for Quick Style Gallery defaults of
    /// built-in styles not yet materialized as `<w:style>` blocks.
    public var latentStyles: [LatentStyle] = []

    /// v0.17.0+ (#51): document-level setting for `<w:evenAndOddHeaders/>` in
    /// settings.xml — when true, headers/footers of type `even` apply to even pages.
    public var evenAndOddHeaders: Bool = false

    /// v0.19.0+ (PsychQuant/che-word-mcp#56): every attribute (including all
    /// `xmlns:*` namespace declarations and `mc:Ignorable`) found on the source
    /// `<w:document>` root element. Populated by `DocxReader.read(from:)` from
    /// the parsed source. Emitted verbatim by `DocxWriter.writeDocument(_:to:)`
    /// so a no-op round-trip preserves the original namespace decl set instead
    /// of collapsing to the hardcoded `xmlns:w` + `xmlns:r` pair (the v3.12.0
    /// failure mode that caused `libxml2` to report "unbound prefix" errors on
    /// every body that referenced `mc:`, `wp:`, `w14:`, etc.).
    ///
    /// Empty for documents constructed via `WordDocument()` initializer without
    /// a source ZIP — the Writer falls back to emitting only `xmlns:w` + `xmlns:r`
    /// so create-from-scratch behavior stays unchanged.
    public var documentRootAttributes: [String: String] = [:]

    public init() {
        self.body = Body()
        self.styles = Style.defaultStyles
        self.properties = DocumentProperties()
        self.numbering = Numbering()
        self.sectionProperties = SectionProperties()
        self.headers = []
        self.footers = []
        self.images = []
        self.hyperlinkReferences = []
        self.comments = CommentsCollection()
        self.revisions = RevisionsCollection()
        self.footnotes = FootnotesCollection()
        self.endnotes = EndnotesCollection()
        self.preservedArchive = nil
    }

    /// Release the source archive's unzip tempDir (added v0.12.0).
    ///
    /// Idempotent: calling on a document whose `preservedArchive == nil` is a
    /// no-op. After `close()`, subsequent `DocxWriter.write(self, to:)` falls
    /// back to scratch mode (no preserve-by-default).
    ///
    /// Callers SHOULD invoke `close()` after the final `DocxWriter.write()` to
    /// avoid leaking the tempDir until process exit.
    public mutating func close() {
        preservedArchive?.cleanup()
        preservedArchive = nil
    }

    /// Manual Equatable conformance excluding `preservedArchive`.
    /// See doc comment on `preservedArchive` for rationale.
    public static func == (lhs: WordDocument, rhs: WordDocument) -> Bool {
        return lhs.body == rhs.body
            && lhs.styles == rhs.styles
            && lhs.properties == rhs.properties
            && lhs.numbering == rhs.numbering
            && lhs.sectionProperties == rhs.sectionProperties
            && lhs.headers == rhs.headers
            && lhs.footers == rhs.footers
            && lhs.images == rhs.images
            && lhs.hyperlinkReferences == rhs.hyperlinkReferences
            && lhs.comments == rhs.comments
            && lhs.revisions == rhs.revisions
            && lhs.footnotes == rhs.footnotes
            && lhs.endnotes == rhs.endnotes
            && lhs.nextBookmarkId == rhs.nextBookmarkId
            && lhs.nextHyperlinkId == rhs.nextHyperlinkId
            && lhs.documentRootAttributes == rhs.documentRootAttributes
    }

    // MARK: - Document Info

    public struct Info {
        public let paragraphCount: Int
        public let characterCount: Int
        public let wordCount: Int
        public let tableCount: Int

        public init(paragraphCount: Int, characterCount: Int, wordCount: Int, tableCount: Int) {
            self.paragraphCount = paragraphCount
            self.characterCount = characterCount
            self.wordCount = wordCount
            self.tableCount = tableCount
        }
    }

    public func getInfo() -> Info {
        let paragraphs = getParagraphs()
        let text = getText()
        let words = text.split(whereSeparator: { $0.isWhitespace || $0.isNewline })

        return Info(
            paragraphCount: paragraphs.count,
            characterCount: text.count,
            wordCount: words.count,
            tableCount: body.tables.count
        )
    }

    // MARK: - Text Operations

    public func getText() -> String {
        var result = ""
        for child in body.children {
            result += textOfBodyChild(child)
        }
        return result.trimmingCharacters(in: .whitespacesAndNewlines)
    }

    /// Recursive helper: text content of any BodyChild including block-level
    /// SDT wrappers (#44 task 3.4). Block-level controls are transparent —
    /// their children's text is concatenated with the SDT wrapper invisible.
    private func textOfBodyChild(_ child: BodyChild) -> String {
        switch child {
        case .paragraph(let para):
            return para.getText() + "\n"
        case .table(let table):
            return table.getText() + "\n"
        case .contentControl(_, let children):
            return children.map { textOfBodyChild($0) }.joined()
        case .bookmarkMarker, .rawBlockElement:
            // Body-level markers (bookmarks / raw block elements) carry no
            // user-visible text. PsychQuant/che-word-mcp#58.
            return ""
        }
    }

    public func getParagraphs() -> [Paragraph] {
        var result: [Paragraph] = []
        for child in body.children {
            collectTopLevelParagraphs(child, into: &result)
        }
        return result
    }

    /// Recursively collect paragraphs that are siblings at the body level,
    /// transparently descending into block-level SDTs (#44 task 3.4).
    /// Does NOT descend into tables — that's `getAllParagraphs`'s job.
    private func collectTopLevelParagraphs(_ child: BodyChild, into result: inout [Paragraph]) {
        switch child {
        case .paragraph(let para):
            result.append(para)
        case .table:
            break
        case .contentControl(_, let children):
            for c in children { collectTopLevelParagraphs(c, into: &result) }
        case .bookmarkMarker, .rawBlockElement:
            // No paragraphs to collect from body-level markers (#58).
            break
        }
    }

    /// 遞迴收集所有段落，包含表格儲存格內的段落
    public func getAllParagraphs() -> [Paragraph] {
        var result: [Paragraph] = []
        for child in body.children {
            collectAllParagraphs(child, into: &result)
        }
        return result
    }

    /// Recursively collect every paragraph anywhere — body siblings,
    /// table cells, block-level SDT children (#44 task 3.4).
    private func collectAllParagraphs(_ child: BodyChild, into result: inout [Paragraph]) {
        switch child {
        case .paragraph(let para):
            result.append(para)
        case .table(let table):
            for row in table.rows {
                for cell in row.cells {
                    result.append(contentsOf: cell.paragraphs)
                }
            }
        case .contentControl(_, let children):
            for c in children { collectAllParagraphs(c, into: &result) }
        case .bookmarkMarker, .rawBlockElement:
            // Body-level markers contain no paragraphs (#58).
            break
        }
    }

    // MARK: - Paragraph Operations

    public mutating func appendParagraph(_ paragraph: Paragraph) {
        body.children.append(.paragraph(paragraph))
        modifiedParts.insert("word/document.xml")
    }

    public mutating func insertParagraph(_ paragraph: Paragraph, at index: Int) {
        let clampedIndex = min(max(0, index), body.children.count)
        body.children.insert(.paragraph(paragraph), at: clampedIndex)
        modifiedParts.insert("word/document.xml")
    }

    public mutating func updateParagraph(at index: Int, text: String) throws {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard index >= 0 && index < paragraphIndices.count else {
            throw WordError.invalidIndex(index)
        }

        let actualIndex = paragraphIndices[index]
        if case .paragraph(var para) = body.children[actualIndex] {
            para.runs = [Run(text: text)]
            body.children[actualIndex] = .paragraph(para)
        }
        modifiedParts.insert("word/document.xml")
    }

    public mutating func deleteParagraph(at index: Int) throws {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard index >= 0 && index < paragraphIndices.count else {
            throw WordError.invalidIndex(index)
        }

        let actualIndex = paragraphIndices[index]
        body.children.remove(at: actualIndex)
        modifiedParts.insert("word/document.xml")
    }

    /// Replace occurrences of `find` with `replacement` across the document.
    ///
    /// Delegates per-paragraph splicing to `TextReplacementEngine` (flatten-then-map
    /// algorithm — cross-run matches succeed, replacement inherits the start run's
    /// formatting).
    ///
    /// Scope:
    /// - `.bodyAndTables` (default) — body paragraphs and table-cell paragraphs.
    /// - `.all` — additionally scans headers, footers, footnotes, endnotes.
    ///
    /// - Parameter options: `ReplaceOptions` — scope, regex, matchCase.
    /// - Returns: Number of replacements performed.
    /// - Throws: `ReplaceError.invalidRegex` when `options.regex == true` and the
    ///   pattern is invalid.
    /// Recursive helper for `replaceText` to descend into block-level SDT
    /// children (#44 task 3.4). Returns total replacements made.
    /// v0.19.0+ (PsychQuant/che-word-mcp#56) Phase 5: walk every editable run
    /// surface on a paragraph (top-level runs + Hyperlink.runs +
    /// FieldSimple.runs + AlternateContent.fallbackRuns) so `replace_text` /
    /// `format_text` no longer silently fail on text living inside structural
    /// wrappers. Returns the number of replacements made on this paragraph.
    private static func replaceInParagraphSurfaces(
        _ para: inout Paragraph,
        find: String,
        with replacement: String,
        options: ReplaceOptions
    ) throws -> Int {
        var count = 0
        count += try TextReplacementEngine.replace(
            runs: &para.runs, find: find, with: replacement, options: options
        )
        for hIdx in 0..<para.hyperlinks.count {
            count += try TextReplacementEngine.replace(
                runs: &para.hyperlinks[hIdx].runs, find: find, with: replacement, options: options
            )
        }
        for fIdx in 0..<para.fieldSimples.count {
            count += try TextReplacementEngine.replace(
                runs: &para.fieldSimples[fIdx].runs, find: find, with: replacement, options: options
            )
        }
        for acIdx in 0..<para.alternateContents.count {
            count += try TextReplacementEngine.replace(
                runs: &para.alternateContents[acIdx].fallbackRuns,
                find: find, with: replacement, options: options
            )
        }
        // PsychQuant/che-word-mcp#63: inline `<w:sdt>` content controls store
        // their inner runs as raw XML in `ContentControl.content`, not typed
        // `[Run]` arrays. Pre-fix `replace_text` silently 0-matched on text
        // wrapped in inline SDTs — common in docx output from external
        // converters (pandoc / Quarto / LaTeX→docx) which wrap cross-ref
        // placeholders (`[tab:foo]`, `[fig:bar]`) in inline SDTs by convention.
        // Issue title pointed at "literal `[ ]` brackets" but the brackets
        // were a coincidence — the trigger is the SDT wrapping.
        for cIdx in 0..<para.contentControls.count {
            count += try replaceInContentControl(
                &para.contentControls[cIdx], find: find, with: replacement, options: options
            )
        }
        return count
    }

    /// Recursively replace text inside an inline `<w:sdt>` content control.
    ///
    /// `ContentControl.content` is verbatim inner XML (typically a sequence of
    /// `<w:r><w:t>...</w:t></w:r>` blocks, possibly mixed with `<w:hyperlink>`,
    /// `<w:fldSimple>`, etc). We mutate text by walking the parsed XML tree's
    /// `<w:t>` descendants and applying cross-element find/replace, preserving
    /// the wrapper's structure (sdtPr, hyperlink/fldSimple wrappers inside,
    /// unrecognized elements). `<w:delText>` and `<w:instrText>` are skipped
    /// because they don't represent displayed text. Nested `<w:sdt>` subtrees
    /// inside the content blob are also skipped — those would have been parsed
    /// out into `cc.children` by `SDTParser.parseSDT`, so we recurse via
    /// `cc.children` instead of duplicating coverage.
    private static func replaceInContentControl(
        _ cc: inout ContentControl,
        find: String,
        with replacement: String,
        options: ReplaceOptions
    ) throws -> Int {
        var count = 0
        if !cc.content.isEmpty, cc.content.contains("<w:t") {
            let result = try TextReplacementEngine.replaceInContentXML(
                cc.content, find: find, with: replacement, options: options
            )
            cc.content = result.xml
            count += result.replacements
        }
        for ccIdx in 0..<cc.children.count {
            count += try replaceInContentControl(
                &cc.children[ccIdx], find: find, with: replacement, options: options
            )
        }
        return count
    }

    private mutating func replaceTextInBodyChildren(
        _ children: inout [BodyChild],
        find: String,
        with replacement: String,
        options: ReplaceOptions
    ) throws -> Int {
        var count = 0
        for i in 0..<children.count {
            switch children[i] {
            case .bookmarkMarker, .rawBlockElement:
                // Body-level markers contain no replaceable text (#58).
                continue
            case .paragraph(var para):
                count += try Self.replaceInParagraphSurfaces(
                    &para, find: find, with: replacement, options: options
                )
                children[i] = .paragraph(para)
            case .table(var table):
                for rowIdx in 0..<table.rows.count {
                    for cellIdx in 0..<table.rows[rowIdx].cells.count {
                        for paraIdx in 0..<table.rows[rowIdx].cells[cellIdx].paragraphs.count {
                            var para = table.rows[rowIdx].cells[cellIdx].paragraphs[paraIdx]
                            count += try Self.replaceInParagraphSurfaces(
                                &para, find: find, with: replacement, options: options
                            )
                            table.rows[rowIdx].cells[cellIdx].paragraphs[paraIdx] = para
                        }
                    }
                }
                children[i] = .table(table)
            case .contentControl(let metadata, var inner):
                count += try replaceTextInBodyChildren(
                    &inner, find: find, with: replacement, options: options
                )
                children[i] = .contentControl(metadata, children: inner)
            }
        }
        return count
    }

    @discardableResult
    public mutating func replaceText(
        find: String,
        with replacement: String,
        options: ReplaceOptions = ReplaceOptions()
    ) throws -> Int {
        var count = 0

        // Body + tables (always scanned). v0.19.0+ (#56) uses
        // `replaceInParagraphSurfaces` so edits inside Hyperlink / FieldSimple /
        // AlternateContent.fallbackRuns are no longer silently dropped.
        for i in 0..<body.children.count {
            switch body.children[i] {
            case .bookmarkMarker, .rawBlockElement:
                // Body-level markers contain no replaceable text (#58).
                continue
            case .paragraph(var para):
                count += try Self.replaceInParagraphSurfaces(
                    &para, find: find, with: replacement, options: options
                )
                body.children[i] = .paragraph(para)
            case .table(var table):
                for rowIdx in 0..<table.rows.count {
                    for cellIdx in 0..<table.rows[rowIdx].cells.count {
                        for paraIdx in 0..<table.rows[rowIdx].cells[cellIdx].paragraphs.count {
                            var para = table.rows[rowIdx].cells[cellIdx].paragraphs[paraIdx]
                            count += try Self.replaceInParagraphSurfaces(
                                &para, find: find, with: replacement, options: options
                            )
                            table.rows[rowIdx].cells[cellIdx].paragraphs[paraIdx] = para
                        }
                    }
                }
                body.children[i] = .table(table)
            case .contentControl(let metadata, var children):
                count += try replaceTextInBodyChildren(
                    &children, find: find, with: replacement, options: options
                )
                body.children[i] = .contentControl(metadata, children: children)
            }
        }

        // Body / table replacements always touch document.xml (when count > 0)
        if count > 0 {
            modifiedParts.insert("word/document.xml")
        }

        // Headers / footers / footnotes / endnotes — only under .all scope
        guard options.scope == .all else { return count }

        // v0.19.5+ (#56 R5-CONT P0 #3): containers route through the same
        // `replaceTextInBodyChildren` recursion the body path uses, so text
        // inside container tables / nested tables / SDT children is reachable.
        // Pre-fix the four loops iterated `.paragraphs` (the R5 P0 #6 flat
        // backward-compat view) — `replaceText(scope: .all)` against a
        // header / footer / note table cell silently returned 0 matches
        // even though the text existed and round-tripped on save (verify
        // R5 P0 #3 / Codex P1 / Regression F1 / DA C1).
        // Local-var copy avoids the Swift exclusivity violation that
        // `&self` (the recursion lock on `replaceTextInBodyChildren`)
        // would otherwise raise for `&headers[i].bodyChildren`.
        for i in 0..<headers.count {
            let beforeCount = count
            var children = headers[i].bodyChildren
            count += try replaceTextInBodyChildren(
                &children, find: find, with: replacement, options: options
            )
            headers[i].bodyChildren = children
            if count > beforeCount {
                modifiedParts.insert("word/\(headers[i].fileName)")
            }
        }
        for i in 0..<footers.count {
            let beforeCount = count
            var children = footers[i].bodyChildren
            count += try replaceTextInBodyChildren(
                &children, find: find, with: replacement, options: options
            )
            footers[i].bodyChildren = children
            if count > beforeCount {
                modifiedParts.insert("word/\(footers[i].fileName)")
            }
        }
        let beforeFootnotes = count
        for i in 0..<footnotes.footnotes.count {
            var children = footnotes.footnotes[i].bodyChildren
            count += try replaceTextInBodyChildren(
                &children, find: find, with: replacement, options: options
            )
            footnotes.footnotes[i].bodyChildren = children
        }
        if count > beforeFootnotes {
            modifiedParts.insert("word/footnotes.xml")
        }
        let beforeEndnotes = count
        for i in 0..<endnotes.endnotes.count {
            var children = endnotes.endnotes[i].bodyChildren
            count += try replaceTextInBodyChildren(
                &children, find: find, with: replacement, options: options
            )
            endnotes.endnotes[i].bodyChildren = children
        }
        if count > beforeEndnotes {
            modifiedParts.insert("word/endnotes.xml")
        }

        return count
    }

    // MARK: - Formatting

    public mutating func formatParagraph(at index: Int, with format: RunProperties) throws {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard index >= 0 && index < paragraphIndices.count else {
            throw WordError.invalidIndex(index)
        }

        let actualIndex = paragraphIndices[index]
        if case .paragraph(var para) = body.children[actualIndex] {
            for i in 0..<para.runs.count {
                para.runs[i].properties.merge(with: format)
            }
            body.children[actualIndex] = .paragraph(para)
        }
        modifiedParts.insert("word/document.xml")
    }

    public mutating func setParagraphFormat(at index: Int, properties: ParagraphProperties) throws {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard index >= 0 && index < paragraphIndices.count else {
            throw WordError.invalidIndex(index)
        }

        let actualIndex = paragraphIndices[index]
        if case .paragraph(var para) = body.children[actualIndex] {
            para.properties.merge(with: properties)
            body.children[actualIndex] = .paragraph(para)
        }
        modifiedParts.insert("word/document.xml")
    }

    public mutating func applyStyle(at index: Int, style: String) throws {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard index >= 0 && index < paragraphIndices.count else {
            throw WordError.invalidIndex(index)
        }

        let actualIndex = paragraphIndices[index]
        if case .paragraph(var para) = body.children[actualIndex] {
            para.properties.style = style
            body.children[actualIndex] = .paragraph(para)
        }
        modifiedParts.insert("word/document.xml")
    }

    // MARK: - Table Operations

    public mutating func appendTable(_ table: Table) {
        body.children.append(.table(table))
        body.tables.append(table)
        modifiedParts.insert("word/document.xml")
    }

    public mutating func insertTable(_ table: Table, at index: Int) {
        let clampedIndex = min(max(0, index), body.children.count)
        body.children.insert(.table(table), at: clampedIndex)
        body.tables.append(table)
        modifiedParts.insert("word/document.xml")
    }

    /// 取得所有表格
    public func getTables() -> [Table] {
        return body.children.compactMap { child in
            if case .table(let table) = child {
                return table
            }
            return nil
        }
    }

    /// 取得表格索引對應到 body.children 的實際索引
    private func getTableIndices() -> [Int] {
        return body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .table = child { return i }
            return nil
        }
    }

    /// 更新表格儲存格內容
    public mutating func updateCell(tableIndex: Int, row: Int, col: Int, text: String) throws {
        let tableIndices = getTableIndices()

        guard tableIndex >= 0 && tableIndex < tableIndices.count else {
            throw WordError.invalidIndex(tableIndex)
        }

        let actualIndex = tableIndices[tableIndex]
        if case .table(var table) = body.children[actualIndex] {
            guard row >= 0 && row < table.rows.count else {
                throw WordError.invalidIndex(row)
            }
            let cellCount = table.rows[row].cells.count
            guard col >= 0 && col < cellCount else {
                let gridSpans = table.rows[row].cells.map { $0.properties.gridSpan ?? 1 }
                let gridCols = gridSpans.reduce(0, +)
                throw WordError.invalidFormat("Invalid col index \(col) for row \(row): row has \(cellCount) cell(s) (grid columns: \(gridCols), spans: \(gridSpans))")
            }

            // 只更新文字，保留 cell properties + run properties（粗體、字型等）
            let cell = table.rows[row].cells[col]
            if let firstRun = cell.paragraphs.first?.runs.first {
                // 保留第一個 run 的格式，只替換文字
                var updatedRun = firstRun
                updatedRun.text = text
                var updatedPara = cell.paragraphs[0]
                updatedPara.runs = [updatedRun]
                table.rows[row].cells[col].paragraphs = [updatedPara]
            } else {
                // 空 cell，直接設文字（保留 cell properties）
                table.rows[row].cells[col].paragraphs = [Paragraph(text: text)]
            }

            body.children[actualIndex] = .table(table)

            // 同步更新 body.tables
            if tableIndex < body.tables.count {
                body.tables[tableIndex] = table
            }
        }
        modifiedParts.insert("word/document.xml")
    }

    /// 刪除表格
    public mutating func deleteTable(at tableIndex: Int) throws {
        let tableIndices = getTableIndices()

        guard tableIndex >= 0 && tableIndex < tableIndices.count else {
            throw WordError.invalidIndex(tableIndex)
        }

        let actualIndex = tableIndices[tableIndex]
        body.children.remove(at: actualIndex)

        // 同步更新 body.tables
        if tableIndex < body.tables.count {
            body.tables.remove(at: tableIndex)
        }
        modifiedParts.insert("word/document.xml")
    }

    /// 合併儲存格（水平合併）
    public mutating func mergeCellsHorizontal(tableIndex: Int, row: Int, startCol: Int, endCol: Int) throws {
        let tableIndices = getTableIndices()

        guard tableIndex >= 0 && tableIndex < tableIndices.count else {
            throw WordError.invalidIndex(tableIndex)
        }

        let actualIndex = tableIndices[tableIndex]
        if case .table(var table) = body.children[actualIndex] {
            guard row >= 0 && row < table.rows.count else {
                throw WordError.invalidIndex(row)
            }
            guard startCol >= 0 && startCol < table.rows[row].cells.count else {
                throw WordError.invalidIndex(startCol)
            }
            guard endCol >= startCol && endCol < table.rows[row].cells.count else {
                throw WordError.invalidIndex(endCol)
            }

            // 設定第一個儲存格的 gridSpan
            let span = endCol - startCol + 1
            table.rows[row].cells[startCol].properties.gridSpan = span

            // 移除被合併的儲存格（從後往前移除以保持索引正確）
            for col in (startCol + 1...endCol).reversed() {
                table.rows[row].cells.remove(at: col)
            }

            body.children[actualIndex] = .table(table)
            if tableIndex < body.tables.count {
                body.tables[tableIndex] = table
            }
        }
        modifiedParts.insert("word/document.xml")
    }

    /// 合併儲存格（垂直合併）
    public mutating func mergeCellsVertical(tableIndex: Int, col: Int, startRow: Int, endRow: Int) throws {
        let tableIndices = getTableIndices()

        guard tableIndex >= 0 && tableIndex < tableIndices.count else {
            throw WordError.invalidIndex(tableIndex)
        }

        let actualIndex = tableIndices[tableIndex]
        if case .table(var table) = body.children[actualIndex] {
            guard startRow >= 0 && startRow < table.rows.count else {
                throw WordError.invalidIndex(startRow)
            }
            guard endRow >= startRow && endRow < table.rows.count else {
                throw WordError.invalidIndex(endRow)
            }

            // 設定第一個儲存格為 restart
            if col < table.rows[startRow].cells.count {
                table.rows[startRow].cells[col].properties.verticalMerge = .restart
            }

            // 設定其餘儲存格為 continue
            for row in (startRow + 1)...endRow {
                if col < table.rows[row].cells.count {
                    table.rows[row].cells[col].properties.verticalMerge = .continue
                }
            }

            body.children[actualIndex] = .table(table)
            if tableIndex < body.tables.count {
                body.tables[tableIndex] = table
            }
        }
        modifiedParts.insert("word/document.xml")
    }

    /// 設定表格樣式（邊框）
    public mutating func setTableBorders(tableIndex: Int, borders: TableBorders) throws {
        let tableIndices = getTableIndices()

        guard tableIndex >= 0 && tableIndex < tableIndices.count else {
            throw WordError.invalidIndex(tableIndex)
        }

        let actualIndex = tableIndices[tableIndex]
        if case .table(var table) = body.children[actualIndex] {
            table.properties.borders = borders
            body.children[actualIndex] = .table(table)
            if tableIndex < body.tables.count {
                body.tables[tableIndex] = table
            }
        }
        modifiedParts.insert("word/document.xml")
    }

    /// 設定儲存格底色
    public mutating func setCellShading(tableIndex: Int, row: Int, col: Int, shading: CellShading) throws {
        let tableIndices = getTableIndices()

        guard tableIndex >= 0 && tableIndex < tableIndices.count else {
            throw WordError.invalidIndex(tableIndex)
        }

        let actualIndex = tableIndices[tableIndex]
        if case .table(var table) = body.children[actualIndex] {
            guard row >= 0 && row < table.rows.count else {
                throw WordError.invalidIndex(row)
            }
            guard col >= 0 && col < table.rows[row].cells.count else {
                throw WordError.invalidIndex(col)
            }

            table.rows[row].cells[col].properties.shading = shading
            body.children[actualIndex] = .table(table)
            if tableIndex < body.tables.count {
                body.tables[tableIndex] = table
            }
        }
        modifiedParts.insert("word/document.xml")
    }

    // MARK: - Style Management

    /// 取得所有樣式
    public func getStyles() -> [Style] {
        return styles
    }

    /// 根據 ID 取得樣式
    public func getStyle(by id: String) -> Style? {
        return styles.first { $0.id == id }
    }

    /// 新增自訂樣式
    public mutating func addStyle(_ style: Style) throws {
        // 檢查是否已存在相同 ID 的樣式
        if styles.contains(where: { $0.id == style.id }) {
            throw WordError.invalidFormat("Style with id '\(style.id)' already exists")
        }
        styles.append(style)
        modifiedParts.insert("word/styles.xml")
    }

    /// 更新樣式
    public mutating func updateStyle(id: String, with updates: StyleUpdate) throws {
        guard let index = styles.firstIndex(where: { $0.id == id }) else {
            throw WordError.invalidFormat("Style '\(id)' not found")
        }

        var style = styles[index]

        if let name = updates.name {
            style.name = name
        }
        if let basedOn = updates.basedOn {
            style.basedOn = basedOn
        }
        if let nextStyle = updates.nextStyle {
            style.nextStyle = nextStyle
        }
        if let isQuickStyle = updates.isQuickStyle {
            style.isQuickStyle = isQuickStyle
        }
        if let paragraphProps = updates.paragraphProperties {
            if style.paragraphProperties == nil {
                style.paragraphProperties = ParagraphProperties()
            }
            style.paragraphProperties?.merge(with: paragraphProps)
        }
        if let runProps = updates.runProperties {
            if style.runProperties == nil {
                style.runProperties = RunProperties()
            }
            style.runProperties?.merge(with: runProps)
        }

        styles[index] = style
        modifiedParts.insert("word/styles.xml")
    }

    /// 刪除樣式（不能刪除預設樣式）
    public mutating func deleteStyle(id: String) throws {
        guard let index = styles.firstIndex(where: { $0.id == id }) else {
            throw WordError.invalidFormat("Style '\(id)' not found")
        }

        // 不能刪除預設樣式
        if styles[index].isDefault {
            throw WordError.invalidFormat("Cannot delete default style '\(id)'")
        }

        // 檢查是否為內建樣式（Normal, Heading1-3, Title, Subtitle）
        let builtInIds = ["Normal", "Heading1", "Heading2", "Heading3", "Title", "Subtitle"]
        if builtInIds.contains(id) {
            throw WordError.invalidFormat("Cannot delete built-in style '\(id)'")
        }

        styles.remove(at: index)
        modifiedParts.insert("word/styles.xml")
    }

    // MARK: - List Operations

    /// 插入項目符號清單
    public mutating func insertBulletList(items: [String], at index: Int? = nil) -> Int {
        let numId = numbering.createBulletList()

        for (itemIndex, text) in items.enumerated() {
            var para = Paragraph(text: text)
            para.properties.numbering = NumberingInfo(numId: numId, level: 0)

            if let index = index {
                insertParagraph(para, at: index + itemIndex)
            } else {
                appendParagraph(para)
            }
        }

        // appendParagraph/insertParagraph already marked document.xml — add numbering.xml.
        modifiedParts.insert("word/numbering.xml")
        return numId
    }

    /// 插入編號清單
    public mutating func insertNumberedList(items: [String], at index: Int? = nil) -> Int {
        let numId = numbering.createNumberedList()

        for (itemIndex, text) in items.enumerated() {
            var para = Paragraph(text: text)
            para.properties.numbering = NumberingInfo(numId: numId, level: 0)

            if let index = index {
                insertParagraph(para, at: index + itemIndex)
            } else {
                appendParagraph(para)
            }
        }

        modifiedParts.insert("word/numbering.xml")
        return numId
    }

    /// 設定段落的清單層級
    public mutating func setListLevel(paragraphIndex: Int, level: Int) throws {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let actualIndex = paragraphIndices[paragraphIndex]
        if case .paragraph(var para) = body.children[actualIndex] {
            guard para.properties.numbering != nil else {
                throw WordError.invalidFormat("Paragraph is not part of a list")
            }

            guard level >= 0 && level <= 8 else {
                throw WordError.invalidParameter("level", "Must be between 0 and 8")
            }

            para.properties.numbering?.level = level
            body.children[actualIndex] = .paragraph(para)
        }
        modifiedParts.insert("word/document.xml")
    }

    /// 將段落添加到現有清單
    public mutating func addToList(paragraphIndex: Int, numId: Int, level: Int = 0) throws {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let actualIndex = paragraphIndices[paragraphIndex]
        if case .paragraph(var para) = body.children[actualIndex] {
            para.properties.numbering = NumberingInfo(numId: numId, level: level)
            body.children[actualIndex] = .paragraph(para)
        }
        modifiedParts.insert("word/document.xml")
    }

    /// 移除段落的清單格式
    public mutating func removeFromList(paragraphIndex: Int) throws {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let actualIndex = paragraphIndices[paragraphIndex]
        if case .paragraph(var para) = body.children[actualIndex] {
            para.properties.numbering = nil
            body.children[actualIndex] = .paragraph(para)
        }
        modifiedParts.insert("word/document.xml")
    }

    // MARK: - Page Settings

    /// 設定頁面大小
    public mutating func setPageSize(_ size: PageSize) {
        sectionProperties.pageSize = size
        modifiedParts.insert("word/document.xml")
    }

    /// 設定頁面大小（使用名稱）
    public mutating func setPageSize(name: String) throws {
        guard let size = PageSize.from(name: name) else {
            throw WordError.invalidParameter("pageSize", "Unknown page size: \(name). Valid options: letter, a4, legal, a3, a5, b5, executive")
        }
        sectionProperties.pageSize = size
        modifiedParts.insert("word/document.xml")
    }

    /// 設定頁邊距
    public mutating func setPageMargins(_ margins: PageMargins) {
        sectionProperties.pageMargins = margins
        modifiedParts.insert("word/document.xml")
    }

    /// 設定頁邊距（使用名稱）
    public mutating func setPageMargins(name: String) throws {
        guard let margins = PageMargins.from(name: name) else {
            throw WordError.invalidParameter("margins", "Unknown margin preset: \(name). Valid options: normal, narrow, moderate, wide")
        }
        sectionProperties.pageMargins = margins
        modifiedParts.insert("word/document.xml")
    }

    /// 設定頁邊距（使用具體數值，單位：twips）
    public mutating func setPageMargins(top: Int? = nil, right: Int? = nil, bottom: Int? = nil, left: Int? = nil) {
        if let top = top {
            sectionProperties.pageMargins.top = top
        }
        if let right = right {
            sectionProperties.pageMargins.right = right
        }
        if let bottom = bottom {
            sectionProperties.pageMargins.bottom = bottom
        }
        if let left = left {
            sectionProperties.pageMargins.left = left
        }
        modifiedParts.insert("word/document.xml")
    }

    /// 設定頁面方向
    public mutating func setPageOrientation(_ orientation: PageOrientation) {
        sectionProperties.orientation = orientation

        // 如果切換方向，也要交換頁面寬高
        if orientation == .landscape && sectionProperties.pageSize.width < sectionProperties.pageSize.height {
            sectionProperties.pageSize = sectionProperties.pageSize.landscape
        } else if orientation == .portrait && sectionProperties.pageSize.width > sectionProperties.pageSize.height {
            sectionProperties.pageSize = sectionProperties.pageSize.landscape
        }
        modifiedParts.insert("word/document.xml")
    }

    /// 插入分頁符
    public mutating func insertPageBreak(at paragraphIndex: Int? = nil) {
        // 分頁符是一個特殊的段落，只包含 <w:br w:type="page"/>
        var para = Paragraph()
        para.hasPageBreak = true

        if let index = paragraphIndex {
            insertParagraph(para, at: index)
        } else {
            appendParagraph(para)
        }
        // append/insert already mark document.xml — explicit insert is idempotent.
    }

    /// 插入分節符
    public mutating func insertSectionBreak(type: SectionBreakType = .nextPage, at paragraphIndex: Int? = nil) {
        // 分節符放在段落屬性中
        var para = Paragraph()
        para.properties.sectionBreak = type

        if let index = paragraphIndex {
            insertParagraph(para, at: index)
        } else {
            appendParagraph(para)
        }
    }

    // MARK: - Header/Footer Operations

    /// 取得下一個可用的關係 ID。v0.12.0+: 在 overlay mode（preservedArchive 非
    /// nil）會掃描原始 `_rels/document.xml.rels` 避免與 preserved unknown rels 衝突。
    private var nextRelationshipId: String {
        // Read original rels XML when in overlay mode, otherwise allocate from
        // typed model only (preserves prior behavior for create_document paths).
        var originalRelsXML = ""
        if let tempDir = preservedArchive?.tempDir {
            let url = tempDir.appendingPathComponent("word/_rels/document.xml.rels")
            originalRelsXML = (try? String(contentsOf: url, encoding: .utf8)) ?? ""
        }
        var reservedIds: [String] = ["rId1", "rId2", "rId3"]
        if !numbering.abstractNums.isEmpty { reservedIds.append("rId4") }
        for header in headers { reservedIds.append(header.id) }
        for footer in footers { reservedIds.append(footer.id) }
        for image in images { reservedIds.append(image.id) }
        for hyperlinkRef in hyperlinkReferences { reservedIds.append(hyperlinkRef.relationshipId) }
        let allocator = RelationshipIdAllocator(
            originalRelsXML: originalRelsXML,
            additionalReservedIds: reservedIds
        )
        return allocator.allocate()
    }

    /// v0.19.5+ (#56 R5-CONT-3 P1 #5): diagnostic for callers that bypass
    /// `addHeader`/`addFooter` and append directly to `document.headers` /
    /// `document.footers` without setting `originalFileName`. Returns
    /// `(scope: "header"|"footer", fileName: String, indices: [Int])`
    /// for every fileName shared by ≥2 containers. Empty when no
    /// collisions. Public so MCP / diagnostic tools can warn before save.
    public var containerFileNameCollisions: [(scope: String, fileName: String, indices: [Int])] {
        var result: [(scope: String, fileName: String, indices: [Int])] = []
        var headerSeen: [String: [Int]] = [:]
        for (i, h) in headers.enumerated() {
            headerSeen[h.fileName, default: []].append(i)
        }
        for (fn, idxs) in headerSeen where idxs.count > 1 {
            result.append((scope: "header", fileName: fn, indices: idxs))
        }
        var footerSeen: [String: [Int]] = [:]
        for (i, f) in footers.enumerated() {
            footerSeen[f.fileName, default: []].append(i)
        }
        for (fn, idxs) in footerSeen where idxs.count > 1 {
            result.append((scope: "footer", fileName: fn, indices: idxs))
        }
        return result
    }

    /// v0.19.5+ (#56 R5-CONT-3 P1 #5): auto-repair container fileName
    /// collisions by reassigning `originalFileName` on the SECOND+
    /// instances using the same allocator the public `addHeader` /
    /// `addFooter` API uses. First instance keeps its existing
    /// (possibly nil → type-default) fileName. Marks every reassigned
    /// container's part dirty so the writer emits to the new path.
    /// Idempotent: calling again on a clean doc does nothing.
    /// Recommended call site: just before save when callers may have
    /// constructed containers via direct `headers.append`.
    public mutating func repairContainerFileNames() {
        var renamed = false
        var seen: Set<String> = []
        for i in 0..<headers.count {
            if seen.contains(headers[i].fileName) {
                headers[i].originalFileName = allocateHeaderFileName(for: headers[i].type)
                modifiedParts.insert("word/\(headers[i].fileName)")
                renamed = true
            }
            seen.insert(headers[i].fileName)
        }
        var seenF: Set<String> = []
        for i in 0..<footers.count {
            if seenF.contains(footers[i].fileName) {
                footers[i].originalFileName = allocateFooterFileName(for: footers[i].type)
                modifiedParts.insert("word/\(footers[i].fileName)")
                renamed = true
            }
            seenF.insert(footers[i].fileName)
        }
        // v0.19.5+ (#56 R5-CONT-4 Logic HIGH §15.4): when a rename
        // occurred, mark word/_rels/document.xml.rels dirty AND
        // [Content_Types].xml dirty too. Pre-fix the rename only marked
        // the new container-part path dirty — but document.xml.rels
        // still referenced the OLD fileName as the rId target. Overlay
        // mode preserved document.xml.rels from source archive →
        // post-save reread doc had rels pointing to wrong path.
        // [Content_Types].xml has Override entries for each container
        // by PartName too — same staleness risk on rename.
        if renamed {
            modifiedParts.insert("word/_rels/document.xml.rels")
            modifiedParts.insert("[Content_Types].xml")
        }
    }

    /// v0.13.5+ (#53): allocate a unique header fileName for the given type
    /// among existing typed-model headers. Returns `headerN.xml` (default),
    /// `headerFirstN.xml`, or `headerEvenN.xml` where N is the smallest
    /// positive integer such that no existing header has the same fileName.
    private func allocateHeaderFileName(for type: HeaderFooterType) -> String {
        let prefix: String
        switch type {
        case .default: prefix = "header"
        case .first: prefix = "headerFirst"
        case .even: prefix = "headerEven"
        }
        let existing = Set(headers.map { $0.fileName })
        var n = 1
        while existing.contains("\(prefix)\(n).xml") { n += 1 }
        return "\(prefix)\(n).xml"
    }

    /// v0.13.5+ (#53): allocate a unique footer fileName. Same logic as
    /// `allocateHeaderFileName(for:)` but scans `footers`.
    private func allocateFooterFileName(for type: HeaderFooterType) -> String {
        let prefix: String
        switch type {
        case .default: prefix = "footer"
        case .first: prefix = "footerFirst"
        case .even: prefix = "footerEven"
        }
        let existing = Set(footers.map { $0.fileName })
        var n = 1
        while existing.contains("\(prefix)\(n).xml") { n += 1 }
        return "\(prefix)\(n).xml"
    }

    /// 新增頁首
    public mutating func addHeader(text: String, type: HeaderFooterType = .default) -> Header {
        let id = nextRelationshipId
        // v0.13.5+ (#53): auto-suffix fileName so multi-instance default-type
        // adds don't collide on the type-based fallback "header1.xml".
        let fileName = allocateHeaderFileName(for: type)
        var header = Header.withText(text, id: id, type: type)
        header.originalFileName = fileName
        headers.append(header)

        // 更新分節屬性中的頁首參照
        if type == .default {
            sectionProperties.headerReference = id
        }

        modifiedParts.insert("word/\(header.fileName)")
        // sectPr (in document.xml) + content-types + rels grow when we add a header.
        modifiedParts.insert("word/document.xml")
        modifiedParts.insert("[Content_Types].xml")
        modifiedParts.insert("word/_rels/document.xml.rels")
        return header
    }

    /// 新增含頁碼的頁首
    public mutating func addHeaderWithPageNumber(type: HeaderFooterType = .default) -> Header {
        let id = nextRelationshipId
        let fileName = allocateHeaderFileName(for: type)
        var header = Header.withPageNumber(id: id, type: type)
        header.originalFileName = fileName
        headers.append(header)

        if type == .default {
            sectionProperties.headerReference = id
        }

        modifiedParts.insert("word/\(header.fileName)")
        modifiedParts.insert("word/document.xml")
        modifiedParts.insert("[Content_Types].xml")
        modifiedParts.insert("word/_rels/document.xml.rels")
        return header
    }

    /// 更新頁首內容
    public mutating func updateHeader(id: String, text: String) throws {
        guard let index = headers.firstIndex(where: { $0.id == id }) else {
            throw WordError.invalidFormat("Header '\(id)' not found")
        }

        var header = headers[index]
        header.paragraphs = [Paragraph(text: text)]
        headers[index] = header
        modifiedParts.insert("word/\(header.fileName)")
    }

    /// 新增頁尾
    public mutating func addFooter(text: String, type: HeaderFooterType = .default) -> Footer {
        let id = nextRelationshipId
        // v0.13.5+ (#53): auto-suffix fileName.
        let fileName = allocateFooterFileName(for: type)
        var footer = Footer.withText(text, id: id, type: type)
        footer.originalFileName = fileName
        footers.append(footer)

        // 更新分節屬性中的頁尾參照
        if type == .default {
            sectionProperties.footerReference = id
        }

        modifiedParts.insert("word/\(footer.fileName)")
        modifiedParts.insert("word/document.xml")
        modifiedParts.insert("[Content_Types].xml")
        modifiedParts.insert("word/_rels/document.xml.rels")
        return footer
    }

    /// 新增含頁碼的頁尾
    public mutating func addFooterWithPageNumber(format: PageNumberFormat = .simple, type: HeaderFooterType = .default) -> Footer {
        let id = nextRelationshipId
        let fileName = allocateFooterFileName(for: type)
        var footer = Footer.withPageNumber(id: id, format: format, type: type)
        footer.originalFileName = fileName
        footers.append(footer)

        if type == .default {
            sectionProperties.footerReference = id
        }

        modifiedParts.insert("word/\(footer.fileName)")
        modifiedParts.insert("word/document.xml")
        modifiedParts.insert("[Content_Types].xml")
        modifiedParts.insert("word/_rels/document.xml.rels")
        return footer
    }

    /// 更新頁尾內容
    public mutating func updateFooter(id: String, text: String) throws {
        guard let index = footers.firstIndex(where: { $0.id == id }) else {
            throw WordError.invalidFormat("Footer '\(id)' not found")
        }

        var footer = footers[index]
        footer.paragraphs = [Paragraph(text: text)]
        footers[index] = footer
        modifiedParts.insert("word/\(footer.fileName)")
    }

    // MARK: - Image Operations

    /// 取得下一個可用的圖片關係 ID
    /// internal (not private) so InsertLocation extension can reuse it.
    var nextImageRelationshipId: String {
        // v0.13.3+ (che-word-mcp#41 defensive hardening): delegate to the
        // allocator-based `nextRelationshipId` which consults original rels in
        // overlay mode. The naïve formula `baseId + headers + footers + images`
        // happened to track the typed model in lockstep for most cases but is
        // fragile against any mismatched assignment (e.g., reader-loaded doc
        // with hyperlinks/comments rels not counted by the formula).
        return nextRelationshipId
    }

    /// 從 Base64 插入圖片
    public mutating func insertImage(
        base64: String,
        fileName: String,
        widthPx: Int,
        heightPx: Int,
        at paragraphIndex: Int? = nil,
        name: String = "Picture",
        description: String = ""
    ) throws -> String {
        let imageId = nextImageRelationshipId

        // 建立圖片參照
        let imageRef = try ImageReference.from(base64: base64, fileName: fileName, id: imageId)
        images.append(imageRef)

        // 建立 Drawing
        var drawing = Drawing.from(widthPx: widthPx, heightPx: heightPx, imageId: imageId, name: name)
        drawing.description = description

        // 建立含圖片的 Run
        let run = Run.withDrawing(drawing)

        // 建立段落
        let para = Paragraph(runs: [run])

        // 插入段落
        if let index = paragraphIndex {
            insertParagraph(para, at: index)
        } else {
            appendParagraph(para)
        }

        // append/insert already mark document.xml; new image bumps media + rels + content_types.
        modifiedParts.insert("word/media/\(imageRef.fileName)")
        modifiedParts.insert("word/_rels/document.xml.rels")
        modifiedParts.insert("[Content_Types].xml")
        return imageId
    }

    /// 從檔案路徑插入圖片
    public mutating func insertImage(
        path: String,
        widthPx: Int,
        heightPx: Int,
        at paragraphIndex: Int? = nil,
        name: String = "Picture",
        description: String = ""
    ) throws -> String {
        let imageId = nextImageRelationshipId

        // 建立圖片參照
        let imageRef = try ImageReference.from(path: path, id: imageId)
        images.append(imageRef)

        // 建立 Drawing
        var drawing = Drawing.from(widthPx: widthPx, heightPx: heightPx, imageId: imageId, name: name)
        drawing.description = description

        // 建立含圖片的 Run
        let run = Run.withDrawing(drawing)

        // 建立段落
        let para = Paragraph(runs: [run])

        // 插入段落
        if let index = paragraphIndex {
            insertParagraph(para, at: index)
        } else {
            appendParagraph(para)
        }

        modifiedParts.insert("word/media/\(imageRef.fileName)")
        modifiedParts.insert("word/_rels/document.xml.rels")
        modifiedParts.insert("[Content_Types].xml")
        return imageId
    }

    /// 更新圖片大小
    public mutating func updateImage(imageId: String, widthPx: Int? = nil, heightPx: Int? = nil) throws {
        // 搜尋所有段落找到含有此圖片的 Run
        for i in 0..<body.children.count {
            if case .paragraph(var para) = body.children[i] {
                for j in 0..<para.runs.count {
                    if var drawing = para.runs[j].drawing, drawing.imageId == imageId {
                        if let w = widthPx {
                            drawing.width = w * 9525  // 轉換為 EMU
                        }
                        if let h = heightPx {
                            drawing.height = h * 9525
                        }
                        para.runs[j].drawing = drawing
                        body.children[i] = .paragraph(para)
                        modifiedParts.insert("word/document.xml")
                        return
                    }
                }
            }
        }
        throw WordError.invalidFormat("Image '\(imageId)' not found")
    }

    /// 設定圖片樣式
    public mutating func setImageStyle(
        imageId: String,
        hasBorder: Bool? = nil,
        borderColor: String? = nil,
        borderWidth: Int? = nil,
        hasShadow: Bool? = nil
    ) throws {
        // 搜尋所有段落找到含有此圖片的 Run
        for i in 0..<body.children.count {
            if case .paragraph(var para) = body.children[i] {
                for j in 0..<para.runs.count {
                    if var drawing = para.runs[j].drawing, drawing.imageId == imageId {
                        if let border = hasBorder {
                            drawing.hasBorder = border
                        }
                        if let color = borderColor {
                            drawing.borderColor = color
                        }
                        if let width = borderWidth {
                            drawing.borderWidth = width
                        }
                        if let shadow = hasShadow {
                            drawing.hasShadow = shadow
                        }
                        para.runs[j].drawing = drawing
                        body.children[i] = .paragraph(para)
                        modifiedParts.insert("word/document.xml")
                        return
                    }
                }
            }
        }
        throw WordError.invalidFormat("Image '\(imageId)' not found")
    }

    /// 刪除圖片
    public mutating func deleteImage(imageId: String) throws {
        // 移除圖片資源
        guard let resourceIndex = images.firstIndex(where: { $0.id == imageId }) else {
            throw WordError.invalidFormat("Image '\(imageId)' not found")
        }
        images.remove(at: resourceIndex)

        // 搜尋並移除含有此圖片的 Run
        for i in 0..<body.children.count {
            if case .paragraph(var para) = body.children[i] {
                para.runs.removeAll { $0.drawing?.imageId == imageId }
                body.children[i] = .paragraph(para)
            }
        }
        modifiedParts.insert("word/document.xml")
        modifiedParts.insert("word/_rels/document.xml.rels")
        modifiedParts.insert("[Content_Types].xml")
    }

    /// 取得所有圖片資訊
    public func getImages() -> [(id: String, fileName: String, widthPx: Int, heightPx: Int)] {
        var result: [(id: String, fileName: String, widthPx: Int, heightPx: Int)] = []

        // 從 images 和 body 中收集圖片資訊
        for image in images {
            // 找對應的 Drawing 來取得尺寸
            var widthPx = 0
            var heightPx = 0

            for child in body.children {
                if case .paragraph(let para) = child {
                    for run in para.runs {
                        if let drawing = run.drawing, drawing.imageId == image.id {
                            widthPx = drawing.widthInPixels
                            heightPx = drawing.heightInPixels
                            break
                        }
                    }
                }
            }

            result.append((id: image.id, fileName: image.fileName, widthPx: widthPx, heightPx: heightPx))
        }

        return result
    }

    // MARK: - Hyperlink Operations

    /// 取得下一個可用的超連結關係 ID
    private var nextHyperlinkRelationshipId: String {
        // 基本 ID 從 rId4 開始
        let baseId = numbering.abstractNums.isEmpty ? 4 : 5
        let usedCount = headers.count + footers.count + images.count + hyperlinkReferences.count
        return "rId\(baseId + usedCount)"
    }

    /// 插入外部超連結
    public mutating func insertHyperlink(
        url: String,
        text: String,
        at paragraphIndex: Int? = nil,
        tooltip: String? = nil
    ) -> String {
        let hyperlinkId = "hyperlink_\(nextHyperlinkId)"
        nextHyperlinkId += 1

        let relationshipId = nextHyperlinkRelationshipId

        // 建立超連結關係（用於 .rels 檔案）
        let reference = HyperlinkReference(relationshipId: relationshipId, url: url)
        hyperlinkReferences.append(reference)

        // 建立超連結
        let hyperlink = Hyperlink.external(
            id: hyperlinkId,
            text: text,
            url: url,
            relationshipId: relationshipId,
            tooltip: tooltip
        )

        // 如果指定了段落索引，加到該段落；否則建立新段落
        if let index = paragraphIndex {
            let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
                if case .paragraph = child { return i }
                return nil
            }

            if index >= 0 && index < paragraphIndices.count {
                let actualIndex = paragraphIndices[index]
                if case .paragraph(var para) = body.children[actualIndex] {
                    para.hyperlinks.append(hyperlink)
                    body.children[actualIndex] = .paragraph(para)
                }
            }
        } else {
            // 建立新段落包含超連結
            var para = Paragraph()
            para.hyperlinks.append(hyperlink)
            appendParagraph(para)
        }

        modifiedParts.insert("word/document.xml")
        modifiedParts.insert("word/_rels/document.xml.rels")
        return hyperlinkId
    }

    /// 插入內部連結（連結到書籤）
    public mutating func insertInternalLink(
        bookmarkName: String,
        text: String,
        at paragraphIndex: Int? = nil,
        tooltip: String? = nil
    ) -> String {
        let hyperlinkId = "hyperlink_\(nextHyperlinkId)"
        nextHyperlinkId += 1

        // 內部連結不需要 relationship
        let hyperlink = Hyperlink.internal(
            id: hyperlinkId,
            text: text,
            bookmarkName: bookmarkName,
            tooltip: tooltip
        )

        if let index = paragraphIndex {
            let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
                if case .paragraph = child { return i }
                return nil
            }

            if index >= 0 && index < paragraphIndices.count {
                let actualIndex = paragraphIndices[index]
                if case .paragraph(var para) = body.children[actualIndex] {
                    para.hyperlinks.append(hyperlink)
                    body.children[actualIndex] = .paragraph(para)
                }
            }
        } else {
            var para = Paragraph()
            para.hyperlinks.append(hyperlink)
            appendParagraph(para)
        }

        modifiedParts.insert("word/document.xml")
        return hyperlinkId
    }

    /// 更新超連結
    public mutating func updateHyperlink(hyperlinkId: String, text: String? = nil, url: String? = nil) throws {
        // v0.19.5+ (#56 R5 P1 #3): walk every paragraph surface across body
        // (incl. tables / nested tables / SDT children), headers, footers,
        // footnotes, endnotes — not just `body.children[i].paragraph`. Pre-fix
        // any hyperlink living inside a header table cell (or any container
        // table) raised `invalidFormat("Hyperlink ... not found")` even though
        // it existed and `list_hyperlinks` could see it. (DA-N4 from R4.)
        // v0.19.5+ (#56 R5-CONT P1 #8): capture the rId during the typed
        // mutation walk so the URL sync can target the OWNING part's rels
        // file (header*.xml.rels / footer*.xml.rels / footnotes.xml.rels /
        // endnotes.xml.rels / document.xml.rels), not always document-scope.
        var capturedRId: String? = nil
        let mutator: (inout Hyperlink) -> Void = { hyperlink in
            if let newText = text {
                hyperlink.text = newText
            }
            if let newUrl = url {
                hyperlink.url = newUrl
            }
            capturedRId = hyperlink.relationshipId
        }
        guard let partKey = applyToHyperlink(id: hyperlinkId, apply: mutator) else {
            throw WordError.invalidFormat("Hyperlink '\(hyperlinkId)' not found")
        }
        modifiedParts.insert(partKey)

        // Sync the matching Relationship URL into the owning part's rels.
        if let newUrl = url, let rId = capturedRId {
            let relsKey = relsPartKey(forBodyPartKey: partKey)
            updateHyperlinkRelTarget(rId: rId, newUrl: newUrl, partKey: partKey)
            modifiedParts.insert(relsKey)
        }
    }

    /// v0.19.5+ (#56 R5-CONT P1 #8): map a body/container part key
    /// (`word/header1.xml`) to its corresponding rels file key
    /// (`word/_rels/header1.xml.rels`). Mirrors the OOXML convention where
    /// `word/<part>.xml` has rels at `word/_rels/<part>.xml.rels`.
    private func relsPartKey(forBodyPartKey partKey: String) -> String {
        // Strip "word/" prefix, append ".rels" inside _rels/.
        let prefix = "word/"
        if partKey.hasPrefix(prefix) {
            let suffix = String(partKey.dropFirst(prefix.count))
            return "word/_rels/\(suffix).rels"
        }
        return "word/_rels/document.xml.rels"
    }

    /// v0.19.5+ (#56 R5-CONT P1 #8): update the hyperlink relationship's
    /// target URL inside the owning part's relationships collection.
    /// Body uses document-scope `hyperlinkReferences`; containers use
    /// their own `relationships`.
    private mutating func updateHyperlinkRelTarget(rId: String, newUrl: String, partKey: String) {
        switch partKey {
        case "word/document.xml":
            if let idx = hyperlinkReferences.firstIndex(where: { $0.relationshipId == rId }) {
                hyperlinkReferences[idx].url = newUrl
            }
        case "word/footnotes.xml":
            updateRelTarget(in: &footnotes.relationships, rId: rId, newUrl: newUrl)
        case "word/endnotes.xml":
            updateRelTarget(in: &endnotes.relationships, rId: rId, newUrl: newUrl)
        default:
            // Header / footer paths: word/header*.xml or word/footer*.xml.
            for hi in 0..<headers.count where DocumentWalker.headerPartKey(for: headers[hi]) == partKey {
                updateRelTarget(in: &headers[hi].relationships, rId: rId, newUrl: newUrl)
            }
            for fi in 0..<footers.count where DocumentWalker.footerPartKey(for: footers[fi]) == partKey {
                updateRelTarget(in: &footers[fi].relationships, rId: rId, newUrl: newUrl)
            }
        }
    }

    private func updateRelTarget(in collection: inout RelationshipsCollection, rId: String, newUrl: String) {
        if let idx = collection.relationships.firstIndex(where: { $0.id == rId }) {
            collection.relationships[idx].target = newUrl
        }
    }

    /// 刪除超連結
    public mutating func deleteHyperlink(hyperlinkId: String) throws {
        // v0.19.5+ (#56 R5 P1 #3): walk every part — same scope expansion as
        // updateHyperlink. Removed hyperlinks living inside headers / footers /
        // footnotes / endnotes / nested tables / SDTs no longer silently miss.
        //
        // v0.19.5+ (#56 R5-CONT-2 P0 #2): mirror `updateHyperlink(url:)`'s
        // R5-CONT P1 #8 per-container rels routing. Pre-fix this method
        // unconditionally removed from `document.hyperlinkReferences` and
        // marked `word/_rels/document.xml.rels` dirty regardless of where
        // the hyperlink lived → for container hyperlinks (whose rels live
        // in `header*.xml.rels` etc.), the container's rels collection
        // kept an orphan rel AND document.xml.rels was wrongly dirtied.
        // Now: route through the owning part's rels, mark only the
        // affected rels file dirty.
        var capturedRId: String? = nil
        guard let partKey = removeHyperlink(id: hyperlinkId, captureRelationshipId: &capturedRId) else {
            throw WordError.invalidFormat("Hyperlink '\(hyperlinkId)' not found")
        }
        modifiedParts.insert(partKey)

        if let rId = capturedRId {
            let relsKey = relsPartKey(forBodyPartKey: partKey)
            removeHyperlinkRelTarget(rId: rId, partKey: partKey)
            modifiedParts.insert(relsKey)
        }
    }

    /// v0.19.5+ (#56 R5-CONT-2 P0 #2): remove the hyperlink relationship
    /// from the owning part's relationships. Body uses document-scope
    /// `hyperlinkReferences`; containers use their own `relationships`.
    /// Mirrors `updateHyperlinkRelTarget` (R5-CONT P1 #8) for the delete
    /// side.
    private mutating func removeHyperlinkRelTarget(rId: String, partKey: String) {
        switch partKey {
        case "word/document.xml":
            hyperlinkReferences.removeAll { $0.relationshipId == rId }
        case "word/footnotes.xml":
            footnotes.relationships.relationships.removeAll { $0.id == rId }
        case "word/endnotes.xml":
            endnotes.relationships.relationships.removeAll { $0.id == rId }
        default:
            for hi in 0..<headers.count where DocumentWalker.headerPartKey(for: headers[hi]) == partKey {
                headers[hi].relationships.relationships.removeAll { $0.id == rId }
            }
            for fi in 0..<footers.count where DocumentWalker.footerPartKey(for: footers[fi]) == partKey {
                footers[fi].relationships.relationships.removeAll { $0.id == rId }
            }
        }

        // v0.19.5+ (#56 R5-CONT-3 P1 #4): defensive sweep of legacy
        // document-scope `hyperlinkReferences` for the rId. Pre-R5-CONT
        // P1 #8 introduced per-container rels, but documents migrated
        // from the older single-rels model may still carry the same rId
        // in document.hyperlinkReferences (legitimate when caller
        // historically used document-scope before the migration). Without
        // this sweep, container deletes leave a doc-scope orphan that
        // never gets cleaned up. Safe because rId scoping is per-part —
        // an orphan document.hyperlinkReferences entry for an rId that
        // matches a container's deleted rel is the migration case.
        if partKey != "word/document.xml" {
            hyperlinkReferences.removeAll { $0.relationshipId == rId }
        }
    }

    // MARK: - §8.3 helpers (v0.19.5+ #56 R5 P1 #3)

    /// Walks every paragraph surface across all parts, applying `apply` to the
    /// first hyperlink whose id matches. Returns the part key (`word/document.xml`,
    /// `word/header*.xml`, etc.) on hit, nil on miss. Body recurses into
    /// tables (incl. nested) + content-control children. Headers / footers /
    /// footnotes / endnotes recurse into their `bodyChildren` tables.
    private mutating func applyToHyperlink(
        id hyperlinkId: String,
        apply: (inout Hyperlink) -> Void
    ) -> String? {
        // Body
        for i in 0..<body.children.count {
            if Self.applyHyperlinkInBodyChild(&body.children[i], id: hyperlinkId, apply: apply) {
                return "word/document.xml"
            }
        }
        // Headers
        for i in 0..<headers.count {
            for j in 0..<headers[i].bodyChildren.count {
                if Self.applyHyperlinkInBodyChild(&headers[i].bodyChildren[j], id: hyperlinkId, apply: apply) {
                    return "word/\(headers[i].fileName)"
                }
            }
        }
        // Footers
        for i in 0..<footers.count {
            for j in 0..<footers[i].bodyChildren.count {
                if Self.applyHyperlinkInBodyChild(&footers[i].bodyChildren[j], id: hyperlinkId, apply: apply) {
                    return "word/\(footers[i].fileName)"
                }
            }
        }
        // Footnotes
        for i in 0..<footnotes.footnotes.count {
            for j in 0..<footnotes.footnotes[i].bodyChildren.count {
                if Self.applyHyperlinkInBodyChild(&footnotes.footnotes[i].bodyChildren[j], id: hyperlinkId, apply: apply) {
                    return "word/footnotes.xml"
                }
            }
        }
        // Endnotes
        for i in 0..<endnotes.endnotes.count {
            for j in 0..<endnotes.endnotes[i].bodyChildren.count {
                if Self.applyHyperlinkInBodyChild(&endnotes.endnotes[i].bodyChildren[j], id: hyperlinkId, apply: apply) {
                    return "word/endnotes.xml"
                }
            }
        }
        return nil
    }

    private static func applyHyperlinkInBodyChild(
        _ child: inout BodyChild,
        id hyperlinkId: String,
        apply: (inout Hyperlink) -> Void
    ) -> Bool {
        switch child {
        case .bookmarkMarker, .rawBlockElement:
            // Body-level markers carry no hyperlinks (#58).
            return false
        case .paragraph(var para):
            if let idx = para.hyperlinks.firstIndex(where: { $0.id == hyperlinkId }) {
                apply(&para.hyperlinks[idx])
                child = .paragraph(para)
                return true
            }
            return false
        case .table(var table):
            for rowIdx in 0..<table.rows.count {
                for cellIdx in 0..<table.rows[rowIdx].cells.count {
                    for paraIdx in 0..<table.rows[rowIdx].cells[cellIdx].paragraphs.count {
                        var para = table.rows[rowIdx].cells[cellIdx].paragraphs[paraIdx]
                        if let idx = para.hyperlinks.firstIndex(where: { $0.id == hyperlinkId }) {
                            apply(&para.hyperlinks[idx])
                            table.rows[rowIdx].cells[cellIdx].paragraphs[paraIdx] = para
                            child = .table(table)
                            return true
                        }
                    }
                    for nestedIdx in 0..<table.rows[rowIdx].cells[cellIdx].nestedTables.count {
                        var nested = BodyChild.table(table.rows[rowIdx].cells[cellIdx].nestedTables[nestedIdx])
                        if Self.applyHyperlinkInBodyChild(&nested, id: hyperlinkId, apply: apply) {
                            if case .table(let mutatedNested) = nested {
                                table.rows[rowIdx].cells[cellIdx].nestedTables[nestedIdx] = mutatedNested
                            }
                            child = .table(table)
                            return true
                        }
                    }
                }
            }
            return false
        case .contentControl(let metadata, var inner):
            for k in 0..<inner.count {
                if Self.applyHyperlinkInBodyChild(&inner[k], id: hyperlinkId, apply: apply) {
                    child = .contentControl(metadata, children: inner)
                    return true
                }
            }
            return false
        }
    }

    /// Walks every paragraph surface across all parts; on the first match,
    /// removes the hyperlink, captures its relationshipId via `inout`, and
    /// returns the owning part key. Symmetric with `applyToHyperlink`.
    private mutating func removeHyperlink(
        id hyperlinkId: String,
        captureRelationshipId: inout String?
    ) -> String? {
        for i in 0..<body.children.count {
            if Self.removeHyperlinkRecursive(in: &body.children[i], id: hyperlinkId,
                                             captureRelationshipId: &captureRelationshipId) {
                return "word/document.xml"
            }
        }
        for i in 0..<headers.count {
            for j in 0..<headers[i].bodyChildren.count {
                if Self.removeHyperlinkRecursive(in: &headers[i].bodyChildren[j], id: hyperlinkId,
                                                 captureRelationshipId: &captureRelationshipId) {
                    return "word/\(headers[i].fileName)"
                }
            }
        }
        for i in 0..<footers.count {
            for j in 0..<footers[i].bodyChildren.count {
                if Self.removeHyperlinkRecursive(in: &footers[i].bodyChildren[j], id: hyperlinkId,
                                                 captureRelationshipId: &captureRelationshipId) {
                    return "word/\(footers[i].fileName)"
                }
            }
        }
        for i in 0..<footnotes.footnotes.count {
            for j in 0..<footnotes.footnotes[i].bodyChildren.count {
                if Self.removeHyperlinkRecursive(in: &footnotes.footnotes[i].bodyChildren[j], id: hyperlinkId,
                                                 captureRelationshipId: &captureRelationshipId) {
                    return "word/footnotes.xml"
                }
            }
        }
        for i in 0..<endnotes.endnotes.count {
            for j in 0..<endnotes.endnotes[i].bodyChildren.count {
                if Self.removeHyperlinkRecursive(in: &endnotes.endnotes[i].bodyChildren[j], id: hyperlinkId,
                                                 captureRelationshipId: &captureRelationshipId) {
                    return "word/endnotes.xml"
                }
            }
        }
        return nil
    }

    private static func removeHyperlinkRecursive(
        in child: inout BodyChild,
        id hyperlinkId: String,
        captureRelationshipId: inout String?
    ) -> Bool {
        switch child {
        case .bookmarkMarker, .rawBlockElement:
            // Body-level markers carry no hyperlinks (#58).
            return false
        case .paragraph(var para):
            if let idx = para.hyperlinks.firstIndex(where: { $0.id == hyperlinkId }) {
                captureRelationshipId = para.hyperlinks[idx].relationshipId
                para.hyperlinks.remove(at: idx)
                child = .paragraph(para)
                return true
            }
            return false
        case .table(var table):
            for rowIdx in 0..<table.rows.count {
                for cellIdx in 0..<table.rows[rowIdx].cells.count {
                    for paraIdx in 0..<table.rows[rowIdx].cells[cellIdx].paragraphs.count {
                        var para = table.rows[rowIdx].cells[cellIdx].paragraphs[paraIdx]
                        if let idx = para.hyperlinks.firstIndex(where: { $0.id == hyperlinkId }) {
                            captureRelationshipId = para.hyperlinks[idx].relationshipId
                            para.hyperlinks.remove(at: idx)
                            table.rows[rowIdx].cells[cellIdx].paragraphs[paraIdx] = para
                            child = .table(table)
                            return true
                        }
                    }
                    for nestedIdx in 0..<table.rows[rowIdx].cells[cellIdx].nestedTables.count {
                        var nested = BodyChild.table(table.rows[rowIdx].cells[cellIdx].nestedTables[nestedIdx])
                        if removeHyperlinkRecursive(in: &nested, id: hyperlinkId, captureRelationshipId: &captureRelationshipId) {
                            if case .table(let mutatedNested) = nested {
                                table.rows[rowIdx].cells[cellIdx].nestedTables[nestedIdx] = mutatedNested
                            }
                            child = .table(table)
                            return true
                        }
                    }
                }
            }
            return false
        case .contentControl(let metadata, var inner):
            for k in 0..<inner.count {
                if removeHyperlinkRecursive(in: &inner[k], id: hyperlinkId, captureRelationshipId: &captureRelationshipId) {
                    child = .contentControl(metadata, children: inner)
                    return true
                }
            }
            return false
        }
    }

    /// Lookup helper for updateHyperlink URL→relationship sync. Returns the
    /// hyperlink's relationshipId if found in any part, nil otherwise.
    private func findRelationshipId(forHyperlinkId hyperlinkId: String) -> String? {
        var match: String? = nil
        DocumentWalker.walkAllParagraphs(in: self) { para, _ in
            if match != nil { return }
            if let h = para.hyperlinks.first(where: { $0.id == hyperlinkId }) {
                match = h.relationshipId
            }
        }
        return match
    }

    /// 列出所有超連結
    /// v0.19.5+ (#56 R5-CONT P0 #7): walks every paragraph across body
    /// (incl. nested tables / SDT children), headers, footers, footnotes,
    /// endnotes via `DocumentWalker.walkAllParagraphs`. Pre-fix only body
    /// top-level paragraphs were listed → the listed-id set was a strict
    /// subset of what `updateHyperlink` / `deleteHyperlink` (R5 P1 #3)
    /// could find, so callers couldn't programmatically discover ids
    /// they were allowed to mutate. Verify R5 P0 #7 (DA C5).
    public func getHyperlinks() -> [(id: String, text: String, url: String?, anchor: String?, type: String)] {
        var result: [(id: String, text: String, url: String?, anchor: String?, type: String)] = []
        DocumentWalker.walkAllParagraphs(in: self) { para, _ in
            for hyperlink in para.hyperlinks {
                let typeStr = hyperlink.type == .external ? "external" : "internal"
                result.append((
                    id: hyperlink.id,
                    text: hyperlink.text,
                    url: hyperlink.url,
                    anchor: hyperlink.anchor,
                    type: typeStr
                ))
            }
        }
        return result
    }

    // MARK: - Bookmark Operations

    /// 插入書籤
    public mutating func insertBookmark(
        name: String,
        at paragraphIndex: Int? = nil
    ) throws -> Int {
        // 驗證書籤名稱
        let normalizedName = Bookmark.normalizeName(name)
        guard Bookmark.validateName(normalizedName) else {
            throw BookmarkError.invalidName(name)
        }

        // 檢查是否已存在同名書籤 (cross-part scope as of v0.19.9+ / #58 A-CONT-3 P0 #3).
        // Pre-A-CONT-3 only walked body — TOC anchors in headers / footers /
        // footnotes / endnotes survived `insertBookmark(name: ...)` and produced
        // silent name collisions. Now walks all 5 part types via the same
        // collector helper used by getBookmarks().
        let allExistingNames = Set(getBookmarks().map { $0.name })
        if allExistingNames.contains(normalizedName) {
            throw BookmarkError.duplicateName(normalizedName)
        }

        let bookmarkId = nextBookmarkId
        nextBookmarkId += 1

        let bookmark = Bookmark(id: bookmarkId, name: normalizedName)

        if let index = paragraphIndex {
            let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
                if case .paragraph = child { return i }
                return nil
            }

            guard index >= 0 && index < paragraphIndices.count else {
                throw WordError.invalidIndex(index)
            }

            let actualIndex = paragraphIndices[index]
            if case .paragraph(var para) = body.children[actualIndex] {
                Self.appendBookmarkSyncingMarkers(to: &para, bookmark: bookmark)
                body.children[actualIndex] = .paragraph(para)
            }
        } else {
            // 加到最後一個段落，如果沒有段落則建立新的
            if let lastIndex = body.children.lastIndex(where: {
                if case .paragraph = $0 { return true }
                return false
            }) {
                if case .paragraph(var para) = body.children[lastIndex] {
                    Self.appendBookmarkSyncingMarkers(to: &para, bookmark: bookmark)
                    body.children[lastIndex] = .paragraph(para)
                }
            } else {
                var para = Paragraph()
                Self.appendBookmarkSyncingMarkers(to: &para, bookmark: bookmark)
                appendParagraph(para)
            }
        }

        modifiedParts.insert("word/document.xml")
        return bookmarkId
    }

    /// 刪除書籤
    public mutating func deleteBookmark(name: String) throws {
        // Body — paragraph-level bookmarks first.
        for i in 0..<body.children.count {
            if case .paragraph(var para) = body.children[i] {
                if let index = para.bookmarks.firstIndex(where: { $0.name == name }) {
                    let removed = para.bookmarks.remove(at: index)
                    // v0.19.2+ (#56 follow-up F2): keep `bookmarkMarkers` in sync
                    // so source-loaded paragraphs (which always go through the
                    // sort-by-position emit path) don't leave behind zombie
                    // `<w:bookmarkStart>` markers with empty `w:name=""` from
                    // the markers-only side of the model.
                    para.bookmarkMarkers.removeAll { $0.id == removed.id }
                    body.children[i] = .paragraph(para)
                    modifiedParts.insert("word/document.xml")
                    return
                }
            }
        }

        // v0.19.8+ (#58 A-CONT-2): try body-level `.bookmarkMarker` entries +
        // container body-level markers (headers / footers / footnotes / endnotes).
        // Pre-A-CONT-2 `getBookmarks()` could list these names but `deleteBookmark`
        // couldn't delete them (always threw `notFound`). Now they're symmetric.
        if Self.tryDeleteBodyLevelBookmark(name: name, from: &body.children) {
            modifiedParts.insert("word/document.xml")
            return
        }
        for i in 0..<headers.count {
            if Self.tryDeleteBodyLevelBookmark(name: name, from: &headers[i].bodyChildren) {
                // v0.19.9+ (#58 A-CONT-3 P0 #1): writer's overlay-mode dirty-gate
                // (DocxWriter.swift:141) checks `dirty.contains("word/\(header.fileName)")`,
                // so we must insert the FULL PATH not the basename. Pre-A-CONT-3
                // this inserted `headers[i].fileName` (basename `"header1.xml"`)
                // which silently failed the dirty-gate → deletion lost on disk.
                modifiedParts.insert("word/\(headers[i].fileName)")
                return
            }
        }
        for i in 0..<footers.count {
            if Self.tryDeleteBodyLevelBookmark(name: name, from: &footers[i].bodyChildren) {
                // v0.19.9+ (#58 A-CONT-3 P0 #1): same correctness fix as headers.
                modifiedParts.insert("word/\(footers[i].fileName)")
                return
            }
        }
        for i in 0..<footnotes.footnotes.count {
            if Self.tryDeleteBodyLevelBookmark(name: name, from: &footnotes.footnotes[i].bodyChildren) {
                modifiedParts.insert("word/footnotes.xml")
                return
            }
        }
        for i in 0..<endnotes.endnotes.count {
            if Self.tryDeleteBodyLevelBookmark(name: name, from: &endnotes.endnotes[i].bodyChildren) {
                modifiedParts.insert("word/endnotes.xml")
                return
            }
        }

        throw BookmarkError.notFound(name)
    }

    /// v0.19.8+ (#58 A-CONT-2): helper for `deleteBookmark`. Scans body-level
    /// `.bookmarkMarker` entries (and recurses into block-level `.contentControl`)
    /// for a `.start` marker matching the given name. On hit, removes both the
    /// `.start` marker and any matching `.end` marker (paired by id) from the
    /// children array. Returns true on hit, false otherwise.
    private static func tryDeleteBodyLevelBookmark(name: String, from children: inout [BodyChild]) -> Bool {
        // First pass — find the matching .start to learn its id.
        var matchedId: Int?
        for i in 0..<children.count {
            if case .bookmarkMarker(let marker) = children[i],
               marker.kind == .start,
               marker.name == name {
                matchedId = marker.id
                break
            }
        }
        if let id = matchedId {
            // Remove BOTH .start (by name+id) and any .end (by id only).
            children.removeAll { child in
                if case .bookmarkMarker(let marker) = child, marker.id == id {
                    return true
                }
                return false
            }
            return true
        }
        // Recurse into block-level SDTs.
        for i in 0..<children.count {
            if case .contentControl(let meta, var inner) = children[i] {
                if tryDeleteBodyLevelBookmark(name: name, from: &inner) {
                    children[i] = .contentControl(meta, children: inner)
                    return true
                }
            }
        }
        return false
    }

    /// v0.19.2+ (#56 follow-up F2): central helper for adding a `Bookmark`
    /// to a paragraph that ALSO appends the matching `BookmarkRangeMarker`
    /// pair (start + end) needed by the sort-by-position emit path.
    ///
    /// Without this sync, source-loaded paragraphs (which trigger sort-by-
    /// position because their existing markers are non-empty) silently drop
    /// any newly-added bookmarks at save time — the typed `bookmarks` entry
    /// is created but the writer only emits markers, never the typed list.
    ///
    /// **Position assignment**: the new markers land just after every
    /// existing positioned child (max position + 1 for start, +2 for end).
    /// This produces a zero-width point bookmark sitting at the paragraph
    /// tail, which matches the API semantics ("name this paragraph") since
    /// `addBookmark` doesn't accept a span argument. Callers needing a
    /// span-bookmark across specific runs should construct
    /// `BookmarkRangeMarker` explicitly via the typed model.
    ///
    /// v0.19.3+ (#56 round 2 P1-4): the marker pair is ONLY appended when
    /// the paragraph already routes to the sort-by-position emit path
    /// (`hasSourcePositionedChildren == true`). For pure API-built paragraphs
    /// the legacy `bookmarks`-only emit path runs a wrap-around pattern
    /// (`<w:bookmarkStart/><w:r>text</w:r><w:bookmarkEnd/>`), preserving the
    /// v3.12.0 semantic that `addBookmark` spans the existing run text. F2
    /// blindly added markers everywhere, downgrading API-path bookmarks to
    /// zero-width point bookmarks at paragraph end — silent behavioral change
    /// for callers expecting span semantics.
    private static func appendBookmarkSyncingMarkers(
        to paragraph: inout Paragraph,
        bookmark: Bookmark
    ) {
        paragraph.bookmarks.append(bookmark)

        // P1-4 routing: only sync markers when the paragraph is already on
        // the sort-by-position path. Pure API paragraphs (no markers / no
        // raw-carriers / no positioned runs/hyperlinks) keep the legacy
        // wrap-around bookmark emit shape via `bookmarks` alone.
        guard paragraph.hasSourcePositionedChildren else { return }

        // Compute the next free position after every existing positioned child.
        // Mirrors the collections enumerated by `Paragraph.toXMLSortedByPosition`.
        // PsychQuant/ooxml-swift#5 (F6): position is now `Int? = nil`. Treat
        // nil as "no explicit position" by mapping to -1, which is below any
        // valid position; max() then ignores nil contributions naturally.
        var positions: [Int] = []
        positions.append(contentsOf: paragraph.runs.compactMap { $0.position })
        positions.append(contentsOf: paragraph.hyperlinks.compactMap { $0.position })
        positions.append(contentsOf: paragraph.fieldSimples.compactMap { $0.position })
        positions.append(contentsOf: paragraph.alternateContents.compactMap { $0.position })
        positions.append(contentsOf: paragraph.bookmarkMarkers.compactMap { $0.position })
        positions.append(contentsOf: paragraph.commentRangeMarkers.compactMap { $0.position })
        positions.append(contentsOf: paragraph.permissionRangeMarkers.compactMap { $0.position })
        positions.append(contentsOf: paragraph.proofErrorMarkers.compactMap { $0.position })
        positions.append(contentsOf: paragraph.smartTags.compactMap { $0.position })
        positions.append(contentsOf: paragraph.customXmlBlocks.compactMap { $0.position })
        positions.append(contentsOf: paragraph.bidiOverrides.compactMap { $0.position })
        positions.append(contentsOf: paragraph.unrecognizedChildren.compactMap { $0.position })
        let nextPosition = (positions.max() ?? -1) + 1

        paragraph.bookmarkMarkers.append(
            BookmarkRangeMarker(kind: .start, id: bookmark.id, position: nextPosition)
        )
        paragraph.bookmarkMarkers.append(
            BookmarkRangeMarker(kind: .end, id: bookmark.id, position: nextPosition + 1)
        )
    }

    /// 列出所有書籤
    public func getBookmarks() -> [(id: Int, name: String, paragraphIndex: Int)] {
        var result: [(id: Int, name: String, paragraphIndex: Int)] = []
        var paragraphCount = 0

        // Body — paragraph-level + body-level markers (with paragraph index).
        for child in body.children {
            switch child {
            case .paragraph(let para):
                for bookmark in para.bookmarks {
                    result.append((id: bookmark.id, name: bookmark.name, paragraphIndex: paragraphCount))
                }
                paragraphCount += 1
            case .bookmarkMarker(let marker):
                // v0.19.7+ (#58 A-CONT): surface body-level bookmark starts
                // (TOC `_Toc<digits>` anchors that wrap multiple paragraphs).
                // Only `.start` markers carry a name; `.end` markers are
                // matched by id and don't represent a separate bookmark.
                // paragraphIndex = -1 sentinel indicates "not inside a paragraph"
                // (the marker sits at body level, between or wrapping paragraphs).
                if marker.kind == .start, let name = marker.name {
                    result.append((id: marker.id, name: name, paragraphIndex: -1))
                }
            case .contentControl(_, let inner):
                // v0.19.8+ (#58 A-CONT-2): recurse into block-level SDT to
                // surface body-level markers nested inside content controls.
                Self.collectBodyLevelBookmarkNamesRecursive(in: inner, into: &result)
            case .table, .rawBlockElement:
                // Tables and raw block elements don't carry typed bookmarks
                // visible to the API. (Bookmarks inside table cells are picked
                // up via the per-paragraph walker on `cell.paragraphs[].bookmarks`
                // — but that lives in the paragraph-level path, not here.
                // Raw block elements are opaque XML by design.)
                break
            }
        }

        // v0.19.8+ (#58 A-CONT-2): walk body-level markers in container parts
        // (headers / footers / footnotes / endnotes). Pre-A-CONT-2 these were
        // preserved on disk but invisible to MCP `list_bookmarks`. Container
        // markers carry no paragraph index in the body-document sense, so
        // they share the `paragraphIndex = -1` sentinel.
        // v0.19.9+ (#58 A-CONT-3 P0 #2): extended to ALSO walk paragraph-level
        // bookmarks inside container paragraphs. Pre-A-CONT-3 only body-level
        // container markers were surfaced; paragraph-level markers (the more
        // common case in real-world Word docs) were silently skipped.
        for header in headers {
            Self.collectAllBookmarksFromContainer(in: header.bodyChildren, into: &result)
        }
        for footer in footers {
            Self.collectAllBookmarksFromContainer(in: footer.bodyChildren, into: &result)
        }
        for footnote in footnotes.footnotes {
            Self.collectAllBookmarksFromContainer(in: footnote.bodyChildren, into: &result)
        }
        for endnote in endnotes.endnotes {
            Self.collectAllBookmarksFromContainer(in: endnote.bodyChildren, into: &result)
        }

        return result
    }

    /// v0.19.9+ (#58 A-CONT-3 P0 #2): unified container walker that surfaces
    /// BOTH paragraph-level bookmarks (`Paragraph.bookmarks`) AND body-level
    /// `.bookmarkMarker` entries. Recurses into block-level `.contentControl`.
    /// All container bookmarks share `paragraphIndex = -1` because the
    /// body-document paragraph index doesn't apply across part boundaries.
    private static func collectAllBookmarksFromContainer(
        in children: [BodyChild],
        into result: inout [(id: Int, name: String, paragraphIndex: Int)]
    ) {
        for child in children {
            switch child {
            case .paragraph(let para):
                for bookmark in para.bookmarks {
                    result.append((id: bookmark.id, name: bookmark.name, paragraphIndex: -1))
                }
            case .bookmarkMarker(let marker):
                if marker.kind == .start, let name = marker.name {
                    result.append((id: marker.id, name: name, paragraphIndex: -1))
                }
            case .contentControl(_, let inner):
                collectAllBookmarksFromContainer(in: inner, into: &result)
            case .table, .rawBlockElement:
                // Tables in containers theoretically contain paragraphs with
                // bookmarks, but that nesting is uncommon and out of A-CONT-3
                // scope. Raw block elements are opaque XML by design.
                continue
            }
        }
    }

    /// v0.19.8+ (#58 A-CONT-2): recursive helper for body-level bookmark name
    /// extraction. Used by `getBookmarks()` and `deleteBookmark(name:)` to walk
    /// body-level markers across body + container parts uniformly. Recurses
    /// into block-level `.contentControl(_, let inner)` children so SDT-nested
    /// markers are surfaced. paragraph-level bookmarks (inside `.paragraph(...)`)
    /// are intentionally NOT collected here — they live in `Paragraph.bookmarks`
    /// and use a different paragraph-index semantic.
    private static func collectBodyLevelBookmarkNamesRecursive(
        in children: [BodyChild],
        into result: inout [(id: Int, name: String, paragraphIndex: Int)]
    ) {
        for child in children {
            switch child {
            case .bookmarkMarker(let marker):
                if marker.kind == .start, let name = marker.name {
                    result.append((id: marker.id, name: name, paragraphIndex: -1))
                }
            case .contentControl(_, let inner):
                collectBodyLevelBookmarkNamesRecursive(in: inner, into: &result)
            case .paragraph, .table, .rawBlockElement:
                continue
            }
        }
    }

    // MARK: - Comment Operations

    /// 插入註解
    public mutating func insertComment(
        text: String,
        author: String,
        paragraphIndex: Int
    ) throws -> Int {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count else {
            throw CommentError.invalidParagraphIndex(paragraphIndex)
        }

        let commentId = comments.nextCommentId()
        let comment = Comment(
            id: commentId,
            author: author,
            text: text,
            paragraphIndex: paragraphIndex
        )

        comments.comments.append(comment)

        // 在段落中添加註解標記
        let actualIndex = paragraphIndices[paragraphIndex]
        if case .paragraph(var para) = body.children[actualIndex] {
            para.commentIds.append(commentId)
            body.children[actualIndex] = .paragraph(para)
        }

        modifiedParts.insert("word/comments.xml")
        modifiedParts.insert("word/document.xml")
        if comments.hasExtendedComments {
            modifiedParts.insert("word/commentsExtended.xml")
        }
        return commentId
    }

    /// 更新註解
    public mutating func updateComment(commentId: Int, text: String) throws {
        guard let index = comments.comments.firstIndex(where: { $0.id == commentId }) else {
            throw CommentError.notFound(commentId)
        }

        comments.comments[index].text = text
        modifiedParts.insert("word/comments.xml")
        if comments.hasExtendedComments {
            modifiedParts.insert("word/commentsExtended.xml")
        }
    }

    /// 刪除註解
    public mutating func deleteComment(commentId: Int) throws {
        guard let index = comments.comments.firstIndex(where: { $0.id == commentId }) else {
            throw CommentError.notFound(commentId)
        }

        // 從段落中移除註解標記
        let paragraphIndex = comments.comments[index].paragraphIndex
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        if paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count {
            let actualIndex = paragraphIndices[paragraphIndex]
            if case .paragraph(var para) = body.children[actualIndex] {
                para.commentIds.removeAll { $0 == commentId }
                body.children[actualIndex] = .paragraph(para)
            }
        }

        comments.comments.remove(at: index)
        modifiedParts.insert("word/comments.xml")
        modifiedParts.insert("word/document.xml")
        if comments.hasExtendedComments {
            modifiedParts.insert("word/commentsExtended.xml")
        }
    }

    /// 列出所有註解
    public func getComments() -> [(id: Int, author: String, text: String, paragraphIndex: Int, date: Date)] {
        return comments.comments.map { comment in
            (id: comment.id, author: comment.author, text: comment.text,
             paragraphIndex: comment.paragraphIndex, date: comment.date)
        }
    }

    /// 列出所有修訂（完整 Revision 結構，含 source / previousFormatDescription）
    ///
    /// 與 `getRevisions()` 不同之處：回傳完整的 `Revision` struct 而非 tuple，
    /// 讓呼叫端可以取得 `source`（body / header / footer / footnote / endnote）
    /// 和 `previousFormatDescription`（格式變更的人類可讀摘要）。
    /// 既有的 `getRevisions()` 保留不動，仍回傳原本的 tuple（僅含 body 修訂）。
    public func getRevisionsFull() -> [Revision] {
        return revisions.revisions
    }

    /// 列出所有註解（完整 Comment 結構，含 parentId / paraId / done）
    ///
    /// 與 `getComments()` 不同之處：回傳完整的 `Comment` struct 而非 tuple，
    /// 讓呼叫端可以取得 `parentId`（建立 reply 的 threading）、`paraId`、`done` 等欄位。
    /// 既有的 `getComments()` 保留不動，仍回傳原本的 tuple，不影響 downstream。
    public func getCommentsFull() -> [Comment] {
        return comments.comments
    }

    // MARK: - Track Changes Operations

    /// 啟用修訂追蹤
    public mutating func enableTrackChanges(author: String = "Unknown") {
        revisions.settings.enabled = true
        revisions.settings.author = author
        revisions.settings.dateTime = Date()
        modifiedParts.insert("word/settings.xml")
    }

    /// 停用修訂追蹤
    public mutating func disableTrackChanges() {
        revisions.settings.enabled = false
        modifiedParts.insert("word/settings.xml")
    }

    /// 檢查修訂追蹤是否啟用
    public func isTrackChangesEnabled() -> Bool {
        return revisions.settings.enabled
    }

    /// Allocate a fresh revision id (max existing + 1, or 1 when empty).
    ///
    /// Mirrors `allocateSdtId()` (v0.15.0): scans the document for the highest
    /// existing revision id and returns one greater. Because `revisions.revisions`
    /// is the single collection populated by `DocxReader` for body, headers,
    /// footers, footnotes, and endnotes, scanning that collection alone is
    /// sufficient.
    ///
    /// The method is idempotent — consecutive calls without appending the
    /// allocated id return the same value (no stale-cache risk).
    public func allocateRevisionId() -> Int {
        let maxId = revisions.revisions.map { $0.id }.max() ?? 0
        return maxId + 1
    }

    /// 3-tier author resolution: explicit arg (when non-nil and non-empty)
    /// wins; otherwise fall back to the track-changes settings author; finally
    /// "Unknown". Used by all v0.18.0 revision-generating mutations.
    fileprivate func resolveAuthor(_ explicit: String?) -> String {
        if let explicit = explicit, !explicit.isEmpty {
            return explicit
        }
        let stored = revisions.settings.author
        if !stored.isEmpty { return stored }
        return "Unknown"
    }

    // MARK: - Track Changes Generators (v0.18.0+, che-word-mcp#45)

    /// Locate the body-children index for the Nth paragraph (skipping tables and
    /// content controls). Returns nil when out of range.
    private func bodyIndexForParagraph(_ paragraphIndex: Int) -> Int? {
        var seen = 0
        for (i, child) in body.children.enumerated() {
            if case .paragraph = child {
                if seen == paragraphIndex { return i }
                seen += 1
            }
        }
        return nil
    }

    /// Insert `text` at character `position` within paragraph `paragraphIndex`,
    /// wrapping the new run with a `<w:ins>` revision. Splits an existing run
    /// when the position falls inside one (preserves the surrounding text +
    /// formatting). Returns the allocated revision id.
    ///
    /// Throws `WordError.trackChangesNotEnabled` when track changes is off
    /// (no auto-enable side effect). Throws `WordError.invalidIndex` for
    /// out-of-range paragraph index or character position.
    public mutating func insertTextAsRevision(
        text: String,
        atParagraph paragraphIndex: Int,
        position: Int,
        author: String? = nil,
        date: Date? = nil
    ) throws -> Int {
        guard isTrackChangesEnabled() else {
            throw WordError.trackChangesNotEnabled
        }
        guard let bodyIdx = bodyIndexForParagraph(paragraphIndex) else {
            throw WordError.invalidIndex(paragraphIndex)
        }
        guard case .paragraph(var paragraph) = body.children[bodyIdx] else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let totalLength = paragraph.runs.reduce(0) { $0 + $1.text.count }
        guard position >= 0, position <= totalLength else {
            throw WordError.invalidIndex(position)
        }

        let revisionId = allocateRevisionId()
        let resolvedAuthor = resolveAuthor(author)
        let resolvedDate = date ?? Date()

        var newRun = Run(text: text)
        newRun.revisionId = revisionId

        if position == 0 {
            paragraph.runs.insert(newRun, at: 0)
        } else if position == totalLength {
            paragraph.runs.append(newRun)
        } else {
            // Split the run that contains `position`.
            var charsSeen = 0
            for runIdx in 0..<paragraph.runs.count {
                let runLength = paragraph.runs[runIdx].text.count
                let charsAfter = charsSeen + runLength
                if position == charsSeen {
                    paragraph.runs.insert(newRun, at: runIdx)
                    break
                } else if position < charsAfter {
                    let offsetInRun = position - charsSeen
                    let original = paragraph.runs[runIdx]
                    let beforeText = String(original.text.prefix(offsetInRun))
                    let afterText = String(original.text.suffix(runLength - offsetInRun))

                    var beforeRun = original
                    beforeRun.text = beforeText
                    var afterRun = original
                    afterRun.text = afterText

                    paragraph.runs[runIdx] = beforeRun
                    paragraph.runs.insert(newRun, at: runIdx + 1)
                    paragraph.runs.insert(afterRun, at: runIdx + 2)
                    break
                }
                charsSeen = charsAfter
            }
        }

        let revision = Revision(
            id: revisionId,
            type: .insertion,
            author: resolvedAuthor,
            date: resolvedDate,
            content: text
        )
        paragraph.revisions.append(revision)
        revisions.revisions.append(revision)

        body.children[bodyIdx] = .paragraph(paragraph)
        modifiedParts.insert("word/document.xml")

        return revisionId
    }

    /// Mark text in `[start, end)` of paragraph `paragraphIndex` as a tracked
    /// deletion. Splits straddling runs at the boundaries so only the deleted
    /// substring carries the revision id; the writer substitutes `<w:t>` with
    /// `<w:delText>` for those runs.
    ///
    /// Cross-paragraph delete is OUT OF SCOPE — `start` and `end` are character
    /// offsets within the named paragraph's text only.
    public mutating func deleteTextAsRevision(
        atParagraph paragraphIndex: Int,
        start: Int,
        end: Int,
        author: String? = nil,
        date: Date? = nil
    ) throws -> Int {
        guard isTrackChangesEnabled() else {
            throw WordError.trackChangesNotEnabled
        }
        guard let bodyIdx = bodyIndexForParagraph(paragraphIndex) else {
            throw WordError.invalidIndex(paragraphIndex)
        }
        guard case .paragraph(var paragraph) = body.children[bodyIdx] else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let totalLength = paragraph.runs.reduce(0) { $0 + $1.text.count }
        guard start >= 0, end >= start, end <= totalLength else {
            throw WordError.invalidIndex(end)
        }

        let revisionId = allocateRevisionId()
        let resolvedAuthor = resolveAuthor(author)
        let resolvedDate = date ?? Date()

        var newRuns: [Run] = []
        var charsSeen = 0
        var deletedText = ""

        for run in paragraph.runs {
            let runStart = charsSeen
            let runEnd = charsSeen + run.text.count

            if runEnd <= start || runStart >= end {
                newRuns.append(run)
            } else {
                let preLen = max(0, start - runStart)
                let postLen = max(0, runEnd - end)
                let midLen = run.text.count - preLen - postLen

                if preLen > 0 {
                    var preRun = run
                    preRun.text = String(run.text.prefix(preLen))
                    newRuns.append(preRun)
                }
                if midLen > 0 {
                    var midRun = run
                    let midText = String(run.text.dropFirst(preLen).prefix(midLen))
                    midRun.text = midText
                    midRun.revisionId = revisionId
                    newRuns.append(midRun)
                    deletedText += midText
                }
                if postLen > 0 {
                    var postRun = run
                    postRun.text = String(run.text.suffix(postLen))
                    newRuns.append(postRun)
                }
            }

            charsSeen = runEnd
        }

        paragraph.runs = newRuns

        var revision = Revision(
            id: revisionId,
            type: .deletion,
            author: resolvedAuthor,
            date: resolvedDate,
            content: deletedText
        )
        revision.originalText = deletedText
        paragraph.revisions.append(revision)
        revisions.revisions.append(revision)

        body.children[bodyIdx] = .paragraph(paragraph)
        modifiedParts.insert("word/document.xml")

        return revisionId
    }

    /// Move a contiguous span of text from `[fromStart, fromEnd)` in
    /// `fromParagraph` to `toPosition` in `toParagraph`, recording the action
    /// as a paired `<w:moveFrom>` / `<w:moveTo>` revision. The two halves get
    /// adjacent ids (returned as `(fromId, toId)` where `toId == fromId + 1`).
    ///
    /// Single-paragraph moves are rejected (`invalidParameter`) — callers should
    /// model them as a delete + insert pair instead.
    public mutating func moveTextAsRevision(
        fromParagraph: Int,
        fromStart: Int,
        fromEnd: Int,
        toParagraph: Int,
        toPosition: Int,
        author: String? = nil,
        date: Date? = nil
    ) throws -> (fromId: Int, toId: Int) {
        guard isTrackChangesEnabled() else {
            throw WordError.trackChangesNotEnabled
        }
        guard fromParagraph != toParagraph else {
            throw WordError.invalidParameter("toParagraph",
                "single-paragraph move is out of scope; use delete + insert instead")
        }
        guard let fromBodyIdx = bodyIndexForParagraph(fromParagraph) else {
            throw WordError.invalidIndex(fromParagraph)
        }
        guard let toBodyIdx = bodyIndexForParagraph(toParagraph) else {
            throw WordError.invalidIndex(toParagraph)
        }
        guard case .paragraph(var fromPara) = body.children[fromBodyIdx] else {
            throw WordError.invalidIndex(fromParagraph)
        }
        guard case .paragraph(var toPara) = body.children[toBodyIdx] else {
            throw WordError.invalidIndex(toParagraph)
        }

        let fromTotal = fromPara.runs.reduce(0) { $0 + $1.text.count }
        guard fromStart >= 0, fromEnd >= fromStart, fromEnd <= fromTotal else {
            throw WordError.invalidIndex(fromEnd)
        }
        let toTotal = toPara.runs.reduce(0) { $0 + $1.text.count }
        guard toPosition >= 0, toPosition <= toTotal else {
            throw WordError.invalidIndex(toPosition)
        }

        let fromId = allocateRevisionId()
        let toId = fromId + 1
        let resolvedAuthor = resolveAuthor(author)
        let resolvedDate = date ?? Date()

        // moveFrom side: split runs at [fromStart, fromEnd), tag middle runs.
        var newFromRuns: [Run] = []
        var charsSeen = 0
        var movedText = ""
        for run in fromPara.runs {
            let runStart = charsSeen
            let runEnd = charsSeen + run.text.count
            if runEnd <= fromStart || runStart >= fromEnd {
                newFromRuns.append(run)
            } else {
                let preLen = max(0, fromStart - runStart)
                let postLen = max(0, runEnd - fromEnd)
                let midLen = run.text.count - preLen - postLen
                if preLen > 0 {
                    var preRun = run
                    preRun.text = String(run.text.prefix(preLen))
                    newFromRuns.append(preRun)
                }
                if midLen > 0 {
                    var midRun = run
                    let midText = String(run.text.dropFirst(preLen).prefix(midLen))
                    midRun.text = midText
                    midRun.revisionId = fromId
                    newFromRuns.append(midRun)
                    movedText += midText
                }
                if postLen > 0 {
                    var postRun = run
                    postRun.text = String(run.text.suffix(postLen))
                    newFromRuns.append(postRun)
                }
            }
            charsSeen = runEnd
        }
        fromPara.runs = newFromRuns

        // moveTo side: insert a new run at toPosition.
        var newRun = Run(text: movedText)
        newRun.revisionId = toId
        if toPosition == 0 {
            toPara.runs.insert(newRun, at: 0)
        } else if toPosition == toTotal {
            toPara.runs.append(newRun)
        } else {
            var seen = 0
            for runIdx in 0..<toPara.runs.count {
                let length = toPara.runs[runIdx].text.count
                let after = seen + length
                if toPosition == seen {
                    toPara.runs.insert(newRun, at: runIdx)
                    break
                } else if toPosition < after {
                    let offsetInRun = toPosition - seen
                    let original = toPara.runs[runIdx]
                    var beforeRun = original
                    beforeRun.text = String(original.text.prefix(offsetInRun))
                    var afterRun = original
                    afterRun.text = String(original.text.suffix(length - offsetInRun))
                    toPara.runs[runIdx] = beforeRun
                    toPara.runs.insert(newRun, at: runIdx + 1)
                    toPara.runs.insert(afterRun, at: runIdx + 2)
                    break
                }
                seen = after
            }
        }

        var fromRevision = Revision(
            id: fromId, type: .moveFrom, author: resolvedAuthor,
            date: resolvedDate, content: movedText
        )
        fromRevision.originalText = movedText
        fromPara.revisions.append(fromRevision)
        revisions.revisions.append(fromRevision)

        var toRevision = Revision(
            id: toId, type: .moveTo, author: resolvedAuthor,
            date: resolvedDate, content: movedText
        )
        toRevision.newText = movedText
        toPara.revisions.append(toRevision)
        revisions.revisions.append(toRevision)

        body.children[fromBodyIdx] = .paragraph(fromPara)
        body.children[toBodyIdx] = .paragraph(toPara)
        modifiedParts.insert("word/document.xml")

        return (fromId: fromId, toId: toId)
    }

    /// Replace `ParagraphProperties` on paragraph `paragraphIndex`, recording
    /// the previous properties as a tracked `<w:pPrChange>` revision. Returns
    /// the allocated revision id.
    public mutating func applyParagraphPropertiesAsRevision(
        atParagraph paragraphIndex: Int,
        newProperties: ParagraphProperties,
        author: String? = nil,
        date: Date? = nil
    ) throws -> Int {
        guard isTrackChangesEnabled() else {
            throw WordError.trackChangesNotEnabled
        }
        guard let bodyIdx = bodyIndexForParagraph(paragraphIndex) else {
            throw WordError.invalidIndex(paragraphIndex)
        }
        guard case .paragraph(var paragraph) = body.children[bodyIdx] else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let revisionId = allocateRevisionId()
        let resolvedAuthor = resolveAuthor(author)
        let resolvedDate = date ?? Date()

        let previousProperties = paragraph.properties
        paragraph.properties = newProperties
        paragraph.previousProperties = previousProperties
        paragraph.paragraphFormatChangeRevisionId = revisionId

        var revision = Revision(
            id: revisionId,
            type: .paragraphChange,
            author: resolvedAuthor,
            date: resolvedDate
        )
        revision.previousFormatDescription = "ParagraphProperties change"
        paragraph.revisions.append(revision)
        revisions.revisions.append(revision)

        body.children[bodyIdx] = .paragraph(paragraph)
        modifiedParts.insert("word/document.xml")

        return revisionId
    }

    /// Replace `RunProperties` of the run at `atRunIndex` in paragraph
    /// `paragraphIndex`, recording the previous properties as a tracked
    /// `<w:rPrChange>` revision. The replacement is in-place — the run keeps
    /// its text and identity, only its formatting changes. Returns the
    /// allocated revision id.
    public mutating func applyRunPropertiesAsRevision(
        atParagraph paragraphIndex: Int,
        atRunIndex runIndex: Int,
        newProperties: RunProperties,
        author: String? = nil,
        date: Date? = nil
    ) throws -> Int {
        guard isTrackChangesEnabled() else {
            throw WordError.trackChangesNotEnabled
        }
        guard let bodyIdx = bodyIndexForParagraph(paragraphIndex) else {
            throw WordError.invalidIndex(paragraphIndex)
        }
        guard case .paragraph(var paragraph) = body.children[bodyIdx] else {
            throw WordError.invalidIndex(paragraphIndex)
        }
        guard runIndex >= 0, runIndex < paragraph.runs.count else {
            throw WordError.invalidIndex(runIndex)
        }

        let revisionId = allocateRevisionId()
        let resolvedAuthor = resolveAuthor(author)
        let resolvedDate = date ?? Date()

        let previousProperties = paragraph.runs[runIndex].properties
        paragraph.runs[runIndex].properties = newProperties
        paragraph.runs[runIndex].formatChangeRevisionId = revisionId

        var revision = Revision(
            id: revisionId,
            type: .formatChange,
            author: resolvedAuthor,
            date: resolvedDate,
            previousFormat: previousProperties
        )
        revision.previousFormat = previousProperties
        paragraph.revisions.append(revision)
        revisions.revisions.append(revision)

        body.children[bodyIdx] = .paragraph(paragraph)
        modifiedParts.insert("word/document.xml")

        return revisionId
    }

    /// 取得所有修訂（僅 body 來源，向後相容）
    ///
    /// 只回傳 `source == .body` 的修訂。Container 來源（header/footer/footnote/endnote）
    /// 的修訂需要透過 `getRevisionsFull()` 取得。
    public func getRevisions() -> [(id: Int, type: String, author: String, paragraphIndex: Int, originalText: String?, newText: String?)] {
        return revisions.revisions
            .filter { $0.source == .body }
            .map { rev in
                (id: rev.id, type: rev.type.rawValue, author: rev.author,
                 paragraphIndex: rev.paragraphIndex, originalText: rev.originalText, newText: rev.newText)
            }
    }

    /// v0.19.4+ (#56 R3-NEW-4): when a Revision was created by the Reader's
    /// hasNonRunChild raw-capture branch, the wrapper lives as a verbatim XML
    /// entry in `paragraph.unrecognizedChildren`. Run-text replacement won't
    /// touch it. Walks every part (body incl. nested tables and content-control
    /// children, headers, footers, footnotes, endnotes) for the matching entry
    /// by name + id and either unwraps (accept) or removes (reject) it.
    ///
    /// Returns the originating part key (e.g. `"word/header1.xml"`,
    /// `"word/footnotes.xml"`) when a match was acted on, or `nil` when not
    /// found. Caller marks `partKey` dirty in `modifiedParts` and propagates
    /// the not-found case to a `RevisionError.notFound` throw rather than
    /// silently returning success.
    ///
    /// `accept` true → replace the wrapper rawXML with just the inner content
    /// (between the opening tag's `>` and the closing `</w:NAME>`).
    /// `accept` false → remove the entry entirely (drops wrapper AND inner).
    private mutating func handleMixedContentWrapperRevision(
        revisionId: Int,
        wrapperName: String,
        accept: Bool
    ) -> String? {
        let openTagPrefix = "<w:\(wrapperName)"
        let openIdMarker = "w:id=\"\(revisionId)\""
        let closeTag = "</w:\(wrapperName)>"

        // Opening-tag-only match: a nested element carrying the same numeric id
        // (e.g. `<w:ins w:id="3"><w:bookmarkStart w:id="5"/></w:ins>`) must NOT
        // false-hit the outer wrapper for revision 5.
        func match(_ child: UnrecognizedChild) -> Bool {
            guard child.name == wrapperName,
                  child.rawXML.hasPrefix(openTagPrefix),
                  let openTagEnd = child.rawXML.firstIndex(of: ">") else { return false }
            let openTag = child.rawXML[child.rawXML.startIndex..<openTagEnd]
            return openTag.contains(openIdMarker)
        }

        func transformParagraph(_ para: inout Paragraph) -> Bool {
            guard let idx = para.unrecognizedChildren.firstIndex(where: match) else { return false }
            if accept {
                let raw = para.unrecognizedChildren[idx].rawXML
                guard let openEnd = raw.firstIndex(of: ">"),
                      let closeStart = raw.range(of: closeTag, options: .backwards)?.lowerBound else {
                    // Malformed wrapper — fall through to removal (safer than emitting it again).
                    para.unrecognizedChildren.remove(at: idx)
                    return true
                }
                let innerStart = raw.index(after: openEnd)
                let inner = String(raw[innerStart..<closeStart])
                if inner.isEmpty {
                    para.unrecognizedChildren.remove(at: idx)
                } else {
                    para.unrecognizedChildren[idx] = UnrecognizedChild(
                        name: "raw",
                        rawXML: inner,
                        position: para.unrecognizedChildren[idx].position
                    )
                }
            } else {
                para.unrecognizedChildren.remove(at: idx)
            }
            // Also drop the matching typed Revision from the per-paragraph list.
            para.revisions.removeAll { $0.id == revisionId }
            return true
        }

        // Body — recurse into tables, nested tables, and content-control children.
        if let partKey = transformInBodyChildren(&body.children, transform: transformParagraph) {
            return partKey
        }

        // v0.19.5+ (#56 R5-CONT P0 #1): containers now route through the same
        // `transformInBodyChildren` recursion so wrappers inside header tables
        // / footer SDTs / footnote nested tables are reachable. Pre-fix the
        // four loops iterated `.paragraphs` (the flat backward-compat
        // computed view), missing anything inside `.table` / `.contentControl`
        // BodyChild cases.

        // Headers
        for hi in 0..<headers.count {
            let key = DocumentWalker.headerPartKey(for: headers[hi])
            if let _ = transformInBodyChildren(&headers[hi].bodyChildren, partKey: key, transform: transformParagraph) {
                return key
            }
        }
        // Footers
        for fi in 0..<footers.count {
            let key = DocumentWalker.footerPartKey(for: footers[fi])
            if let _ = transformInBodyChildren(&footers[fi].bodyChildren, partKey: key, transform: transformParagraph) {
                return key
            }
        }
        // Footnotes
        for fni in 0..<footnotes.footnotes.count {
            if let _ = transformInBodyChildren(
                &footnotes.footnotes[fni].bodyChildren,
                partKey: DocumentWalker.footnotesPartKey,
                transform: transformParagraph
            ) {
                return DocumentWalker.footnotesPartKey
            }
        }
        // Endnotes
        for eni in 0..<endnotes.endnotes.count {
            if let _ = transformInBodyChildren(
                &endnotes.endnotes[eni].bodyChildren,
                partKey: DocumentWalker.endnotesPartKey,
                transform: transformParagraph
            ) {
                return DocumentWalker.endnotesPartKey
            }
        }
        return nil
    }

    /// Recursively transforms a `[BodyChild]` slice in-place. Returns the
    /// supplied `partKey` on first hit, or nil if no paragraph in the slice
    /// matched. Used by the part-spanning wrapper helper to handle paragraphs
    /// that live inside tables, nested tables, or block-level content-control
    /// children.
    ///
    /// v0.19.5+ (#56 R5-CONT P0 #1): now parameterized over `partKey` so
    /// container `bodyChildren` (header / footer / footnote / endnote) can
    /// reuse the same recursion. Pre-fix the body-only hardcoded
    /// `DocumentWalker.bodyPartKey` return forced container loops to roll
    /// their own (incomplete) iteration over `.paragraphs`, missing wrappers
    /// inside container tables.
    private func transformInBodyChildren(
        _ children: inout [BodyChild],
        partKey: String = DocumentWalker.bodyPartKey,
        transform: (inout Paragraph) -> Bool
    ) -> String? {
        for i in 0..<children.count {
            switch children[i] {
            case .bookmarkMarker, .rawBlockElement:
                // Body-level markers contain no paragraphs to transform (#58).
                continue
            case .paragraph(var para):
                if transform(&para) {
                    children[i] = .paragraph(para)
                    return partKey
                }
            case .table(var table):
                var hit = false
                for r in 0..<table.rows.count {
                    for c in 0..<table.rows[r].cells.count {
                        for p in 0..<table.rows[r].cells[c].paragraphs.count {
                            if transform(&table.rows[r].cells[c].paragraphs[p]) {
                                hit = true
                            }
                        }
                        // Recurse into nested tables.
                        for nt in 0..<table.rows[r].cells[c].nestedTables.count {
                            var nestedChildren: [BodyChild] = [.table(table.rows[r].cells[c].nestedTables[nt])]
                            if let _ = transformInBodyChildren(&nestedChildren, partKey: partKey, transform: transform) {
                                if case .table(let updated) = nestedChildren[0] {
                                    table.rows[r].cells[c].nestedTables[nt] = updated
                                }
                                hit = true
                            }
                        }
                    }
                }
                if hit {
                    children[i] = .table(table)
                    return partKey
                }
            case .contentControl(let cc, var inner):
                if let _ = transformInBodyChildren(&inner, partKey: partKey, transform: transform) {
                    children[i] = .contentControl(cc, children: inner)
                    return partKey
                }
            }
        }
        return nil
    }

    /// 接受修訂
    public mutating func acceptRevision(revisionId: Int) throws {
        guard let index = revisions.revisions.firstIndex(where: { $0.id == revisionId }) else {
            throw RevisionError.notFound(revisionId)
        }

        let revision = revisions.revisions[index]

        // v0.19.4+ (#56 R3-NEW-4): mixed-content wrapper revisions live in
        // `paragraph.unrecognizedChildren` as raw XML. Strip / unwrap the raw
        // entry before falling through to the legacy run-based logic (which
        // would no-op on raw XML and silently leave the wrapper in place).
        if revision.isMixedContentWrapper {
            let wrapperName: String
            switch revision.type {
            case .insertion: wrapperName = "ins"
            case .deletion: wrapperName = "del"
            case .moveFrom: wrapperName = "moveFrom"
            case .moveTo: wrapperName = "moveTo"
            default: wrapperName = "ins"
            }
            // Accept = keep inner content (typical for insertion / moveTo) or
            // drop both wrapper + inner (typical for deletion / moveFrom).
            let keepInner = (revision.type == .insertion || revision.type == .moveTo)
            guard let partKey = handleMixedContentWrapperRevision(
                revisionId: revisionId,
                wrapperName: wrapperName,
                accept: keepInner
            ) else {
                // No matching unrecognizedChild entry in any part. Surface the
                // mismatch instead of silently removing the typed Revision.
                throw RevisionError.notFound(revisionId)
            }
            revisions.revisions.remove(at: index)
            modifiedParts.insert(partKey)
            return
        }

        // v0.19.5+ (#56 R5-CONT P0 #5): typed `.deletion` now routes by
        // `revision.source` instead of indexing body.children unconditionally.
        // Pre-fix the body-only branch could either silently no-op (paragraph
        // index out of body bounds) or DELETE THE WRONG body paragraph (if
        // the index happened to fall in body range), then mark the wrong
        // part dirty. Container `.deletion` revisions therefore corrupted
        // body content and lost their own marker without surfacing — strictly
        // worse than R4's notFound (which at least reported failure).
        // See verify R5 P0 #5 (DA C2 + H2).
        // v0.19.5+ (#56 R5-CONT-4 P0 #1): mirror R5-CONT-3 §15.1+§15.2's
        // reject-side clearMarker pattern for ALL accept-side typed
        // branches. Pre-fix accept removed only document.revisions[id]
        // — paragraph.revisions[id] / run.revisionId / paraFormatChangeId /
        // run.formatChangeRevisionId stayed intact → Paragraph.toXML
        // still wrapped runs in <w:ins>/<w:del>/<w:rPrChange>/<w:pPrChange>
        // on save. Confirmed silent corruption by R5-CONT-3 verify (DA +
        // Codex independent confirm). §15.6 matrix-pin's `if operation
        // == "reject"` guard documented this as expected — R5-CONT-4
        // also removes that guard (§17.2).
        let revisionId_ = revision.id
        let clearAllMarkers: (inout Paragraph) -> Void = { para in
            para.revisions.removeAll { $0.id == revisionId_ }
            if para.paragraphFormatChangeRevisionId == revisionId_ {
                para.paragraphFormatChangeRevisionId = nil
            }
            for j in 0..<para.runs.count {
                if para.runs[j].revisionId == revisionId_ {
                    para.runs[j].revisionId = nil
                }
                if para.runs[j].formatChangeRevisionId == revisionId_ {
                    para.runs[j].formatChangeRevisionId = nil
                }
            }
        }

        var partKeyForDirty = "word/document.xml"
        switch revision.type {
        case .insertion:
            // 接受插入：移除標記，保留文字（文字已在文件中）。
            // R5-CONT-4 P0 #1: ALSO clear paragraph + run revision-id
            // refs so Paragraph.toXML stops wrapping runs in <w:ins>.
            if let key = applyToParagraph(at: revision.paragraphIndex, in: revision.source, mutate: clearAllMarkers) {
                partKeyForDirty = key
            } else {
                // Soft-fail: the typed Revision exists but no matching
                // paragraph carries the revision-id (legitimate when
                // the marker was already cleared). Don't throw — the
                // operation is semantically "no-op cleanup" not "error".
                partKeyForDirty = try sourceToPartKey(revision.source, revisionId: revisionId)
            }
        case .deletion:
            // 接受刪除：實際移除被標記為刪除的文字 AND clear marker.
            let originalText = revision.originalText
            let removeAndClear: (inout Paragraph) -> Void = { para in
                if let txt = originalText {
                    for j in 0..<para.runs.count {
                        para.runs[j].text = para.runs[j].text.replacingOccurrences(of: txt, with: "")
                    }
                }
                clearAllMarkers(&para)
            }
            if let key = applyToParagraph(at: revision.paragraphIndex, in: revision.source, mutate: removeAndClear) {
                partKeyForDirty = key
            } else {
                // No matching paragraph in the source part — surface the
                // failure instead of silently dropping the typed Revision.
                throw RevisionError.notFound(revisionId)
            }
        case .formatting, .paragraphChange, .formatChange, .moveFrom, .moveTo:
            // R5-CONT-4 P0 #1: clear paragraph/run revision-id refs
            // (rPrChange / pPrChange / move-from/to wrapper emit) so
            // toXML stops wrapping. Soft-fail on miss (legitimate
            // already-cleared state).
            if let key = applyToParagraph(at: revision.paragraphIndex, in: revision.source, mutate: clearAllMarkers) {
                partKeyForDirty = key
            } else {
                partKeyForDirty = try sourceToPartKey(revision.source, revisionId: revisionId)
            }
        }

        // 移除修訂記錄
        revisions.revisions.remove(at: index)
        modifiedParts.insert(partKeyForDirty)
    }

    /// v0.19.5+ (#56 R5-CONT P0 #5): map a `RevisionSource` to the part
    /// key the writer's overlay-mode dirty-gate checks. Mirrors
    /// `DocumentWalker.headerPartKey` / `footerPartKey` for `.header(id:)` /
    /// `.footer(id:)` and routes notes / body to their canonical paths.
    ///
    /// v0.19.5+ (#56 R5-CONT-3 P1 #3): throws `RevisionError.notFound`
    /// when source is `.header(id:X)` / `.footer(id:X)` but no container
    /// with that rId exists. Pre-fix the silent fallback to
    /// `"word/document.xml"` masked orphan-revision logic bugs (e.g.,
    /// container was renamed between propagation and accept, or caller
    /// constructed `.header(id: "rId-typo")` by mistake → wrong-part
    /// dirty + revision still removed). Surface the failure so the
    /// caller can rollback the typed Revision removal.
    private func sourceToPartKey(_ source: RevisionSource, revisionId: Int) throws -> String {
        switch source {
        case .body:
            return "word/document.xml"
        case .header(let id):
            if let h = headers.first(where: { $0.id == id }) {
                return DocumentWalker.headerPartKey(for: h)
            }
            throw RevisionError.notFound(revisionId)
        case .footer(let id):
            if let f = footers.first(where: { $0.id == id }) {
                return DocumentWalker.footerPartKey(for: f)
            }
            throw RevisionError.notFound(revisionId)
        case .footnote:
            return DocumentWalker.footnotesPartKey
        case .endnote:
            return DocumentWalker.endnotesPartKey
        }
    }

    /// v0.19.5+ (#56 R5-CONT P0 #5): apply `mutate` to the paragraph at the
    /// given `paragraphIndex` within the part identified by `source`.
    /// Walks `bodyChildren` flat-by-source-paragraph-index for body and
    /// containers — matches the index semantics
    /// `propagateRevisionsFromBodyChildren` writes (each visited paragraph
    /// gets the same `paragraphIndex` argument the helper receives).
    /// Returns the part key on hit, nil on miss.
    private mutating func applyToParagraph(
        at paragraphIndex: Int,
        in source: RevisionSource,
        mutate: (inout Paragraph) -> Void
    ) -> String? {
        switch source {
        case .body:
            return Self.applyToFlatParagraph(at: paragraphIndex, in: &body.children, mutate: mutate, partKey: "word/document.xml")
        case .header(let id):
            guard let hi = headers.firstIndex(where: { $0.id == id }) else { return nil }
            let key = DocumentWalker.headerPartKey(for: headers[hi])
            return Self.applyToFlatParagraph(at: paragraphIndex, in: &headers[hi].bodyChildren, mutate: mutate, partKey: key)
        case .footer(let id):
            guard let fi = footers.firstIndex(where: { $0.id == id }) else { return nil }
            let key = DocumentWalker.footerPartKey(for: footers[fi])
            return Self.applyToFlatParagraph(at: paragraphIndex, in: &footers[fi].bodyChildren, mutate: mutate, partKey: key)
        case .footnote(let id):
            guard let fni = footnotes.footnotes.firstIndex(where: { $0.id == id }) else { return nil }
            return Self.applyToFlatParagraph(at: paragraphIndex, in: &footnotes.footnotes[fni].bodyChildren, mutate: mutate, partKey: DocumentWalker.footnotesPartKey)
        case .endnote(let id):
            guard let eni = endnotes.endnotes.firstIndex(where: { $0.id == id }) else { return nil }
            return Self.applyToFlatParagraph(at: paragraphIndex, in: &endnotes.endnotes[eni].bodyChildren, mutate: mutate, partKey: DocumentWalker.endnotesPartKey)
        }
    }

    /// Static recursive helper for `applyToParagraph`. Walks bodyChildren
    /// counting `.paragraph` cases (incl. those inside `.table` cells,
    /// nested tables, and `.contentControl` children) until the
    /// `paragraphIndex`-th paragraph is reached, applies `mutate`, and
    /// returns `partKey`. Counter is value-typed (passed by reference via
    /// inout) so the recursion can short-circuit cleanly.
    private static func applyToFlatParagraph(
        at paragraphIndex: Int,
        in children: inout [BodyChild],
        mutate: (inout Paragraph) -> Void,
        partKey: String
    ) -> String? {
        var counter = 0
        var hit = false
        applyToFlatParagraphRecursive(
            in: &children,
            target: paragraphIndex,
            counter: &counter,
            hit: &hit,
            mutate: mutate
        )
        return hit ? partKey : nil
    }

    private static func applyToFlatParagraphRecursive(
        in children: inout [BodyChild],
        target: Int,
        counter: inout Int,
        hit: inout Bool,
        mutate: (inout Paragraph) -> Void
    ) {
        for i in 0..<children.count {
            if hit { return }
            switch children[i] {
            case .bookmarkMarker, .rawBlockElement:
                // Body-level markers don't count as paragraphs and contain
                // no paragraphs to mutate (#58).
                continue
            case .paragraph(var para):
                if counter == target {
                    mutate(&para)
                    children[i] = .paragraph(para)
                    hit = true
                    return
                }
                counter += 1
            case .table(var table):
                for r in 0..<table.rows.count {
                    if hit { break }
                    for c in 0..<table.rows[r].cells.count {
                        if hit { break }
                        for p in 0..<table.rows[r].cells[c].paragraphs.count {
                            if hit { break }
                            if counter == target {
                                var para = table.rows[r].cells[c].paragraphs[p]
                                mutate(&para)
                                table.rows[r].cells[c].paragraphs[p] = para
                                hit = true
                                break
                            }
                            counter += 1
                        }
                        // Recurse into nested tables.
                        for nt in 0..<table.rows[r].cells[c].nestedTables.count {
                            if hit { break }
                            var nestedChildren: [BodyChild] = [.table(table.rows[r].cells[c].nestedTables[nt])]
                            applyToFlatParagraphRecursive(in: &nestedChildren, target: target, counter: &counter, hit: &hit, mutate: mutate)
                            if hit, case .table(let updated) = nestedChildren[0] {
                                table.rows[r].cells[c].nestedTables[nt] = updated
                            }
                        }
                    }
                }
                if hit {
                    children[i] = .table(table)
                    return
                }
            case .contentControl(let cc, var inner):
                applyToFlatParagraphRecursive(in: &inner, target: target, counter: &counter, hit: &hit, mutate: mutate)
                if hit {
                    children[i] = .contentControl(cc, children: inner)
                    return
                }
            }
        }
    }

    /// 拒絕修訂
    public mutating func rejectRevision(revisionId: Int) throws {
        guard let index = revisions.revisions.firstIndex(where: { $0.id == revisionId }) else {
            throw RevisionError.notFound(revisionId)
        }

        let revision = revisions.revisions[index]

        // v0.19.4+ (#56 R3-NEW-4): mirror of acceptRevision's mixed-content
        // handling. Reject of an insertion drops the wrapper AND inner content
        // (matches Word's "reject all" semantics — inserted content disappears).
        // Reject of a deletion would restore the deleted content; we restore by
        // keeping the inner content (the wrapper went around what was already
        // in the paragraph). moveFrom / moveTo follow the same pattern as
        // deletion / insertion respectively.
        if revision.isMixedContentWrapper {
            let wrapperName: String
            switch revision.type {
            case .insertion: wrapperName = "ins"
            case .deletion: wrapperName = "del"
            case .moveFrom: wrapperName = "moveFrom"
            case .moveTo: wrapperName = "moveTo"
            default: wrapperName = "ins"
            }
            // Reject = drop inner for insertion / moveTo (the new content is
            // rejected); keep inner for deletion / moveFrom (the deletion is
            // rejected, so the original content is preserved).
            let keepInner = (revision.type == .deletion || revision.type == .moveFrom)
            guard let partKey = handleMixedContentWrapperRevision(
                revisionId: revisionId,
                wrapperName: wrapperName,
                accept: keepInner
            ) else {
                throw RevisionError.notFound(revisionId)
            }
            revisions.revisions.remove(at: index)
            modifiedParts.insert(partKey)
            return
        }

        // v0.19.5+ (#56 R5-CONT-2 P0 #3): mirror R5-CONT P0 #5's
        // acceptRevision typed `.deletion` source-routing for the reject
        // side. Pre-fix the typed `.insertion` branch indexed
        // `body.children` regardless of `revision.source` → for a
        // container-source revision, `rejectRevision` either silently
        // no-op'd OR DELETED BODY TEXT matching `newText` (silent
        // body corruption + revision marker vanished + wrong part
        // dirty). Strictly worse than the R5 verify P0 #1 / R5-CONT
        // P0 #5 patterns because the asymmetric behavior was never
        // documented as a known limitation.
        var partKeyForDirty = "word/document.xml"
        switch revision.type {
        case .insertion:
            // Reject insertion: remove the inserted text AND clear typed
            // Revision marker on paragraph + run level so the writer
            // doesn't re-emit the <w:ins> wrapper around now-empty runs.
            //
            // v0.19.5+ (#56 R5-CONT-3 §15.6 matrix-pin caught): pre-fix
            // R5-CONT-2 §13.3 only ran removeText but left
            // paragraph.revisions[id] + run.revisionId intact. The
            // wrapper would re-emit on save with empty text — silent
            // file-state inconsistency. Extends §15.1's clearMarker
            // pattern to the insertion-reject path.
            let newText = revision.newText
            let removeAndClear: (inout Paragraph) -> Void = { para in
                if let txt = newText {
                    for j in 0..<para.runs.count {
                        para.runs[j].text = para.runs[j].text.replacingOccurrences(of: txt, with: "")
                    }
                }
                para.revisions.removeAll { $0.id == revisionId }
                for j in 0..<para.runs.count {
                    if para.runs[j].revisionId == revisionId {
                        para.runs[j].revisionId = nil
                    }
                }
            }
            if let key = applyToParagraph(at: revision.paragraphIndex, in: revision.source, mutate: removeAndClear) {
                partKeyForDirty = key
            } else {
                throw RevisionError.notFound(revisionId)
            }
        case .deletion:
            // Reject deletion: restore the deleted text by clearing the
            // typed Revision marker on the paragraph + run level. The text
            // already lives in the runs (marked but not removed); the
            // wrapper emit is driven by paragraph.revisions + run.revisionId,
            // so we MUST clear both for the file to converge with the API
            // state. Honor source for dirty-tracking.
            //
            // v0.19.5+ (#56 R5-CONT-3 P0 #1): pre-fix this branch only set
            // partKeyForDirty and let `revisions.revisions.remove(at: index)`
            // clean the document-scope list. paragraph.revisions still
            // contained the Revision id; run.revisionId still referenced it;
            // Paragraph.toXML() (Paragraph.swift:330-345) still grouped
            // matching runs into <w:del> wrappers on save → silent
            // inconsistency between API state ("rejected") and file state
            // (still wrapped). DA C1 / R5-CONT-2 verify P0.
            let revisionId = revision.id
            let clearMarker: (inout Paragraph) -> Void = { para in
                para.revisions.removeAll { $0.id == revisionId }
                for j in 0..<para.runs.count {
                    if para.runs[j].revisionId == revisionId {
                        para.runs[j].revisionId = nil
                    }
                }
            }
            if let key = applyToParagraph(at: revision.paragraphIndex, in: revision.source, mutate: clearMarker) {
                partKeyForDirty = key
            } else {
                throw RevisionError.notFound(revisionId)
            }
        case .formatting, .paragraphChange, .formatChange, .moveFrom, .moveTo:
            // v0.19.5+ (#56 R5-CONT-3 P1 #2): mirror P0 #1's clear-marker
            // pattern. These types don't have body-text mutation to undo,
            // but the marker DOES need to be cleared to converge file
            // state with API state (otherwise the writer keeps emitting
            // the wrapper). pPrChange uses paragraph.paragraphFormatChangeRevisionId;
            // rPrChange uses run.formatChangeRevisionId; moveFrom/moveTo
            // share the run.revisionId pattern with insertion/deletion.
            let revisionId = revision.id
            let clearAllMarkers: (inout Paragraph) -> Void = { para in
                para.revisions.removeAll { $0.id == revisionId }
                if para.paragraphFormatChangeRevisionId == revisionId {
                    para.paragraphFormatChangeRevisionId = nil
                }
                for j in 0..<para.runs.count {
                    if para.runs[j].revisionId == revisionId {
                        para.runs[j].revisionId = nil
                    }
                    if para.runs[j].formatChangeRevisionId == revisionId {
                        para.runs[j].formatChangeRevisionId = nil
                    }
                }
            }
            if let key = applyToParagraph(at: revision.paragraphIndex, in: revision.source, mutate: clearAllMarkers) {
                partKeyForDirty = key
            } else {
                // For format/move types where the typed Revision exists but
                // no matching paragraph carries any revision-id reference
                // (legitimate when the marker was already cleared), honor
                // source for dirty-tracking and fall through to the
                // remove-from-document-revisions step. Don't throw notFound
                // because the operation is semantically "no-op" not "error".
                partKeyForDirty = try sourceToPartKey(revision.source, revisionId: revisionId)
            }
        }

        // 移除修訂記錄
        revisions.revisions.remove(at: index)
        modifiedParts.insert(partKeyForDirty)
    }

    /// 接受所有修訂（legacy non-throwing API; preserved for backward compat)
    ///
    /// v0.19.5+ (#56 R5 P1 #5): per-revision errors are still swallowed
    /// here for backward compatibility (this signature is consumed by
    /// downstream `che-word-mcp` per the R5 design's zero-MCP-source-change
    /// discipline). Callers that need to surface aggregate failure SHALL
    /// use `tryAcceptAllRevisions()` instead — same semantics but throws
    /// `RevisionError.partialFailure([Int])` if any per-revision helper
    /// failed.
    public mutating func acceptAllRevisions() {
        try? tryAcceptAllRevisions()
    }

    /// 拒絕所有修訂（legacy non-throwing API; preserved for backward compat）
    ///
    /// See `acceptAllRevisions` — same backward-compat rationale. Use
    /// `tryRejectAllRevisions()` when aggregate failure surfacing matters.
    public mutating func rejectAllRevisions() {
        try? tryRejectAllRevisions()
    }

    /// v0.19.5+ (#56 R5 P1 #5): throwing variant of `acceptAllRevisions`.
    /// Per-revision errors are aggregated; if any failed, throws
    /// `RevisionError.partialFailure([Int])` listing the failing ids.
    /// Successful sibling revisions are still applied (partial-success
    /// semantics) so a single orphan id cannot block the rest. Closes
    /// DA-N9 — the silent-corruption mode where `try?` swallowed all
    /// failures and let the caller assume all-clear.
    public mutating func tryAcceptAllRevisions() throws {
        var failedIds: [Int] = []
        // 從後往前接受，避免索引問題
        for revision in revisions.revisions.reversed() {
            do {
                try acceptRevision(revisionId: revision.id)
            } catch {
                failedIds.append(revision.id)
            }
        }
        if !failedIds.isEmpty {
            throw RevisionError.partialFailure(failedIds)
        }
    }

    /// v0.19.5+ (#56 R5 P1 #5): throwing variant of `rejectAllRevisions`.
    /// See `tryAcceptAllRevisions` — same aggregate-failure semantics.
    public mutating func tryRejectAllRevisions() throws {
        var failedIds: [Int] = []
        // 從後往前拒絕，避免索引問題
        for revision in revisions.revisions.reversed() {
            do {
                try rejectRevision(revisionId: revision.id)
            } catch {
                failedIds.append(revision.id)
            }
        }
        if !failedIds.isEmpty {
            throw RevisionError.partialFailure(failedIds)
        }
    }

    // MARK: - Footnote Operations

    /// 插入腳註
    public mutating func insertFootnote(
        text: String,
        paragraphIndex: Int
    ) throws -> Int {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count else {
            throw FootnoteError.invalidParagraphIndex(paragraphIndex)
        }

        let footnoteId = footnotes.nextFootnoteId()
        let footnote = Footnote(
            id: footnoteId,
            text: text,
            paragraphIndex: paragraphIndex
        )

        footnotes.footnotes.append(footnote)

        // 在段落中添加腳註參照
        let actualIndex = paragraphIndices[paragraphIndex]
        if case .paragraph(var para) = body.children[actualIndex] {
            para.footnoteIds.append(footnoteId)
            body.children[actualIndex] = .paragraph(para)
        }

        modifiedParts.insert("word/footnotes.xml")
        modifiedParts.insert("word/document.xml")
        return footnoteId
    }

    /// 刪除腳註
    public mutating func deleteFootnote(footnoteId: Int) throws {
        guard let index = footnotes.footnotes.firstIndex(where: { $0.id == footnoteId }) else {
            throw FootnoteError.notFound(footnoteId)
        }

        // 從段落中移除腳註參照
        let paragraphIndex = footnotes.footnotes[index].paragraphIndex
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        if paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count {
            let actualIndex = paragraphIndices[paragraphIndex]
            if case .paragraph(var para) = body.children[actualIndex] {
                para.footnoteIds.removeAll { $0 == footnoteId }
                body.children[actualIndex] = .paragraph(para)
            }
        }

        footnotes.footnotes.remove(at: index)
        modifiedParts.insert("word/footnotes.xml")
        modifiedParts.insert("word/document.xml")
    }

    /// 列出所有腳註
    public func getFootnotes() -> [(id: Int, text: String, paragraphIndex: Int)] {
        return footnotes.footnotes.map { footnote in
            (id: footnote.id, text: footnote.text, paragraphIndex: footnote.paragraphIndex)
        }
    }

    // MARK: - Endnote Operations

    /// 插入尾註
    public mutating func insertEndnote(
        text: String,
        paragraphIndex: Int
    ) throws -> Int {
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        guard paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count else {
            throw EndnoteError.invalidParagraphIndex(paragraphIndex)
        }

        let endnoteId = endnotes.nextEndnoteId()
        let endnote = Endnote(
            id: endnoteId,
            text: text,
            paragraphIndex: paragraphIndex
        )

        endnotes.endnotes.append(endnote)

        // 在段落中添加尾註參照
        let actualIndex = paragraphIndices[paragraphIndex]
        if case .paragraph(var para) = body.children[actualIndex] {
            para.endnoteIds.append(endnoteId)
            body.children[actualIndex] = .paragraph(para)
        }

        modifiedParts.insert("word/endnotes.xml")
        modifiedParts.insert("word/document.xml")
        return endnoteId
    }

    /// 刪除尾註
    public mutating func deleteEndnote(endnoteId: Int) throws {
        guard let index = endnotes.endnotes.firstIndex(where: { $0.id == endnoteId }) else {
            throw EndnoteError.notFound(endnoteId)
        }

        // 從段落中移除尾註參照
        let paragraphIndex = endnotes.endnotes[index].paragraphIndex
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }

        if paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count {
            let actualIndex = paragraphIndices[paragraphIndex]
            if case .paragraph(var para) = body.children[actualIndex] {
                para.endnoteIds.removeAll { $0 == endnoteId }
                body.children[actualIndex] = .paragraph(para)
            }
        }

        endnotes.endnotes.remove(at: index)
        modifiedParts.insert("word/endnotes.xml")
        modifiedParts.insert("word/document.xml")
    }

    /// 列出所有尾註
    public func getEndnotes() -> [(id: Int, text: String, paragraphIndex: Int)] {
        return endnotes.endnotes.map { endnote in
            (id: endnote.id, text: endnote.text, paragraphIndex: endnote.paragraphIndex)
        }
    }

}

// MARK: - Body

public struct Body: Equatable {
    public var children: [BodyChild] = []
    public var tables: [Table] = []
}

public enum BodyChild: Equatable {
    case paragraph(Paragraph)
    case table(Table)
    /// v0.15.0+ (#44 task 3.4): block-level Structured Document Tag wrapping
    /// one or more body children (paragraphs / tables / nested SDTs). The
    /// outer `<w:sdt>` appears directly inside `<w:body>` or `<w:tc>` rather
    /// than inside a `<w:p>`. The control's metadata is on `ContentControl.sdt`;
    /// `children` are the body elements that lived inside `<w:sdtContent>`.
    case contentControl(ContentControl, children: [BodyChild])
    /// v0.19.6+ (PsychQuant/che-word-mcp#58): body-level `<w:bookmarkStart>` /
    /// `<w:bookmarkEnd>` (e.g., TOC `_Toc<digits>` anchors that span multiple
    /// paragraphs). Pre-fix `parseBodyChildren` silently dropped these via the
    /// switch's `default: continue` branch. The marker carries `id` and (for
    /// `.start`) `name`; `position` is 0 because there is no enclosing paragraph.
    case bookmarkMarker(BookmarkRangeMarker)
    /// v0.19.6+ (#58): catch-all for any other unrecognized direct child of
    /// `<w:body>` (other EG_BlockLevelElts members like `<w:moveFromRangeStart>`,
    /// body-level `<w:commentRangeStart>`, vendor extensions, etc.). Preserved
    /// as raw XML so future element kinds round-trip byte-equivalent without
    /// per-element parser/writer branches. Same architectural pattern as
    /// `Run.rawElements` (v0.14.0+, #52).
    case rawBlockElement(RawElement)
}

// MARK: - Document Properties

public struct DocumentProperties: Equatable {
    public var title: String?
    public var subject: String?
    public var creator: String?
    public var keywords: String?
    public var description: String?
    public var lastModifiedBy: String?
    public var revision: Int?
    public var created: Date?
    public var modified: Date?
}

// MARK: - Advanced Features Extensions

extension WordDocument {
    // MARK: - Table of Contents

    /// 插入目錄
    public mutating func insertTableOfContents(
        at index: Int? = nil,
        title: String? = "Contents",
        headingLevels: ClosedRange<Int> = 1...3,
        includePageNumbers: Bool = true,
        useHyperlinks: Bool = true
    ) {
        let toc = TableOfContents(
            title: title,
            headingLevels: headingLevels,
            includePageNumbers: includePageNumbers,
            useHyperlinks: useHyperlinks
        )

        // 建立包含 TOC XML 的段落
        var para = Paragraph()
        var run = Run(text: "")
        run.properties.rawXML = toc.toXML()
        para.runs = [run]

        if let idx = index {
            insertParagraph(para, at: idx)
        } else {
            // 預設插入到開頭
            body.children.insert(.paragraph(para), at: 0)
        }
        modifiedParts.insert("word/document.xml")
    }

    // MARK: - Form Controls

    /// 插入文字欄位
    public mutating func insertTextField(
        at paragraphIndex: Int,
        name: String,
        defaultValue: String? = nil,
        maxLength: Int? = nil
    ) throws {
        guard paragraphIndex >= 0 && paragraphIndex < getParagraphs().count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let field = FormTextField(name: name, defaultValue: defaultValue, maxLength: maxLength)

        // 找到段落並添加欄位
        var childIndex = 0
        for (i, child) in body.children.enumerated() {
            if case .paragraph = child {
                if childIndex == paragraphIndex {
                    var run = Run(text: "")
                    run.properties.rawXML = field.toXML()
                    if case .paragraph(var para) = body.children[i] {
                        para.runs.append(run)
                        body.children[i] = .paragraph(para)
                    }
                    modifiedParts.insert("word/document.xml")
                    return
                }
                childIndex += 1
            }
        }
    }

    /// 插入核取方塊
    public mutating func insertCheckbox(
        at paragraphIndex: Int,
        name: String,
        isChecked: Bool = false
    ) throws {
        guard paragraphIndex >= 0 && paragraphIndex < getParagraphs().count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let checkbox = FormCheckbox(name: name, isChecked: isChecked)

        var childIndex = 0
        for (i, child) in body.children.enumerated() {
            if case .paragraph = child {
                if childIndex == paragraphIndex {
                    var run = Run(text: "")
                    run.properties.rawXML = checkbox.toXML()
                    if case .paragraph(var para) = body.children[i] {
                        para.runs.append(run)
                        body.children[i] = .paragraph(para)
                    }
                    modifiedParts.insert("word/document.xml")
                    return
                }
                childIndex += 1
            }
        }
    }

    /// 插入下拉選單
    public mutating func insertDropdown(
        at paragraphIndex: Int,
        name: String,
        options: [String],
        selectedIndex: Int = 0
    ) throws {
        guard paragraphIndex >= 0 && paragraphIndex < getParagraphs().count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let dropdown = FormDropdown(name: name, options: options, selectedIndex: selectedIndex)

        var childIndex = 0
        for (i, child) in body.children.enumerated() {
            if case .paragraph = child {
                if childIndex == paragraphIndex {
                    var run = Run(text: "")
                    run.properties.rawXML = dropdown.toXML()
                    if case .paragraph(var para) = body.children[i] {
                        para.runs.append(run)
                        body.children[i] = .paragraph(para)
                    }
                    modifiedParts.insert("word/document.xml")
                    return
                }
                childIndex += 1
            }
        }
    }

    // MARK: - Mathematical Equations

    /// Insert a math equation at the given `InsertLocation` (PsychQuant/che-word-mcp#84).
    ///
    /// Mirrors the anchor support already on `insertImage(at: InsertLocation, ...)`
    /// and `insertParagraph(_:at:)`. Display-mode equations create a new
    /// paragraph at the resolved location (all `InsertLocation` cases
    /// supported). Inline-mode equations only support `.paragraphIndex`
    /// per che-word-mcp#67 F2 (inline anchor semantics are ambiguous —
    /// "append OMML run into existing para containing this text" vs
    /// "insert new para before/after the matching para"); other anchors
    /// throw `InsertLocationError.inlineModeRequiresParagraphIndexForAnchor`
    /// with the rejected anchor kind (post-#23; pre-#91 abused
    /// `.invalidParagraphIndex(-1)` as a sentinel which lied about the failure
    /// shape).
    ///
    /// - Throws: `InsertLocationError` on anchor resolution failure;
    ///   `InsertLocationError.inlineModeRequiresParagraphIndexForAnchor` for
    ///   inline-mode non-`paragraphIndex` anchors;
    ///   `InsertLocationError.invalidParagraphIndex(idx)` for inline-mode
    ///   `.paragraphIndex(idx)` with out-of-range `idx` (post-#91; pre-#91 this
    ///   silently no-op'd via the deprecated `Int?` overload).
    public mutating func insertEquation(
        at location: InsertLocation,
        latex: String,
        displayMode: Bool = false
    ) throws {
        if !displayMode {
            // che-word-mcp#67 F2: inline equation rejects non-paragraphIndex
            // anchors. che-word-mcp#91 Defect 2: explicit bounds-check before
            // insertion because the legacy `Int?` overload silently no-ops on
            // out-of-range index (asymmetric vs display mode's throw).
            //
            // che-word-mcp#91 verify F1 (convergent finding from Codex P1 +
            // Devil's Advocate DA-1): bounds-check MUST count only top-level
            // `.paragraph` body children, NOT `getParagraphs().count` which
            // recurses into block-level `.contentControl` (Document.swift:222).
            // Inline insertion enumerates direct top-level paragraphs only (no
            // SDT descent), so a narrower count is what the operation can
            // actually reach.
            // Pre-corrective bug: doc shape `[.contentControl(_, [.paragraph])]`
            // has `getParagraphs().count == 1`, so `.paragraphIndex(0)` passed
            // the old guard but the legacy path found zero matches and silently
            // no-op'd — exactly the failure class Defect 2 was meant to fix.
            if case .paragraphIndex(let idx) = location {
                guard appendInlineEquationRun(atTopLevelParagraphIndex: idx, latex: latex) else {
                    throw InsertLocationError.invalidParagraphIndex(idx)
                }
                return
            }
            // che-word-mcp#91 Defect 1: dedicated error case (was
            // `.invalidParagraphIndex(-1)` sentinel pre-fix).
            throw InsertLocationError.inlineModeRequiresParagraphIndexForAnchor(location.anchorKindName)
        }

        // Display mode: build the equation paragraph then route through
        // insertParagraph(_:at:) which handles all 6 InsertLocation cases.
        let equation = MathEquation(latex: latex, displayMode: true)
        var run = Run(text: "")
        let omml = equation.toXML()
        run.properties.rawXML = omml
        // Verify findings (PsychQuant/che-word-mcp#85 BLOCKING #2 from
        // batched verify of e53fa00): also set top-level `run.rawXML` so
        // the new `Paragraph.flattenedDisplayText()` OMML walk sees this
        // freshly-inserted equation BEFORE a save → reload cycle. Without
        // this, sequential anchor lookups (the canonical batch-CLI
        // workflow — rescue script Phase 5) would skip equations inserted
        // earlier in the same session because flatten reads `run.rawXML`
        // (single-source-of-truth for read-side OMML, populated by
        // `DocxReader.parseRun`) but `properties.rawXML` is the write-side
        // sink that only round-trips through disk re-parse.
        run.rawXML = omml
        var para = Paragraph()
        para.runs = [run]
        para.properties.alignment = .center
        try insertParagraph(para, at: location)
    }

    /// Insert a math equation at the given paragraph index (legacy overload).
    ///
    /// **Deprecated in v0.21.5** — use the `InsertLocation` overload above
    /// for anchor-aware insertion. This Int?-only signature will be removed
    /// in v0.22 (alongside other v0.22 deprecations: `Hyperlink.text` setter,
    /// `Paragraph.commentIds` field).
    @available(*, deprecated, message: "Use insertEquation(at: InsertLocation, latex:, displayMode:) for anchor-aware insertion. WARNING: legacy inline overload can silently no-op on out-of-range indexes and SDT-nested paragraphs; the new overload throws structured errors. Will be removed in v0.22.")
    public mutating func insertEquation(
        at paragraphIndex: Int? = nil,
        latex: String,
        displayMode: Bool = false
    ) {
        if displayMode {
            // 獨立區塊公式，建立新段落
            let equation = MathEquation(latex: latex, displayMode: true)
            var para = Paragraph()
            var run = Run(text: "")
            run.properties.rawXML = equation.toXML()
            para.runs = [run]
            para.properties.alignment = .center

            if let idx = paragraphIndex {
                insertParagraph(para, at: idx)
            } else {
                appendParagraph(para)
            }
        } else {
            // 行內公式，加入到現有段落
            if let idx = paragraphIndex {
                _ = appendInlineEquationRun(atTopLevelParagraphIndex: idx, latex: latex)
            }
        }
    }

    /// Append an inline equation run to the Nth top-level body paragraph.
    ///
    /// Inline equations intentionally target only direct `.paragraph` body
    /// children. Block-level SDT descendants are excluded to match the
    /// post-#91 contract and avoid `getParagraphs()` recursion drift. Returning
    /// `false` lets the new `InsertLocation` overload throw while the deprecated
    /// `Int?` overload preserves its legacy no-op behavior (#21/#22).
    @discardableResult
    private mutating func appendInlineEquationRun(
        atTopLevelParagraphIndex paragraphIndex: Int,
        latex: String
    ) -> Bool {
        guard paragraphIndex >= 0 else { return false }

        var childIndex = 0
        for i in body.children.indices {
            guard case .paragraph(var para) = body.children[i] else { continue }
            if childIndex == paragraphIndex {
                let equation = MathEquation(latex: latex, displayMode: false)
                let omml = equation.toXML()
                var run = Run(text: "")
                run.properties.rawXML = omml
                run.rawXML = omml
                para.runs.append(run)
                body.children[i] = .paragraph(para)
                modifiedParts.insert("word/document.xml")
                return true
            }
            childIndex += 1
        }
        return false
    }

    // MARK: - Advanced Paragraph Formatting

    /// 設定段落邊框
    public mutating func setParagraphBorder(
        at paragraphIndex: Int,
        border: ParagraphBorder
    ) throws {
        guard paragraphIndex >= 0 && paragraphIndex < getParagraphs().count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        var childIndex = 0
        for (i, child) in body.children.enumerated() {
            if case .paragraph(var para) = child {
                if childIndex == paragraphIndex {
                    para.properties.border = border
                    body.children[i] = .paragraph(para)
                    modifiedParts.insert("word/document.xml")
                    return
                }
                childIndex += 1
            }
        }
    }

    /// 設定段落底色
    public mutating func setParagraphShading(
        at paragraphIndex: Int,
        fill: String,
        pattern: ShadingPattern? = nil
    ) throws {
        guard paragraphIndex >= 0 && paragraphIndex < getParagraphs().count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        var childIndex = 0
        for (i, child) in body.children.enumerated() {
            if case .paragraph(var para) = child {
                if childIndex == paragraphIndex {
                    para.properties.shading = CellShading(fill: fill, pattern: pattern)
                    body.children[i] = .paragraph(para)
                    modifiedParts.insert("word/document.xml")
                    return
                }
                childIndex += 1
            }
        }
    }

    /// 設定字元間距
    public mutating func setCharacterSpacing(
        at paragraphIndex: Int,
        spacing: Int? = nil,
        position: Int? = nil,
        kern: Int? = nil
    ) throws {
        guard paragraphIndex >= 0 && paragraphIndex < getParagraphs().count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        let charSpacing = CharacterSpacing(spacing: spacing, position: position, kern: kern)

        var childIndex = 0
        for (i, child) in body.children.enumerated() {
            if case .paragraph(var para) = child {
                if childIndex == paragraphIndex {
                    for j in 0..<para.runs.count {
                        para.runs[j].properties.characterSpacing = charSpacing
                    }
                    body.children[i] = .paragraph(para)
                    modifiedParts.insert("word/document.xml")
                    return
                }
                childIndex += 1
            }
        }
    }

    /// 設定文字效果
    public mutating func setTextEffect(
        at paragraphIndex: Int,
        effect: TextEffect
    ) throws {
        guard paragraphIndex >= 0 && paragraphIndex < getParagraphs().count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        var childIndex = 0
        for (i, child) in body.children.enumerated() {
            if case .paragraph(var para) = child {
                if childIndex == paragraphIndex {
                    for j in 0..<para.runs.count {
                        para.runs[j].properties.textEffect = effect
                    }
                    body.children[i] = .paragraph(para)
                    modifiedParts.insert("word/document.xml")
                    return
                }
                childIndex += 1
            }
        }
    }

    // MARK: - Drawing Operations

    /// 插入繪圖元素（圖片）到指定段落
    public mutating func insertDrawing(_ drawing: Drawing, at paragraphIndex: Int) throws {
        guard paragraphIndex >= 0 && paragraphIndex <= getParagraphs().count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        // 建立包含繪圖的 Run
        let drawingRun = Run.withDrawing(drawing)

        // 如果段落索引有效，將繪圖添加到該段落
        var childIndex = 0
        for (i, child) in body.children.enumerated() {
            if case .paragraph(var para) = child {
                if childIndex == paragraphIndex {
                    para.runs.append(drawingRun)
                    body.children[i] = .paragraph(para)
                    modifiedParts.insert("word/document.xml")
                    return
                }
                childIndex += 1
            }
        }

        // 如果超出範圍，創建新段落
        var newPara = Paragraph()
        newPara.runs = [drawingRun]
        body.children.append(.paragraph(newPara))
        modifiedParts.insert("word/document.xml")
    }

    // MARK: - Field Code Operations

    /// 插入欄位代碼到指定段落
    public mutating func insertFieldCode<F: FieldCode>(_ field: F, at paragraphIndex: Int) throws {
        guard paragraphIndex >= 0 && paragraphIndex <= getParagraphs().count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        // 產生欄位 XML 並包裝成 Run
        let fieldXML = field.toFieldXML()
        var fieldRun = Run(text: "")
        fieldRun.rawXML = fieldXML  // 使用 raw XML 方式

        var childIndex = 0
        for (i, child) in body.children.enumerated() {
            if case .paragraph(var para) = child {
                if childIndex == paragraphIndex {
                    para.runs.append(fieldRun)
                    body.children[i] = .paragraph(para)
                    modifiedParts.insert("word/document.xml")
                    return
                }
                childIndex += 1
            }
        }

        // 如果超出範圍，創建新段落
        var newPara = Paragraph()
        newPara.runs = [fieldRun]
        body.children.append(.paragraph(newPara))
        modifiedParts.insert("word/document.xml")
    }

    // MARK: - Content Control (SDT) Operations

    /// 插入內容控制項到指定段落
    ///
    /// **v0.15.0+ (#44 task 3.0)**: SDT is set as a `Paragraph.contentControls`
    /// entry (proper sibling of runs), not as a `Run.rawXML` blob (which
    /// produced malformed `<w:p><w:r><w:sdt>...</w:sdt></w:r></w:p>`). The
    /// new structure emits as proper Word XML: `<w:p><w:sdt>...</w:sdt></w:p>`.
    /// SDTParser (#44 task 3.1) reads this back into ContentControl.
    public mutating func insertContentControl(_ control: ContentControl, at paragraphIndex: Int) throws {
        guard paragraphIndex >= 0 && paragraphIndex <= getParagraphs().count else {
            throw WordError.invalidIndex(paragraphIndex)
        }

        // SDT-only paragraph: no runs, single ContentControl child.
        var sdtPara = Paragraph()
        sdtPara.contentControls = [control]

        // 插入到指定位置
        var childIndex = 0
        var insertPosition = body.children.count

        for (i, child) in body.children.enumerated() {
            if case .paragraph = child {
                if childIndex == paragraphIndex {
                    insertPosition = i
                    break
                }
                childIndex += 1
            }
        }

        body.children.insert(.paragraph(sdtPara), at: insertPosition)
        modifiedParts.insert("word/document.xml")
    }

    // MARK: - Repeating Section Operations

    /// 插入重複區段到指定位置
    public mutating func insertRepeatingSection(_ section: RepeatingSection, at index: Int) throws {
        guard index >= 0 && index <= body.children.count else {
            throw WordError.invalidIndex(index)
        }

        // 產生重複區段 XML
        let sectionXML = section.toXML()
        var sectionPara = Paragraph()
        var sectionRun = Run(text: "")
        sectionRun.rawXML = sectionXML
        sectionPara.runs = [sectionRun]

        // 插入到指定位置
        body.children.insert(.paragraph(sectionPara), at: min(index, body.children.count))
        modifiedParts.insert("word/document.xml")
    }

    // MARK: - SDT ID Allocation (che-word-mcp-content-controls-read-write task 2.3)

    /// 配置下一個可用的 SDT id，採用「全文件最大 id + 1」策略。
    ///
    /// 掃描 body 所有 paragraph 內 run 的 rawXML（目前 SDT 仍以 rawXML blob 儲存，
    /// task 3.1 落地後會擴充掃描已解析的 ContentControl 樹）。若文件沒有任何 SDT
    /// 則回傳 1。
    ///
    /// - Note: 目前未實作 per-session 快取；每次呼叫都做全文件掃描。
    ///   對小型文件（< 100 SDTs）效能可接受；大型文件可在 task 3.x 完成後加上快取。
    public func allocateSdtId() -> Int {
        var maxId = 0
        for child in body.children {
            maxId = max(maxId, Self.extractMaxSdtIdFromBodyChild(child))
        }
        return maxId + 1
    }

    /// Recursive scan of a BodyChild for the maximum SDT id. Handles
    /// paragraph-level SDTs (v0.15.0+ structured + legacy rawXML) and
    /// block-level SDTs (task 3.4). Tables traversed for paragraphs inside
    /// cells.
    private static func extractMaxSdtIdFromBodyChild(_ child: BodyChild) -> Int {
        switch child {
        case .bookmarkMarker, .rawBlockElement:
            // Body-level markers contain no SDTs (#58).
            return 0
        case .paragraph(let paragraph):
            var maxId = 0
            for run in paragraph.runs {
                guard let raw = run.rawXML, raw.contains("<w:sdt>") else { continue }
                maxId = max(maxId, extractMaxSdtId(from: raw))
            }
            for control in paragraph.contentControls {
                maxId = max(maxId, extractMaxSdtIdFromControl(control))
            }
            return maxId
        case .table(let table):
            var maxId = 0
            for row in table.rows {
                for cell in row.cells {
                    for para in cell.paragraphs {
                        maxId = max(maxId, extractMaxSdtIdFromBodyChild(.paragraph(para)))
                    }
                }
            }
            return maxId
        case .contentControl(let control, let children):
            var maxId = control.sdt.id ?? 0
            for c in children {
                maxId = max(maxId, extractMaxSdtIdFromBodyChild(c))
            }
            return maxId
        }
    }

    /// Recursively scan a ContentControl tree for the maximum SDT id.
    /// Task 3.0 helper; replaces rawXML-only scan.
    private static func extractMaxSdtIdFromControl(_ control: ContentControl) -> Int {
        var maxId = control.sdt.id ?? 0
        for child in control.children {
            maxId = max(maxId, extractMaxSdtIdFromControl(child))
        }
        return maxId
    }

    // MARK: - Style Mutations (#44 styles-sections-numbering-foundations Phase 3)

    /// Traverse the `basedOn` reference chain from `styleId` upward to root.
    ///
    /// Returns ordered chain starting with the queried style itself, ending
    /// at a root style (one with no `basedOn`). When the chain contains a
    /// cycle, traversal stops at the first revisited style and the prefix
    /// is returned (callers MAY detect a cycle when the chain length seems
    /// short relative to expectation).
    ///
    /// Returns empty array when `styleId` does not exist.
    public func getStyleInheritanceChain(styleId: String) -> [Style] {
        guard styles.contains(where: { $0.id == styleId }) else { return [] }
        var chain: [Style] = []
        var visited: Set<String> = []
        var currentId: String? = styleId
        while let id = currentId {
            if visited.contains(id) { break }
            visited.insert(id)
            guard let style = styles.first(where: { $0.id == id }) else { break }
            chain.append(style)
            currentId = style.basedOn
        }
        return chain
    }

    /// Set bidirectional `<w:link>` between a paragraph and a character style.
    ///
    /// - Throws: `WordError.styleNotFound(id)` when either id is missing.
    ///   `WordError.typeMismatch(expected:actual:)` when the paragraph style is
    ///   not `.paragraph` type or character style is not `.character` type.
    public mutating func linkStyles(paragraphStyleId: String, characterStyleId: String) throws {
        guard let pIdx = styles.firstIndex(where: { $0.id == paragraphStyleId }) else {
            throw WordError.styleNotFound(paragraphStyleId)
        }
        guard let cIdx = styles.firstIndex(where: { $0.id == characterStyleId }) else {
            throw WordError.styleNotFound(characterStyleId)
        }
        guard styles[pIdx].type == .paragraph else {
            throw WordError.typeMismatch(expected: "paragraph", actual: styles[pIdx].type.rawValue)
        }
        guard styles[cIdx].type == .character else {
            throw WordError.typeMismatch(expected: "character", actual: styles[cIdx].type.rawValue)
        }
        styles[pIdx].linkedStyleId = characterStyleId
        styles[cIdx].linkedStyleId = paragraphStyleId
        modifiedParts.insert("word/styles.xml")
    }

    /// Add or replace a localized name alias on a style.
    ///
    /// If an alias with the same `lang` already exists, it is replaced rather
    /// than duplicated.
    ///
    /// - Throws: `WordError.styleNotFound(styleId)` when the style is missing.
    public mutating func addStyleNameAlias(styleId: String, lang: String, name: String) throws {
        guard let idx = styles.firstIndex(where: { $0.id == styleId }) else {
            throw WordError.styleNotFound(styleId)
        }
        if let existing = styles[idx].aliases.firstIndex(where: { $0.lang == lang }) {
            styles[idx].aliases[existing].name = name
        } else {
            styles[idx].aliases.append(StyleAlias(lang: lang, name: name))
        }
        modifiedParts.insert("word/styles.xml")
    }

    /// Replace the document's `latentStyles` collection wholesale.
    /// Pass an empty array to remove the `<w:latentStyles>` block entirely.
    public mutating func setLatentStyles(_ entries: [LatentStyle]) {
        latentStyles = entries
        modifiedParts.insert("word/styles.xml")
    }

    // MARK: - Table / Hyperlink / Header Mutations (#44 tables-hyperlinks-headers-builtin)

    /// Find the body table at index `tableIndex` (0-based, only top-level tables).
    /// Throws `WordError.invalidIndex` when out of bounds.
    private func findTable(at tableIndex: Int) throws -> Table {
        let tables = body.children.compactMap { c -> Table? in
            if case .table(let t) = c { return t }
            return nil
        }
        guard tableIndex >= 0 && tableIndex < tables.count else {
            throw WordError.invalidIndex(tableIndex)
        }
        return tables[tableIndex]
    }

    /// Replace the table at `tableIndex` with a mutated version.
    private mutating func replaceTable(at tableIndex: Int, with newTable: Table) throws {
        var seen = 0
        for i in 0..<body.children.count {
            if case .table = body.children[i] {
                if seen == tableIndex {
                    body.children[i] = .table(newTable)
                    modifiedParts.insert("word/document.xml")
                    return
                }
                seen += 1
            }
        }
        throw WordError.invalidIndex(tableIndex)
    }

    /// Apply or replace a `<w:tblStylePr>` block on the table.
    public mutating func setTableConditionalStyle(
        tableIndex: Int,
        type: TableConditionalStyleType,
        properties: TableConditionalStyleProperties
    ) throws {
        var table = try findTable(at: tableIndex)
        if let existing = table.conditionalStyles.firstIndex(where: { $0.type == type }) {
            table.conditionalStyles[existing] = TableConditionalStyle(type: type, properties: properties)
        } else {
            table.conditionalStyles.append(TableConditionalStyle(type: type, properties: properties))
        }
        try replaceTable(at: tableIndex, with: table)
    }

    /// Set explicit table layout (`<w:tblLayout w:type>`).
    public mutating func setTableLayout(tableIndex: Int, type: TableLayout) throws {
        var table = try findTable(at: tableIndex)
        table.explicitLayout = type
        try replaceTable(at: tableIndex, with: table)
    }

    /// Mark a row as header-row (`<w:tblHeader/>`) so it repeats on page break.
    public mutating func setHeaderRow(tableIndex: Int, rowIndex: Int) throws {
        var table = try findTable(at: tableIndex)
        guard rowIndex >= 0 && rowIndex < table.rows.count else {
            throw WordError.invalidIndex(rowIndex)
        }
        table.rows[rowIndex].properties.isHeader = true
        try replaceTable(at: tableIndex, with: table)
    }

    /// Set table-level left indent (`<w:tblInd>`) in twips.
    public mutating func setTableIndent(tableIndex: Int, value: Int) throws {
        var table = try findTable(at: tableIndex)
        table.tableIndent = value
        try replaceTable(at: tableIndex, with: table)
    }

    /// Insert a new table inside the cell at (rowIndex, colIndex) of the
    /// parent table. Throws `nestedTooDeep` when nesting would exceed depth 5.
    public mutating func insertNestedTable(
        parentTableIndex: Int,
        rowIndex: Int,
        colIndex: Int,
        rows: Int,
        cols: Int
    ) throws {
        var table = try findTable(at: parentTableIndex)
        guard rowIndex >= 0 && rowIndex < table.rows.count else {
            throw WordError.invalidIndex(rowIndex)
        }
        guard colIndex >= 0 && colIndex < table.rows[rowIndex].cells.count else {
            throw WordError.invalidIndex(colIndex)
        }
        // Depth check: count existing nesting depth in target cell.
        let parentDepth = Self.computeMaxNestDepth(in: table.rows[rowIndex].cells[colIndex])
        guard parentDepth + 1 <= 5 else {
            throw WordError.nestedTooDeep(depth: parentDepth + 1, max: 5)
        }
        let nested = Table(rowCount: rows, columnCount: cols)
        table.rows[rowIndex].cells[colIndex].nestedTables.append(nested)
        try replaceTable(at: parentTableIndex, with: table)
    }

    /// Recursively compute the max nesting depth inside a cell. Returns 0
    /// when the cell has no nested tables.
    private static func computeMaxNestDepth(in cell: TableCell) -> Int {
        guard !cell.nestedTables.isEmpty else { return 0 }
        var maxDepth = 0
        for nested in cell.nestedTables {
            for row in nested.rows {
                for child in row.cells {
                    let depth = 1 + computeMaxNestDepth(in: child)
                    if depth > maxDepth { maxDepth = depth }
                }
            }
            // At least one nested table itself counts as depth 1.
            if maxDepth == 0 { maxDepth = 1 }
        }
        return maxDepth
    }

    /// Set the tooltip on an existing hyperlink identified by id. Pass `nil`
    /// to clear. Throws `hyperlinkNotFound` when the id does not match.
    public mutating func setHyperlinkTooltip(hyperlinkId: String, tooltip: String?) throws {
        var found = false
        for childIdx in 0..<body.children.count {
            guard case .paragraph(var para) = body.children[childIdx] else { continue }
            if let hlIdx = para.hyperlinks.firstIndex(where: { $0.id == hyperlinkId }) {
                para.hyperlinks[hlIdx].tooltip = tooltip
                body.children[childIdx] = .paragraph(para)
                found = true
                break
            }
        }
        guard found else { throw WordError.hyperlinkNotFound(hyperlinkId) }
        modifiedParts.insert("word/document.xml")
    }

    /// Add a header part of the given type, returning the assigned filename
    /// (e.g., `header2.xml`). Marks both the new header file AND document.xml
    /// dirty (sectPr changes).
    @discardableResult
    public mutating func addHeaderOfType(text: String, type: HeaderFooterType) throws -> String {
        // Find next available header index.
        let usedIndices = headers.compactMap { h -> Int? in
            // file name format: "headerN.xml"
            let fn = h.fileName
            let stripped = fn.replacingOccurrences(of: "header", with: "").replacingOccurrences(of: ".xml", with: "")
            return Int(stripped)
        }
        let nextIdx = (usedIndices.max() ?? 0) + 1
        let rId = "rId\(1000 + nextIdx)"  // simple synthetic rId; conflict-checked elsewhere
        let header = Header(
            id: rId,
            paragraphs: [Paragraph(text: text)],
            type: type,
            originalFileName: "header\(nextIdx).xml"
        )
        // Replace any existing header of same type
        if let existingIdx = headers.firstIndex(where: { $0.type == type }) {
            headers[existingIdx] = header
        } else {
            headers.append(header)
        }
        modifiedParts.insert("word/\(header.fileName)")
        modifiedParts.insert("word/document.xml")
        return header.fileName
    }

    /// Toggle `<w:evenAndOddHeaders/>` in settings.xml.
    public mutating func setEvenAndOddHeaders(_ enabled: Bool) {
        evenAndOddHeaders = enabled
        modifiedParts.insert("word/settings.xml")
    }

    /// Clone the source header's content into a new header part of the given
    /// type. Returns the new file name. Used by unlink-from-previous semantics.
    @discardableResult
    public mutating func cloneHeaderForSection(
        sourceFileName: String,
        targetSectionIndex: Int,
        type: HeaderFooterType
    ) throws -> String {
        guard let source = headers.first(where: { $0.fileName == sourceFileName }) else {
            throw WordError.parseError("source header \(sourceFileName) not found")
        }
        // Allocate next index.
        let usedIndices = headers.compactMap { h -> Int? in
            let stripped = h.fileName.replacingOccurrences(of: "header", with: "").replacingOccurrences(of: ".xml", with: "")
            return Int(stripped)
        }
        let nextIdx = (usedIndices.max() ?? 0) + 1
        let rId = "rId\(1000 + nextIdx)"
        let cloned = Header(
            id: rId,
            paragraphs: source.paragraphs,  // deep copy via value semantics
            type: type,
            originalFileName: "header\(nextIdx).xml"
        )
        headers.append(cloned)
        modifiedParts.insert("word/\(cloned.fileName)")
        modifiedParts.insert("word/document.xml")
        return cloned.fileName
    }

    // MARK: - Section Mutations (#44 styles-sections-numbering-foundations Phase 5)

    /// SectionInfo: summary returned by `getAllSections`.
    public struct SectionInfo: Equatable {
        public let sectionIndex: Int
        public let paragraphRange: ClosedRange<Int>
        public let pageSize: PageSize
        public let pageMargins: PageMargins
        public let orientation: PageOrientation
        public let columns: Int
        public let lineNumbers: LineNumbers?
        public let verticalAlignment: SectionVerticalAlignment?
        public let pageNumberFormat: SectionPageNumberFormat?
        public let sectionBreakType: SectionBreakType?
        public let titlePageDistinct: Bool
        public let headerReferences: HeaderFooterReferences
        public let footerReferences: HeaderFooterReferences
    }

    /// Set line numbering on the document's section properties.
    /// Currently the document holds a single `sectionProperties` — `sectionIndex`
    /// must be 0 until multi-section support lands.
    ///
    /// Throws `WordError.invalidIndex(sectionIndex)` for any non-zero index.
    public mutating func setSectionLineNumbers(
        sectionIndex: Int,
        countBy: Int,
        start: Int? = nil,
        restart: LineNumberRestart = .continuous
    ) throws {
        guard sectionIndex == 0 else { throw WordError.invalidIndex(sectionIndex) }
        sectionProperties.lineNumbers = LineNumbers(countBy: countBy, start: start, restart: restart)
        modifiedParts.insert("word/document.xml")
    }

    /// Set the vertical alignment of section content.
    public mutating func setSectionVerticalAlignment(
        sectionIndex: Int,
        alignment: SectionVerticalAlignment
    ) throws {
        guard sectionIndex == 0 else { throw WordError.invalidIndex(sectionIndex) }
        sectionProperties.verticalAlignment = alignment
        modifiedParts.insert("word/document.xml")
    }

    /// Set page number format (decimal / Roman / letter) and optional start value.
    public mutating func setSectionPageNumberFormat(
        sectionIndex: Int,
        start: Int? = nil,
        format: SectionPageNumberFormat
    ) throws {
        guard sectionIndex == 0 else { throw WordError.invalidIndex(sectionIndex) }
        sectionProperties.pageNumberFormat = format
        sectionProperties.pageNumberStartValue = start
        modifiedParts.insert("word/document.xml")
    }

    /// Set section break type (nextPage / continuous / evenPage / oddPage).
    public mutating func setSectionBreakType(
        sectionIndex: Int,
        type: SectionBreakType
    ) throws {
        guard sectionIndex == 0 else { throw WordError.invalidIndex(sectionIndex) }
        sectionProperties.sectionBreakType = type
        modifiedParts.insert("word/document.xml")
    }

    /// Toggle `<w:titlePg/>` in the section's properties.
    public mutating func setTitlePageDistinct(
        sectionIndex: Int,
        enabled: Bool
    ) throws {
        guard sectionIndex == 0 else { throw WordError.invalidIndex(sectionIndex) }
        sectionProperties.titlePageDistinct = enabled
        modifiedParts.insert("word/document.xml")
    }

    /// Assign per-type header / footer rId references on a section.
    /// Provided keys overwrite; omitted keys are left unchanged.
    public mutating func setSectionHeaderFooterReferences(
        sectionIndex: Int,
        headerDefault: String? = nil,
        headerFirst: String? = nil,
        headerEven: String? = nil,
        footerDefault: String? = nil,
        footerFirst: String? = nil,
        footerEven: String? = nil
    ) throws {
        guard sectionIndex == 0 else { throw WordError.invalidIndex(sectionIndex) }
        if let h = headerDefault { sectionProperties.headerReferences.defaultRef = h }
        if let h = headerFirst { sectionProperties.headerReferences.firstRef = h }
        if let h = headerEven { sectionProperties.headerReferences.evenRef = h }
        if let f = footerDefault { sectionProperties.footerReferences.defaultRef = f }
        if let f = footerFirst { sectionProperties.footerReferences.firstRef = f }
        if let f = footerEven { sectionProperties.footerReferences.evenRef = f }
        modifiedParts.insert("word/document.xml")
    }

    /// Return one SectionInfo per section. Currently always one entry
    /// (single-section document model — multi-section split lands later).
    public func getAllSections() -> [SectionInfo] {
        let paraIndices = body.children.enumerated().compactMap { (i, c) -> Int? in
            if case .paragraph = c { return i }
            return nil
        }
        let lastIdx = max(paraIndices.count - 1, 0)
        return [
            SectionInfo(
                sectionIndex: 0,
                paragraphRange: 0...lastIdx,
                pageSize: sectionProperties.pageSize,
                pageMargins: sectionProperties.pageMargins,
                orientation: sectionProperties.orientation,
                columns: sectionProperties.columns,
                lineNumbers: sectionProperties.lineNumbers,
                verticalAlignment: sectionProperties.verticalAlignment,
                pageNumberFormat: sectionProperties.pageNumberFormat,
                sectionBreakType: sectionProperties.sectionBreakType,
                titlePageDistinct: sectionProperties.titlePageDistinct,
                headerReferences: sectionProperties.headerReferences,
                footerReferences: sectionProperties.footerReferences
            )
        ]
    }

    // MARK: - Numbering Mutations (#44 styles-sections-numbering-foundations Phase 4)

    /// Create a new abstractNum + paired num in numbering.xml.
    ///
    /// - Throws: `WordError.invalidIndex(0)` when `levels` is empty,
    ///   `WordError.invalidIndex(levels.count)` when more than 9 levels.
    /// - Returns: the new numId allocated for the paired `<w:num>`.
    public mutating func createNumberingDefinition(levels: [Level]) throws -> Int {
        guard !levels.isEmpty else { throw WordError.invalidIndex(0) }
        guard levels.count <= 9 else { throw WordError.invalidIndex(levels.count) }

        let abstractNumId = numbering.nextAbstractNumId
        let numId = numbering.nextNumId
        numbering.abstractNums.append(AbstractNum(abstractNumId: abstractNumId, levels: levels))
        numbering.nums.append(Num(numId: numId, abstractNumId: abstractNumId))
        modifiedParts.insert("word/numbering.xml")
        return numId
    }

    /// Add a `<w:lvlOverride>` to the named num. Replaces any existing
    /// override for the same level on that num.
    ///
    /// - Throws: `WordError.numIdNotFound(numId)` when the num does not exist.
    public mutating func overrideNumberingLevel(numId: Int, level: Int, startValue: Int) throws {
        guard let idx = numbering.nums.firstIndex(where: { $0.numId == numId }) else {
            throw WordError.numIdNotFound(numId)
        }
        if let existing = numbering.nums[idx].lvlOverrides.firstIndex(where: { $0.ilvl == level }) {
            numbering.nums[idx].lvlOverrides[existing].startOverride = startValue
        } else {
            numbering.nums[idx].lvlOverrides.append(LvlOverride(ilvl: level, startOverride: startValue))
        }
        modifiedParts.insert("word/numbering.xml")
    }

    /// Attach a numId+level to the paragraph at the given top-level body index.
    ///
    /// Marks BOTH `word/numbering.xml` (numId reference touched) AND
    /// `word/document.xml` (paragraph mutated) dirty.
    ///
    /// - Throws: `WordError.numIdNotFound(numId)` when the num is missing,
    ///   `WordError.invalidIndex(paragraphIndex)` when the paragraph index
    ///   is out of bounds.
    public mutating func assignNumberingToParagraph(paragraphIndex: Int, numId: Int, level: Int) throws {
        guard numbering.nums.contains(where: { $0.numId == numId }) else {
            throw WordError.numIdNotFound(numId)
        }
        let paragraphIndices = body.children.enumerated().compactMap { (i, child) -> Int? in
            if case .paragraph = child { return i }
            return nil
        }
        guard paragraphIndex >= 0 && paragraphIndex < paragraphIndices.count else {
            throw WordError.invalidIndex(paragraphIndex)
        }
        let actualIndex = paragraphIndices[paragraphIndex]
        if case .paragraph(var para) = body.children[actualIndex] {
            para.properties.numbering = NumberingInfo(numId: numId, level: level)
            body.children[actualIndex] = .paragraph(para)
        }
        modifiedParts.insert("word/document.xml")
        modifiedParts.insert("word/numbering.xml")
    }

    /// Continue an existing list — assign the same numId as a previously
    /// numbered paragraph to a new paragraph at level 0.
    ///
    /// - Throws: `WordError.numIdNotFound(numId)` or `WordError.invalidIndex`.
    public mutating func continueList(paragraphIndex: Int, previousListNumId: Int) throws {
        try assignNumberingToParagraph(paragraphIndex: paragraphIndex, numId: previousListNumId, level: 0)
    }

    /// Start a new list referencing an existing abstractNum, returning the
    /// freshly-allocated numId.
    ///
    /// - Throws: `WordError.abstractNumIdNotFound` or `WordError.invalidIndex`.
    @discardableResult
    public mutating func startNewList(paragraphIndex: Int, abstractNumId: Int, level: Int = 0) throws -> Int {
        guard numbering.abstractNums.contains(where: { $0.abstractNumId == abstractNumId }) else {
            throw WordError.abstractNumIdNotFound(abstractNumId)
        }
        let numId = numbering.nextNumId
        numbering.nums.append(Num(numId: numId, abstractNumId: abstractNumId))
        // Now assign — assignNumberingToParagraph marks both parts dirty.
        try assignNumberingToParagraph(paragraphIndex: paragraphIndex, numId: numId, level: level)
        return numId
    }

    /// Sweep `<w:num>` entries whose numId is not referenced by any paragraph
    /// (anywhere — top-level, inside tables, inside block-level SDTs).
    ///
    /// AbstractNums are templates and are NOT GCed even when unreferenced.
    ///
    /// - Returns: deleted numIds in numId order.
    @discardableResult
    public mutating func gcOrphanNumbering() -> [Int] {
        var referenced: Set<Int> = []
        func collectFromBodyChild(_ child: BodyChild) {
            switch child {
            case .bookmarkMarker, .rawBlockElement:
                // Body-level markers carry no numbering references (#58).
                return
            case .paragraph(let p):
                if let numInfo = p.properties.numbering { referenced.insert(numInfo.numId) }
            case .table(let table):
                for row in table.rows {
                    for cell in row.cells {
                        for cellPara in cell.paragraphs {
                            if let numInfo = cellPara.properties.numbering {
                                referenced.insert(numInfo.numId)
                            }
                        }
                    }
                }
            case .contentControl(_, let children):
                for c in children { collectFromBodyChild(c) }
            }
        }
        for c in body.children { collectFromBodyChild(c) }
        // Headers / footers also have paragraphs that may carry numbering refs.
        for header in headers {
            for p in header.paragraphs {
                if let numInfo = p.properties.numbering { referenced.insert(numInfo.numId) }
            }
        }
        for footer in footers {
            for p in footer.paragraphs {
                if let numInfo = p.properties.numbering { referenced.insert(numInfo.numId) }
            }
        }

        let allNumIds = numbering.nums.map { $0.numId }.sorted()
        let orphans = allNumIds.filter { !referenced.contains($0) }
        guard !orphans.isEmpty else { return [] }
        numbering.nums.removeAll { orphans.contains($0.numId) }
        modifiedParts.insert("word/numbering.xml")
        return orphans
    }

    // MARK: - Content Control Mutations (#44 Phase 4)

    /// Replace the text content of a text-bearing SDT identified by id.
    ///
    /// Locates the ContentControl anywhere in the document tree (paragraph
    /// children, block-level wrappers, nested children). Replaces the
    /// content with a single run containing `newText`. The SDT's `<w:sdtPr>`
    /// (tag, alias, type, lockType, placeholder) is preserved untouched.
    ///
    /// - Throws: `WordError.contentControlNotFound(id)` when no SDT has the
    ///   given id. `WordError.unsupportedSDTType(type)` when the target's
    ///   type cannot hold plain text (picture, dropDownList, comboBox,
    ///   checkbox, group, repeatingSection).
    public mutating func updateContentControl(id: Int, newText: String) throws {
        let escaped = Self.escapeContentControlText(newText)
        let newContent = "<w:p><w:r><w:t xml:space=\"preserve\">\(escaped)</w:t></w:r></w:p>"

        let result = try mutateContentControl(id: id) { control in
            switch control.sdt.type {
            case .picture, .dropDownList, .comboBox, .checkbox, .group,
                 .repeatingSection, .repeatingSectionItem:
                throw WordError.unsupportedSDTType(control.sdt.type)
            case .richText, .plainText, .date, .bibliography, .citation:
                control.content = newContent
            }
        }
        guard result else { throw WordError.contentControlNotFound(id) }
        modifiedParts.insert("word/document.xml")
    }

    /// Replace the entire `<w:sdtContent>` region of a ContentControl with
    /// the supplied XML fragment. Validates the input against an element
    /// whitelist before applying.
    ///
    /// - Throws: `WordError.contentControlNotFound(id)` when not found.
    ///   `WordError.disallowedElement(name)` when `contentXML` contains
    ///   `<w:sdt>`, `<w:body>`, `<w:sectPr>`, or an XML declaration.
    public mutating func replaceContentControlContent(id: Int, contentXML: String) throws {
        if let bad = Self.disallowedElement(in: contentXML) {
            throw WordError.disallowedElement(bad)
        }
        let result = try mutateContentControl(id: id) { control in
            control.content = contentXML
        }
        guard result else { throw WordError.contentControlNotFound(id) }
        modifiedParts.insert("word/document.xml")
    }

    /// Remove a ContentControl from the document tree.
    ///
    /// - Parameters:
    ///   - id: SDT id to remove.
    ///   - keepContent: when `true` (default), the SDT's content is unwrapped
    ///     into the parent container at the SDT's former position. When
    ///     `false`, the SDT and its content are removed entirely.
    /// - Throws: `WordError.contentControlNotFound(id)` when not found.
    public mutating func deleteContentControl(id: Int, keepContent: Bool = true) throws {
        let removed = removeContentControl(id: id, keepContent: keepContent, in: &body.children)
        guard removed else { throw WordError.contentControlNotFound(id) }
        modifiedParts.insert("word/document.xml")
    }

    // MARK: - Content Control Mutation Helpers

    /// Walk the body tree and apply `mutate` to the first ContentControl with
    /// matching id (paragraph-level, block-level, or nested). Returns true
    /// when found.
    private mutating func mutateContentControl(
        id: Int,
        mutate: (inout ContentControl) throws -> Void
    ) throws -> Bool {
        for i in 0..<body.children.count {
            switch body.children[i] {
            case .bookmarkMarker, .rawBlockElement:
                // Body-level markers contain no content controls (#58).
                continue
            case .paragraph(var para):
                if try Self.mutateInControlList(id: id, controls: &para.contentControls, mutate: mutate) {
                    body.children[i] = .paragraph(para)
                    return true
                }
            case .table(var table):
                if try mutateInTable(id: id, table: &table, mutate: mutate) {
                    body.children[i] = .table(table)
                    return true
                }
            case .contentControl(var outer, var children):
                if outer.sdt.id == id {
                    try mutate(&outer)
                    body.children[i] = .contentControl(outer, children: children)
                    return true
                }
                if try mutateInBodyChildList(id: id, children: &children, mutate: mutate) {
                    body.children[i] = .contentControl(outer, children: children)
                    return true
                }
            }
        }
        return false
    }

    private mutating func mutateInTable(
        id: Int,
        table: inout Table,
        mutate: (inout ContentControl) throws -> Void
    ) throws -> Bool {
        for r in 0..<table.rows.count {
            for c in 0..<table.rows[r].cells.count {
                for p in 0..<table.rows[r].cells[c].paragraphs.count {
                    var para = table.rows[r].cells[c].paragraphs[p]
                    if try Self.mutateInControlList(id: id, controls: &para.contentControls, mutate: mutate) {
                        table.rows[r].cells[c].paragraphs[p] = para
                        return true
                    }
                }
            }
        }
        return false
    }

    private mutating func mutateInBodyChildList(
        id: Int,
        children: inout [BodyChild],
        mutate: (inout ContentControl) throws -> Void
    ) throws -> Bool {
        for i in 0..<children.count {
            switch children[i] {
            case .bookmarkMarker, .rawBlockElement:
                // Body-level markers contain no content controls (#58).
                continue
            case .paragraph(var para):
                if try Self.mutateInControlList(id: id, controls: &para.contentControls, mutate: mutate) {
                    children[i] = .paragraph(para)
                    return true
                }
            case .table(var table):
                if try mutateInTable(id: id, table: &table, mutate: mutate) {
                    children[i] = .table(table)
                    return true
                }
            case .contentControl(var outer, var inner):
                if outer.sdt.id == id {
                    try mutate(&outer)
                    children[i] = .contentControl(outer, children: inner)
                    return true
                }
                if try mutateInBodyChildList(id: id, children: &inner, mutate: mutate) {
                    children[i] = .contentControl(outer, children: inner)
                    return true
                }
            }
        }
        return false
    }

    private static func mutateInControlList(
        id: Int,
        controls: inout [ContentControl],
        mutate: (inout ContentControl) throws -> Void
    ) throws -> Bool {
        for i in 0..<controls.count {
            if controls[i].sdt.id == id {
                try mutate(&controls[i])
                return true
            }
            // Recurse into nested children
            if try mutateInControlList(id: id, controls: &controls[i].children, mutate: mutate) {
                return true
            }
        }
        return false
    }

    /// Remove a ContentControl by id from any container in the body tree.
    /// Returns true when found and removed.
    private func removeContentControl(
        id: Int,
        keepContent: Bool,
        in children: inout [BodyChild]
    ) -> Bool {
        for i in 0..<children.count {
            switch children[i] {
            case .bookmarkMarker, .rawBlockElement:
                // Body-level markers contain no content controls (#58).
                continue
            case .paragraph(var para):
                if Self.removeFromControlList(id: id, keepContent: keepContent, controls: &para.contentControls) {
                    children[i] = .paragraph(para)
                    return true
                }
            case .table(var table):
                for r in 0..<table.rows.count {
                    for c in 0..<table.rows[r].cells.count {
                        for p in 0..<table.rows[r].cells[c].paragraphs.count {
                            var para = table.rows[r].cells[c].paragraphs[p]
                            if Self.removeFromControlList(id: id, keepContent: keepContent, controls: &para.contentControls) {
                                table.rows[r].cells[c].paragraphs[p] = para
                                children[i] = .table(table)
                                return true
                            }
                        }
                    }
                }
            case .contentControl(let outer, var inner):
                if outer.sdt.id == id {
                    if keepContent {
                        children.replaceSubrange(i...i, with: inner)
                    } else {
                        children.remove(at: i)
                    }
                    return true
                }
                if removeContentControl(id: id, keepContent: keepContent, in: &inner) {
                    children[i] = .contentControl(outer, children: inner)
                    return true
                }
            }
        }
        return false
    }

    /// Remove a ContentControl by id from a list of paragraph-level controls
    /// (recursing into nested children). When `keepContent=true`, the
    /// removed control's content is wrapped in a Run.rawXML carrier and
    /// added to the parent paragraph — but since we don't have access to
    /// the parent paragraph's runs here, the content is dropped on
    /// paragraph-level removal. Keep-content for paragraph-level SDTs is
    /// best handled at the BodyChild.contentControl level (block-level SDT).
    /// Returns true when found.
    private static func removeFromControlList(
        id: Int,
        keepContent: Bool,
        controls: inout [ContentControl]
    ) -> Bool {
        for i in 0..<controls.count {
            if controls[i].sdt.id == id {
                controls.remove(at: i)
                return true
            }
            // Recurse into nested children
            if removeFromControlList(id: id, keepContent: keepContent, controls: &controls[i].children) {
                return true
            }
        }
        return false
    }

    /// Detect disallowed elements in user-supplied content XML.
    /// Returns the offending element name, or nil if input is acceptable.
    /// Used by `replaceContentControlContent` whitelist check.
    private static func disallowedElement(in xml: String) -> String? {
        let trimmed = xml.trimmingCharacters(in: .whitespacesAndNewlines)
        if trimmed.hasPrefix("<?xml") {
            return "<?xml"
        }
        // Each entry: (literal prefix to detect, name reported in the error).
        // Element names in errors are without leading "<" so callers can match
        // by element local name (e.g., "w:sdt" matches the SDT spec wording).
        let banned: [(detect: String, name: String)] = [
            ("<w:sdt", "w:sdt"),
            ("<w:body", "w:body"),
            ("<w:sectPr", "w:sectPr"),
        ]
        for entry in banned where xml.contains(entry.detect) {
            return entry.name
        }
        return nil
    }

    private static func escapeContentControlText(_ s: String) -> String {
        return s
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
    }

    /// 從一段包含 SDT 的 rawXML 中抽出所有 `<w:id w:val="N"/>` 的最大值。
    /// 因為 rawXML 以 `<w:sdt>` 為界，所有 `<w:id w:val=>` 都屬於 SDT 家族
    /// （Bookmark / Comment 使用不同屬性形式，不會誤判）。
    private static func extractMaxSdtId(from xml: String) -> Int {
        var maxId = 0
        let pattern = #"<w:id w:val=\"(\d+)\"\s*/>"#
        guard let regex = try? NSRegularExpression(pattern: pattern) else { return 0 }
        let range = NSRange(xml.startIndex..., in: xml)
        regex.enumerateMatches(in: xml, range: range) { match, _, _ in
            guard let match = match,
                  let valueRange = Range(match.range(at: 1), in: xml),
                  let value = Int(xml[valueRange]) else { return }
            maxId = max(maxId, value)
        }
        return maxId
    }
}
