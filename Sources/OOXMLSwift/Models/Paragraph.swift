import Foundation

/// 段落 (Paragraph) - Word 文件的基本結構單元
public struct Paragraph: Equatable {
    public var runs: [Run]
    public var properties: ParagraphProperties
    public var hasPageBreak: Bool = false      // 是否為分頁符段落
    public var bookmarks: [Bookmark] = []      // 段落內的書籤
    public var hyperlinks: [Hyperlink] = []    // 段落內的超連結
    public var commentIds: [Int] = []          // 段落關聯的註解 ID
    public var footnoteIds: [Int] = []         // 段落內的腳註 ID
    public var endnoteIds: [Int] = []          // 段落內的尾註 ID
    public var revisions: [Revision] = []      // 段落內的修訂記錄（w:ins/w:del）
    public var semantic: SemanticAnnotation?  // 語義標註

    /// v0.18.0+ (che-word-mcp#45): id of a `Revision` (type `.paragraphChange`)
    /// in `revisions` whose `previousFormat` is captured by
    /// `previousProperties` below. When set, `Paragraph.toXML()` emits
    /// `<w:pPrChange>` inside this paragraph's `<w:pPr>` block.
    public var paragraphFormatChangeRevisionId: Int?

    /// v0.18.0+ (che-word-mcp#45): pre-mutation paragraph properties for a
    /// tracked `<w:pPrChange>` revision. Paired with
    /// `paragraphFormatChangeRevisionId`.
    public var previousProperties: ParagraphProperties?

    /// v0.15.0+ (che-word-mcp#44, task 3.0): structured Content Controls
    /// (SDTs) appearing as siblings of runs inside this paragraph. Emitted
    /// after `runs` in `toXML()`. The previous architecture stuffed entire
    /// `<w:sdt>...</w:sdt>` strings into `Run.rawXML`, producing malformed
    /// `<w:p><w:r><w:sdt>...</w:sdt></w:r></w:p>` (SDT inside Run). Now SDTs
    /// emit as proper siblings of runs: `<w:p><w:r>...</w:r><w:sdt>...</w:sdt></w:p>`.
    public var contentControls: [ContentControl] = []

    public init(runs: [Run] = [], properties: ParagraphProperties = ParagraphProperties()) {
        self.runs = runs
        self.properties = properties
    }

    /// 便利初始化器：直接用文字建立段落
    public init(text: String, properties: ParagraphProperties = ParagraphProperties()) {
        self.runs = [Run(text: text)]
        self.properties = properties
    }

    /// 取得段落純文字
    public func getText() -> String {
        var text = runs.map { $0.text }.joined()
        // 加入超連結文字
        for hyperlink in hyperlinks {
            text += hyperlink.text
        }
        return text
    }
}

// MARK: - Paragraph Properties

/// 段落格式屬性
public struct ParagraphProperties: Equatable {
    public var alignment: Alignment?
    public var spacing: Spacing?
    public var indentation: Indentation?
    public var style: String?                  // 樣式名稱 (e.g., "Heading1")
    public var numbering: NumberingInfo?       // 編號/項目符號
    public var keepNext: Bool = false          // 與下段同頁
    public var keepLines: Bool = false         // 段落不分頁
    public var pageBreakBefore: Bool = false   // 段落前分頁
    public var sectionBreak: SectionBreakType? // 分節符類型
    public var border: ParagraphBorder?        // 段落邊框
    public var shading: ParagraphShading?      // 段落底色

    public init() {}

    /// 合併格式（覆蓋非 nil 值）
    mutating func merge(with other: ParagraphProperties) {
        if let alignment = other.alignment { self.alignment = alignment }
        if let spacing = other.spacing { self.spacing = spacing }
        if let indentation = other.indentation { self.indentation = indentation }
        if let style = other.style { self.style = style }
        if let numbering = other.numbering { self.numbering = numbering }
        if other.keepNext { self.keepNext = true }
        if other.keepLines { self.keepLines = true }
        if other.pageBreakBefore { self.pageBreakBefore = true }
        if let sectionBreak = other.sectionBreak { self.sectionBreak = sectionBreak }
        if let border = other.border { self.border = border }
        if let shading = other.shading { self.shading = shading }
    }
}

// MARK: - Supporting Types

/// 對齊方式
public enum Alignment: String, Codable {
    case left = "left"
    case center = "center"
    case right = "right"
    case both = "both"      // 左右對齊（兩端對齊）
    case distribute = "distribute"  // 分散對齊
}

/// 段落對齊（Alignment 的別名）
public typealias ParagraphAlignment = Alignment

/// 段落間距
public struct Spacing: Equatable {
    public var before: Int?        // 段前間距 (1/20 點，twips)
    public var after: Int?         // 段後間距 (1/20 點)
    public var line: Int?          // 行高 (1/240 點 或 百分比)
    public var lineRule: LineRule? // 行高規則

    public init(before: Int? = nil, after: Int? = nil, line: Int? = nil, lineRule: LineRule? = nil) {
        self.before = before
        self.after = after
        self.line = line
        self.lineRule = lineRule
    }

    /// 便利方法：建立點數間距
    public static func points(before: Double? = nil, after: Double? = nil, lineSpacing: Double? = nil) -> Spacing {
        var spacing = Spacing()
        if let before = before {
            spacing.before = Int(before * 20)  // 轉換為 twips
        }
        if let after = after {
            spacing.after = Int(after * 20)
        }
        if let lineSpacing = lineSpacing {
            spacing.line = Int(lineSpacing * 240)  // 固定行高
            spacing.lineRule = .exact
        }
        return spacing
    }
}

/// 行高規則
public enum LineRule: String, Codable {
    case auto = "auto"          // 單行/1.5行/雙行
    case exact = "exact"        // 固定行高
    case atLeast = "atLeast"    // 最小行高
}

/// 縮排
public struct Indentation: Equatable {
    public var left: Int?          // 左縮排 (twips)
    public var right: Int?         // 右縮排 (twips)
    public var firstLine: Int?     // 首行縮排 (twips)
    public var hanging: Int?       // 凸排 (twips)

    public init(left: Int? = nil, right: Int? = nil, firstLine: Int? = nil, hanging: Int? = nil) {
        self.left = left
        self.right = right
        self.firstLine = firstLine
        self.hanging = hanging
    }

    /// 便利方法：建立字元縮排（假設 1 字元 = 240 twips）
    public static func characters(left: Int? = nil, right: Int? = nil, firstLine: Int? = nil) -> Indentation {
        var indent = Indentation()
        if let left = left {
            indent.left = left * 240
        }
        if let right = right {
            indent.right = right * 240
        }
        if let firstLine = firstLine {
            indent.firstLine = firstLine * 240
        }
        return indent
    }
}

/// 編號資訊
public struct NumberingInfo: Equatable {
    public var numId: Int          // 編號定義 ID
    public var level: Int          // 編號層級 (0-8)

    public init(numId: Int, level: Int = 0) {
        self.numId = numId
        self.level = level
    }
}

// MARK: - XML 生成

extension Paragraph {
    /// 轉換為 OOXML XML 字串
    public func toXML() -> String {
        var xml = "<w:p>"

        // Paragraph Properties
        let propsXML = properties.toXML()
        let formatChangeRevision: Revision? = {
            guard let id = paragraphFormatChangeRevisionId else { return nil }
            return revisions.first { $0.id == id }
        }()
        let needsPPr = !propsXML.isEmpty
            || properties.sectionBreak != nil
            || formatChangeRevision != nil
        if needsPPr {
            xml += "<w:pPr>\(propsXML)"

            // 分節符放在段落屬性中
            if let sectionBreak = properties.sectionBreak {
                xml += "<w:sectPr><w:type w:val=\"\(sectionBreak.rawValue)\"/></w:sectPr>"
            }

            // v0.18.0+ (che-word-mcp#45): pPrChange tracks paragraph format change
            if let revision = formatChangeRevision, let prev = previousProperties {
                xml += revision.toOpeningXML()
                xml += "<w:pPr>\(prev.toXML())</w:pPr>"
                xml += revision.toClosingXML()
            }

            xml += "</w:pPr>"
        }

        // 分頁符
        if hasPageBreak {
            xml += "<w:r><w:br w:type=\"page\"/></w:r>"
        }

        // 書籤開始標記
        for bookmark in bookmarks {
            xml += bookmark.toBookmarkStartXML()
        }

        // 註解範圍開始標記
        for commentId in commentIds {
            xml += "<w:commentRangeStart w:id=\"\(commentId)\"/>"
        }

        // Runs (v0.18.0: grouped by Run.revisionId so a single revision wrapping
        // multiple consecutive runs emits one <w:ins>/<w:del>/etc pair, not one
        // wrapper per run; deletion-type wrappers also substitute <w:t> with
        // <w:delText> per OOXML spec; runs flagged with formatChangeRevisionId
        // emit <w:rPrChange> inside their <w:rPr>)
        for group in Self.groupRunsByRevisionId(runs) {
            if let revisionId = group.revisionId,
               let revision = revisions.first(where: { $0.id == revisionId }) {
                xml += revision.toOpeningXML()
                for run in group.runs {
                    xml += Self.emitRun(run, asDelText: revision.type == .deletion,
                                        paragraphRevisions: revisions)
                }
                xml += revision.toClosingXML()
            } else {
                for run in group.runs {
                    xml += Self.emitRun(run, asDelText: false,
                                        paragraphRevisions: revisions)
                }
            }
        }

        // v0.15.0+ (#44 task 3.0): Content Controls as proper siblings of runs.
        for control in contentControls {
            xml += control.toXML()
        }

        // 超連結
        for hyperlink in hyperlinks {
            xml += hyperlink.toXML()
        }

        // 註解範圍結束標記和參照
        for commentId in commentIds {
            xml += "<w:commentRangeEnd w:id=\"\(commentId)\"/>"
            xml += "<w:r><w:commentReference w:id=\"\(commentId)\"/></w:r>"
        }

        // 腳註參照
        for footnoteId in footnoteIds {
            xml += "<w:r><w:rPr><w:rStyle w:val=\"FootnoteReference\"/></w:rPr><w:footnoteReference w:id=\"\(footnoteId)\"/></w:r>"
        }

        // 尾註參照
        for endnoteId in endnoteIds {
            xml += "<w:r><w:rPr><w:rStyle w:val=\"EndnoteReference\"/></w:rPr><w:endnoteReference w:id=\"\(endnoteId)\"/></w:r>"
        }

        // 書籤結束標記
        for bookmark in bookmarks {
            xml += bookmark.toBookmarkEndXML()
        }

        xml += "</w:p>"
        return xml
    }

    // MARK: - Revision Grouping (v0.18.0+)

    /// Internal grouping of consecutive runs sharing the same `revisionId`.
    fileprivate struct RunGroup {
        let revisionId: Int?
        var runs: [Run]
    }

    /// Walk `runs` once, coalescing consecutive entries that share the same
    /// `revisionId` (including the all-nil case). Used by `toXML()` to emit
    /// a single `<w:ins>`/`<w:del>` wrapper around each contiguous run group
    /// instead of one wrapper per run.
    fileprivate static func groupRunsByRevisionId(_ runs: [Run]) -> [RunGroup] {
        var groups: [RunGroup] = []
        for run in runs {
            if !groups.isEmpty, groups[groups.count - 1].revisionId == run.revisionId {
                groups[groups.count - 1].runs.append(run)
            } else {
                groups.append(RunGroup(revisionId: run.revisionId, runs: [run]))
            }
        }
        return groups
    }

    /// Emit a single run's XML with optional revision decorations:
    ///   - `asDelText == true`: substitute `<w:t>` → `<w:delText>`
    ///   - `run.formatChangeRevisionId` set + matching Revision in
    ///     `paragraphRevisions`: emit `<w:rPrChange>` inside `<w:rPr>`
    /// Falls through to `Run.toXML()` when neither decoration applies and the
    /// run has no rawXML override.
    fileprivate static func emitRun(_ run: Run, asDelText: Bool,
                                    paragraphRevisions: [Revision]) -> String {
        if let rawXML = run.rawXML { return rawXML }
        if let rawXML = run.properties.rawXML { return rawXML }

        let formatRevision: Revision? = {
            guard let id = run.formatChangeRevisionId else { return nil }
            return paragraphRevisions.first { $0.id == id }
        }()

        if !asDelText && formatRevision == nil {
            return run.toXML()
        }

        var xml = "<w:r>"

        let propsXML = run.properties.toXML()
        if !propsXML.isEmpty || formatRevision != nil {
            xml += "<w:rPr>\(propsXML)"
            if let revision = formatRevision, let prev = revision.previousFormat {
                xml += revision.toOpeningXML()
                xml += prev.toChangeXML()
                xml += revision.toClosingXML()
            }
            xml += "</w:rPr>"
        }

        if let drawing = run.drawing {
            xml += drawing.toXML()
        } else if !run.text.isEmpty || (run.rawElements?.isEmpty ?? true) {
            if asDelText {
                xml += "<w:delText xml:space=\"preserve\">\(escapeRunText(run.text))</w:delText>"
            } else {
                xml += "<w:t xml:space=\"preserve\">\(escapeRunText(run.text))</w:t>"
            }
        }

        if let rawElements = run.rawElements {
            for raw in rawElements {
                xml += raw.xml
            }
        }

        xml += "</w:r>"
        return xml
    }

    fileprivate static func escapeRunText(_ string: String) -> String {
        return string
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
            .replacingOccurrences(of: "'", with: "&apos;")
    }
}

extension ParagraphProperties {
    /// 轉換為 OOXML XML 字串
    public func toXML() -> String {
        var parts: [String] = []

        // 樣式
        if let style = style {
            parts.append("<w:pStyle w:val=\"\(style)\"/>")
        }

        // 編號
        if let numbering = numbering {
            parts.append("<w:numPr>")
            parts.append("<w:ilvl w:val=\"\(numbering.level)\"/>")
            parts.append("<w:numId w:val=\"\(numbering.numId)\"/>")
            parts.append("</w:numPr>")
        }

        // 對齊
        if let alignment = alignment {
            parts.append("<w:jc w:val=\"\(alignment.rawValue)\"/>")
        }

        // 間距
        if let spacing = spacing {
            var attrs: [String] = []
            if let before = spacing.before {
                attrs.append("w:before=\"\(before)\"")
            }
            if let after = spacing.after {
                attrs.append("w:after=\"\(after)\"")
            }
            if let line = spacing.line {
                attrs.append("w:line=\"\(line)\"")
            }
            if let lineRule = spacing.lineRule {
                attrs.append("w:lineRule=\"\(lineRule.rawValue)\"")
            }
            if !attrs.isEmpty {
                parts.append("<w:spacing \(attrs.joined(separator: " "))/>" )
            }
        }

        // 縮排
        if let indentation = indentation {
            var attrs: [String] = []
            if let left = indentation.left {
                attrs.append("w:left=\"\(left)\"")
            }
            if let right = indentation.right {
                attrs.append("w:right=\"\(right)\"")
            }
            if let firstLine = indentation.firstLine {
                attrs.append("w:firstLine=\"\(firstLine)\"")
            }
            if let hanging = indentation.hanging {
                attrs.append("w:hanging=\"\(hanging)\"")
            }
            if !attrs.isEmpty {
                parts.append("<w:ind \(attrs.joined(separator: " "))/>" )
            }
        }

        // 分頁控制
        if keepNext {
            parts.append("<w:keepNext/>")
        }
        if keepLines {
            parts.append("<w:keepLines/>")
        }
        if pageBreakBefore {
            parts.append("<w:pageBreakBefore/>")
        }

        // 段落邊框
        if let border = border {
            parts.append(border.toXML())
        }

        // 段落底色
        if let shading = shading {
            parts.append(shading.toXML())
        }

        return parts.joined()
    }
}
