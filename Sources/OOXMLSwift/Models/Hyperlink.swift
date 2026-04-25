import Foundation

// MARK: - Hyperlink

/// 超連結（外部 URL 或內部書籤連結）
///
/// v0.19.0+ (PsychQuant/che-word-mcp#56) hybrid model:
/// - `runs: [Run]` — typed editable surface; `replace_text` / `format_text`
///   walk this to apply tool-mediated edits to text inside the hyperlink
///   (the v3.12.0 silent failure mode).
/// - `text: String` — computed property `runs.map { $0.text }.joined()`
///   so existing 218 MCP tools that read `hyperlink.text` keep working.
///   Setter collapses to single Run (matches pre-fix observable behavior).
/// - `rawAttributes` / `rawChildren` — raw passthrough escape hatch for
///   unrecognized `<w:hyperlink>` attributes / direct children so they
///   survive a no-op round-trip even when not modeled.
/// - `position: Int` — source-document order, used by Phase 4
///   sort-by-position emit in `Paragraph.toXML()`.
public struct Hyperlink: Equatable {
    public var id: String              // 唯一識別碼（用於管理）
    public var relationshipId: String? // 關係 ID（外部連結使用 rId）
    public var anchor: String?         // 書籤名稱（內部連結使用）
    public var url: String?            // 外部 URL
    /// v0.19.0+ (#56): typed editable runs that replace the previous single
    /// `text: String` storage. `text` is now a computed convenience property
    /// projecting joined run text.
    public var runs: [Run] = []
    public var tooltip: String?        // 滑鼠懸停提示
    /// v0.17.0+ (#50): controls `w:history` attribute. `true` (default) = link
    /// is added to "visited" history; `false` = emit `w:history="0"`.
    public var history: Bool = true

    /// v0.19.0+ (#56): unmodeled `<w:hyperlink>` attribute → value passthrough
    /// (e.g., `w:tgtFrame`, `w:docLocation`, vendor extensions). Writer emits
    /// these alongside the typed attribute set so attribute-level losses are
    /// impossible.
    public var rawAttributes: [String: String] = [:]

    /// v0.19.0+ (#56): direct children of `<w:hyperlink>` that are not Runs
    /// (e.g., nested SDTs, future extensions). Stored as verbatim XML strings
    /// so unknown content survives a round-trip without typed modeling.
    public var rawChildren: [String] = []

    /// v0.19.0+ (#56): source-document order index for Phase 4 sort-by-position
    /// emit. Default 0 for hyperlinks created via initializers (those rely on
    /// the existing `Paragraph.toXML()` paths until Phase 4 lands).
    public var position: Int = 0

    /// v0.19.0+ (#56): displayed text computed as the joined run text. Setter
    /// collapses `runs` to a single Run carrying the assigned string (matches
    /// the pre-fix observable behavior — assigning `text = "foo"` to a
    /// multi-run hyperlink already replaced the entire visible text).
    public var text: String {
        get { runs.map { $0.text }.joined() }
        set { runs = [Run(text: newValue)] }
    }

    // 連結類型
    public var type: HyperlinkType {
        if relationshipId != nil {
            return .external
        } else if anchor != nil {
            return .internal
        }
        return .external
    }

    public init(id: String, text: String, url: String, relationshipId: String, tooltip: String? = nil, history: Bool = true) {
        self.id = id
        self.runs = [Run(text: text)]
        self.url = url
        self.relationshipId = relationshipId
        self.anchor = nil
        self.tooltip = tooltip
        self.history = history
    }

    public init(id: String, text: String, anchor: String, tooltip: String? = nil, history: Bool = true) {
        self.id = id
        self.runs = [Run(text: text)]
        self.anchor = anchor
        self.relationshipId = nil
        self.url = nil
        self.tooltip = tooltip
        self.history = history
    }

    /// v0.19.0+ (#56): full-control initializer for the typed Reader path.
    /// Used by `DocxReader` to populate the hybrid surface from source XML.
    public init(
        id: String,
        runs: [Run],
        relationshipId: String? = nil,
        anchor: String? = nil,
        url: String? = nil,
        tooltip: String? = nil,
        history: Bool = true,
        rawAttributes: [String: String] = [:],
        rawChildren: [String] = [],
        position: Int = 0
    ) {
        self.id = id
        self.runs = runs
        self.relationshipId = relationshipId
        self.anchor = anchor
        self.url = url
        self.tooltip = tooltip
        self.history = history
        self.rawAttributes = rawAttributes
        self.rawChildren = rawChildren
        self.position = position
    }

    /// 建立外部連結
    public static func external(id: String, text: String, url: String, relationshipId: String, tooltip: String? = nil) -> Hyperlink {
        return Hyperlink(id: id, text: text, url: url, relationshipId: relationshipId, tooltip: tooltip)
    }

    /// 建立內部連結（連到書籤）
    public static func `internal`(id: String, text: String, bookmarkName: String, tooltip: String? = nil) -> Hyperlink {
        return Hyperlink(id: id, text: text, anchor: bookmarkName, tooltip: tooltip)
    }

    /// Relationship 類型（用於 .rels 檔案）
    public static let relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
}

/// 超連結類型
public enum HyperlinkType: Equatable {
    case external   // 外部 URL
    case `internal` // 內部書籤連結
}

// MARK: - Hyperlink Reference (用於 document.xml.rels)

/// 超連結關係（儲存在 document.xml.rels）
public struct HyperlinkReference: Equatable {
    public var relationshipId: String  // rId
    public var url: String            // 目標 URL

    public init(relationshipId: String, url: String) {
        self.relationshipId = relationshipId
        self.url = url
    }
}

// MARK: - XML Generation

extension Hyperlink {
    /// 轉換為 OOXML XML（放在段落內）
    func toXML() -> String {
        var xml = "<w:hyperlink"

        // 外部連結使用 r:id，內部連結使用 w:anchor
        if let rId = relationshipId {
            xml += " r:id=\"\(rId)\""
        } else if let anchor = anchor {
            xml += " w:anchor=\"\(escapeXML(anchor))\""
        }

        // 提示文字
        if let tooltip = tooltip {
            xml += " w:tooltip=\"\(escapeXML(tooltip))\""
        }

        // v0.17.0+ (#50): w:history (only emit when false; true is default)
        if !history {
            xml += " w:history=\"0\""
        }

        xml += ">"

        // 連結文字（帶有藍色底線樣式）
        xml += """
        <w:r>
            <w:rPr>
                <w:rStyle w:val="Hyperlink"/>
                <w:color w:val="0563C1"/>
                <w:u w:val="single"/>
            </w:rPr>
            <w:t xml:space="preserve">\(escapeXML(text))</w:t>
        </w:r>
        """

        xml += "</w:hyperlink>"
        return xml
    }

    private func escapeXML(_ string: String) -> String {
        return string
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
            .replacingOccurrences(of: "'", with: "&apos;")
    }
}
