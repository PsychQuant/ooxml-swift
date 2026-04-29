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

    /// v0.19.3+ (#56 round 2 P0-3): unified ordered children list preserving
    /// source-document order between `<w:r>` and non-run children. Reader
    /// populates this from source XML (so e.g. `<w:r>A</w:r><w:sdt>X</w:sdt><w:r>B</w:r>`
    /// round-trips A→SDT→B, not A→B→SDT). Writer prefers `children` if
    /// non-empty; falls back to legacy `runs` then `rawChildren` ordering for
    /// API-built hyperlinks (which never populate `children`).
    public var children: [HyperlinkChild] = []

    /// v0.19.0+ (#56): source-document order index for Phase 4 sort-by-position
    /// emit. Default 0 for hyperlinks created via initializers (those rely on
    /// the existing `Paragraph.toXML()` paths until Phase 4 lands).
    public var position: Int? = nil

    /// v0.19.0+ (#56): displayed text computed as the joined run text. Setter
    /// collapses `runs` to a single Run carrying the assigned string (matches
    /// the pre-fix observable behavior — assigning `text = "foo"` to a
    /// multi-run hyperlink already replaced the entire visible text).
    public var text: String {
        get { runs.map { $0.text }.joined() }
        @available(*, deprecated, message: "Mutates runs destructively (loses formatting / rawElements). Use .runs directly to preserve formatting; assign a single Run to replace, append/insert Runs to extend.")
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
        self.runs = [Hyperlink.makeStyledRun(text: text)]
        self.url = url
        self.relationshipId = relationshipId
        self.anchor = nil
        self.tooltip = tooltip
        self.history = history
    }

    public init(id: String, text: String, anchor: String, tooltip: String? = nil, history: Bool = true) {
        self.id = id
        self.runs = [Hyperlink.makeStyledRun(text: text)]
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
        children: [HyperlinkChild] = [],
        position: Int? = nil
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
        self.children = children
        self.position = position
    }

    /// v0.19.3+ (#56 round 2 P0-1): build a Hyperlink-styled Run for the API
    /// path. Centralizes the visual style (Hyperlink character style + 0563C1
    /// blue + single underline) so all API-constructed hyperlinks render
    /// consistently in Word, matching the v0.19.1 hardcoded template.
    fileprivate static func makeStyledRun(text: String) -> Run {
        return Run(
            text: text,
            properties: RunProperties(
                underline: .single,
                color: "0563C1",
                rStyle: "Hyperlink"
            )
        )
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

/// v0.19.3+ (#56 round 2 P0-3): tagged child entry preserving source-document
/// order between `<w:r>` and non-run children inside `<w:hyperlink>`.
public enum HyperlinkChild: Equatable {
    case run(Run)
    case rawXML(String)
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
    /// Reserved attribute names that have a typed surface on `Hyperlink` —
    /// emitted from the typed fields, never from `rawAttributes`. Anything
    /// else in `rawAttributes` is appended verbatim (alphabetically sorted
    /// for determinism).
    private static let typedAttributeNames: Set<String> = [
        "r:id", "w:anchor", "w:tooltip", "w:history",
    ]

    /// 轉換為 OOXML XML（放在段落內）
    ///
    /// v0.19.2+ (#56 follow-up F1): rewrite to honour the v0.19.0 hybrid model:
    /// - Iterate `runs` (preserving each `RunProperties` and any per-run
    ///   `rawXML`/`rawElements` via `Run.toXML()`) instead of collapsing to a
    ///   single hardcoded `<w:r>` with bake-in `Hyperlink` style.
    /// - Emit `rawAttributes` (sorted) so vendor / unmodeled attributes
    ///   (`w:tgtFrame`, `w:docLocation`, etc.) round-trip byte-equivalent.
    /// - Append `rawChildren` verbatim after runs so non-Run direct children
    ///   (nested SDT, future extensions) survive.
    /// - Fallback path for `runs.isEmpty` — happens when `Hyperlink` is built
    ///   via the API initializers using `text:` (which now populates a single
    ///   Run) but defensively also when caller blanks the runs collection.
    ///   Emits the legacy hardcoded styled run carrying empty text so the
    ///   output stays valid OOXML.
    func toXML() -> String {
        var xml = "<w:hyperlink"

        // Typed attributes first (deterministic order, matches pre-fix output).
        if let rId = relationshipId {
            xml += " r:id=\"\(escapeXML(rId))\""
        } else if let anchor = anchor {
            xml += " w:anchor=\"\(escapeXML(anchor))\""
        }
        if let tooltip = tooltip {
            xml += " w:tooltip=\"\(escapeXML(tooltip))\""
        }
        // v0.17.0+ (#50): w:history (only emit when false; true is default)
        if !history {
            xml += " w:history=\"0\""
        }

        // v0.19.2+ (#56 F1): unmodeled passthrough attributes from source.
        // Skip any whose name collides with a typed attribute we already
        // emitted (prevents duplicate-attribute corruption).
        for (name, value) in rawAttributes.sorted(by: { $0.key < $1.key }) {
            guard !Self.typedAttributeNames.contains(name) else { continue }
            xml += " \(name)=\"\(escapeXML(value))\""
        }

        xml += ">"

        // v0.19.3+ (#56 round 2 P0-3): if Reader populated `children`, prefer
        // walking it in source-document order so `<w:r>A</w:r><w:sdt>X</w:sdt><w:r>B</w:r>`
        // round-trips A→SDT→B (not A→B→SDT).
        //
        // v0.19.4+ (#56 R3-NEW-1): mutation detection — when `runs` no longer
        // matches the run sequence derived from `children`, an API mutation
        // (`Hyperlink.text` setter / `replaceText` / `updateHyperlink` /
        // `format_text` / property-only edits like `runs[0].properties.bold`)
        // has touched `runs` and the saved XML must reflect that. Fall through
        // to the `runs` + `rawChildren` path so mutations are visible on save.
        //
        // v0.19.5+ (#56 R5 P1 #1): upgrade detection from joined-text comparison
        // to deep `[Run]` equality (synthesized `Equatable` covers `text` +
        // `properties`). Pre-fix the text-only check missed property-only
        // mutations (e.g., `runs[0].properties.bold = true` with same text)
        // and equal-length text swaps that re-derive the same joined string;
        // both silently dropped on save. Trade-off restated: when a mutation
        // touches a hyperlink whose `children` carries non-run elements
        // (rare), the non-run order is lost — `design.md` deems this
        // acceptable since silent edit-failure has the wider blast radius.
        let childrenRuns = children.compactMap { child -> Run? in
            if case .run(let run) = child { return run } else { return nil }
        }
        let childrenAuthoritative = !children.isEmpty && childrenRuns == runs

        if childrenAuthoritative {
            for child in children {
                switch child {
                case .run(let run):
                    xml += run.toXML()
                case .rawXML(let raw):
                    xml += raw
                }
            }
        } else if runs.isEmpty && rawChildren.isEmpty {
            // Empty fallback: emit the hardcoded Hyperlink-styled run so the
            // wrapper stays valid OOXML even when the caller cleared content.
            xml += """
            <w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t xml:space="preserve"></w:t></w:r>
            """
        } else {
            for run in runs {
                xml += run.toXML()
            }
            for raw in rawChildren {
                xml += raw
            }
        }

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
