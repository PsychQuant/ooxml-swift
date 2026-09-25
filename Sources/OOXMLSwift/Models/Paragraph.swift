import Foundation

/// 段落 (Paragraph) - Word 文件的基本結構單元
public struct Paragraph: Equatable {
    /// v0.31.0+ (Spectra `paragraph-tree-projection-impl`, `word-aligned-state-sync`
    /// Phase 1 task 2.1): when non-nil, this Paragraph is a tree-backed view
    /// over the wrapped `<w:p>` element. Getters walk `xmlNode.children` at
    /// access time; setters mutate the tree directly (Phase 1 stub) and call
    /// `xmlNode.markDirty()` so the writer re-serializes from typed fields.
    /// When nil (legacy detached mode), getters/setters operate on the stored
    /// properties below. `XmlNode` is a class, so two value-copies of the
    /// same tree-backed Paragraph share the same underlying tree state.
    public var xmlNode: XmlNode?

    /// Legacy stored backing for `runs` used in detached mode.
    /// In tree-backed mode the public `runs` accessor walks `xmlNode.children`
    /// instead and this buffer is ignored. Renamed from the previous public
    /// `runs` stored property in v0.31.0 (`paragraph-tree-projection-impl`).
    internal var _legacyRuns: [Run] = []
    public var properties: ParagraphProperties
    public var hasPageBreak: Bool = false      // 是否為分頁符段落
    public var bookmarks: [Bookmark] = []      // 段落內的書籤
    public var hyperlinks: [Hyperlink] = []    // 段落內的超連結
    @available(*, deprecated, message: "Use commentRangeMarkers (source of truth since Phase 4) or the computed commentRangeIds. Stored commentIds is no longer populated by Reader since v0.21.4 and will be removed in v0.22.")
    public var commentIds: [Int] = []          // 段落關聯的註解 ID（v0.21.4+ deprecated — Reader 不再填值）
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

    /// v0.19.0+ (PsychQuant/che-word-mcp#56): position-indexed range markers for
    /// `<w:bookmarkStart>` / `<w:bookmarkEnd>` parsed from source. Parallel to
    /// `bookmarks` (the typed name+id model) — the marker carries source-order
    /// position so Phase 4's sort-by-position emit can re-emit at original
    /// relative offsets. Empty for paragraphs created via initializer / API
    /// (those rely on `bookmarks` + the existing wrap-around `toXML` emit path).
    public var bookmarkMarkers: [BookmarkRangeMarker] = []

    /// v0.19.0+ (#56) Phase 3: typed `<w:fldSimple>` wrappers parsed from source.
    /// Holds field expression (`instr`) + rendered runs so `replace_text` /
    /// `format_text` can apply edits inside SEQ Table captions, REF
    /// cross-references, etc.
    public var fieldSimples: [FieldSimple] = []

    /// v0.19.0+ (#56) Phase 3: typed `<mc:AlternateContent>` wrappers parsed
    /// from source. Holds verbatim `rawXML` for byte-equivalent re-emit plus
    /// typed `fallbackRuns` for tool-mediated read access to `<mc:Fallback>`
    /// content (math transliterations, drawing flat text, etc.).
    public var alternateContents: [AlternateContent] = []

    /// v0.19.0+ (#56) Phase 4: position-indexed `<w:commentRangeStart>` /
    /// `<w:commentRangeEnd>` markers. Parallel to `commentIds` (legacy
    /// position-less collection) — `commentIds` stays for backward compat
    /// while `commentRangeMarkers` carries the source position needed for
    /// sort-by-position emit.
    public var commentRangeMarkers: [CommentRangeMarker] = []

    /// v0.21.4+ (PsychQuant/ooxml-swift#6, F9): canonical comment-id list
    /// derived from `commentRangeMarkers`. Replaces the deprecated stored
    /// `commentIds` field. Returns the unique ids in order of first appearance,
    /// reflecting both Reader-loaded markers and any markers appended post-load.
    public var commentRangeIds: [Int] {
        var seen = Set<Int>()
        var result: [Int] = []
        for marker in commentRangeMarkers {
            if seen.insert(marker.id).inserted {
                result.append(marker.id)
            }
        }
        return result
    }

    /// v0.19.0+ (#56) Phase 4: position-indexed `<w:permStart>` / `<w:permEnd>`
    /// markers (editor permission gates).
    public var permissionRangeMarkers: [PermissionRangeMarker] = []

    /// v0.19.0+ (#56) Phase 4: position-indexed `<w:proofErr>` markers
    /// (Word Proofing UI annotations).
    public var proofErrorMarkers: [ProofErrorMarker] = []

    /// v0.19.0+ (#56) Phase 4: position-indexed `<w:smartTag>` raw-carriers.
    public var smartTags: [SmartTagBlock] = []

    /// v0.19.0+ (#56) Phase 4: position-indexed `<w:customXml>` raw-carriers.
    public var customXmlBlocks: [CustomXmlBlock] = []

    /// v0.19.0+ (#56) Phase 4: position-indexed `<w:dir>` / `<w:bdo>` raw-carriers.
    public var bidiOverrides: [BidiOverrideBlock] = []

    /// v0.19.0+ (#56) Phase 4: fallback for any `<w:p>` child whose local
    /// name does not match any typed parser or registered raw-carrier.
    /// Surfaces ECMA-376 spec gaps so the round-trip test suite can XCTFail
    /// with the unknown element name (per design decision "ecma-376 `<w:p>`
    /// schema as the completeness checklist").
    public var unrecognizedChildren: [UnrecognizedChild] = []

    /// v0.20.3+ (#66 sub-stack E): `w14:paraId` attribute on `<w:p>` opening
    /// tag — Word's revision-tracking GUID anchoring paragraph identity for
    /// collaborative editing and comment threading. 8-character hex token
    /// (NOT RFC 4122 UUID); preserved as opaque String to avoid format
    /// rejection. Pre-fix accounted for ~95% of w14:* token loss in the
    /// NTPU thesis fixture.
    public var w14ParaId: String?

    /// v0.20.3+ (#66 sub-stack E): `w14:textId` attribute on `<w:p>` opening
    /// tag — Word's text-content revision GUID. Independent of `w14ParaId`
    /// (Word can emit either alone). Same opaque-String preservation rules.
    public var w14TextId: String?

    public init(runs: [Run] = [], properties: ParagraphProperties = ParagraphProperties()) {
        self._legacyRuns = runs
        self.properties = properties
    }

    /// 便利初始化器：直接用文字建立段落
    public init(text: String, properties: ParagraphProperties = ParagraphProperties()) {
        self._legacyRuns = [Run(text: text)]
        self.properties = properties
    }

    /// v0.31.0+ Tree-backed initializer. Wraps an existing `<w:p>` xmlNode
    /// so getters walk its children and setters mutate the tree directly.
    ///
    /// The legacy stored fields (`_legacyRuns`, `properties`, `bookmarks`, …)
    /// are initialized to their empty defaults; in tree-backed mode they are
    /// shadowed by computed accessors that read from `xmlNode.children`.
    /// Callers MUST NOT rely on those stored fields when `xmlNode != nil`.
    ///
    /// Semantic validation (asserting the node is a `<w:p>` element) is left
    /// to callers; this initializer accepts any element xmlNode so unit tests
    /// can synthesize fixtures without paying the schema-check cost.
    public init(xmlNode: XmlNode) {
        self.xmlNode = xmlNode
        self.properties = ParagraphProperties()
    }

    /// v0.31.0+ Mode-aware view of `<w:r>` runs.
    ///
    /// - Tree-backed: walks `xmlNode.children` at every access and returns one
    ///   `Run` per `<w:r>` child in document order, with `text` populated by
    ///   concatenating the `<w:t>` descendants. **Phase 1 stub:** the returned
    ///   `Run` values are NOT themselves tree-backed — task 2.2 of
    ///   `word-aligned-state-sync` adds `Run(xmlNode:)` and rewires this getter
    ///   so further accessors propagate freshly. Existing tests only assert on
    ///   `count`, so this stub is sufficient for Phase 1.
    /// - Detached: returns the legacy stored buffer.
    ///
    /// Setter writes to `_legacyRuns` in both modes. In tree-backed mode the
    /// write is a "ghost" buffer (not visible through the getter); proper
    /// tree-mutating routing arrives with the op-log path in Phase 2.
    public var runs: [Run] {
        get {
            guard let node = xmlNode else { return _legacyRuns }
            return node.children.compactMap { child -> Run? in
                guard child.kind == .element, child.localName == "r" else { return nil }
                let text = Self.concatenatedText(in: child)
                return Run(text: text)
            }
        }
        set { _legacyRuns = newValue }
    }

    /// v0.31.0+ Mode-aware view of the paragraph's plain-text content.
    ///
    /// - Tree-backed getter: concatenates `textContent` of every `<w:t>`
    ///   descendant inside each `<w:r>` child of the wrapped `<w:p>`, in
    ///   document order. No caching — re-walks on every access.
    /// - Detached getter: joins `_legacyRuns.map { $0.text }`.
    ///
    /// The setter is finalized in task 1.4 — see the dedicated implementation
    /// below for the Phase 1 stub semantics (tree-mutating in tree-backed
    /// mode, replace-runs in detached mode).
    public var text: String {
        get {
            if let node = xmlNode {
                return node.children
                    .filter { $0.kind == .element && $0.localName == "r" }
                    .map { Self.concatenatedText(in: $0) }
                    .joined()
            }
            return _legacyRuns.map { $0.text }.joined()
        }
        set {
            if let node = xmlNode {
                let textNode = XmlNode.text(newValue)
                let wt = XmlNode.element(prefix: "w", localName: "t", children: [textNode])
                let wr = XmlNode.element(prefix: "w", localName: "r", children: [wt])
                node.children = [wr]
                node.markDirty()
            } else {
                _legacyRuns = [Run(text: newValue)]
            }
        }
    }

    /// Helper: concatenate every `<w:t>` descendant text inside a `<w:r>`
    /// xmlNode (or any element whose subtree contains `<w:t>` text leaves).
    /// Walks the immediate children only — `<w:r>` children are typically
    /// `<w:t>`, `<w:rPr>`, `<w:tab>`, `<w:br>`, etc.
    private static func concatenatedText(in runNode: XmlNode) -> String {
        var out = ""
        for child in runNode.children where child.kind == .element && child.localName == "t" {
            for grand in child.children where grand.kind == .text {
                out += grand.textContent
            }
        }
        return out
    }

    /// v0.31.0+ Stable identifier for this paragraph.
    ///
    /// - Tree-backed: returns `xmlNode.stableID` if any OOXML stable-ID attribute
    ///   is present (e.g. `"w14:paraId=0ABC1234"`); otherwise falls back to
    ///   `"lib:<UUID>"` when the reader assigned a library-generated UUID;
    ///   otherwise `nil`.
    /// - Detached (legacy): always returns `nil`.
    ///
    /// Format note: `XmlNode.stableID` already prefixes with the attribute name
    /// (e.g. `"w14:paraId=..."`), so we pass it through verbatim. Library
    /// UUIDs get the `"lib:"` prefix here so consumers can distinguish them
    /// from native OOXML stable IDs at the Paragraph layer.
    public var id: String? {
        guard let node = xmlNode else { return nil }
        if let stable = node.stableID { return stable }
        if let lib = node.libraryUUID { return "lib:\(lib.uuidString)" }
        return nil
    }

    /// 取得段落純文字。
    ///
    /// As of v0.22.x (PsychQuant/ooxml-swift#43), this method delegates to
    /// `flattenedDisplayText()` so the two text-extraction paths return
    /// identical strings. Pre-#43 this was a legacy implementation that
    /// only joined `runs.map { $0.text }` + `hyperlink.text`, missing:
    ///
    /// - Direct-child OMML in `unrecognizedChildren` (Pandoc display math
    ///   `<m:oMath>` / `<m:oMathPara>` as direct child of `<w:p>`,
    ///   per #99 / #100 / #101 / #102)
    /// - Per-run OMath visibleText embedded in `Run.rawXML` (per #85 / #92)
    /// - Field codes (`<w:fldSimple>`) and alternate content fallback runs
    /// - Content controls (`<w:sdt>`)
    ///
    /// The pre-#43 divergence caused silent off-by-N gaps in
    /// `che-word-mcp__search_text` because that MCP tool delegated to
    /// `getText()` while anchor-matching paths used `flattenedDisplayText()`.
    /// Unifying behind a single canonical implementation prevents future
    /// callers from re-discovering the same trap.
    ///
    /// Callers who specifically want runs-only text (no OMath, no
    /// hyperlinks) should call `runs.map { $0.text }.joined()` directly
    /// rather than relying on the legacy pre-#43 behavior.
    public func getText() -> String {
        return flattenedDisplayText()
    }
}

// MARK: - Equatable (mode-aware identity vs content)

extension Paragraph {
    /// v0.31.0+ Custom Equatable replacing auto-synthesized conformance per
    /// `paragraph-tree-projection-impl` Decision 5.
    ///
    /// Behavior depends on the storage mode of both sides:
    ///
    /// 1. **Both tree-backed**: identity equality on the wrapped `xmlNode`
    ///    reference (`===`). Op-log addresses paragraphs by id (== identity);
    ///    content equality on different elements would silently merge log
    ///    entries that target different paragraphs.
    /// 2. **Both detached**: content equality across the legacy stored fields,
    ///    preserving pre-v0.31 auto-synthesized behavior the che-word-mcp
    ///    test suite depends on.
    /// 3. **Mixed (one tree-backed, one detached)**: always `false`. The two
    ///    storage modes are not interchangeable; comparing across them is
    ///    almost certainly a caller mistake worth surfacing.
    public static func == (lhs: Paragraph, rhs: Paragraph) -> Bool {
        switch (lhs.xmlNode, rhs.xmlNode) {
        case let (a?, b?):
            return a === b
        case (nil, nil):
            return contentEquals(lhs, rhs)
        default:
            return false
        }
    }

    /// Detached-mode content equality across all legacy stored fields.
    /// Mirrors what auto-synthesized `Equatable` would have compared.
    /// `commentIds` is included for behavioral parity even though it is
    /// deprecated (the existing warning is consistent with Document.swift's
    /// other call sites; the field is slated for removal in v0.22).
    private static func contentEquals(_ lhs: Paragraph, _ rhs: Paragraph) -> Bool {
        return lhs._legacyRuns == rhs._legacyRuns
            && lhs.properties == rhs.properties
            && lhs.hasPageBreak == rhs.hasPageBreak
            && lhs.bookmarks == rhs.bookmarks
            && lhs.hyperlinks == rhs.hyperlinks
            && lhs.commentIds == rhs.commentIds
            && lhs.footnoteIds == rhs.footnoteIds
            && lhs.endnoteIds == rhs.endnoteIds
            && lhs.revisions == rhs.revisions
            && lhs.semantic == rhs.semantic
            && lhs.paragraphFormatChangeRevisionId == rhs.paragraphFormatChangeRevisionId
            && lhs.previousProperties == rhs.previousProperties
            && lhs.contentControls == rhs.contentControls
            && lhs.bookmarkMarkers == rhs.bookmarkMarkers
            && lhs.fieldSimples == rhs.fieldSimples
            && lhs.alternateContents == rhs.alternateContents
            && lhs.commentRangeMarkers == rhs.commentRangeMarkers
            && lhs.permissionRangeMarkers == rhs.permissionRangeMarkers
            && lhs.proofErrorMarkers == rhs.proofErrorMarkers
            && lhs.smartTags == rhs.smartTags
            && lhs.customXmlBlocks == rhs.customXmlBlocks
            && lhs.bidiOverrides == rhs.bidiOverrides
            && lhs.unrecognizedChildren == rhs.unrecognizedChildren
            && lhs.w14ParaId == rhs.w14ParaId
            && lhs.w14TextId == rhs.w14TextId
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
    /// 段落邊框。v3.13.0+ (#176)：任何指派（setter、直接指派、`border?.top?.color
    /// = …` 這類經由 optional chaining 的修改）都會清掉 `sourceBorder`——之後一律
    /// typed 輸出，即使新值剛好等於讀取時的有損投影（revooxmlc MEDIUM-1）。
    public var border: ParagraphBorder? {        // 段落邊框
        didSet { sourceBorder = nil }
    }
    /// 段落底色。指派語意同 `border`（清掉 `sourceShading`）。
    public var shading: ParagraphShading? {      // 段落底色
        didSet { sourceShading = nil }
    }

    /// v3.13.0+ (#176)：讀取時的來源 `<w:pBdr>` / `<w:shd>` 原文。typed 模型表達
    /// 不了來源的全部內容（`ParagraphBorderStyle` 沒有 themeColor/themeTint/
    /// themeShade/shadow/frame，`ParagraphBorder` 沒有 `bar` 邊，
    /// `ParagraphBorderType` / `ShadingPattern` 只涵蓋少數 `ST_Border` / `ST_Shd`
    /// 值，`CellShading` 沒有 theme* 屬性），所以只要 `border` / `shading` 在讀取
    /// 之後**沒有被指派過**，`toXML()` 就原樣輸出這份原文；被指派過（`didSet`
    /// 清掉原文）就走 typed 輸出；設成 nil 就不輸出。reader 必須先指派 typed
    /// 投影、再設原文（順序反過來原文會被自己的指派清掉）。
    internal var sourceBorder: SourcePreservedPPrChild<ParagraphBorder>?
    internal var sourceShading: SourcePreservedPPrChild<ParagraphShading>?

    /// v0.20.2+ (#65 sub-stack D): paragraph-mark RunProperties — the
    /// `<w:rPr>` direct child of `<w:pPr>` that controls pilcrow ¶ glyph
    /// formatting (font, size, color, language, kerning) per ECMA-376
    /// §17.3.1.27 CT_PPrBase. Schema is identical to run-level CT_RPr,
    /// so parser reuses `parseRunProperties` verbatim and emit reuses
    /// `RunProperties.toXML()` — typed extraction (rFonts 4-axis, noProof,
    /// kern, lang 3-axis) and raw passthrough (rawChildren for w14:* effects)
    /// come for free from sub-stack C. Pre-fix this nested rPr was silently
    /// dropped at parse time, accounting for ~50% of the residual `<w:lang>`
    /// loss in NTPU thesis fixture round-trip (16.66% → < 5% target).
    public var markRunProperties: RunProperties?

    /// v3.12.0+ (#168): `<w:pPr>` child elements the typed model has no field
    /// for — `<w:kinsoku>`, `<w:snapToGrid>`, `<w:widowControl>`,
    /// `<w:wordWrap>`, and the rest of CT_PPrBase's long tail. Captured
    /// verbatim by `DocxReader.parseParagraphProperties` (same "if not typed,
    /// preserve as raw" principle as `Run.rawElements`, v0.14.0+/#52) and
    /// re-emitted by `toXML()`.
    ///
    /// Why this exists: a typed edit through an API like `updateCell` /
    /// `updateCellParagraph` forces the whole `word/document.xml` to be
    /// regenerated from the typed model (`markTypedDirty`, see #168) —
    /// including paragraphs the edit never touched. Before this field, any
    /// pPr child outside the typed vocabulary silently vanished from EVERY
    /// paragraph in the document, not just the edited one. This field makes
    /// the typed model a (mostly) lossless carrier for those children so the
    /// regeneration no longer drops them.
    ///
    /// v3.13.0+ (#175)：`toXML()` 依元素名稱把每個 raw 子元素排進 ECMA-376
    /// `CT_PPr` 的 schema 位置（`canonicalPPrPosition`），與 typed 欄位交錯，
    /// 不再一律擠在 typed 欄位之後、`<w:rPr>` 之前。只有 qualified name 是
    /// `w:<schema 元素名>` 的才算 schema 已知；擴充命名空間（`w14:`、`w15:`…）
    /// 與 `mc:AlternateContent` 等表外元素走「錨定」規則，見
    /// `canonicalPPrPosition` 的說明。
    ///
    /// `sectPr`, `pPrChange`, `pBdr`, and `shd` are deliberately excluded from
    /// raw capture (see `DocxReader.recognizedPPrChildNames`)：`pBdr`/`shd`
    /// 自 v3.13.0（#176）起讀進 typed 欄位 `border`/`shading`（未被改動時原樣
    /// 輸出來源原文，見 `sourceBorder`），raw 再捕捉一次會輸出兩份；
    /// `sectPr`/`pPrChange` 由 `Paragraph` 的 pPr 包裝層在 `toXML()` 的輸出
    /// 之後另外輸出。
    public var rawChildren: [RawElement] = []

    /// v3.13.0+ (#175)：讀取時為每個表外（非 `w:` schema）raw 子元素記下的來源
    /// 錨點，依來源順序。輸出時以**元素身分**對應——XML 相等、依序取第一筆
    /// 還沒用過的紀錄——而不是按 `rawChildren` 的索引，所以呼叫端增刪
    /// `rawChildren` 不會讓別的元素對到錯的錨點（revooxmlc LOW-1）。沒有表外
    /// 子元素的段落不記，讓它與呼叫端自組的段落仍然 `==`。
    internal var rawChildSourceAnchors: [RawChildSourceAnchor] = []

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
        // #176：取用 `other` 的值是一次指派（`didSet` 清掉自己的來源原文），
        // 來源原文跟著值走：`other` 是讀進來、沒被指派過的，就帶著它的原文；
        // `other` 是呼叫端自組的（沒有原文），就是 typed 輸出——即使值剛好等於
        // 自己的投影（revooxmlc MEDIUM-1）。
        if let border = other.border {
            self.border = border
            self.sourceBorder = other.sourceBorder
        }
        if let shading = other.shading {
            self.shading = shading
            self.sourceShading = other.sourceShading
        }
        if let markRunProperties = other.markRunProperties { self.markRunProperties = markRunProperties }
        if !other.rawChildren.isEmpty {
            self.rawChildren = other.rawChildren
            self.rawChildSourceAnchors = other.rawChildSourceAnchors
        }
    }
}

/// v3.13.0+ (#175)：一個 `rawChildren` 元素在來源 `<w:pPr>` 中的錨點——
/// 它之前最近的一個 schema 已知子元素（`w:<CT_PPr 元素名>`）的 local name；
/// `nil` 代表來源中它之前沒有任何 schema 已知子元素。
internal struct RawChildSourceAnchor: Equatable {
    let xml: String
    let anchor: String?
}

/// v3.13.0+ (#176)：一個由 typed 欄位承載、但 typed 模型表達不了全部內容的
/// pPr 子元素——來源原文（`.nodeCompactEmptyElement` 序列化）與讀取當下的
/// typed 投影。
internal struct SourcePreservedPPrChild<Projection: Equatable>: Equatable {
    let xml: String
    let projection: Projection
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
    /// v0.19.0+ (PsychQuant/che-word-mcp#56) Phase 4: detect whether this
    /// paragraph carries source-loaded position-indexed children (markers,
    /// new typed wrappers, raw-carriers).
    ///
    /// v0.19.3+ (#56 round 2 P0-8): also treat any run or hyperlink with a
    /// non-zero `position` as a source-loaded signal. Pre-fix, a paragraph
    /// like `<w:r>A</w:r><w:hyperlink>L</w:hyperlink><w:r>B</w:r>` (no
    /// markers, no carriers, just runs + hyperlink at positions 0/1/2) went
    /// to the legacy path, which emits all runs first then all hyperlinks —
    /// silently re-ordering the visible text to "A B L". Including these
    /// signals routes the paragraph to `toXMLSortedByPosition`, preserving
    /// source order. API-built paragraphs leave positions at 0 and stay on
    /// the legacy path.
    internal var hasSourcePositionedChildren: Bool {
        return !bookmarkMarkers.isEmpty
            || !fieldSimples.isEmpty
            || !alternateContents.isEmpty
            || !commentRangeMarkers.isEmpty
            || !permissionRangeMarkers.isEmpty
            || !proofErrorMarkers.isEmpty
            || !smartTags.isEmpty
            || !customXmlBlocks.isEmpty
            || !bidiOverrides.isEmpty
            || !unrecognizedChildren.isEmpty
            || runs.contains(where: { ($0.position ?? 0) > 0 })
            || hyperlinks.contains(where: { ($0.position ?? 0) > 0 })
            // v0.19.4+ (#56 R3-NEW-2): paragraph-level <w:sdt> with source
            // position participates in sort-by-position emit.
            // PsychQuant/ooxml-swift#5 (F6): position is now `Int? = nil` —
            // nil means "append-mode" (API-built, no explicit source order),
            // so it does NOT trigger sort path; only explicit positive
            // positions do. Mirror legacy `> 0` semantic via `?? 0`.
            || contentControls.contains(where: { ($0.position ?? 0) > 0 })
    }

    /// 轉換為 OOXML XML 字串
    public func toXML() -> String {
        if hasSourcePositionedChildren {
            return toXMLSortedByPosition()
        }
        return toXMLLegacy()
    }

    /// v0.21.4+ (PsychQuant/ooxml-swift#6, F8): throwing variant of `toXML()`
    /// used by the save path (`DocxWriter.xmlForBodyChild`). Inspects
    /// `alternateContents` for any `AlternateContent.fallbackRunsModified`
    /// flag set to `true` and throws
    /// `RoundtripError.unserializedFallbackEdit(position:)` per the dirty-tracking
    /// contract documented on `AlternateContent.fallbackRuns`. The non-throwing
    /// `toXML()` retains its current behaviour for in-memory inspection /
    /// debug callers (deviation note recorded in
    /// `openspec/changes/roundtrip-loud-fail/tasks.md` Group 3).
    public func toXMLThrowing() throws -> String {
        for ac in alternateContents where ac.fallbackRunsModified {
            throw RoundtripError.unserializedFallbackEdit(position: ac.position ?? 0)
        }
        return toXML()
    }

    /// v0.20.3+ (#66 sub-stack E): build the `<w:p>` opening tag with optional
    /// w14:paraId / w14:textId attributes. Shared by both emit paths.
    /// Independent attributes — Word can emit either alone. Values are
    /// escaped via `escapeXMLAttribute` (mirrors every other attribute emit
    /// in this file, e.g., ParagraphProperties.toXML() pStyle handling) to
    /// prevent XML injection if a caller sets an unsanitized GUID value.
    private func openingPTag() -> String {
        var attrs: [String] = []
        if let paraId = w14ParaId {
            attrs.append("w14:paraId=\"\(escapeXMLAttribute(paraId))\"")
        }
        if let textId = w14TextId {
            attrs.append("w14:textId=\"\(escapeXMLAttribute(textId))\"")
        }
        if attrs.isEmpty {
            return "<w:p>"
        }
        return "<w:p \(attrs.joined(separator: " "))>"
    }

    /// Legacy emit path used by API-built paragraphs (no source-loaded markers).
    /// Mirrors v3.12.0 behavior: bookmarks / commentRanges / runs / SDTs /
    /// hyperlinks / footnoteRefs / endnoteRefs / bookmark-end in fixed order.
    fileprivate func toXMLLegacy() -> String {
        var xml = openingPTag()

        // Paragraph Properties
        let propsXML = properties.toXML()
        let formatChangeRevision: Revision? = {
            guard let id = paragraphFormatChangeRevisionId else { return nil }
            return revisions.first { $0.id == id }
        }()
        // `needsPPr` gate implements "drop empty <w:pPr> on emit" — an
        // all-default paragraph emits nothing for pPr (OOXML treats absence ≡
        // <w:pPr/>). This invariant is asserted by Issue4PPrRegressionGuardTests
        // .testEmptyPPrSelfClosingProducesNoUnrecognizedAndDropsEmptyBlock; the
        // sorted-by-position emit path at L502 mirrors this gate. If you extend
        // either condition, update both sites + the test docstring to match.
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

    // MARK: - Sort-by-position emit (Phase 4)

    /// v0.19.0+ (PsychQuant/che-word-mcp#56) Phase 4: emit `<w:p>` content in
    /// source-document order using each child's `position` field. Used when
    /// the paragraph carries any of the new position-indexed collections
    /// (bookmarkMarkers, fieldSimples, alternateContents, the 6 raw-carriers,
    /// unrecognizedChildren). Implements design decision "Position-index
    /// ordering, not enum refactor": each parallel array contributes
    /// `(position, xml-string)` tuples that are sorted then emitted.
    ///
    /// Coexistence with legacy paragraph fields:
    /// - pPr emits first (pre-content per ECMA-376).
    /// - v0.19.3+ (#56 round 2 P0-4 + P0-5): legacy position-less collections
    ///   (`hasPageBreak` / `bookmarks` / `commentIds` / `footnoteIds` /
    ///   `endnoteIds` / `contentControls`) now emit at the legacy positions
    ///   relative to the sort window — `hasPageBreak` + `bookmarks-start` +
    ///   `commentRangeStart` BEFORE the sort children, then `contentControls`
    ///   + `commentRangeEnd` + `commentReference` + `footnoteReference` +
    ///   `endnoteReference` + `bookmarks-end` AFTER. Pre-fix the doc-comment
    ///   claimed they would emit AFTER but the code dropped them entirely,
    ///   so any paragraph with source markers + `insert_comment` /
    ///   `insert_footnote` / `insert_content_control` silently lost them on
    ///   save. Sort path is now functionally a superset of the legacy path:
    ///   anything legacy could emit, sort can too.
    /// - The position-indexed `bookmarkMarkers` / `commentRangeMarkers` /
    ///   `permissionRangeMarkers` continue to drive their own
    ///   `<w:bookmarkStart>` / `<w:bookmarkEnd>` / etc emit through the
    ///   sort window, so source-loaded paragraphs (which use these instead
    ///   of the legacy collections) keep byte-equivalent round-trip.
    fileprivate func toXMLSortedByPosition() -> String {
        var xml = openingPTag()

        // Paragraph Properties (mirrors legacy path).
        let propsXML = properties.toXML()
        let formatChangeRevision: Revision? = {
            guard let id = paragraphFormatChangeRevisionId else { return nil }
            return revisions.first { $0.id == id }
        }()
        // `needsPPr` gate — sorted-by-position counterpart of L371 (legacy emit
        // path). Same drop-empty-pPr invariant locked in by
        // Issue4PPrRegressionGuardTests
        // .testEmptyPPrSelfClosingProducesNoUnrecognizedAndDropsEmptyBlock. Keep
        // both sites in sync; if you add a new condition (e.g.,
        // `|| isCustomDirectFormatting`), update both + the test docstring.
        let needsPPr = !propsXML.isEmpty
            || properties.sectionBreak != nil
            || formatChangeRevision != nil
        if needsPPr {
            xml += "<w:pPr>\(propsXML)"
            if let sectionBreak = properties.sectionBreak {
                xml += "<w:sectPr><w:type w:val=\"\(sectionBreak.rawValue)\"/></w:sectPr>"
            }
            if let revision = formatChangeRevision, let prev = previousProperties {
                xml += revision.toOpeningXML()
                xml += "<w:pPr>\(prev.toXML())</w:pPr>"
                xml += revision.toClosingXML()
            }
            xml += "</w:pPr>"
        }

        // v0.19.3+ (#56 round 2 P0-5): legacy pre-content collections. Mirrors
        // legacy `toXMLLegacy` ordering. Skipped per-collection when the
        // positioned variant is populated to avoid double-emit — Reader keeps
        // `bookmarks` populated for backward-compat reads while also populating
        // `bookmarkMarkers`, so emitting both would duplicate every source
        // bookmark on round-trip.
        if hasPageBreak {
            xml += "<w:r><w:br w:type=\"page\"/></w:r>"
        }
        if bookmarkMarkers.isEmpty {
            for bookmark in bookmarks {
                xml += bookmark.toBookmarkStartXML()
            }
        }
        // v0.19.4+ (#56 R3-NEW-3): per-id gate, not blanket isEmpty.
        // Pre-fix `if commentRangeMarkers.isEmpty` skipped the entire legacy
        // emit when source had any commentRangeMarker, dropping new commentIds
        // added via insertComment on source-loaded paragraphs. Per-id gate
        // emits start markers only for commentIds NOT already covered by a
        // source-loaded commentRangeMarker, preserving R2 P0-5 no-double-emit
        // semantics while restoring R3-NEW-3 insertComment marker output.
        let commentIdsCoveredByMarkers = Set(commentRangeMarkers.map { $0.id })
        for commentId in commentIds where !commentIdsCoveredByMarkers.contains(commentId) {
            xml += "<w:commentRangeStart w:id=\"\(commentId)\"/>"
        }

        // Build (position, payload) pairs from every position-indexed collection.
        // v0.19.2+ (#56 F3): runs are kept as `.run(Run)` rather than pre-emitted
        // strings so the post-sort pass can group consecutive same-revisionId
        // runs and wrap them in <w:ins>/<w:del>/<w:moveFrom>/<w:moveTo>. Without
        // this grouping, source-loaded paragraphs with revision tracking would
        // emit individual <w:r>...</w:r> with no enclosing revision wrapper —
        // i.e., revision history silently wiped on round-trip.
        var positioned: [(position: Int, entry: PositionedEntry)] = []

        // v0.19.5+ (#56 R5-CONT P0 #6): runs / hyperlinks / fieldSimples /
        // alternateContents now mirror what contentControls (R5 P0 #2) and
        // bookmarkMarkers / commentRangeMarkers / etc. do — only `position
        // > 0` (source-loaded) entries join the sorted-emit list. API-built
        // entries (default `position == 0`) are emitted in the legacy
        // post-content section so they land AT THE END of paragraph text
        // rather than sorting BEFORE source-loaded children. Pre-fix an
        // appended Run with default position 0 sorted before any
        // `position >= 1` source run → text rendered at paragraph head
        // instead of the intended append position (verify R5 P0 #6 / DA C3).
        for run in runs where (run.position ?? 0) > 0 {
            positioned.append((run.position!, .run(run)))
        }
        for hyperlink in hyperlinks where (hyperlink.position ?? 0) > 0 {
            positioned.append((hyperlink.position!, .xml(hyperlink.toXML())))
        }
        for field in fieldSimples where (field.position ?? 0) > 0 {
            positioned.append((field.position!, .xml(Self.emitFieldSimple(field))))
        }
        for ac in alternateContents where (ac.position ?? 0) > 0 {
            positioned.append((ac.position!, .xml(ac.rawXML)))
        }
        // v0.19.5+ (#56 R5-CONT-2 P0 #4): apply `where position > 0` filter
        // to the 8 remaining positioned collections, completing the sweep
        // R5-CONT P0 #6 started for runs / hyperlinks / fieldSimples /
        // alternateContents. Pre-fix these went into the sort list
        // unconditionally → API-built marker (constructed with position 0)
        // sorted BEFORE every source-loaded child (position >= 1) and
        // landed at paragraph head. Post-content section below emits the
        // position-0 entries in append position to mirror the
        // contentControls + runs / hyperlinks / fieldSimples /
        // alternateContents handling.
        for marker in bookmarkMarkers where (marker.position ?? 0) > 0 {
            positioned.append((marker.position!, .xml(Self.emitBookmarkMarker(marker, paragraph: self))))
        }
        for marker in commentRangeMarkers where (marker.position ?? 0) > 0 {
            positioned.append((marker.position!, .xml(Self.emitCommentRangeMarker(marker))))
        }
        for marker in permissionRangeMarkers where (marker.position ?? 0) > 0 {
            positioned.append((marker.position!, .xml(Self.emitPermissionRangeMarker(marker))))
        }
        for marker in proofErrorMarkers where (marker.position ?? 0) > 0 {
            positioned.append((marker.position!, .xml("<w:proofErr w:type=\"\(marker.type.rawValue)\"/>")))
        }
        for tag in smartTags where (tag.position ?? 0) > 0 {
            positioned.append((tag.position!, .xml(tag.rawXML)))
        }
        for block in customXmlBlocks where (block.position ?? 0) > 0 {
            positioned.append((block.position!, .xml(block.rawXML)))
        }
        for block in bidiOverrides where (block.position ?? 0) > 0 {
            positioned.append((block.position!, .xml(block.rawXML)))
        }
        for child in unrecognizedChildren where (child.position ?? 0) > 0 {
            positioned.append((child.position!, .xml(child.rawXML)))
        }
        // v0.19.4+ (#56 R3-NEW-2): paragraph-level <w:sdt> with source position
        // joins the sorted emit. Position-0 controls are API-built and stay on
        // the post-content legacy path below for backward-compatibility.
        for control in contentControls where (control.position ?? 0) > 0 {
            positioned.append((control.position!, .xml(control.toXML())))
        }

        // Stable sort by position. Equal positions retain insertion order.
        positioned.sort { $0.position < $1.position }

        // Walk the sorted list. Whenever consecutive entries are `.run(_)`
        // sharing the same `revisionId`, emit them inside a single revision
        // wrapper. `.xml(_)` entries flush as-is.
        var i = 0
        while i < positioned.count {
            switch positioned[i].entry {
            case .xml(let s):
                xml += s
                i += 1
            case .run(let firstRun):
                // Collect run group: consecutive .run entries with same revisionId.
                var group: [Run] = [firstRun]
                var j = i + 1
                while j < positioned.count {
                    if case .run(let next) = positioned[j].entry, next.revisionId == firstRun.revisionId {
                        group.append(next)
                        j += 1
                    } else {
                        break
                    }
                }

                if let revId = firstRun.revisionId,
                   let revision = revisions.first(where: { $0.id == revId }) {
                    xml += revision.toOpeningXML()
                    let asDelText = revision.type == .deletion
                    for r in group {
                        xml += Self.emitRun(r, asDelText: asDelText, paragraphRevisions: revisions)
                    }
                    xml += revision.toClosingXML()
                } else {
                    for r in group {
                        xml += Self.emitRun(r, asDelText: false, paragraphRevisions: revisions)
                    }
                }
                i = j
            }
        }

        // v0.19.3+ (#56 round 2 P0-4 + P0-5): legacy post-content collections.
        // Mirrors `toXMLLegacy` ordering so any caller that mutated the legacy
        // single-list collections on a sort-routed paragraph still gets their
        // children emitted. For source-loaded paragraphs these collections are
        // empty (parser populates the positioned variants instead).
        //
        // v0.19.4+ (#56 R3-NEW-2): only emit position-0 contentControls here
        // (API-built). Position>0 controls were already emitted in the sorted
        // list above at their source position; emitting them again would
        // duplicate the SDT in the output.
        for control in contentControls where (control.position ?? 0) == 0 {
            xml += control.toXML()
        }
        // v0.19.5+ (#56 R5-CONT P0 #6): symmetric post-content emit for
        // API-built (position == 0) runs / hyperlinks / fieldSimples /
        // alternateContents. Pre-fix these defaulted to position 0 in
        // the positioned-list and sorted BEFORE every source child
        // (position >= 1), so an MCP `insertText` against a source-loaded
        // paragraph silently relocated text to the paragraph head. Now they
        // emit AFTER source-loaded children, matching the append semantics
        // that contentControls (R5 P0 #2) and the legacy emit path use.
        for run in runs where (run.position ?? 0) == 0 {
            xml += Self.emitRun(run, asDelText: false, paragraphRevisions: revisions)
        }
        for hyperlink in hyperlinks where (hyperlink.position ?? 0) == 0 {
            xml += hyperlink.toXML()
        }
        for field in fieldSimples where (field.position ?? 0) == 0 {
            xml += Self.emitFieldSimple(field)
        }
        for ac in alternateContents where (ac.position ?? 0) == 0 {
            xml += ac.rawXML
        }
        // v0.19.5+ (#56 R5-CONT-2 P0 #4): post-content emit for the 8
        // remaining position-0 collections — completes the sweep
        // R5-CONT P0 #6 only half-covered. Symmetric with the above.
        for marker in bookmarkMarkers where (marker.position ?? 0) == 0 {
            xml += Self.emitBookmarkMarker(marker, paragraph: self)
        }
        for marker in commentRangeMarkers where (marker.position ?? 0) == 0 {
            xml += Self.emitCommentRangeMarker(marker)
        }
        for marker in permissionRangeMarkers where (marker.position ?? 0) == 0 {
            xml += Self.emitPermissionRangeMarker(marker)
        }
        for marker in proofErrorMarkers where (marker.position ?? 0) == 0 {
            xml += "<w:proofErr w:type=\"\(marker.type.rawValue)\"/>"
        }
        for tag in smartTags where (tag.position ?? 0) == 0 {
            xml += tag.rawXML
        }
        for block in customXmlBlocks where (block.position ?? 0) == 0 {
            xml += block.rawXML
        }
        for block in bidiOverrides where (block.position ?? 0) == 0 {
            xml += block.rawXML
        }
        for child in unrecognizedChildren where (child.position ?? 0) == 0 {
            xml += child.rawXML
        }
        // v0.19.4+ (#56 R3-NEW-3): per-id gate (see pre-content emit above).
        // commentReference is unique to the legacy emit path — it is NOT
        // generated by CommentRangeMarker — so the per-id check is critical:
        // skipping a covered id avoids double-emit of end markers AND prevents
        // duplicating an inline reference run that the source-loaded paragraph
        // already carries (parsed as a Run with rawElements).
        for commentId in commentIds where !commentIdsCoveredByMarkers.contains(commentId) {
            xml += "<w:commentRangeEnd w:id=\"\(commentId)\"/>"
            xml += "<w:r><w:commentReference w:id=\"\(commentId)\"/></w:r>"
        }
        for footnoteId in footnoteIds {
            xml += "<w:r><w:rPr><w:rStyle w:val=\"FootnoteReference\"/></w:rPr><w:footnoteReference w:id=\"\(footnoteId)\"/></w:r>"
        }
        for endnoteId in endnoteIds {
            xml += "<w:r><w:rPr><w:rStyle w:val=\"EndnoteReference\"/></w:rPr><w:endnoteReference w:id=\"\(endnoteId)\"/></w:r>"
        }
        if bookmarkMarkers.isEmpty {
            for bookmark in bookmarks {
                xml += bookmark.toBookmarkEndXML()
            }
        }

        xml += "</w:p>"
        return xml
    }

    /// v0.19.2+ (#56 F3): Tagged payload for sort-by-position emit so the
    /// post-sort pass can recognize runs (and group them by revisionId) vs
    /// pre-rendered XML fragments.
    fileprivate enum PositionedEntry {
        case run(Run)
        case xml(String)
    }

    /// Emit a `<w:bookmarkStart>` or `<w:bookmarkEnd>` from a marker. For
    /// `.start`, look up the matching `Bookmark` in `paragraph.bookmarks`
    /// to retrieve the name (the marker only carries id).
    fileprivate static func emitBookmarkMarker(_ marker: BookmarkRangeMarker, paragraph: Paragraph) -> String {
        switch marker.kind {
        case .start:
            let name = paragraph.bookmarks.first(where: { $0.id == marker.id })?.name ?? ""
            return "<w:bookmarkStart w:id=\"\(marker.id)\" w:name=\"\(escapeXMLAttribute(name))\"/>"
        case .end:
            return "<w:bookmarkEnd w:id=\"\(marker.id)\"/>"
        }
    }

    fileprivate static func emitCommentRangeMarker(_ marker: CommentRangeMarker) -> String {
        switch marker.kind {
        case .start:
            return "<w:commentRangeStart w:id=\"\(marker.id)\"/>"
        case .end:
            return "<w:commentRangeEnd w:id=\"\(marker.id)\"/>"
        }
    }

    fileprivate static func emitPermissionRangeMarker(_ marker: PermissionRangeMarker) -> String {
        switch marker.kind {
        case .start:
            var attrs = "w:id=\"\(escapeXMLAttribute(marker.id))\""
            if let editorGroup = marker.editorGroup {
                attrs += " w:edGrp=\"\(escapeXMLAttribute(editorGroup))\""
            }
            if let editor = marker.editor {
                attrs += " w:ed=\"\(escapeXMLAttribute(editor))\""
            }
            return "<w:permStart \(attrs)/>"
        case .end:
            return "<w:permEnd w:id=\"\(escapeXMLAttribute(marker.id))\"/>"
        }
    }

    fileprivate static func emitFieldSimple(_ field: FieldSimple) -> String {
        var xml = "<w:fldSimple w:instr=\"\(escapeXMLAttribute(field.instr))\""
        for (name, value) in field.rawAttributes.sorted(by: { $0.key < $1.key }) {
            xml += " \(name)=\"\(escapeXMLAttribute(value))\""
        }
        xml += ">"
        for run in field.runs {
            xml += run.toXML()
        }
        xml += "</w:fldSimple>"
        return xml
    }

    // v0.19.5+ (#56 R5 P0 #3): the prior fileprivate static escapeXMLAttribute
    // was deleted. Call sites now route through the shared internal
    // `escapeXMLAttribute(_:)` from `IO/XMLAttributeEscape.swift`. The shared
    // helper additionally escapes `'` → `&apos;` (the prior local copy
    // missed it), closing the apostrophe-injection gap and standardizing on
    // `&apos;` (not `&#39;`) for byte-equivalence with Word's emit.

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
    /// v3.13.0+ (#175)：ECMA-376 `CT_PPr` 子元素的 schema 位置，key 是 `w:`
    /// 命名空間下的 local name。順序抄自 ISO/IEC 29500-4:2016 transitional
    /// `wml.xsd` 的 `CT_PPrBase` sequence（33 個）加上 `CT_PPr` extension 的
    /// `rPr`, `sectPr`, `pPrChange`，並與 python-docx `CT_PPr._tag_seq` 核對過。
    /// `PPrSchemaOrderTests.schemaSequence` 以字面常數獨立寫了一份，表抄錯會被
    /// 測試抓到。
    ///
    /// `toXML()` 以這張表排序 typed 欄位與 `rawChildren`。表外的 raw 子元素
    /// （qualified name 不是 `w:<表內名稱>`：擴充命名空間、`mc:AlternateContent`、
    /// 或 `w14:jc` 這種與表內同名但不同命名空間的元素）一律依「錨定」規則放：
    ///
    /// 1. 讀進來的表外元素（有 `rawChildSourceAnchors` 紀錄，以元素身分對應）：
    ///    錨點 = **來源 `<w:pPr>`** 中它之前最近的 schema 已知兄弟（typed 或 raw
    ///    都算）。緊接在錨點的 schema 位置之後（同一位置的已知元素之後、下一個
    ///    schema 位置之前）；來源中它之前沒有已知兄弟 → pPr 最前面。錨點是
    ///    「位置」不是元素：錨點元素之後被 typed 編輯或從 `rawChildren` 移除，
    ///    表外元素仍停在那個位置；在 `rawChildren` 其他地方增刪元素也不影響它。
    /// 2. 沒有紀錄（呼叫端自組或新加入）的 `mc:AlternateContent`：放在它的
    ///    `mc:Choice`／`mc:Fallback` 第一個 `w:` 子元素的 schema 位置——經 MCE
    ///    處理、換成該分支內容後仍是合法順序（分支只包一種元素時）。
    /// 3. 其他沒有紀錄的表外元素：錨定在 **`rawChildren` 陣列**中它之前最近的
    ///    schema 已知 raw 子元素之後；前面沒有 → pPr 最前面。
    /// 4. 同一位置的多個表外元素維持 `rawChildren` 中的先後。
    ///
    /// `sectPr` 與 `pPrChange` 由 `Paragraph` 的 pPr 包裝層在本函式輸出之後
    /// 輸出（`CT_PPr` 最後兩個位置），錨定在它們之後的未知元素因此會落在
    /// `rPr` 之後、`sectPr` 之前。
    internal static let canonicalPPrPosition: [String: Int] = [
        "pStyle": 1, "keepNext": 2, "keepLines": 3, "pageBreakBefore": 4,
        "framePr": 5, "widowControl": 6, "numPr": 7, "suppressLineNumbers": 8,
        "pBdr": 9, "shd": 10, "tabs": 11, "suppressAutoHyphens": 12,
        "kinsoku": 13, "wordWrap": 14, "overflowPunct": 15, "topLinePunct": 16,
        "autoSpaceDE": 17, "autoSpaceDN": 18, "bidi": 19, "adjustRightInd": 20,
        "snapToGrid": 21, "spacing": 22, "ind": 23, "contextualSpacing": 24,
        "mirrorIndents": 25, "suppressOverlap": 26, "jc": 27, "textDirection": 28,
        "textAlignment": 29, "textboxTightWrap": 30, "outlineLvl": 31, "divId": 32,
        "cnfStyle": 33, "rPr": 34, "sectPr": 35, "pPrChange": 36,
    ]

    /// 錨點不存在（規則 3）時的位置：早於所有 schema 位置。
    private static let leadingPosition = 0

    /// raw XML 開頭元素的 schema local name；qualified name 不是 `w:<表內名稱>`
    /// 時回傳 nil（= 表外元素）。與 `DocxReader` 一致以 `w:` 前綴認定
    /// WordprocessingML 命名空間——reader 的 typed 解析本來就只認 `w:` 前綴。
    internal static func schemaLocalName(ofRawXML xml: String) -> String? {
        guard let qualified = rawQualifiedName(xml),
              qualified.hasPrefix("w:") else { return nil }
        let local = String(qualified.dropFirst(2))
        return canonicalPPrPosition[local] == nil ? nil : local
    }

    /// `mc:AlternateContent`（任何前綴）的 `mc:Choice`／`mc:Fallback` 第一個直接
    /// 子元素若是 `w:<CT_PPr 元素>`，回傳其中最小的 schema 位置；否則 nil。
    /// 只看每個分支的第一個子元素：`<w:rPr><w:snapToGrid/></w:rPr>` 這種巢狀
    /// 內容裡的 `snapToGrid` 是 rPr 的子元素，不能拿來定位。
    private static func wrappedSchemaPosition(ofAlternateContent xml: String) -> Int? {
        guard let qualified = rawQualifiedName(xml),
              qualified.split(separator: ":").last == "AlternateContent" else { return nil }
        let pattern = #"<[A-Za-z_][\w.-]*:(?:Choice|Fallback)\b[^>]*>\s*<w:([A-Za-z]+)"#
        guard let regex = try? NSRegularExpression(pattern: pattern) else { return nil }
        let ns = xml as NSString
        return regex.matches(in: xml, range: NSRange(location: 0, length: ns.length))
            .compactMap { canonicalPPrPosition[ns.substring(with: $0.range(at: 1))] }
            .min()
    }

    private static func rawQualifiedName(_ xml: String) -> String? {
        let trimmed = xml.drop { $0.isWhitespace }
        guard trimmed.first == "<" else { return nil }
        let name = trimmed.dropFirst().prefix { !$0.isWhitespace && $0 != "/" && $0 != ">" }
        guard let first = name.first, first != "!", first != "?" else { return nil }
        return String(name)
    }

    /// 轉換為 OOXML XML 字串
    ///
    /// v3.13.0+ (#175)：所有子元素依 `canonicalPPrPosition` 的 `CT_PPr` schema
    /// 順序輸出——typed 欄位與 `rawChildren` 放進同一條排序管線，表外 raw 子元素
    /// 依錨定規則放置。修正前的固定順序是 pStyle → numPr → jc → spacing → ind →
    /// keepNext → keepLines → pageBreakBefore → pBdr → shd → rawChildren → rPr。
    public func toXML() -> String {
        // (position, tier, sequence, xml)：tier 0 = schema 已知元素，tier 1 =
        // 錨定在該位置之後的表外元素；sequence 保持同位置內的加入順序。
        var slots: [(position: Int, tier: Int, sequence: Int, xml: String)] = []
        func add(_ name: String, _ xml: String) {
            slots.append((Self.canonicalPPrPosition[name]!, 0, slots.count, xml))
        }

        // 樣式
        if let style = style {
            add("pStyle", "<w:pStyle w:val=\"\(escapeXMLAttribute(style))\"/>")
        }

        // 分頁控制
        if keepNext {
            add("keepNext", "<w:keepNext/>")
        }
        if keepLines {
            add("keepLines", "<w:keepLines/>")
        }
        if pageBreakBefore {
            add("pageBreakBefore", "<w:pageBreakBefore/>")
        }

        // 編號
        if let numbering = numbering {
            add("numPr", "<w:numPr><w:ilvl w:val=\"\(numbering.level)\"/>"
                + "<w:numId w:val=\"\(numbering.numId)\"/></w:numPr>")
        }

        // 段落邊框 / 段落底色。v3.13.0+ (#176)：讀取後沒被指派過（來源原文還在）
        // → 原樣輸出原文（theme 色、bar 邊、enum 外的 val 一個都不少）；否則 typed。
        // 投影相等的檢查是防呆：`didSet` 已保證有原文時值就是讀取當下的投影。
        if let border = border {
            add("pBdr", sourceBorder.flatMap { $0.projection == border ? $0.xml : nil }
                ?? border.toXML())
        }
        if let shading = shading {
            add("shd", sourceShading.flatMap { $0.projection == shading ? $0.xml : nil }
                ?? shading.toXML())
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
                add("spacing", "<w:spacing \(attrs.joined(separator: " "))/>")
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
                add("ind", "<w:ind \(attrs.joined(separator: " "))/>")
            }
        }

        // 對齊
        if let alignment = alignment {
            add("jc", "<w:jc w:val=\"\(alignment.rawValue)\"/>")
        }

        // v0.20.2+ (#65 sub-stack D): paragraph-mark <w:rPr> (CT_PPr 的 rPr
        // 位置). RunProperties.toXML() returns the inner content (rStyle,
        // b, rFonts, lang, rawChildren, etc.) without the outer wrapper, so
        // we wrap with <w:rPr>...</w:rPr> here. Skip emission when nil OR
        // when the inner content is empty — matches the "no synthetic empty
        // <w:rPr/>" gate discipline established in sub-stack B-CONT-2.
        if let markProps = markRunProperties {
            let inner = markProps.toXML()
            if !inner.isEmpty {
                add("rPr", "<w:rPr>\(inner)</w:rPr>")
            }
        }

        // v3.12.0+ (#168) / v3.13.0+ (#175): pPr children the typed model
        // doesn't recognize (kinsoku, snapToGrid, widowControl, …), captured
        // verbatim by `DocxReader.parseParagraphProperties`, each slotted by
        // element name; 表外元素依 `canonicalPPrPosition` 說明的錨定規則放置。
        var unusedRecords = rawChildSourceAnchors
        var precedingKnown: String? = nil
        for raw in rawChildren {
            if let local = Self.schemaLocalName(ofRawXML: raw.xml) {
                slots.append((Self.canonicalPPrPosition[local]!, 0, slots.count, raw.xml))
                precedingKnown = local
                continue
            }
            func position(after anchor: String?) -> Int {
                anchor.flatMap { Self.canonicalPPrPosition[$0] } ?? Self.leadingPosition
            }
            if let record = unusedRecords.firstIndex(where: { $0.xml == raw.xml }) {
                // 規則 1：讀進來的表外元素，以元素身分對應來源錨點。
                slots.append((position(after: unusedRecords.remove(at: record).anchor), 1, slots.count, raw.xml))
            } else if let wrapped = Self.wrappedSchemaPosition(ofAlternateContent: raw.xml) {
                // 規則 2：沒有紀錄的 mc:AlternateContent 佔它包著的元素的位置。
                slots.append((wrapped, 0, slots.count, raw.xml))
            } else {
                // 規則 3：沒有紀錄的其他表外元素，錨定在陣列中前一個已知元素。
                slots.append((position(after: precedingKnown), 1, slots.count, raw.xml))
            }
        }

        slots.sort {
            ($0.position, $0.tier, $0.sequence) < ($1.position, $1.tier, $1.sequence)
        }
        return slots.map(\.xml).joined()
    }
}
