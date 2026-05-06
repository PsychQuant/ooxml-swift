import Foundation

/// 文字運行 (Run) - Word 文件中的最小文字單元
/// 一個 Run 包含具有相同格式的連續文字
/// v0.14.0+ (che-word-mcp#52): typed carrier for an unknown OOXML element
/// preserved verbatim within a Run. `name` is the local element name
/// (e.g., `"pict"`, `"object"`); `xml` is the verbatim serialized XML.
public struct RawElement: Equatable {
    public let name: String
    public let xml: String
    public init(name: String, xml: String) {
        self.name = name
        self.xml = xml
    }
}

public struct Run {
    /// v0.31.1+ (Spectra `sibling-types-tree-projection-impl`,
    /// `word-aligned-state-sync` Phase 1 task 2.2): when non-nil, this Run is
    /// a tree-backed view over the wrapped `<w:r>` element. Getters walk
    /// `xmlNode.children` at access time; the `text` setter mutates the tree
    /// directly (Phase 1 stub) and calls `xmlNode.markDirty()` so the writer
    /// re-serializes from typed fields. When nil (legacy detached mode),
    /// getters/setters operate on the legacy stored buffers below.
    /// `XmlNode` is a class, so two value-copies of the same tree-backed Run
    /// share the same underlying tree state. Mirrors the pattern shipped for
    /// `Paragraph` in v0.31.0 (`paragraph-tree-projection-impl`).
    public var xmlNode: XmlNode?

    /// Legacy stored backing for `text` used in detached mode.
    /// In tree-backed mode the public `text` accessor walks `xmlNode.children`
    /// instead and this buffer is ignored. Renamed from the previous public
    /// `text` stored property in v0.31.1.
    internal var _legacyText: String = ""

    /// Legacy stored backing for `properties` used in detached mode.
    /// In tree-backed mode the public `properties` accessor returns the
    /// `RunProperties()` Phase 1 stub default; the setter ghost-writes here
    /// (full tree-walking parser arrives in Phase 2). Renamed from the
    /// previous public `properties` stored property in v0.31.1.
    internal var _legacyProperties: RunProperties = RunProperties()

    public var drawing: Drawing?  // 圖片繪圖元素
    public var rawXML: String?    // 原始 XML（用於欄位代碼、SDT 等進階功能）
    public var semantic: SemanticAnnotation?  // 語義標註

    /// v0.14.0+ (che-word-mcp#52): preserves unknown OOXML child elements of
    /// `<w:r>` (e.g., `<w:pict>` VML watermarks, `<w:object>` OLE embeds,
    /// `<w:ruby>` annotations) by carrying their verbatim XML. Populated by
    /// `DocxReader.parseRun` for any child element whose local name is NOT
    /// among the typed kinds (`rPr`, `t`, `drawing`, `oMath`, `oMathPara`).
    /// Emitted verbatim by `Run.toXML()` after typed children. Default `nil`
    /// preserves Equatable equality for programmatic Run construction.
    public var rawElements: [RawElement]?

    /// v0.18.0+ (che-word-mcp#45): id of the `Revision` in the enclosing
    /// `Paragraph.revisions` collection that this run belongs to (insertion,
    /// deletion, moveFrom, moveTo). When set, `Paragraph.toXML()` groups
    /// consecutive runs with the same `revisionId` and wraps them with the
    /// matching `<w:ins>` / `<w:del>` / `<w:moveFrom>` / `<w:moveTo>` element.
    public var revisionId: Int?

    /// v0.18.0+ (che-word-mcp#45): id of the `Revision` (type `.formatChange`)
    /// in the enclosing `Paragraph.revisions` whose `previousFormat` describes
    /// this run's pre-mutation properties. Orthogonal to `revisionId` — a run
    /// can be a tracked insertion AND carry an independent format-change
    /// revision. When set, `Paragraph.toXML()` emits `<w:rPrChange>` inside
    /// the run's `<w:rPr>` block so Word UI shows the format change as a
    /// tracked revision.
    public var formatChangeRevisionId: Int?

    /// v0.19.0+ (PsychQuant/che-word-mcp#56) Phase 4: source-document order
    /// index for `Paragraph.toXML()` sort-by-position emit. Populated by
    /// `DocxReader.parseParagraph` for direct `<w:r>` children only — runs
    /// inside revision wrappers (`<w:ins>` / `<w:del>` / etc.) leave this
    /// at 0 because their emit order is governed by the enclosing wrapper.
    /// Default 0 keeps backward compat with API-built Runs.
    public var position: Int? = nil

    public init(text: String, properties: RunProperties = RunProperties()) {
        self._legacyText = text
        self._legacyProperties = properties
        self.drawing = nil
        self.rawXML = nil
        self.semantic = nil
    }

    /// v0.31.1+ Tree-backed initializer. Wraps an existing `<w:r>` xmlNode so
    /// getters walk its children and the `text` setter mutates the tree
    /// directly.
    ///
    /// The legacy stored fields (`_legacyText`, `_legacyProperties`,
    /// `drawing`, `rawXML`, `semantic`, `rawElements`, `revisionId`,
    /// `formatChangeRevisionId`, `position`) are initialized to their empty
    /// defaults; in tree-backed mode `text` and `properties` are shadowed by
    /// computed accessors that read from `xmlNode.children`. Callers MUST NOT
    /// rely on those stored fields when `xmlNode != nil`.
    ///
    /// Semantic validation (asserting the node is a `<w:r>` element) is left
    /// to callers; this initializer accepts any element xmlNode so unit tests
    /// can synthesize fixtures without paying the schema-check cost.
    public init(xmlNode: XmlNode) {
        self.xmlNode = xmlNode
        // All legacy stored fields keep their default values; the computed
        // `text` / `properties` accessors below shadow them when tree-backed.
    }

    /// v0.31.1+ Stable identifier for this run.
    ///
    /// - Tree-backed: returns `xmlNode.stableID` if any OOXML stable-ID
    ///   attribute is present (e.g. `"w:id=42"` for revision runs); otherwise
    ///   falls back to `"lib:<UUID>"` when the reader assigned a
    ///   library-generated UUID; otherwise `nil`.
    /// - Detached (legacy): always returns `nil`.
    ///
    /// `<w:r>` does not natively carry `w14:paraId` / `w:bookmarkId`; the
    /// `XmlNode.stableID` precedence list still resolves `w:id` (Run revision
    /// IDs) and `r:id` (relationship-bearing wrappers). The op log addresses
    /// runs by these surrogate IDs or the `lib:` UUID fallback.
    public var id: String? {
        guard let node = xmlNode else { return nil }
        if let stable = node.stableID { return stable }
        if let lib = node.libraryUUID { return "lib:\(lib.uuidString)" }
        return nil
    }

    /// v0.31.1+ Mode-aware view of the run's plain-text content.
    ///
    /// - Tree-backed getter: concatenates `textContent` of every `<w:t>`
    ///   direct child of the wrapped `<w:r>` xmlNode, in document order.
    ///   No caching — re-walks on every access.
    /// - Detached getter: returns the legacy stored `_legacyText`.
    ///
    /// Tree-backed setter (Phase 1 stub): replaces the wrapped `<w:r>`'s
    /// existing `<w:t>` direct children with a single new `<w:t>X</w:t>`
    /// element while preserving every non-`<w:t>` sibling (`<w:rPr>`,
    /// `<w:tab>`, `<w:br>`, `<w:drawing>`, …). Calls `xmlNode.markDirty()`
    /// so `XmlTreeWriter` re-serializes the run from typed fields. Phase 2
    /// of `word-aligned-state-sync` (target v0.32.0) routes this through the
    /// op log to preserve formatting more faithfully.
    ///
    /// Detached setter: writes to `_legacyText` directly (matches pre-v0.31.1
    /// behavior).
    public var text: String {
        get {
            if let node = xmlNode {
                var out = ""
                for child in node.children where child.kind == .element && child.localName == "t" {
                    for grand in child.children where grand.kind == .text {
                        out += grand.textContent
                    }
                }
                return out
            }
            return _legacyText
        }
        set {
            if let node = xmlNode {
                // Build the replacement <w:t>X</w:t>.
                let textNode = XmlNode.text(newValue)
                let wt = XmlNode.element(prefix: "w", localName: "t", children: [textNode])
                // Preserve every non-<w:t> sibling (<w:rPr>, <w:tab>, <w:br>, …)
                // and replace the existing <w:t> children with the new one.
                // Per spec: "the wrapped xmlNode's children SHALL contain
                // exactly one <w:t>New</w:t> element among any non-<w:t>
                // siblings preserved".
                var rebuilt: [XmlNode] = []
                var inserted = false
                for child in node.children {
                    if child.kind == .element && child.localName == "t" {
                        if !inserted {
                            rebuilt.append(wt)
                            inserted = true
                        }
                        // Drop additional <w:t> children — only one new one survives.
                    } else {
                        rebuilt.append(child)
                    }
                }
                if !inserted {
                    // No pre-existing <w:t> child: append the new one at the end.
                    rebuilt.append(wt)
                }
                node.children = rebuilt
                node.markDirty()
            } else {
                _legacyText = newValue
            }
        }
    }

    /// v0.31.1+ Mode-aware view of the run's `<w:rPr>` properties.
    ///
    /// - Tree-backed getter: **Phase 1 stub** — returns the default
    ///   `RunProperties()` regardless of the wrapped `<w:rPr>` child contents.
    ///   Full `<w:rPr>` tree-walking parser is deferred to Phase 2 because
    ///   `RunProperties` is a 14+ field struct whose round-trip parser already
    ///   exists in `DocxReader` and is non-trivial to mirror here.
    /// - Detached getter: returns the legacy stored `_legacyProperties`.
    ///
    /// Tree-backed setter (Phase 1 stub): ghost-writes to `_legacyProperties`
    /// (no tree mutation). Phase 2 op-log routing replaces this with proper
    /// `<w:rPr>` reconstruction.
    ///
    /// Detached setter: writes to `_legacyProperties` directly.
    public var properties: RunProperties {
        get {
            if xmlNode != nil {
                return RunProperties()
            }
            return _legacyProperties
        }
        set {
            // Ghost-write in both modes; Phase 1 limitation documented above.
            _legacyProperties = newValue
        }
    }
}

// MARK: - Equatable (mode-aware identity vs content)

extension Run: Equatable {
    /// v0.31.1+ Custom Equatable replacing auto-synthesized conformance per
    /// `sibling-types-tree-projection-impl` Decision 5.
    ///
    /// Behavior depends on the storage mode of both sides:
    ///
    /// 1. **Both tree-backed**: identity equality on the wrapped `xmlNode`
    ///    reference (`===`). Op-log addresses runs by id (== identity);
    ///    content equality on different elements would silently merge log
    ///    entries that target different runs.
    /// 2. **Both detached**: content equality across the legacy stored fields,
    ///    preserving pre-v0.31.1 auto-synthesized behavior the che-word-mcp
    ///    test suite depends on.
    /// 3. **Mixed (one tree-backed, one detached)**: always `false`. The two
    ///    storage modes are not interchangeable; comparing across them is
    ///    almost certainly a caller mistake worth surfacing.
    public static func == (lhs: Run, rhs: Run) -> Bool {
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
    /// Mirrors what auto-synthesized `Equatable` would have compared on the
    /// pre-v0.31.1 struct shape.
    private static func contentEquals(_ lhs: Run, _ rhs: Run) -> Bool {
        return lhs._legacyText == rhs._legacyText
            && lhs._legacyProperties == rhs._legacyProperties
            && lhs.drawing == rhs.drawing
            && lhs.rawXML == rhs.rawXML
            && lhs.semantic == rhs.semantic
            && lhs.rawElements == rhs.rawElements
            && lhs.revisionId == rhs.revisionId
            && lhs.formatChangeRevisionId == rhs.formatChangeRevisionId
            && lhs.position == rhs.position
    }
}

// MARK: - Run Properties

/// v0.20.0+ (#60): 4-axis `<w:rFonts>` properties. ECMA-376 §17.3.2 RPrBase
/// distinguishes Latin (`w:ascii`), High-ANSI (`w:hAnsi`), East-Asian
/// (`w:eastAsia`), and Complex Script (`w:cs`) font assignments because
/// different scripts may need different fonts (e.g., Times New Roman for
/// Latin + DFKai-SB for traditional Chinese eastAsia + Mangal for Devanagari
/// cs). Pre-v0.20.0 `RunProperties.fontName: String?` collapsed all 4 into
/// one value; this struct restores the distinction. `fontName` field kept
/// for backward compat — when both are set, `rFonts` wins.
public struct RFontsProperties: Equatable {
    public var ascii: String?
    public var hAnsi: String?
    public var eastAsia: String?
    public var cs: String?
    /// `w:hint` controls which axis is used when text crosses script boundaries
    /// without explicit per-character font. Common values: `default`, `eastAsia`, `cs`.
    public var hint: String?

    public init(ascii: String? = nil, hAnsi: String? = nil,
                eastAsia: String? = nil, cs: String? = nil, hint: String? = nil) {
        self.ascii = ascii; self.hAnsi = hAnsi
        self.eastAsia = eastAsia; self.cs = cs; self.hint = hint
    }
}

/// v0.20.0+ (#60): 3-axis `<w:lang>` properties. ECMA-376 §17.3.2 separates
/// Latin language tag (`w:val`), East-Asian (`w:eastAsia`), and Bidi (`w:bidi`).
public struct LanguageProperties: Equatable {
    public var val: String?
    public var eastAsia: String?
    public var bidi: String?

    public init(val: String? = nil, eastAsia: String? = nil, bidi: String? = nil) {
        self.val = val; self.eastAsia = eastAsia; self.bidi = bidi
    }
}

/// Run 格式屬性
public struct RunProperties: Equatable {
    public var bold: Bool = false
    public var italic: Bool = false
    public var underline: UnderlineType?
    public var strikethrough: Bool = false
    public var fontSize: Int?              // 半點 (24 = 12pt)
    public var fontName: String?           // legacy single-axis; mirrors rFonts.ascii. Use rFonts for 4-axis preservation.
    public var color: String?              // RGB hex (e.g., "FF0000")
    public var highlight: HighlightColor?
    public var verticalAlign: VerticalAlign?
    public var characterSpacing: CharacterSpacing?  // 字元間距
    public var textEffect: TextEffect?              // 文字效果
    public var rawXML: String?                      // 原始 XML（用於進階功能如 SDT）

    /// v0.19.3+ (#56 round 2 P0-1): style reference name emitted as
    /// `<w:rStyle w:val="..."/>` — must appear FIRST inside `<w:rPr>` per
    /// ECMA-376 §17.3.2 ordering (CT_RPr's first child). Hyperlink-styled
    /// runs use `"Hyperlink"`, footnote refs use `"FootnoteReference"`,
    /// endnote refs use `"EndnoteReference"`.
    public var rStyle: String?

    /// v0.20.0+ (#60): 4-axis `<w:rFonts>` (ascii / hAnsi / eastAsia / cs).
    /// When set, takes precedence over `fontName`. When `fontName` is set
    /// and `rFonts` is nil, writer emits all 4 axes with the same value
    /// (legacy behavior). See `RFontsProperties` doc comment.
    public var rFonts: RFontsProperties?

    /// v0.20.0+ (#60): `<w:noProof/>` — suppress spell/grammar check on this run.
    public var noProof: Bool = false

    /// v0.20.0+ (#60): `<w:kern w:val="N"/>` — minimum font size threshold for kerning.
    /// OOXML uses half-points (e.g., kern=32 means kern only at 16pt+).
    public var kern: Int?

    /// v0.20.0+ (#60): 3-axis `<w:lang>` (val / eastAsia / bidi).
    public var lang: LanguageProperties?

    /// v0.20.0+ (#60): unrecognized direct children of `<w:rPr>`. Captured
    /// verbatim for byte-equivalent emission. Common content: `<w14:textOutline>`,
    /// `<w14:textFill>`, `<w14:glow>`, `<w14:shadow>`, `<w14:reflection>`,
    /// `<w14:scene3d>`, `<w14:props3d>`, `<w14:ligatures>`, `<w14:numForm>`,
    /// `<w14:numSpacing>`, `<w14:stylisticSets>`, `<w14:cntxtAlts>`. Same
    /// architectural pattern as `Run.rawElements` (v0.14.0+, #52).
    public var rawChildren: [RawElement]?

    public init() {}

    public init(bold: Bool = false,
         italic: Bool = false,
         underline: UnderlineType? = nil,
         strikethrough: Bool = false,
         fontSize: Int? = nil,
         fontName: String? = nil,
         color: String? = nil,
         highlight: HighlightColor? = nil,
         verticalAlign: VerticalAlign? = nil,
         rStyle: String? = nil) {
        self.bold = bold
        self.italic = italic
        self.underline = underline
        self.strikethrough = strikethrough
        self.fontSize = fontSize
        self.fontName = fontName
        self.color = color
        self.highlight = highlight
        self.verticalAlign = verticalAlign
        self.rStyle = rStyle
    }

    /// 合併格式（覆蓋非 nil 值）
    mutating func merge(with other: RunProperties) {
        if other.bold { self.bold = true }
        if other.italic { self.italic = true }
        if let underline = other.underline { self.underline = underline }
        if other.strikethrough { self.strikethrough = true }
        if let fontSize = other.fontSize { self.fontSize = fontSize }
        if let fontName = other.fontName { self.fontName = fontName }
        if let color = other.color { self.color = color }
        if let highlight = other.highlight { self.highlight = highlight }
        if let verticalAlign = other.verticalAlign { self.verticalAlign = verticalAlign }
        if let characterSpacing = other.characterSpacing { self.characterSpacing = characterSpacing }
        if let textEffect = other.textEffect { self.textEffect = textEffect }
        if let rawXML = other.rawXML { self.rawXML = rawXML }
        if let rStyle = other.rStyle { self.rStyle = rStyle }
        // v0.20.0+ (#60): merge new typed fields. `rFonts` overrides whole struct
        // when set (per-axis merge would silently mask source values from the
        // base properties — safer to overwrite atomically).
        if let rFonts = other.rFonts { self.rFonts = rFonts }
        if other.noProof { self.noProof = true }
        if let kern = other.kern { self.kern = kern }
        if let lang = other.lang { self.lang = lang }
        if let rawChildren = other.rawChildren { self.rawChildren = rawChildren }
    }
}

// MARK: - Enums

/// 底線類型
public enum UnderlineType: String, Codable {
    case single = "single"
    case double = "double"
    case dotted = "dotted"
    case dashed = "dash"
    case wave = "wave"
    case thick = "thick"
    case words = "words"        // 只在文字下，空格無底線
}

/// 螢光標記顏色
public enum HighlightColor: String, Codable {
    case yellow = "yellow"
    case green = "green"
    case cyan = "cyan"
    case magenta = "magenta"
    case blue = "blue"
    case red = "red"
    case darkBlue = "darkBlue"
    case darkCyan = "darkCyan"
    case darkGreen = "darkGreen"
    case darkMagenta = "darkMagenta"
    case darkRed = "darkRed"
    case darkYellow = "darkYellow"
    case lightGray = "lightGray"
    case darkGray = "darkGray"
    case black = "black"
    case white = "white"
}

/// 垂直對齊（上標/下標）
public enum VerticalAlign: String, Codable {
    case superscript = "superscript"
    case `subscript` = "subscript"
    case baseline = "baseline"
}

// MARK: - XML 生成

extension Run {
    /// 轉換為 OOXML XML 字串
    func toXML() -> String {
        // 如果 Run 本身有原始 XML，直接輸出（用於欄位代碼、SDT 等）
        if let rawXML = self.rawXML {
            return rawXML
        }

        // 如果 RunProperties 有原始 XML，也直接輸出
        if let rawXML = properties.rawXML {
            return rawXML
        }

        var xml = "<w:r>"

        // Run Properties
        let propsXML = properties.toXML()
        if !propsXML.isEmpty {
            xml += "<w:rPr>\(propsXML)</w:rPr>"
        }

        // Drawing (圖片) - 如果有圖片，優先輸出圖片
        if let drawing = drawing {
            xml += drawing.toXML()
        } else if !text.isEmpty || (rawElements?.isEmpty ?? true) {
            // v0.14.0+ (che-word-mcp#52): when a Run carries only rawElements
            // (e.g., VML watermark with no text child), suppress the synthetic
            // empty <w:t> emission. Empty <w:t> + <w:pict> would inject a
            // spurious empty text node into Word output. Only emit <w:t> when
            // we have actual text OR when there are no rawElements to emit.
            //
            // PsychQuant/ooxml-swift#5 (F13): autosense `xml:space="preserve"`.
            // Pre-fix this attribute was emitted unconditionally — harmless
            // but non-canonical for runs with empty / single-internal-space
            // text. Post-fix the attribute appears only when text contains
            // semantically significant whitespace (leading / trailing /
            // 2+ consecutive whitespace chars). Single internal whitespace
            // is XML-normalised so the flag adds noise; empty text needs no
            // protection at all.
            let needsPreserve = Self.needsXMLSpacePreserve(text)
            let openTag = needsPreserve ? "<w:t xml:space=\"preserve\">" : "<w:t>"
            xml += "\(openTag)\(escapeXML(text))</w:t>"
        }

        // v0.14.0+ (che-word-mcp#52): emit preserved unknown elements in
        // source-document order, after typed children but before </w:r>.
        if let rawElements = rawElements {
            for raw in rawElements {
                xml += raw.xml
            }
        }

        xml += "</w:r>"

        return xml
    }

    /// PsychQuant/ooxml-swift#5 (F13): determine whether a `<w:t>` payload
    /// needs the `xml:space="preserve"` attribute. XML normalises single
    /// internal whitespace; semantic whitespace needs the flag to survive
    /// round-trip through Word's parser. Trigger when text starts/ends with
    /// whitespace OR contains two-or-more consecutive whitespace characters.
    /// Empty text returns false (nothing to protect).
    fileprivate static func needsXMLSpacePreserve(_ text: String) -> Bool {
        if text.isEmpty { return false }
        if text.first?.isWhitespace == true { return true }
        if text.last?.isWhitespace == true { return true }
        // Two-or-more consecutive whitespace chars (regex `\s\s`).
        return text.range(of: #"\s\s"#, options: .regularExpression) != nil
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

extension RunProperties {
    /// ECMA-376 §17.3.2.28 `CT_RPr` canonical child position table, keyed by
    /// localName (no namespace prefix). Used by `toXML()` to slot every typed
    /// emit AND every `rawChildren` element at its schema-mandated index.
    ///
    /// Elements not in this table (vendor extensions like `w14:ligatures`,
    /// `w14:numForm`, `w16cid:*`) are emitted at `unknownTailPosition` —
    /// after every known EG_RPrBase child but before any `<w:rPrChange>`.
    ///
    /// PsychQuant/ooxml-swift#61 / kiki830621/collaboration_guo_analysis#20
    /// (v0.26.0): generalised v0.25.0's "reorder typed parts" fix to also
    /// cover `rawChildren` elements (`bCs`, `webHidden`, `iCs`, `vanish`,
    /// `dstrike`, `caps`, `smallCaps`, etc.). Pre-fix these landed at the
    /// tail of `<w:rPr>` regardless of canonical position, leaving
    /// `bCs`-out-of-order at a 4% violation rate even after v0.25.0.
    fileprivate static let canonicalRPrPosition: [String: Int] = [
        "rStyle": 1, "rFonts": 2,
        "b": 3, "bCs": 4, "i": 5, "iCs": 6,
        "caps": 7, "smallCaps": 8,
        "strike": 9, "dstrike": 10,
        "outline": 11, "shadow": 12, "emboss": 13, "imprint": 14,
        "noProof": 15, "snapToGrid": 16,
        "vanish": 17, "webHidden": 18,
        "color": 19, "spacing": 20, "w": 21, "kern": 22, "position": 23,
        "sz": 24, "szCs": 25,
        "highlight": 26, "u": 27, "effect": 28,
        "bdr": 29, "shd": 30, "fitText": 31,
        "vertAlign": 32,
        "rtl": 33, "cs": 34, "em": 35, "lang": 36,
        "eastAsianLayout": 37, "specVanish": 38, "oMath": 39,
        "rPrChange": 9999  // CT_RPr permits zero/one rPrChange after EG_RPrBase.
    ]

    /// Position assigned to elements not present in `canonicalRPrPosition` —
    /// vendor extensions emit at end of EG_RPrBase, before any rPrChange.
    fileprivate static let unknownTailPosition = 9000

    /// 轉換為 OOXML XML 字串
    ///
    /// All children are emitted in ECMA-376 §17.3.2.28 `CT_RPr` canonical
    /// sequence. Both typed fields (`rStyle`, `rFonts`, `bold`, …) and
    /// `rawChildren` (verbatim XML blobs from the parser) are slotted into
    /// the same positional pipeline — no element is appended unconditionally
    /// at the tail.
    ///
    /// **Why ordering matters**: macOS Word's strict OOXML validator rejects
    /// docx files when too many `<w:rPr>` blocks have inverted child order.
    /// v0.22+ regressed to a 65% violation rate; v0.25.0 fixed typed-emit
    /// order (down to 6%); v0.26.0 (this version) extends the fix to
    /// `rawChildren` placement, driving the rate to ~0%.
    func toXML() -> String {
        // (canonicalPos, sortKey, xml) — `sortKey` breaks ties for same-position
        // emits (e.g. duplicate kern when both characterSpacing.kern and
        // self.kern are set; szCs follows sz at the next slot anyway).
        var slots: [(pos: Int, sortKey: Int, xml: String)] = []
        var nextSortKey = 0
        func add(_ pos: Int, _ xml: String) {
            slots.append((pos, nextSortKey, xml))
            nextSortKey += 1
        }

        // 1. rStyle
        // v0.19.3+ (#56 round 2 P0-1): Hyperlinks rely on this for the
        // "Hyperlink" character style to apply correctly in Word.
        if let rStyle = rStyle {
            // v0.19.4+ (#56 R3-NEW-6): escape source-string before attribute
            // interpolation so a malicious `x"/><inj/>` payload cannot inject
            // sibling elements at write time.
            add(1, "<w:rStyle w:val=\"\(escapeXMLAttribute(rStyle))\"/>")
        }

        // 2. rFonts
        // v0.20.0+ (#60): 4-axis preservation when `rFonts` struct is set;
        // fall back to legacy `fontName` (single value mirrored to all 4 axes)
        // when only the legacy field is set.
        if let rFonts = rFonts {
            var attrs: [String] = []
            if let ascii = rFonts.ascii { attrs.append("w:ascii=\"\(escapeXMLAttribute(ascii))\"") }
            if let hAnsi = rFonts.hAnsi { attrs.append("w:hAnsi=\"\(escapeXMLAttribute(hAnsi))\"") }
            if let eastAsia = rFonts.eastAsia { attrs.append("w:eastAsia=\"\(escapeXMLAttribute(eastAsia))\"") }
            if let cs = rFonts.cs { attrs.append("w:cs=\"\(escapeXMLAttribute(cs))\"") }
            if let hint = rFonts.hint { attrs.append("w:hint=\"\(escapeXMLAttribute(hint))\"") }
            if !attrs.isEmpty {
                add(2, "<w:rFonts \(attrs.joined(separator: " "))/>")
            }
        } else if let fontName = fontName {
            // v0.19.4+ (#56 R3-NEW-6 audit): fontName flows into 4 attributes;
            // escape once to avoid same injection sink as rStyle.
            let n = escapeXMLAttribute(fontName)
            add(2, "<w:rFonts w:ascii=\"\(n)\" w:hAnsi=\"\(n)\" w:eastAsia=\"\(n)\" w:cs=\"\(n)\"/>")
        }

        // 3. b · 5. i · 9. strike · 15. noProof
        if bold { add(3, "<w:b/>") }
        if italic { add(5, "<w:i/>") }
        if strikethrough { add(9, "<w:strike/>") }
        if noProof { add(15, "<w:noProof/>") }

        // 19. color
        if let color = color {
            // v0.19.4+ (#56 R3-NEW-6 audit): color is a hex string in normal
            // use but the field is `String?` so escape defensively.
            add(19, "<w:color w:val=\"\(escapeXMLAttribute(color))\"/>")
        }

        // 20-23. CharacterSpacing block (spacing/w/kern/position)
        // Inline-decompose so each sub-element lands at its canonical slot
        // (spacing 20 → kern 22 → position 23). v0.26.0+ (#61 follow-up)
        // replaces the prior monolithic CharacterSpacing.toXML() append which
        // emitted spacing→position→kern (also out of order for kern↔position).
        if let cs = characterSpacing {
            if let spacing = cs.spacing {
                add(20, "<w:spacing w:val=\"\(spacing)\"/>")
            }
            if let kern = cs.kern {
                add(22, "<w:kern w:val=\"\(kern)\"/>")
            }
            if let position = cs.position {
                add(23, "<w:position w:val=\"\(position)\"/>")
            }
        }

        // 22. typed kern (v0.20.0+ #60).
        // Note: if both `characterSpacing.kern` and `self.kern` are set,
        // two `<w:kern>` elements appear — pre-existing data-model overlap
        // unchanged by this fix; Word tolerates duplicates and the stable
        // sort preserves API call order between them.
        if let kern = kern {
            add(22, "<w:kern w:val=\"\(kern)\"/>")
        }

        // 24/25. sz / szCs (OOXML 使用半點 / half-points)
        if let fontSize = fontSize {
            add(24, "<w:sz w:val=\"\(fontSize)\"/>")
            add(25, "<w:szCs w:val=\"\(fontSize)\"/>")
        }

        // 26. highlight · 27. u
        if let highlight = highlight {
            add(26, "<w:highlight w:val=\"\(highlight.rawValue)\"/>")
        }
        if let underline = underline {
            add(27, "<w:u w:val=\"\(underline.rawValue)\"/>")
        }

        // 28. effect (TextEffect — TextEffect.none returns empty so guard).
        if let textEffect = textEffect {
            let effectXML = textEffect.toXML()
            if !effectXML.isEmpty {
                add(28, effectXML)
            }
        }

        // 32. vertAlign
        if let verticalAlign = verticalAlign {
            add(32, "<w:vertAlign w:val=\"\(verticalAlign.rawValue)\"/>")
        }

        // 36. lang (v0.20.0+ #60)
        if let lang = lang {
            var attrs: [String] = []
            if let val = lang.val { attrs.append("w:val=\"\(escapeXMLAttribute(val))\"") }
            if let eastAsia = lang.eastAsia { attrs.append("w:eastAsia=\"\(escapeXMLAttribute(eastAsia))\"") }
            if let bidi = lang.bidi { attrs.append("w:bidi=\"\(escapeXMLAttribute(bidi))\"") }
            if !attrs.isEmpty {
                add(36, "<w:lang \(attrs.joined(separator: " "))/>")
            }
        }

        // rawChildren (vendor extensions, `<w:bCs>`, `<w:webHidden>`,
        // `<w:rPrChange>`, etc.). v0.26.0+ (#61 follow-up): each rawChild is
        // looked up by localName in `canonicalRPrPosition`. Known schema
        // elements land at their canonical slot (e.g. `bCs` at 4, `webHidden`
        // at 18); unknown vendor extensions go to `unknownTailPosition` (after
        // every known EG_RPrBase child but before rPrChange). Pre-fix all
        // rawChildren landed at the very end regardless of position, causing
        // 417 residual violations after v0.25.0 typed-emit reorder.
        //
        // Source-document insertion order is preserved within same-position
        // groups via the stable `sortKey` tie-break.
        if let rawChildren = rawChildren {
            for raw in rawChildren {
                let pos = Self.canonicalRPrPosition[raw.name] ?? Self.unknownTailPosition
                add(pos, raw.xml)
            }
        }

        // Stable sort by canonical position (sortKey breaks same-position ties).
        slots.sort { lhs, rhs in
            if lhs.pos != rhs.pos { return lhs.pos < rhs.pos }
            return lhs.sortKey < rhs.sortKey
        }

        return slots.map { $0.xml }.joined()
    }
}

