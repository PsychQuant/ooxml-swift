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

public struct Run: Equatable {
    public var text: String
    public var properties: RunProperties
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
    public var position: Int = 0

    public init(text: String, properties: RunProperties = RunProperties()) {
        self.text = text
        self.properties = properties
        self.drawing = nil
        self.rawXML = nil
        self.semantic = nil
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
            xml += "<w:t xml:space=\"preserve\">\(escapeXML(text))</w:t>"
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
    /// 轉換為 OOXML XML 字串
    func toXML() -> String {
        var parts: [String] = []

        // v0.19.3+ (#56 round 2 P0-1): rStyle MUST be first inside <w:rPr>
        // per ECMA-376 §17.3.2 CT_RPr child order. Hyperlinks rely on this
        // for the "Hyperlink" character style to apply correctly in Word.
        if let rStyle = rStyle {
            // v0.19.4+ (#56 R3-NEW-6): escape source-string before attribute
            // interpolation so a malicious `x"/><inj/>` payload cannot inject
            // sibling elements at write time.
            parts.append("<w:rStyle w:val=\"\(escapeXMLAttribute(rStyle))\"/>")
        }
        if bold {
            parts.append("<w:b/>")
        }
        if italic {
            parts.append("<w:i/>")
        }
        if let underline = underline {
            parts.append("<w:u w:val=\"\(underline.rawValue)\"/>")
        }
        if strikethrough {
            parts.append("<w:strike/>")
        }
        if let fontSize = fontSize {
            // OOXML 使用半點 (half-points)
            parts.append("<w:sz w:val=\"\(fontSize)\"/>")
            parts.append("<w:szCs w:val=\"\(fontSize)\"/>")  // 複雜文字大小
        }
        // v0.20.0+ (#60): emit `<w:rFonts>` with 4-axis preservation when
        // `rFonts` struct is set; fall back to legacy `fontName` (single value
        // mirrored to all 4 axes) when only the legacy field is set.
        if let rFonts = rFonts {
            var attrs: [String] = []
            if let ascii = rFonts.ascii { attrs.append("w:ascii=\"\(escapeXMLAttribute(ascii))\"") }
            if let hAnsi = rFonts.hAnsi { attrs.append("w:hAnsi=\"\(escapeXMLAttribute(hAnsi))\"") }
            if let eastAsia = rFonts.eastAsia { attrs.append("w:eastAsia=\"\(escapeXMLAttribute(eastAsia))\"") }
            if let cs = rFonts.cs { attrs.append("w:cs=\"\(escapeXMLAttribute(cs))\"") }
            if let hint = rFonts.hint { attrs.append("w:hint=\"\(escapeXMLAttribute(hint))\"") }
            if !attrs.isEmpty {
                parts.append("<w:rFonts \(attrs.joined(separator: " "))/>")
            }
        } else if let fontName = fontName {
            // v0.19.4+ (#56 R3-NEW-6 audit): fontName flows into 4 attributes;
            // escape once to avoid same injection sink as rStyle.
            let n = escapeXMLAttribute(fontName)
            parts.append("<w:rFonts w:ascii=\"\(n)\" w:hAnsi=\"\(n)\" w:eastAsia=\"\(n)\" w:cs=\"\(n)\"/>")
        }
        if let color = color {
            // v0.19.4+ (#56 R3-NEW-6 audit): color is a hex string in normal
            // use but the field is `String?` so escape defensively.
            parts.append("<w:color w:val=\"\(escapeXMLAttribute(color))\"/>")
        }
        if let highlight = highlight {
            parts.append("<w:highlight w:val=\"\(highlight.rawValue)\"/>")
        }
        if let verticalAlign = verticalAlign {
            parts.append("<w:vertAlign w:val=\"\(verticalAlign.rawValue)\"/>")
        }
        if let characterSpacing = characterSpacing {
            parts.append(characterSpacing.toXML())
        }
        if let textEffect = textEffect {
            parts.append(textEffect.toXML())
        }

        // v0.20.0+ (#60): emit new typed fields.
        if noProof {
            parts.append("<w:noProof/>")
        }
        if let kern = kern {
            parts.append("<w:kern w:val=\"\(kern)\"/>")
        }
        if let lang = lang {
            var attrs: [String] = []
            if let val = lang.val { attrs.append("w:val=\"\(escapeXMLAttribute(val))\"") }
            if let eastAsia = lang.eastAsia { attrs.append("w:eastAsia=\"\(escapeXMLAttribute(eastAsia))\"") }
            if let bidi = lang.bidi { attrs.append("w:bidi=\"\(escapeXMLAttribute(bidi))\"") }
            if !attrs.isEmpty {
                parts.append("<w:lang \(attrs.joined(separator: " "))/>")
            }
        }

        // v0.20.0+ (#60): replay rawChildren in source-document order, AFTER
        // typed children but BEFORE closing `</w:rPr>`. Matches `Run.rawElements`
        // architectural pattern (v0.14.0+, #52).
        if let rawChildren = rawChildren {
            for raw in rawChildren {
                parts.append(raw.xml)
            }
        }

        return parts.joined()
    }
}

