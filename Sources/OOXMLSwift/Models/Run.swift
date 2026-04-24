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

    public init(text: String, properties: RunProperties = RunProperties()) {
        self.text = text
        self.properties = properties
        self.drawing = nil
        self.rawXML = nil
        self.semantic = nil
    }
}

// MARK: - Run Properties

/// Run 格式屬性
public struct RunProperties: Equatable {
    public var bold: Bool = false
    public var italic: Bool = false
    public var underline: UnderlineType?
    public var strikethrough: Bool = false
    public var fontSize: Int?              // 半點 (24 = 12pt)
    public var fontName: String?
    public var color: String?              // RGB hex (e.g., "FF0000")
    public var highlight: HighlightColor?
    public var verticalAlign: VerticalAlign?
    public var characterSpacing: CharacterSpacing?  // 字元間距
    public var textEffect: TextEffect?              // 文字效果
    public var rawXML: String?                      // 原始 XML（用於進階功能如 SDT）

    public init() {}

    public init(bold: Bool = false,
         italic: Bool = false,
         underline: UnderlineType? = nil,
         strikethrough: Bool = false,
         fontSize: Int? = nil,
         fontName: String? = nil,
         color: String? = nil,
         highlight: HighlightColor? = nil,
         verticalAlign: VerticalAlign? = nil) {
        self.bold = bold
        self.italic = italic
        self.underline = underline
        self.strikethrough = strikethrough
        self.fontSize = fontSize
        self.fontName = fontName
        self.color = color
        self.highlight = highlight
        self.verticalAlign = verticalAlign
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
        if let fontName = fontName {
            parts.append("<w:rFonts w:ascii=\"\(fontName)\" w:hAnsi=\"\(fontName)\" w:eastAsia=\"\(fontName)\" w:cs=\"\(fontName)\"/>")
        }
        if let color = color {
            parts.append("<w:color w:val=\"\(color)\"/>")
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

        return parts.joined()
    }
}
