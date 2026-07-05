// WordStyle.swift
// word-aligned-state-sync Phase 4 task 5.3 — typed style references
// (`mdocx-grammar`: "Style references via typed enum with
// define-on-first-use"). Raw strings are rejected at compile time because
// `style:` only accepts `WordStyle`.

import OOXMLSwift

/// A typed style reference. Predefined members mirror `WordStyleMap`
/// (`.heading1` ↔ styleId "Heading1"); authors define custom styles as
/// static extensions (`extension WordStyle { static let titleBrown = … }`).
public struct WordStyle: Equatable, Sendable {
    public let styleId: String
    public let font: String?
    public let fontSize: Int?
    public let color: String?
    public let bold: Bool?
    public let italic: Bool?

    public init(styleId: String, font: String? = nil, fontSize: Int? = nil,
                color: String? = nil, bold: Bool? = nil, italic: Bool? = nil) {
        self.styleId = styleId
        self.font = font
        self.fontSize = fontSize
        self.color = color
        self.bold = bold
        self.italic = italic
    }

    /// The `defineStyle` payload emitted on first use.
    var payload: StylePayload {
        StylePayload(styleId: styleId, name: nil, font: font,
                     fontSize: fontSize, color: color, bold: bold, italic: italic)
    }

    // Predefined members (kept in lockstep with OOXMLSwift.WordStyleMap).
    public static let heading1 = WordStyle(styleId: "Heading1")
    public static let heading2 = WordStyle(styleId: "Heading2")
    public static let heading3 = WordStyle(styleId: "Heading3")
    public static let heading4 = WordStyle(styleId: "Heading4")
    public static let heading5 = WordStyle(styleId: "Heading5")
    public static let heading6 = WordStyle(styleId: "Heading6")
    public static let heading7 = WordStyle(styleId: "Heading7")
    public static let heading8 = WordStyle(styleId: "Heading8")
    public static let heading9 = WordStyle(styleId: "Heading9")
    public static let quote = WordStyle(styleId: "Quote")
    public static let listItem = WordStyle(styleId: "ListItem")
    public static let title = WordStyle(styleId: "Title")
    public static let subtitle = WordStyle(styleId: "Subtitle")
}
