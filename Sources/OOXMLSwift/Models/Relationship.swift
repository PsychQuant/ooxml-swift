import Foundation

// MARK: - Relationship

/// OOXML 關係定義（用於解析 .rels 檔案）
///
/// v0.19.5+ (#56 R5-CONT P1 #8): mutable + Equatable so per-container
/// `RelationshipsCollection` can live on `Header` / `Footer` / `Footnote` /
/// `Endnote` (which are themselves Equatable). `target` is var so URL
/// updates inside container hyperlinks can rewrite the rels Target without
/// rebuilding the whole Relationship.
public struct Relationship: Equatable {
    public let id: String           // 關係 ID (rId1, rId2, ...)
    public let type: RelationshipType
    public var target: String       // 目標路徑 (media/image1.png) or URL
    /// v0.19.5+ (#56 R5-CONT P1 #8): hyperlink relationships carry
    /// `TargetMode="External"`. Captured here so DocxWriter can re-emit
    /// the attribute byte-equivalently for hyperlink rels in
    /// `header*.xml.rels` / `footer*.xml.rels` / etc.
    public var targetMode: String?
    /// v0.19.5+ (#56 R5-CONT-2 P1 #6): preserve the raw Type attribute
    /// string so unknown vendor-extension types (VML / OLE / Word
    /// extension rels) round-trip byte-equivalent. Pre-fix `type` got
    /// `.unknown` (rawValue "") for any unrecognized type, and the writer
    /// emitted `Type=""` — invalid OOXML rels (Word strict mode rejects).
    /// Now: writer prefers `rawType` when non-nil over `type.rawValue`.
    /// Reader populates this from the source `Type` attribute regardless
    /// of recognition. Empty string when API-built (writer falls back to
    /// `type.rawValue`).
    public var rawType: String

    public init(id: String, type: RelationshipType, target: String, targetMode: String? = nil, rawType: String? = nil) {
        self.id = id
        self.type = type
        self.target = target
        self.targetMode = targetMode
        // Default rawType from the typed enum's rawValue; reader overrides
        // with the source's literal Type attribute via the explicit
        // parameter (preserves unknown types verbatim).
        self.rawType = rawType ?? type.rawValue
    }
}

// MARK: - Relationship Type

/// 關係類型
public enum RelationshipType: String {
    case image = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    case hyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    case styles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    case numbering = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
    case settings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    case header = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    case footer = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
    case footnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
    case endnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
    case comments = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    case theme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
    case fontTable = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
    case webSettings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
    case unknown = ""

    public init(rawValue: String) {
        switch rawValue {
        case RelationshipType.image.rawValue:
            self = .image
        case RelationshipType.hyperlink.rawValue:
            self = .hyperlink
        case RelationshipType.styles.rawValue:
            self = .styles
        case RelationshipType.numbering.rawValue:
            self = .numbering
        case RelationshipType.settings.rawValue:
            self = .settings
        case RelationshipType.header.rawValue:
            self = .header
        case RelationshipType.footer.rawValue:
            self = .footer
        case RelationshipType.footnotes.rawValue:
            self = .footnotes
        case RelationshipType.endnotes.rawValue:
            self = .endnotes
        case RelationshipType.comments.rawValue:
            self = .comments
        case RelationshipType.theme.rawValue:
            self = .theme
        case RelationshipType.fontTable.rawValue:
            self = .fontTable
        case RelationshipType.webSettings.rawValue:
            self = .webSettings
        default:
            self = .unknown
        }
    }
}

// MARK: - Relationships Collection

/// 關係集合（來自 .rels 檔案）
public struct RelationshipsCollection: Equatable {
    public var relationships: [Relationship] = []

    public init() {}

    /// 根據 ID 取得關係
    public func get(by id: String) -> Relationship? {
        return relationships.first { $0.id == id }
    }

    /// 取得所有圖片關係
    public var imageRelationships: [Relationship] {
        return relationships.filter { $0.type == .image }
    }

    /// 取得所有超連結關係
    public var hyperlinkRelationships: [Relationship] {
        return relationships.filter { $0.type == .hyperlink }
    }
}
