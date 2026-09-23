// Paragraph.swift
// word-aligned-state-sync Phase 4 task 5.3.

/// Paragraph result-builder container. Maps to OOXML `<w:p>`; the mandatory
/// `id:` maps to `w14:paraId` ("Mandatory explicit identifiers on structural
/// elements" — omitting `id:` is a compile error because no id-less
/// initializer exists). Style kinds go through `style:` (`mdocx-grammar`:
/// "No semantic shortcuts for OOXML-style attributes").
public struct Paragraph {
    public enum NumberingReference: Int, Equatable, Sendable {
        case numbered1 = 1
        case bulletA = 2
    }

    public let id: String
    public let style: WordStyle?
    public let numbering: NumberingReference?
    public let level: Int?
    public let children: [InlineChild]

    public init(id: String, style: WordStyle? = nil,
                numbering: NumberingReference? = nil, level: Int? = nil,
                @WordBuilder content: () -> [InlineChild]) {
        self.id = id
        self.style = style
        self.numbering = numbering
        self.level = level
        self.children = content()
    }

    public init(id: String, style: WordStyle? = nil,
                numbering: NumberingReference? = nil, level: Int? = nil) {
        self.id = id
        self.style = style
        self.numbering = numbering
        self.level = level
        self.children = []
    }
}
