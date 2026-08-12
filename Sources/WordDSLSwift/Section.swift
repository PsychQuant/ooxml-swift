// Section.swift
// word-aligned-state-sync Phase 4 task 5.3.

/// Section container in the DSL (`mdocx-grammar`: "Section as DSL container
/// with compile-time marker inversion"). v0.34 slice: single-section body of
/// paragraphs; section-property serialization (`<w:sectPr>` marker inversion,
/// `type:` parameter) activates with multi-section support in 5.5.
public struct Section {
    public enum SectionType: String, Equatable, Sendable {
        case continuous
        case nextPage
        case nextColumn
        case evenPage
        case oddPage
    }

    public let id: String
    public let type: SectionType?
    public let children: [SectionChild]

    public init(id: String, type: SectionType? = nil,
                @SectionBuilder content: () -> [SectionChild]) {
        self.id = id
        self.type = type
        self.children = content()
    }

    public init(id: String, type: SectionType? = nil) {
        self.id = id
        self.type = type
        self.children = []
    }
}
