// Section.swift
// word-aligned-state-sync Phase 4 task 5.3.

/// Section container in the DSL (`mdocx-grammar`: "Section as DSL container
/// with compile-time marker inversion"). v0.34 slice: single-section body of
/// paragraphs; section-property serialization (`<w:sectPr>` marker inversion,
/// `type:` parameter) activates with multi-section support in 5.5.
public struct Section {
    public let id: String
    public let paragraphs: [Paragraph]

    public init(id: String, @WordBuilder content: () -> [Paragraph]) {
        self.id = id
        self.paragraphs = content()
    }

    public init(id: String) {
        self.id = id
        self.paragraphs = []
    }
}
