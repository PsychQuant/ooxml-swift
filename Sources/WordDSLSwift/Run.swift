// Run.swift
// word-aligned-state-sync Phase 4 task 5.3.

/// Inline text-plus-formatting bundle. Maps to OOXML `<w:r>`. All inline
/// formatting is expressed via constructor parameters — the DSL provides no
/// `Bold(...)` / `Italic(...)` wrappers ("Flat Run with implicit String
/// literal inline grammar"). Plain `String` literals in a paragraph body
/// convert implicitly to unstyled runs.
public struct Run: Equatable, Sendable {
    public let text: String
    public let bold: Bool?
    public let italic: Bool?
    public let color: String?

    public init(_ text: String, bold: Bool? = nil, italic: Bool? = nil,
                color: String? = nil) {
        self.text = text
        self.bold = bold
        self.italic = italic
        self.color = color
    }
}
