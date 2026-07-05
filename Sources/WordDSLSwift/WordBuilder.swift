// WordBuilder.swift
// word-aligned-state-sync Phase 4 task 5.3 — the result builder driving all
// container syntax of the `.mdocx` DSL (`mdocx-grammar` normative surface).
//
// v0.34 slice: WordDocument { Section { Paragraph { String / Run / atoms } } }.
// Hyperlink / Bookmark / Table containers activate in the 5.5 coverage
// iteration; their types exist (placeholders) but are not yet builder inputs.

/// Ordered inline content of a paragraph body. Plain `String` literals become
/// `.text` (the implicit unstyled-Run conversion mandated by "Flat Run with
/// implicit String literal inline grammar").
public enum InlineChild: Equatable, Sendable {
    case text(String)
    case run(Run)
    case tab
    case lineBreak
    case noBreakHyphen
}

@resultBuilder
public enum WordBuilder {

    // MARK: identity expressions (containers)

    public static func buildExpression(_ section: Section) -> Section { section }
    public static func buildExpression(_ paragraph: Paragraph) -> Paragraph { paragraph }

    // MARK: inline expressions (implicit String → unstyled Run per grammar)

    public static func buildExpression(_ text: String) -> [InlineChild] { [.text(text)] }
    public static func buildExpression(_ run: Run) -> [InlineChild] { [.run(run)] }
    public static func buildExpression(_ atom: Tab) -> [InlineChild] { _ = atom; return [.tab] }
    public static func buildExpression(_ atom: Break) -> [InlineChild] { _ = atom; return [.lineBreak] }
    public static func buildExpression(_ atom: NoBreakHyphen) -> [InlineChild] { _ = atom; return [.noBreakHyphen] }

    // MARK: blocks

    public static func buildBlock(_ sections: Section...) -> [Section] { sections }
    public static func buildBlock(_ paragraphs: Paragraph...) -> [Paragraph] { paragraphs }
    public static func buildBlock(_ inline: [InlineChild]...) -> [InlineChild] {
        inline.flatMap { $0 }
    }
}
