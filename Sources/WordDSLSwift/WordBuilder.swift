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
public enum InlineChild {
    case text(String)
    case run(Run)
    case tab
    case lineBreak
    case noBreakHyphen
    case bookmark(Bookmark)
    case hyperlink(Hyperlink)
}

/// Ordered block-level content of a section body.
public enum SectionChild {
    case paragraph(Paragraph)
    case component(type: String, id: String, body: [Paragraph])
    case table(Table)
    case bookmarkStart(BookmarkStart)
    case bookmarkEnd(BookmarkEnd)
}

@resultBuilder
public enum WordBuilder {

    // MARK: identity expressions (containers)

    public static func buildExpression(_ section: Section) -> Section { section }
    public static func buildExpression(_ paragraph: Paragraph) -> Paragraph { paragraph }
    public static func buildExpression(_ row: TableRow) -> TableRow { row }
    public static func buildExpression(_ cell: TableCell) -> TableCell { cell }

    // MARK: inline expressions (implicit String → unstyled Run per grammar)

    public static func buildExpression(_ text: String) -> [InlineChild] { [.text(text)] }
    public static func buildExpression(_ run: Run) -> [InlineChild] { [.run(run)] }
    public static func buildExpression(_ atom: Tab) -> [InlineChild] { _ = atom; return [.tab] }
    public static func buildExpression(_ atom: Break) -> [InlineChild] { _ = atom; return [.lineBreak] }
    public static func buildExpression(_ atom: NoBreakHyphen) -> [InlineChild] { _ = atom; return [.noBreakHyphen] }
    public static func buildExpression(_ bookmark: Bookmark) -> [InlineChild] { [.bookmark(bookmark)] }
    public static func buildExpression(_ hyperlink: Hyperlink) -> [InlineChild] { [.hyperlink(hyperlink)] }

    // MARK: blocks

    public static func buildBlock(_ sections: Section...) -> [Section] { sections }
    public static func buildBlock(_ paragraphs: Paragraph...) -> [Paragraph] { paragraphs }
    /// Single-paragraph component body (`() -> Paragraph`, fixture 07 shape).
    public static func buildBlock(_ paragraph: Paragraph) -> Paragraph { paragraph }
    public static func buildBlock(_ rows: TableRow...) -> [TableRow] { rows }
    public static func buildBlock(_ cells: TableCell...) -> [TableCell] { cells }
    public static func buildBlock(_ inline: [InlineChild]...) -> [InlineChild] {
        inline.flatMap { $0 }
    }
}


/// Result builder for the HETEROGENEOUS Section body (paragraphs, tables,
/// components, bookmark markers). Split from `WordBuilder` because
/// `Paragraph` statements appear in two contexts with different collection
/// types (Section body → `[SectionChild]`, TableCell body → `[Paragraph]`);
/// a single builder with overloaded `buildExpression` is ambiguous under
/// the result-builder transform. The DSL surface is unchanged — builder
/// attributes are invisible in `.mdocx` source.
@resultBuilder
public enum SectionBuilder {
    public static func buildExpression(_ paragraph: Paragraph) -> [SectionChild] {
        [.paragraph(paragraph)]
    }
    public static func buildExpression<C: WordComponent>(_ component: C) -> [SectionChild] {
        [.component(type: String(describing: C.self), id: component.id,
                    body: [component.body()])]
    }
    public static func buildExpression(_ table: Table) -> [SectionChild] { [.table(table)] }
    public static func buildExpression(_ marker: BookmarkStart) -> [SectionChild] {
        [.bookmarkStart(marker)]
    }
    public static func buildExpression(_ marker: BookmarkEnd) -> [SectionChild] {
        [.bookmarkEnd(marker)]
    }
    public static func buildBlock(_ children: [SectionChild]...) -> [SectionChild] {
        children.flatMap { $0 }
    }
}
