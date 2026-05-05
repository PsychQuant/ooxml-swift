// Phase 7 placeholder.
// Full type design lives in `openspec/specs/mdocx-grammar/spec.md`
// (Spectra change `mdocx-syntax`). Implementation is the responsibility of
// `word-aligned-state-sync` Phase 7.

/// Swift `@resultBuilder` for the DSL. Drives all container syntax
/// (`WordDocument { ... }`, `Section { ... }`, `Paragraph { ... }`,
/// `Hyperlink { ... }`, `Bookmark { ... }`, `Table` / `TableRow` /
/// `TableCell` block bodies, and `WordComponent.body`). Phase 7 fills in
/// `buildBlock`, `buildExpression`, `buildPartialBlock`, etc.
@resultBuilder
public enum WordBuilder {
    public static func buildBlock() {}
}
