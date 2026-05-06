// Phase 4 placeholder.
// Full type design lives in `openspec/specs/mdocx-grammar/spec.md`
// (Spectra change `mdocx-syntax`). Implementation is the responsibility of
// `word-aligned-state-sync` Phase 4 (Script transcoder).

/// Swift `@resultBuilder` for the DSL. Drives all container syntax
/// (`WordDocument { ... }`, `Section { ... }`, `Paragraph { ... }`,
/// `Hyperlink { ... }`, `Bookmark { ... }`, `Table` / `TableRow` /
/// `TableCell` block bodies, and `WordComponent.body`). Phase 4 fills in
/// `buildBlock`, `buildExpression`, `buildPartialBlock`, etc.
@resultBuilder
public enum WordBuilder {
    public static func buildBlock() {}
}
