// Phase 4 placeholder.
// Full type design lives in `openspec/specs/mdocx-grammar/spec.md`
// (Spectra change `mdocx-syntax`). Implementation is the responsibility of
// `word-aligned-state-sync` Phase 4 (Script transcoder).

/// Section container in the DSL. Compiler inverts container syntax into the
/// OOXML `<w:sectPr>` marker pattern at serialization time
/// (see Decision 6 in design.md).
public struct Section {
    public let id: String
    public init(id: String) { self.id = id }
}
