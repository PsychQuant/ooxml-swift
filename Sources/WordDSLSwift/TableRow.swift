// Phase 7 placeholder.
// Full type design lives in `openspec/specs/mdocx-grammar/spec.md`
// (Spectra change `mdocx-syntax`). Implementation is the responsibility of
// `word-aligned-state-sync` Phase 7.

/// Table row result-builder container. Maps to OOXML `<w:tr>`. Children
/// are `TableCell` instances.
public struct TableRow {
    public let id: String
    public init(id: String) { self.id = id }
}
