// Phase 4 placeholder.
// Full type design lives in `openspec/specs/mdocx-grammar/spec.md`
// (Spectra change `mdocx-syntax`). Implementation is the responsibility of
// `word-aligned-state-sync` Phase 4 (Script transcoder).

/// Table cell result-builder container. Maps to OOXML `<w:tc>`. Children
/// are block-level content (typically `Paragraph` instances).
public struct TableCell {
    public let id: String
    public init(id: String) { self.id = id }
}
