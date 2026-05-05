// Phase 7 placeholder.
// Full type design lives in `openspec/specs/mdocx-grammar/spec.md`
// (Spectra change `mdocx-syntax`). Implementation is the responsibility of
// `word-aligned-state-sync` Phase 7.

/// Paragraph result-builder container. Maps to OOXML `<w:p>`. Style references
/// (heading levels, list items, etc.) are passed via `style:` parameter
/// (no `Heading1` / `Heading2` etc. wrapper types — see Decision 5).
public struct Paragraph {
    public let id: String
    public init(id: String) { self.id = id }
}
