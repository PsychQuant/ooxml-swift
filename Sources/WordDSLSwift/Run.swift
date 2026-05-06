// Phase 4 placeholder.
// Full type design lives in `openspec/specs/mdocx-grammar/spec.md`
// (Spectra change `mdocx-syntax`). Implementation is the responsibility of
// `word-aligned-state-sync` Phase 4 (Script transcoder).

/// Inline text-plus-formatting bundle. Maps to OOXML `<w:r>`. All formatting
/// (bold, italics, color, etc.) is expressed via `Run` constructor parameters
/// (no `Bold(...)` / `Italic(...)` wrapper types — see Decision 5). Plain
/// `String` literals inside a paragraph body are implicitly converted to
/// unstyled `Run` instances.
public struct Run {
    public let text: String
    public init(_ text: String) { self.text = text }
}
