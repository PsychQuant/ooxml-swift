// Phase 4 placeholder.
// Full type design lives in `openspec/specs/mdocx-grammar/spec.md`
// (Spectra change `mdocx-syntax`). Implementation is the responsibility of
// `word-aligned-state-sync` Phase 4 (Script transcoder).

/// Inline line-break atom. Maps to OOXML `<w:br/>`. Standalone child of a
/// paragraph body, parallel to `Run` and `String` (see Decision 2).
public struct Break {
    public init() {}
}
