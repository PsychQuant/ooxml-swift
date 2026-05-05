// Phase 7 placeholder.
// Full type design lives in `openspec/specs/mdocx-grammar/spec.md`
// (Spectra change `mdocx-syntax`). Implementation is the responsibility of
// `word-aligned-state-sync` Phase 7.

/// Inline no-break-hyphen atom. Maps to OOXML `<w:noBreakHyphen/>`.
/// Standalone child of a paragraph body, parallel to `Run` and `String`
/// (see Decision 2).
public struct NoBreakHyphen {
    public init() {}
}
