// Phase 4 placeholder.
// Full type design lives in `openspec/specs/mdocx-grammar/spec.md`
// (Spectra change `mdocx-syntax`). Implementation is the responsibility of
// `word-aligned-state-sync` Phase 4 (Script transcoder).

/// Hyperlink result-builder container. Maps to OOXML `<w:hyperlink>`.
/// Target is discriminated by `HyperlinkTarget` enum (`.url` / `.anchor` /
/// `.mailto`). Body contains inline content.
public struct Hyperlink {
    public init() {}
}
