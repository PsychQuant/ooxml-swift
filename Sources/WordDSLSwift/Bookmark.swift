// Phase 7 placeholder.
// Full type design lives in `openspec/specs/mdocx-grammar/spec.md`
// (Spectra change `mdocx-syntax`). Implementation is the responsibility of
// `word-aligned-state-sync` Phase 7.

/// Bookmark result-builder container (default form). Maps to a paired
/// OOXML `<w:bookmarkStart>` / `<w:bookmarkEnd>` around the body content.
/// Cross-paragraph spans use the standalone `BookmarkStart` / `BookmarkEnd`
/// escape hatch instead of the container form.
public struct Bookmark {
    public let id: String
    public init(id: String) { self.id = id }
}
