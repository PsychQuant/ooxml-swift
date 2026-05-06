// Phase 4 placeholder.
// Full type design lives in `openspec/specs/mdocx-grammar/spec.md`
// (Spectra change `mdocx-syntax`). Implementation is the responsibility of
// `word-aligned-state-sync` Phase 4 (Script transcoder).

/// Table result-builder container — three-layer structure mirrors OOXML.
/// `Table` (`<w:tbl>`) contains `TableRow` (`<w:tr>`) which contains
/// `TableCell` (`<w:tc>`).
public struct Table {
    public let id: String
    public init(id: String) { self.id = id }
}
