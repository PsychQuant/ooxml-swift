// Phase 4 placeholder.
// Full type design lives in `openspec/specs/mdocx-grammar/spec.md`
// (Spectra change `mdocx-syntax`). Implementation is the responsibility of
// `word-aligned-state-sync` Phase 4 (Script transcoder).

/// Protocol for user-defined DSL components. Each `WordComponent` instance
/// emits a paired `BeginComponent` / `EndComponent` op-log envelope around
/// the operations produced by its body so reverse direction can reconstruct
/// the call site (see Decision 7).
public protocol WordComponent {
    var id: String { get }
}
