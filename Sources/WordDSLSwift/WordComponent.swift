// WordComponent.swift
// word-aligned-state-sync Phase 4 task 5.5 — user-defined components
// (`mdocx-grammar`: "Component-aware op log via BeginComponent and
// EndComponent"). Shape locked by fixture 07: a component stores its `id`
// and a builder closure producing its body paragraph.
public protocol WordComponent {
    var id: String { get }
    var body: () -> Paragraph { get }
}
