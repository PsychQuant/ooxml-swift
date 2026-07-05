// Hyperlink.swift
// mdocx-grammar: "Hyperlinks are containers with target enum". v0.34: DSL
// types compile (fixture 12); op emission awaits an authoring-op channel
// (the hyperlink lowering path lives in EditAlgebra, not yet reachable from
// buildLog) — buildLog throws loudly rather than dropping the link.

public enum HyperlinkTarget: Equatable, Sendable {
    case url(String)
    case anchor(String)
    case mailto(String)
}

public struct Hyperlink {
    public let target: HyperlinkTarget
    public let children: [InlineChild]
    public init(to target: HyperlinkTarget, @WordBuilder content: () -> [InlineChild]) {
        self.target = target
        self.children = content()
    }
}
