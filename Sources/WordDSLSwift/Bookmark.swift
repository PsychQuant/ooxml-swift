// Bookmark.swift
// mdocx-grammar: "Bookmarks default to container with paired-marker escape
// hatch". v0.34: DSL types compile (fixtures 13a/13b); op emission awaits
// the reducer's insertBookmark implementation — buildLog throws loudly
// (DSLEmissionError) rather than dropping the bookmark silently.

/// Container form: `Bookmark(id: "intro") { "text" }` inside a paragraph.
public struct Bookmark {
    public let id: String
    public let children: [InlineChild]
    public init(id: String, @WordBuilder content: () -> [InlineChild]) {
        self.id = id
        self.children = content()
    }
}

/// Cross-paragraph escape hatch (block-level start marker).
public struct BookmarkStart {
    public let id: String
    public init(id: String) { self.id = id }
}

/// Cross-paragraph escape hatch (block-level end marker).
public struct BookmarkEnd {
    public let id: String
    public init(id: String) { self.id = id }
}
