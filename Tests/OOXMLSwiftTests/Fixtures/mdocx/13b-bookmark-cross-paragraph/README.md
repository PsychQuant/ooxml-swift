# Fixture 13b: `bookmark-cross-paragraph`

## Covers

`mdocx-grammar` Requirement **"Bookmarks default to container with paired-marker escape hatch"** (Requirement 13), the **paired-marker escape hatch** for cross-paragraph spans.

Sister fixture: **13a-bookmark-container** covers the container form for single-element spans.

## What this fixture pins

When a bookmark spans across paragraphs (the start and end markers can't share a `<w:p>` parent), the DSL provides standalone `BookmarkStart(id:)` and `BookmarkEnd(id:)` as siblings of paragraphs. The matching `id:` values pair them; the compiler emits a diagnostic if a `BookmarkStart` lacks a matching `BookmarkEnd` (or vice versa).

## Edge case captured

Bookmark span covers TWO paragraphs (`p1` + `p2`). `BookmarkStart` placed before paragraph 1 in the section body; `BookmarkEnd` placed after paragraph 2. Both share `id: "ch1_span"`.

The OOXML output places `<w:bookmarkStart>` at the start of paragraph 1's content (or as direct `<w:body>` child) and `<w:bookmarkEnd>` at the end of paragraph 2's content (or as direct `<w:body>` child). Either placement is OOXML-valid; this fixture's golden uses inline placement (start inside p1, end inside p2 after the run).

## What to inspect when this fixture fails

| Failure | Likely cause |
|---------|--------------|
| Phase B `docx byte-diff` | Implementation auto-paired wrong start/end (id values not respected); OR placed both markers in the same paragraph (collapsed cross-paragraph span); OR omitted one marker entirely |
| Compile error in Phase A | `BookmarkStart(id: "X")` without matching `BookmarkEnd(id: "X")` — that's enforced by the spec; if Phase A reports it, the test source has a real mismatch |
