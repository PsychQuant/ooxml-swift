# Fixture 13a: `bookmark-container`

## Covers

`mdocx-grammar` Requirement **"Bookmarks default to container with paired-marker escape hatch"** (Requirement 13), the **container form**.

Sister fixture: **13b-bookmark-cross-paragraph** covers the paired-marker escape hatch for cross-paragraph spans.

## What this fixture pins

`Bookmark(id: "name") { ... }` is the default form. Used for the common case where a bookmark spans contiguous content within one parent. The DSL wraps the body in `<w:bookmarkStart w:id="0" w:name="intro_text"/>` ... `<w:bookmarkEnd w:id="0"/>` markers. Start/end IDs are auto-paired by the implementation.

## Edge case captured

Bookmark wraps a single Run inside one paragraph — the simplest possible container span. Demonstrates that container form requires no escape hatch when start + end are in the same parent.

## What to inspect when this fixture fails

| Failure | Likely cause |
|---------|--------------|
| Phase B `docx byte-diff` | Mismatched `w:id` between bookmarkStart and bookmarkEnd (must pair); OR wrong `w:name` (must equal DSL `id:` value, not auto-renumbered); OR markers placed outside the run they wrap |
