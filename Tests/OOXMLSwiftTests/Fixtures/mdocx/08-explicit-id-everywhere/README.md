# Fixture 08: `explicit-id-everywhere`

## Covers

`mdocx-grammar` Requirement **"Mandatory explicit identifiers on structural elements"** (Requirement 8).

## What this fixture pins

Every structural DSL element (Section, Paragraph, Bookmark, Hyperlink-as-anchor-target) carries an explicit `id:` parameter at the call site. The compiler refuses any source where an `id:` is omitted. Each `id:` value maps to the OOXML stable identifier (`w14:paraId` on paragraphs, `w:bookmarkId` + `w:name` on bookmarks).

## Edge case captured

Cross-referencing IDs: the Hyperlink targets `body_anchor` (an internal anchor); the Bookmark provides that anchor name. Reverse round-trip MUST recover both IDs verbatim — a regression that re-numbers IDs would break the cross-reference.

## What to inspect when this fixture fails

| Failure | Likely cause |
|---------|--------------|
| Phase B `docx byte-diff` | Implementation re-numbered IDs without updating cross-references; OR omitted `id:` requirement on one element type; OR wrote `id:` to wrong OOXML attribute |
