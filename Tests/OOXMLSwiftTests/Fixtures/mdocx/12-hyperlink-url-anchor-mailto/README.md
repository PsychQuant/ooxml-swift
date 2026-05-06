# Fixture 12: `hyperlink-url-anchor-mailto`

## Covers

`mdocx-grammar` Requirement **"Hyperlinks are containers with target enum"** (Requirement 12).

## What this fixture pins

`Hyperlink(to: HyperlinkTarget)` with three case variants:

- `.url(String)` → `<w:hyperlink r:id="rId10">` + relationship of type `hyperlink` to `https://example.com` (TargetMode=External) in `word/_rels/document.xml.rels`.
- `.anchor(String)` → `<w:hyperlink w:anchor="ch1_intro">` (no rels entry — internal anchor reference).
- `.mailto(String)` → `<w:hyperlink r:id="rId11">` + relationship to `mailto:hello@example.com` (TargetMode=External).

## Edge case captured

All three target cases in one paragraph — proves the enum branches all serialize correctly side-by-side, and that internal-anchor links don't pollute the relationships file.

## What to inspect when this fixture fails

| Failure | Likely cause |
|---------|--------------|
| Phase B `docx byte-diff` | Implementation merged URL + mailto into one rels entry (incorrect — they are distinct relationships); OR put internal anchor in rels (incorrect — anchors are intra-doc); OR wrong `r:id` prefix |
