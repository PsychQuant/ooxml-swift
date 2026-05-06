# Fixture 06: `two-sections-marker-inversion`

## Covers

`mdocx-grammar` Requirement **"Section as DSL container with compile-time marker inversion"** (Requirement 6, the deliberate exception to the OOXML-mirror naming rule per Decision 6).

## What this fixture pins

Two sequential `Section { ... }` containers at the DSL level invert into OOXML's marker pattern:

- First section: `<w:sectPr w:type="continuous"/>` lives inside the last paragraph's `<w:pPr>`.
- Second section (last): `<w:sectPr w:type="nextPage"/>` lives as a direct child of `<w:body>` after the last paragraph.

The DSL preserves human-readable container syntax (`Section { ... }`) while the writer transforms it into the marker pattern OOXML mandates.

## Edge case captured

Two sections, two different `type:` values (`.continuous` and `.nextPage`), demonstrating that section-properties travel through the inversion intact. The reverse transcoder reads the inverted OOXML back into two `Section` blocks (per the spec's "reverse direction reconstructs container syntax" scenario).

## What to inspect when this fixture fails

| Failure | Likely cause |
|---------|--------------|
| Phase B `docx byte-diff` | First section's `<w:sectPr>` placed wrong (must be inside the last paragraph's `<w:pPr>`); OR last section's `<w:sectPr>` placed inside a paragraph instead of direct `<w:body>` child; OR `w:type` attribute lost during inversion |
