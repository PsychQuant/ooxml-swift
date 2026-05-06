# `_normalizer/` — vendored references for `MdocxFixtureNormalizer`

This directory holds reference XML files that `MdocxFixtureNormalizer` compares against during normalization. Files here are NOT fixtures — they are stripping baselines.

## `word-default-theme.xml`

The bytes that the normalizer recognises as "the canonical Word default theme." When a docx's `word/theme/theme1.xml` is byte-equal to this file, the normalizer strips it from the docx (treating it as identity-noise that should not factor into byte-diff comparison). When a docx's theme1.xml differs from this file, the normalizer preserves it (treating it as an intentional non-default theme).

### What "canonical" means here

This file is the bytes that **macdoc declares as the default theme**. It corresponds to the standard Office theme structure shipped by Microsoft Word since 2007 (clrScheme + fontScheme + fmtScheme + objectDefaults + extraClrSchemeLst), with the well-known Office colour palette (windowText/window dk1/lt1, 44546A/E7E6E6 dk2/lt2, accent1-6 5B9BD5 / ED7D31 / A5A5A5 / FFC000 / 4472C4 / 70AD47, hlink 0563C1, folHlink 954F72) and the Calibri Light / Calibri major/minor font scheme.

Different Word versions emit byte-different theme1.xml files even for the same "Office" theme (whitespace, attribute order, additional Word-version-specific elements). This file pins **one specific byte sequence** as the macdoc canonical default. The normalizer's stripping rule is **byte-exact**, not "structurally equivalent to Office theme" — that would require an XML parser pass on every comparison and would make the rule semantically slippery.

### How fixtures interact with this file

Hand-crafted fixture `<slug>.docx` files in this corpus are **minimal** — they typically OMIT `word/theme/theme1.xml` entirely. The normalizer's theme-stripping rule mostly serves Phase B execution: when a future `WordDSLSwift` implementation produces docx output that includes a theme matching this file, the normalizer strips it before byte-diff against the (theme-less) golden.

If a fixture intentionally exercises non-default themes, its docx will include a theme1.xml that diverges from this file; the normalizer preserves that intentional theme.

### Updating this file

This file is changed by:

1. Adding a Phase B test fixture that exercises theme-default detection.
2. Discovery during Phase B activation that Word-version drift breaks the byte-exact check.

In both cases, update with care: changing the bytes here can flip a fixture from "stripped" to "preserved" silently. Always re-run the full corpus after editing. Document the change in this README.

### Why not derive from a real Word save

Per `mdocx-fixture-corpus` design Decision 1 (Hand-crafted XML for golden docx, not Word-saved or ooxml-swift-emitted), we avoid Word-saved files as ground truth. This file is the same discipline applied to the normalizer's reference: a hand-curated, deterministic, version-controlled byte sequence we declare as the noise baseline, rather than a Word-version-specific snapshot.
