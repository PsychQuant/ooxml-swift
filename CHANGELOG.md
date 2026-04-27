# Changelog

All notable changes to ooxml-swift will be documented in this file.

## Skipped versions

- **v0.19.4** (never tagged) — The R3 stack-completion content originally targeted v0.19.4. After the round-3 fix landed, the round-4 6-AI verify (https://github.com/PsychQuant/che-word-mcp/issues/56#issuecomment-4321562429) returned BLOCK with 6 new P0 + 7 P1 findings (walker-asymmetry follow-ups, `position == 0` sentinel collision, attribute-escape sweep gap, block-level SDT typed-Revision propagation, container-symmetric `replaceText`, container `<w:tbl>` parser drop). v0.19.4 was held back. v0.19.5 ships the R3 stack content (preserved verbatim below) **plus** the R5 stack-completion fixes (6 P0 + 5 P1, additive — no breaking change versus the v0.19.4 contract). No v0.19.4 git tag, no v0.19.4 GitHub Release.

## [0.20.2] - 2026-04-27

### Added — Sub-stack D of paragraph-level content-equality (closes #65)

`ParagraphProperties` now extracts and round-trips the `<w:rPr>` direct child of `<w:pPr>` (paragraph-mark formatting per ECMA-376 §17.3.1.27 CT_PPrBase) — the rPr that controls pilcrow ¶ glyph appearance (font, size, color, language tag, kerning).

#### What was lost pre-fix

`parseParagraphProperties` only extracted typed `<w:pPr>` direct children (pStyle, jc, spacing, ind, numPr). The nested `<w:rPr>` was silently dropped at parse time — accounting for ~50% of the residual `<w:lang>` loss in the NTPU thesis fixture round-trip.

#### How it's fixed

Reuse `parseRunProperties(from:)` verbatim. The schema is identical to run-level CT_RPr, so all of sub-stack C's typed extraction (`RFontsProperties` 4-axis, `<w:noProof>`, `<w:kern>`, `LanguageProperties` 3-axis) and raw passthrough (`rawChildren` for `w14:*` effects like `<w14:textOutline>`, `<w14:glow>`, `<w14:textFill>`) come for free. Zero schema duplication.

```swift
// New field on ParagraphProperties (Models/Paragraph.swift)
public var markRunProperties: RunProperties?

// Parser extension (IO/DocxReader.swift parseParagraphProperties)
if let markRPr = element.elements(forName: "w:rPr").first {
    props.markRunProperties = parseRunProperties(from: markRPr)
}

// Writer emits inside <w:pPr>...</w:pPr> with empty-gate discipline
if let markProps = markRunProperties, !markProps.toXML().isEmpty {
    parts.append("<w:rPr>\(markProps.toXML())</w:rPr>")
}
```

#### Measured impact (NTPU thesis fixture, post-D)

| Preservation class | Pre-D | Post-D | Improvement |
|---|---|---|---|
| `<w:lang ` retention | 50% | **98.89%** | +48.89 pp |
| `<w:rFonts>` retention | 88% | 98.77% | +10.77 pp |
| `<w:noProof>` retention | 92% | 100% | +8 pp |
| `<w:kern>` retention | 84% | 99.93% | +15.93 pp |
| `document.xml` size loss | 16.66% | **10.95%** | -5.71 pp |

#### Matrix-pin floor ratchets

`testDocumentContentEqualityInvariant` ratchets in lockstep:
- `<w:lang ` floor: 0.45 → **0.95**
- `<w:rFonts` floor: 0.85 → **0.95**
- `<w:noProof` floor: 0.90 → **0.95**
- `<w:kern ` floor: 0.80 → **0.95**
- `sizeLossRatio` ceiling: 0.175 → **0.12**
- `w14:` floor unchanged (0.04) — sub-stack E (#66) ratchets to 0.95

Any future regression in run-level OR paragraph-level RunProperties handling now fails the matrix-pin.

#### Tests added (4)

- `testParagraphMarkRunPropertiesPreservedThroughRoundtrip` — payload-parity for `<w:lang>` 3-axis with structural assertion that emission stays inside `<w:pPr>`
- `testParagraphMarkRFontsFourAxisPreservedThroughRoundtrip` — 4-axis font preservation for pilcrow CJK glyph
- `testParagraphMarkW14NamespaceEffectsPreservedAsRawChildren` — `<w14:textOutline>` raw-children passthrough in pPr context
- `testParagraphWithoutMarkRunPropertiesEmitsNoRPr` — negative test with full-pPr range-slicing assertion (catches synthetic empty `<w:rPr>` after typed children)

Suite: 682 → 686 tests pass / 0 failures / 1 skipped.

#### Architecture context

Sub-stack D of the `che-word-mcp-paragraph-level-content-equality` Spectra change (bundles #65 + #66). The cross-cutting matrix-pin established in sub-stack C is ratcheted, not duplicated — the architectural principle "if not typed, preserve as raw" extends from run-level (sub-stack C) to paragraph-mark level (sub-stack D). Sub-stack E (#66) will extend to paragraph w14:paraId/textId attributes, completing the path to `< 5%` round-trip loss.

#### Backward compatibility

`markRunProperties` is optional (default nil). All pre-existing callers continue to work — paragraphs without source pPr/rPr emit no synthetic empty wrappers thanks to the writer's `!inner.isEmpty` gate.

## [0.20.1] - 2026-04-27

### Fixed — Sub-stack C-CONT of #58/#59/#60: trim `recognizedRprChildren` to actually-extracted set

The sub-stack C 6-AI verify (run on v0.20.0) returned mixed verdicts:
- R1 PASS (no warnings)
- R2 PASS-WITH-WARNINGS (P2 same finding)
- R5 PASS-WITH-WARNINGS but escalated finding to P0
- Codex BLOCK on the same P0 + 3 NEW P1

**Triple-confirmed P0** (R2 + R5 + Codex independently): `recognizedRprChildren` Set in `parseRunProperties` listed ~16+ rPr child kinds as "recognized" but parseRunProperties had NO extraction for them. Result: silent drop on read because they neither become typed fields NOR get captured into `rawChildren`.

**Affected elements** (all very common in real-world Word documents):
- `<w:spacing>` (character spacing — typeset documents)
- `<w:caps>` / `<w:smallCaps>` (small-caps formatting)
- `<w:position>` (vertical position offset)
- `<w:shd>` (run-level shading / highlighting)
- `<w:bdr>` (run-level border)
- `<w:em>` (CJK emphasis marks)
- `<w:effect>` (text effects: shimmer, blink, etc.)
- `<w:vanish>` / `<w:specVanish>` / `<w:webHidden>` (visibility flags)
- `<w:outline>` / `<w:shadow>` / `<w:emboss>` / `<w:imprint>` (legacy text effects)
- `<w:snapToGrid>` / `<w:fitText>` (layout flags)
- `<w:rtl>` (right-to-left direction)
- `<w:bCs>` / `<w:iCs>` / `<w:dstrike>` (complex-script + double-strikethrough variants)

#### Fix

Trimmed `recognizedRprChildren` to ONLY actually-typed-extracted-or-emitted kinds: `rStyle, b, i, u, strike, sz, szCs, rFonts, color, highlight, vertAlign, noProof, kern, lang, rPrChange`. Everything else falls through to `rawChildren` and round-trips byte-equivalent via the writer's rawChildren replay.

`szCs` retained in set because the writer typed-emits it via `fontSize` (Run.swift:259) — including in rawChildren would cause double emission. `rPrChange` retained because typed-handled at the run level (parseRPrChangeFromRunInline at DocxReader.swift:1532+).

#### Round-trip size impact (additional improvement)

Thesis fixture `document.xml`:
- Pre-fix v0.19.x: 32% loss
- Sub-stack C v0.20.0: 17.75% loss (improvement of 14.25 pp from typed-rPr extraction)
- **Sub-stack C-CONT v0.20.1: 16.66% loss** (additional 1.09 pp from rawChildren capture of previously-silent-dropped elements)

Matrix-pin `testDocumentContentEqualityInvariant` floor tightened from 0.19 → 0.175 to reflect new baseline. Future paragraph-mark rPr fix (out-of-scope) should drop loss to < 5%.

### Deferred (sub-stack C-CONT MAY-tier — Codex P1, separate follow-up SDD)

- **Schema-order rawChildren tail-append** (Codex P1) — `rawChildren` are tail-appended after typed children, but ECMA-376 CT_RPr has schema-order constraints. `<w:b/><w14:textOutline/><w:i/>` becomes `<w:b/><w:i/>...<w14:textOutline/>`. Word tolerates; schema-strict validators may flag. Requires bigger refactor (preserve child-event list with source order). Tracked.
- **`characterSpacing` / `textEffect` parser-side gap** — typed fields exist on `RunProperties` (Run.swift:115-116) and are typed-emitted by toXML, but parseRunProperties has NO extraction. Source `<w:spacing>` / `<w:effect>` now correctly fall through to rawChildren (post-trim), but the typed setters are no-ops for source-loaded docs. Add typed extraction OR remove the typed fields. Tracked.
- **`eastAsianLayout` / `oMath`** in rawChildren — fall through correctly post-trim. Schema-order concern same as above.
- **Static `recognizedRprChildren` Set** (Codex P2) — currently constructed per parseRunProperties call; converting to `static let` would eliminate hot-path allocation. Performance optimization, no correctness impact.
- **Ratio-floor maintenance** (Codex P1) — current matrix-pin floors are calibrated to baseline; future ratchets need explicit follow-up. Add a `// TODO: ratchet floor when paragraph-mark rPr lands` comment per floor + spec follow-up SDD.

### Methodology lesson (6th refinement)

R2 found this as P2 ("inline comment is false; pre-existing parity gap, not regression"). R5 escalated to P0 by recognizing the affected elements are common (`<w:caps>`, `<w:spacing>`, etc.). Codex confirmed the P0 with code trace + identified 3 additional P1 concerns.

The methodology pattern: **a P2-graded finding from one reviewer can become P0 when another reviewer applies real-world impact lens**. Severity-grading is a function of (a) bug presence + (b) blast radius. Use 6-AI verify's diversity to surface the maximum blast radius for each bug class.

### Spectra change

Ships sub-stack C-CONT of `che-word-mcp-issue-58-59-60-document-content-preservation`. After this hotfix, sub-stack C's #60 closure is verified clean for the ACTUAL field-loss audit scope. Out-of-scope items (paragraph-mark rPr + w14:paraId/textId + Codex P1s) tracked as separate follow-up SDD.

## [0.20.0] - 2026-04-27

### Added — Sub-stack C of #58/#59/#60 (closes #60 RunProperties field-loss audit)

Sub-stack C is the architectural completion of the "if not typed, preserve as raw" principle that started in sub-stack A (#58 BodyChild) and continued in sub-stack B (#59 WhitespaceOverlay). This release adds typed RunProperties fields for rFonts (4-axis), noProof, kern, and lang — plus a generic `rawChildren` passthrough for unrecognized direct rPr children (e.g., `<w14:textOutline>`, `<w14:textFill>`, `<w14:glow>`). The matrix-pin gains preservation-class-3 assertions making it LOAD-BEARING for any future RunProperties regression.

#### #60 root cause

`RunProperties.fontName: String?` collapsed the 4-axis `<w:rFonts w:ascii=".." w:hAnsi=".." w:eastAsia=".." w:cs="..">` into a single value. ECMA-376 §17.3.2 RPrBase distinguishes Latin (`w:ascii`), High-ANSI (`w:hAnsi`), East-Asian (`w:eastAsia`), and Complex Script (`w:cs`) font assignments because different scripts may need different fonts (e.g., Times New Roman for Latin + DFKai-SB for traditional Chinese eastAsia + Mangal for Devanagari cs). Pre-fix: parser captured ascii into fontName; writer emitted all 4 axes with that single value. Round-trip silently replaced eastAsia/cs fonts with the ascii value.

Plus `<w:noProof/>`, `<w:kern w:val="32"/>`, `<w:lang w:val="..">` (3-axis), and w14:* effects (`<w14:textOutline>`, `<w14:textFill>`, `<w14:glow>`, etc.) were silently dropped on read because parseRunProperties had no extraction case for them.

#### Fix

**New typed structs** in `Run.swift`:
- `RFontsProperties` — 4 axes (ascii / hAnsi / eastAsia / cs) + hint
- `LanguageProperties` — 3 axes (val / eastAsia / bidi)

**RunProperties extensions**:
- `var rFonts: RFontsProperties?` — when set, takes precedence over legacy `fontName`
- `var noProof: Bool = false`
- `var kern: Int?`
- `var lang: LanguageProperties?`
- `var rawChildren: [RawElement]?` — unrecognized direct rPr children (matches `Run.rawElements` pattern from v0.14.0/#52)

**Backward compatibility**: legacy `fontName: String?` retained. When `rFonts` is nil and `fontName` is set, writer emits 4-axis with same value (current behavior). When `rFonts` is set, writer emits per-axis values. parseRunProperties mirrors `rFonts.ascii → fontName` for legacy callers.

**parseRunProperties** in `DocxReader.swift:2228` extended with extraction for the new fields plus a `recognizedRprChildren` Set covering 30+ typed rPr kinds; collects unrecognized direct rPr children into `rawChildren`.

**RunProperties.toXML()** emits new typed fields in ECMA-376 source order, then replays `rawChildren` after typed children but before closing `</w:rPr>`.

#### Matrix-pin extension (§3.9 — LOAD-BEARING)

`testDocumentContentEqualityInvariant` extended with preservation-class-3 ratio-floor assertions for `<w:rFonts>` (0.85), `<w:noProof>` (0.90), `<w:lang>` (0.45), `<w:kern>` (0.80), `w14:*` (0.04). Floors calibrated to current measured baseline; ANY regression in run-level rPr preservation trips the matrix-pin.

#### Out-of-scope (revealed by matrix-pin, separate follow-up)

The matrix-pin uncovered two pre-existing bugs that are NOT in #60 scope:

1. **`ParagraphProperties` lacks `markRunProperties` field** — the `<w:pPr><w:rPr>...</w:rPr></w:pPr>` (paragraph-mark formatting controlling pilcrow appearance) is silently dropped at parse time. Accounts for ~50% of `<w:lang>` loss in thesis fixture round-trip.
2. **`Paragraph` parser doesn't preserve `w14:paraId`/`w14:textId`** attributes on `<w:p>` (Word's revision-tracking GUIDs). Accounts for ~95% of w14:* token loss (2214 of 2359 tokens are these two attributes).

Both tracked as follow-up SDD. The ratio-floor assertions stay load-bearing for sub-stack C scope while not blocking on these out-of-scope drops.

#### Round-trip size impact

Thesis fixture `document.xml`:
- Pre-fix (v0.19.x): 1473896 → 1006805 bytes (32% loss)
- Post-sub-stack-C: 1473896 → 1212279 bytes (17.75% loss — improvement of 14.25 percentage points)
- Future paragraph-mark rPr fix (out-of-scope) should drop loss to < 5%

### Tests

3 new tests in `Tests/OOXMLSwiftTests/Issue58_60ContentPreservationTests.swift`:

- `testRFontsFourAxisPreservedThroughRoundtrip` (§3.1 — 4-axis preservation)
- `testNoProofAndKernPreservedThroughRoundtrip` (§3.2 — typed extraction for noProof + kern)
- `testW14NamespaceEffectsPreservedAsRawChildren` (§3.3 — w14:* via rawChildren passthrough)

Plus `testDocumentContentEqualityInvariant` matrix-pin extended with §3.9 + §3.11 (preservation-class-3 ratio floors + size sanity check).

Suite total: 682 tests pass / 1 skipped / 0 failures (679 sub-stack B-CONT-2-CONT baseline + 3 new sub-stack C tests).

### API additions (v0.20.0, additive — no breaking change vs v0.19.13)

- `public struct RFontsProperties: Equatable` (4 axes + hint)
- `public struct LanguageProperties: Equatable` (3 axes)
- `public var RunProperties.rFonts: RFontsProperties?`
- `public var RunProperties.noProof: Bool`
- `public var RunProperties.kern: Int?`
- `public var RunProperties.lang: LanguageProperties?`
- `public var RunProperties.rawChildren: [RawElement]?`

Legacy `RunProperties.fontName: String?` kept and behavior preserved for callers that don't use the new `rFonts` field.

### Spectra change

This release ships sub-stack C of `che-word-mcp-issue-58-59-60-document-content-preservation`. Closes #60 (RunProperties field-loss audit) and the cross-cutting matrix-pin (`testDocumentContentEqualityInvariant`). The architectural completion of the "if not typed, preserve as raw" principle.

## [0.19.13] - 2026-04-27

### CRITICAL HOTFIX — Sub-stack B-CONT-2-CONT: revert TIER-0 over-fix that broke `<w:del>` round-trip

The sub-stack B-CONT-2 6-AI verify (run on v0.19.12) returned BLOCK with **R2 + R5 INDEPENDENTLY confirming** a critical content-loss bug introduced by v0.19.12's TIER-0 fix. v0.19.12 silently strips `<w:del>` deleted-text content on every round-trip — affects ALL tracked-change documents with deletions.

#### Bug analysis

v0.19.12 added `"delText"` to `parseRun`'s `recognizedRunChildren` Set to fix R5's prior P0-1 (delTextCounter desync via 2x advance). This was mechanically correct for the counter desync but broke a load-bearing invariant in the writer:

- Pre-v0.19.12: parseRun captured `<w:delText>` into `Run.rawElements`. Writer's gate at `Paragraph.swift:787` (`!run.text.isEmpty || (run.rawElements?.isEmpty ?? true)`) evaluated `false || false` → SKIP synthetic emission. Then `for raw in rawElements { xml += raw.xml }` emitted `<w:delText>content</w:delText>` verbatim. ONE emission, content preserved. (R5's prior P0-2 was correctly falsified for this state.)
- v0.19.12 (BROKEN): added "delText" to recognizedRunChildren → rawElements stayed empty. Writer's gate evaluated `false || true` → TRUE → emit synthetic `<w:delText xml:space="preserve">{run.text}</w:delText>` where run.text="" (parseRun's `<w:t>` loop never sees delText). Output: empty `<w:delText></w:delText>` with content destroyed.

The §2.33 test only counted opening tags (1=1 pre/post), and §2.34 only checked in-memory `Revision.originalText` (populated by the explicit `<w:del>` loop, independent of run.text). Both passed falsely.

#### Fix

Reverted "delText" from `recognizedRunChildren` (back to `["rPr", "t", "drawing", "oMath", "oMathPara"]`). Added `includeDelText: Bool = true` parameter to `advanceWhitespaceCounter(forSkippedXML:)`. parseRun's rawElements loop passes `includeDelText: false` when `localName == "delText"` — the explicit `<w:del>` loop already advances delTextCounter for each delText, so this prevents the double-advance without removing delText from rawElements (which the writer needs).

#### Methodology lesson (5th refinement)

Sub-stack A: matrix-pin needs symmetric assertions across container variants.
Sub-stack B: design ≠ fixtures with real content.
Sub-stack B-CONT: real-world OOXML content classes must be IN fixtures.
Sub-stack B-CONT-2: when adding a counter-advance helper at "all" raw-capture sites, audit ALL `xmlString` references INCLUDING parseRun's own `rawElements` path.
**Sub-stack B-CONT-2-CONT (now)**: when fixing a counter-desync via "skip the element from raw-capture path", verify the WRITER still has access to the element's content via SOMEWHERE — load-bearing invariants in protective gates can be silently broken by upstream changes. Tests must assert end-to-end content preservation, not opening-tag counts or in-memory state.

The §2.33 test was retained from B-CONT-2 as a regression guard for the writer-gate invariant — it should have caught the regression but missed it because the assertion was too narrow. New test `testDelTextContentPreservedThroughRoundTrip` (B-CONT-2-CONT) asserts the actual deleted-text content survives round-trip.

### Tests

1 new test in `Tests/OOXMLSwiftTests/Issue58_60ContentPreservationTests.swift`:

- `testDelTextContentPreservedThroughRoundTrip` (B-CONT-2-CONT — content-preservation guard, not just opening-tag count)

Suite total: 679 tests pass / 1 skipped / 0 failures (678 sub-stack B-CONT-2 baseline + 1 new content guard).

### Severity

**v0.19.12 must NOT be used in production**. Affects all `<w:del>` round-trips. v0.19.13 closes the regression.

### Spectra change

Ships sub-stack B-CONT-2-CONT of `che-word-mcp-issue-58-59-60-document-content-preservation`. Sub-stack C (#60 RunProperties audit) ships next as v0.20.0 + v3.14.0.

## [0.19.12] - 2026-04-27

### Fixed — Sub-stack B-CONT-2 of #58/#59/#60: close delText counter desync + 5 missed raw-capture sites

The sub-stack B-CONT 6-AI verify ([#59 comment 4324076688](https://github.com/PsychQuant/che-word-mcp/issues/59#issuecomment-4324076688)) returned BLOCK with 4-reviewer convergence on:

#### B-CONT-2 TIER-0 — `<w:delText>` counter desync (R5 finding, partial)

**R5's prediction (P0-1 confirmed)**: parseRun's `recognizedRunChildren = ["rPr", "t", "drawing", "oMath", "oMathPara"]` did NOT include `"delText"`. When parseRun was called for a `<w:r>` inside `<w:del>`:
1. Explicit delText loop at `DocxReader.swift:970-993` advanced `delTextCounter` by 1 per delText
2. parseRun's rawElements loop at line 1849-1865 ALSO captured delText into `Run.rawElements`, AND called `advanceWhitespaceCounter` → advanced `delTextCounter` AGAIN

Result: `delTextCounter = 2N` instead of `N`. Every subsequent whitespace `<w:delText>` query landed at wrong index → silent loss for documents with multiple `<w:del>` blocks.

**R5's prediction (P0-2 falsified by code trace)**: R5 also predicted writer-side duplicate emission (`<w:del>` containing `"abc"` → writer producing `<w:delText>abc</w:delText><w:delText>abc</w:delText>`). Test §2.33 confirmed this DOESN'T happen — writer's gate at `Paragraph.swift:787` (`!run.text.isEmpty || (run.rawElements?.isEmpty ?? true)`) skips the explicit `<w:delText>` emission when rawElements covers it. Devil's Advocate found a real bug (P0-1) but mis-graded the severity (P0-2 was false). Test §2.33 retained as regression guard for the writer-gate invariant.

**Fix**: added `"delText"` to `recognizedRunChildren` Set at `DocxReader.swift:1847`. parseRun's rawElements loop now skips delText (already captured by explicit loop). Test §2.34 (`testDeleteTextCounterStaysSyncedAcrossMultipleDels`) GREEN.

#### B-CONT-2 TIER-1 — 5+ missed raw-capture counter-desync sites

B-CONT instrumented 7 raw-capture sites; sub-stack B-CONT verify (R2 + Codex) found 5 missed siblings:

- `parseContainerChildBodyChildren` raw fallback (Codex P0): unrecognized container body-level children with inner `<w:t>` desynced counter
- `parseHyperlink` rawChildren branch (R2 P0): hyperlinks with nested non-`<w:r>` children (e.g., `<w:fldSimple>`)
- `parseFieldSimple` non-`<w:r>` silent skip (R2 P0): also independent content-loss bug; minimum fix: counter advance
- `parseParagraph` `case "smartTag"` / `"customXml"` / `"dir"` / `"bdo"` raw-carriers (R2 P0): all four typed raw-carrier blocks

**Fix**: added `Self.advanceWhitespaceCounter(forSkippedXML: ...)` call at each missed site (5 sites covering 8 cases counting the 4 paragraph raw-carrier branches). 3 representative tests (§2.36-§2.38) cover container-raw-fallback + hyperlink-raw-children + smartTag classes.

### Tests

5 new tests in `Tests/OOXMLSwiftTests/Issue58_60ContentPreservationTests.swift`:

- `testDelTextEmittedExactlyOncePerSourceElement` (B-CONT-2 TIER-0 — R5 P0-2 regression guard for writer-gate)
- `testDeleteTextCounterStaysSyncedAcrossMultipleDels` (B-CONT-2 TIER-0 — R5 P0-1 actual bug)
- `testWhitespaceOverlayContainerRawFallbackDoesNotDesyncCounter` (B-CONT-2 TIER-1 — Codex P0)
- `testWhitespaceOverlayHyperlinkRawChildrenDoesNotDesyncCounter` (B-CONT-2 TIER-1 — R2 P0)
- `testWhitespaceOverlaySmartTagDoesNotDesyncCounter` (B-CONT-2 TIER-1 — R2 P0, representative)

Plus helper `countDelTextElements(in:)` for §2.33's writer-output verification.

Suite total: 678 tests pass / 1 skipped / 0 failures (673 sub-stack B-CONT baseline + 5 B-CONT-2 new tests).

### Methodology lesson (4th refinement, partially confirmed)

R5's "Devil's Advocate worst-case" prediction was 50% accurate on this round: P0-1 (counter desync) was a real bug; P0-2 (writer-side duplicate emission) was a false alarm caught by the writer-gate invariant. **Methodology refinement**: adversarial reviewers can correctly identify a bug class but mis-grade severity by missing protective gates elsewhere in the codebase. Verify-cycle response should TEST predictions (not assume them) — this saved us from a misframed BLOCK and added a regression guard for the writer-gate behavior.

The actual recurring pattern remains: each sub-cycle compresses the prior cycle's blind spot. B-CONT instrumented 7 raw-capture sites; B-CONT-2 found 5+ siblings of the same class. Long-term fix is the central raw-capture helper (§2.43, deferred) but matrix-pin fixture upgrades (§2.44, deferred) and sub-stack C content-equality matrix-pin extensions (§2.45/§2.46, deferred to sub-stack C scope) should catch future regressions of this class.

### Deferred (B-CONT-2 TIER-2, MAY-tier)

- §2.43 Central raw-capture helper refactor — high-value but adds touchpoint risk; future additions still require manual call. Tracked.
- §2.44 `buildAllPartsWhitespaceFixture` upgrade with real-world content classes — sterile fixture remains; per-test class coverage suffices. Long-term consolidation tracked.
- §2.45 / §2.46 Container-part + delText parity in matrix-pin — sub-stack C scope addition.
- Sub-stack B-CONT MAY-tier: static state concurrency hazard (R5 + Codex P1, deferred), single-quoted `xml:space='preserve'` (R5 P2, Word doesn't emit), perf gate (Codex P2, tracked).

### API additions (v0.19.12, additive — no breaking change vs v0.19.11)

No new public API. Only internal change: `recognizedRunChildren` includes `"delText"`.

### Spectra change

Ships sub-stack B-CONT-2 of `che-word-mcp-issue-58-59-60-document-content-preservation`. Sub-stack C (#60 RunProperties audit) ships next as v0.20.0 + v3.14.0.

## [0.19.11] - 2026-04-27

### Fixed — Sub-stack B-CONT of #58/#59/#60: close 4 P0 + 3 P1 from sub-stack B 6-AI verify

The sub-stack B 6-AI verify ([#59 comment 4323956207](https://github.com/PsychQuant/che-word-mcp/issues/59#issuecomment-4323956207)) returned BLOCK with 4-reviewer convergence on a P0 counter-desync class (R2 Logic + R5 Devil's Advocate + Codex). Two root causes converge to the same observable bug (recovered whitespace lands on wrong element OR is silently lost):

#### B-CONT P0 root cause A — prefix-match collision (R2 + R5 + Codex)

`WhitespaceOverlay.swift:54`'s `xml.range(of: "<w:t", ...)` was a prefix match. It also fired on `<w:tab>`, `<w:tabs>`, `<w:tbl>`, `<w:tblPr>`, `<w:tblGrid>`, `<w:tblW>`, `<w:tc>`, `<w:tcPr>`, `<w:tcW>`, `<w:tr>`, `<w:trPr>`, `<w:trHeight>`, `<w:tblBorders>`, `<w:tcBorders>`, `<w:tblCellMar>`, `<w:tblLayout>`, `<w:tblLook>`, `<w:tblStyle>`, etc. The DOM walker `element.elements(forName: "w:t")` is exact-match. Counter desynced immediately in any document with tables or tabs (basically every real Word file, including the thesis fixture).

R5's empirical probe: `<w:tab/> + <w:t xml:space="preserve">     </w:t> + <w:t>after</w:t>` → overlay records whitespace at index 2; parseRun queries index 1 (DOM doesn't see `<w:tab/>` as `<w:t>`) → nil → whitespace LOST.

**Fix**: tag-name boundary check after matching `<w:t` — only count when next char is `>`, ` `, `\t`, `\n`, `\r`, or `/`. Same boundary check applied to the new `countWtElements` and `countDelTextElements` helpers.

#### B-CONT P0 root cause B — skipped raw subtrees (Codex + R2)

When a parsed structure is stored as raw XML (parser doesn't descend into it), the byte scanner still counts `<w:t>` elements inside but `parseRun` never visits them — counter desyncs per skipped subtree. Affected raw-capture sites:

- `parseAlternateContent` skips `<mc:Choice>` branch (`<mc:Fallback>` is the only branch parsed)
- `parseInsRevisionWrapper` raw-captures `<w:ins>/<w:del>/<w:moveFrom>/<w:moveTo>` wrappers with non-run children (via `hasNonRunChild` check)
- `parseBodyChildren` `.rawBlockElement` capture (sub-stack A's catch-all)
- `parseParagraph` unrecognized-child catch-all
- `parseRun` `rawElements` capture for unknown direct `<w:r>` children (e.g., nested `<mc:AlternateContent>`)

**Fix**: parser-side counter advance via new `DocxReader.advanceWhitespaceCounter(forSkippedXML:)` helper. At each raw-capture site, count `<w:t>` (and `<w:delText>`) elements in the skipped subtree's xmlString and advance both counters accordingly. Keeps scanner's source-order index in sync with parser's actual visit count.

#### B-CONT P0 secondary — pathological skip-over (R2 + R5)

Pre-fix: when prefix-match falsely fired on `<w:tbl>`, scanner searched forward for `</w:t>` and consumed the next legitimate one, swallowing real `<w:t>` elements between false-match and consumed-close. Disappears automatically once boundary check (root-cause-A fix) lands.

#### B-CONT P1 — §2.7 matrix-pin landing (R1 + Codex)

The `<w:t>` total-character parity assertion in `testDocumentContentEqualityInvariant` was a placeholder comment in sub-stack B (tasks.md §2.7 was checked done despite the assertion never landing — surfaced by R1 + Codex). New helper `sumWtElementCharCount(in:)` walks `<w:t>` elements with same boundary check as scanner, sums inner-text length. Matrix-pin now asserts equality against thesis fixture — catches future overlay regressions before 6-AI verify.

#### B-CONT P1 — `<w:delText>` overlay coverage (R5)

`<w:delText xml:space="preserve">[whitespace]</w:delText>` was permanently lost on read because (a) overlay only scanned `<w:t`, (b) parseRun's delText loop at `DocxReader.swift:970` read `delText.stringValue` directly with no overlay consult.

**Fix**: extended `WhitespaceOverlay` with second scanner pass for `<w:delText` (mirror of `<w:t>` scan with same boundary + xml:space + decoded-whitespace logic). Added `delTextWhitespaceByIndex` map + `delText(forElementSequenceIndex:)` accessor. Added `WhitespaceParseContext.delTextCounter`. Updated parseRun's delText loop to consult overlay when stringValue.isEmpty. Extended `advanceWhitespaceCounter(forSkippedXML:)` to also advance delTextCounter.

#### B-CONT P1 — comments trimming destroyed recovered whitespace (Codex)

`parseComments` at `DocxReader.swift:2978` called `text.trimmingCharacters(in: .whitespacesAndNewlines)` — destroyed recovered overlay text for any whitespace-only comment AND silently stripped meaningful leading/trailing whitespace from regular comments.

**Fix**: removed the trim. Safe because the XPath walk only reads `<w:t>` inner content; never includes incidental XML pretty-printing whitespace between sibling tags.

#### B-CONT P1 — entity-encoded whitespace not recognized (R5)

Scanner's `inner.allSatisfy({ $0.isWhitespace })` ran on RAW XML bytes — `&#x09;&#x09;` (two tabs) sees `&`, `#`, `x`, `0`, `9` which aren't `Character.isWhitespace`, so the element wasn't stored. Foundation later decoded the entities then stripped → permanent loss.

**Fix**: new `WhitespaceOverlay.decodeXMLEntities(in:)` helper handling numeric decimal (`&#9;`), hex (`&#x09;`, `&#xA0;`), and named (`&nbsp;`) entities. Modified main + delText scanners to decode `innerText` before whitespace check. Stored value is the decoded text — parseRun consult returns proper characters.

### Tests

7 new tests in `Tests/OOXMLSwiftTests/Issue58_60ContentPreservationTests.swift`:

- `testWhitespaceOverlayPrefixMatchTabDoesNotDesyncCounter` (B-CONT P0 root-cause-A — `<w:tab/>` adjacent)
- `testWhitespaceOverlayPrefixMatchTableDoesNotDesyncCounter` (B-CONT P0 root-cause-A — empty-cell table; covers pathological skip-over too)
- `testWhitespaceOverlayMcAlternateContentDoesNotDesyncCounter` (B-CONT P0 root-cause-B — Choice/Fallback both counted but only Fallback parsed)
- `testWhitespaceOverlayInsRevisionWrapperDoesNotDesyncCounter` (B-CONT P0 root-cause-B — raw-captured `<w:ins>` with `<w:bookmarkStart>`)
- `testWhitespaceOnlyCommentPreservedNotTrimmed` (B-CONT P1 Codex — comment trim fix)
- `testEntityEncodedWhitespacePreserved` (B-CONT P1 R5 — `&#x09;&#x09;` decode)
- `testDeleteTextWhitespaceRoundTrips` (B-CONT P1 R5 — `<w:delText>` overlay coverage)

Plus `testDocumentContentEqualityInvariant` extended with §2.23 `<w:t>` total-character parity matrix-pin.

Suite total: 673 tests pass / 1 skipped / 0 failures (666 sub-stack B baseline + 7 B-CONT new tests).

### Methodology lesson

Sub-stack A taught: matrix-pin needs symmetric assertions baked in from design (across container variants). Sub-stack B taught: matrix-pin baked in from design ≠ matrix-pin fixtures with real content (sterile fixtures hid every P0). Sub-stack B-CONT confirms: real-world OOXML content classes (tables, alternate-content, revision wrappers, entity-encoded characters) must be IN the fixtures, not separate test files. Each sub-cycle compresses the prior cycle's blind spot into a tighter design discipline.

### Deferred (B-CONT MAY-tier — P1/P2 not closed)

- Static state concurrency hazard (R5 + Codex P1) — `currentWhitespaceContext` is unsynchronized process-wide state. Documented constraint; works only because DocxReader is single-threaded by convention. Larger refactor (~30-line change) to thread context as parameter through 11 parseRun call sites. Tracked for follow-up SDD.
- Single-quoted `xml:space='preserve'` (R5 P2) — not emitted by Word's serializer; documented as accepted limitation.
- Performance gate (Codex P2) — no fixture benchmark for byte-scan cost. Tracked.

### API additions (v0.19.11, additive — no breaking change vs v0.19.10)

All `internal`. No public API surface change.

- `WhitespaceOverlay.delText(forElementSequenceIndex:)`
- `WhitespaceOverlay.countWtElements(in:)`
- `WhitespaceOverlay.countDelTextElements(in:)`
- `WhitespaceOverlay.decodeXMLEntities(in:)`
- `DocxReader.advanceWhitespaceCounter(forSkippedXML:)`
- `DocxReader.WhitespaceParseContext.delTextCounter`

### Spectra change

This release ships sub-stack B-CONT of `che-word-mcp-issue-58-59-60-document-content-preservation`. Sub-stack C (#60 RunProperties audit) ships next as v0.20.0 + v3.14.0.

## [0.19.10] - 2026-04-27

### Fixed — Sub-stack B of #58/#59/#60: WhitespaceOverlay for Foundation XMLDocument parser limitation

Closes [PsychQuant/che-word-mcp#59](https://github.com/PsychQuant/che-word-mcp/issues/59) — Foundation `XMLDocument` strips whitespace-only `<w:t xml:space="preserve">[whitespace]</w:t>` text node `stringValue` to "" regardless of the `xml:space` attribute AND regardless of `XMLNode.Options.nodePreserveWhitespace` parse option. This is a structural limitation of Foundation's libxml2-backed parser on macOS, not a configuration bug — verified by isolated probe in [#59 diagnosis](https://github.com/PsychQuant/che-word-mcp/issues/59).

The probe on the thesis fixture confirmed: source has 346 whitespace-only `<w:t>` elements (683 chars total); pre-fix Reader recovered 190 (349 chars). 334 chars silently lost on read alone — exactly matching the issue's reported round-trip loss.

#### Architectural approach: pre-parse byte-stream overlay (NOT parser swap)

`WhitespaceOverlay` (new type at `Sources/OOXMLSwift/IO/WhitespaceOverlay.swift`) does a pre-parse byte-stream scan over raw OOXML XML bytes. For each `<w:t xml:space="preserve">[whitespace]</w:t>` element encountered in DOM document order, it records the whitespace content keyed by element sequence index. `parseRun` (and `parseComments`) consult the overlay when `t.stringValue.isEmpty` to recover the lost whitespace bytes.

**Why not switch parsers**: 1-2 weeks of work + new dependency + affects all 10 `XMLDocument(data:)` call sites in DocxReader.swift. Whitespace overlay is contained, surgical, and follows the same architectural pattern as `WordDocument.modifiedParts` overlay (the v0.13.0 byte-preservation architecture).

#### Per-part WhitespaceParseContext

Each of the 6 `<w:t>`-bearing parts (`document.xml`, `header*.xml`, `footer*.xml`, `footnotes.xml`, `endnotes.xml`, `comments.xml`) gets its own `WhitespaceParseContext` (overlay + monotonic per-`<w:t>` counter). `DocxReader.withWhitespaceContext(_:_:)` sets the active context for the duration of a part-parse via static state + defer-cleanup. This avoids threading `inout` parameters through 8 `parseRun` call sites.

`parseRun` consults the active context for each `<w:t>` element. If the context is non-nil and `t.stringValue` is empty, the overlay's recovered text replaces it; if non-empty, the original is used; counter advances either way. `parseComments` does its own XPath walk over `<w:t>` nodes (doesn't go through `parseRun`), so it has the same overlay-consult logic inline.

#### Methodology lesson — comprehensive matrix-pin from design

Sub-stack A's 4 sub-cycles (A → A-CONT → A-CONT-2 → A-CONT-3) demonstrated that matrix-pins added reactively to verify findings always lag the bug by one round. Sub-stack B's matrix-pin test (`testWhitespacePreservedAcrossAllSixPartTypes`) exercises ALL 6 part types in a single fixture from the start — body + header1 + footer1 + footnotes + endnotes + comments — so the convergence cycle is shorter from design.

Pre-implementation: 6 RED assertions in 1 fixture. Post-implementation: 6 GREEN assertions. The next-round verify can't surface symmetric-sibling regressions because all 6 are exercised by the same test.

### Tests

- 2 new tests in `Tests/OOXMLSwiftTests/Issue58_60ContentPreservationTests.swift`:
  - `testWhitespaceOnlyTextRunsRoundTripInBody` (#59 P0 — body-level whitespace recovery isolated)
  - `testWhitespacePreservedAcrossAllSixPartTypes` (#59 cross-part matrix-pin — all 6 part types in one fixture)
- Suite total: 666 tests pass / 1 skipped / 0 failures (664 sub-stack A baseline + 2 sub-stack B new tests)

### API additions (v0.19.10, additive — no breaking change vs v0.19.9 contract)

- `internal struct WhitespaceOverlay` (new file `Sources/OOXMLSwift/IO/WhitespaceOverlay.swift`)
- `internal final class DocxReader.WhitespaceParseContext`
- `internal static var DocxReader.currentWhitespaceContext: WhitespaceParseContext?`
- `internal static func DocxReader.withWhitespaceContext<T>(_:_:)`
- All `internal` — no public API surface change.

### Spectra change

This release ships sub-stack B of `che-word-mcp-issue-58-59-60-document-content-preservation`. Sub-stack C (#60 RunProperties audit) ships next as v0.20.0 + v3.14.0. Sub-stack A's deferred A-CONT-4 follow-ups (paragraph-level container delete state-inconsistency + body SDT recursion asymmetry + insertBookmark perf) are tracked but out of scope for sub-stack B/C.

## [0.19.9] - 2026-04-27

### Fixed — Sub-stack A-CONT-3 of #58 (correctness regression + API symmetry from A-CONT-2 verify)

The sub-stack A-CONT-2 6-AI verify ([report](https://github.com/PsychQuant/che-word-mcp/issues/58#issuecomment-4323715199)) returned BLOCK with 3 P0 + 2 P1 + 4 P2 (3 of 4 reviewers concur — R2 Logic + R5 Devil's Advocate + Codex; R1 Requirements PASS). Maintainer authorized MUST + SHOULD tier scope (3 P0); P1 + P2 deferred.

This is sub-cycle 4 for #58 (A → A-CONT → A-CONT-2 → A-CONT-3). Same trajectory as R5 → R5-CONT-4 (5 sub-cycles for #56).

#### A-CONT-3 P0 #1 — `deleteBookmark` dirty-key path mismatch (silent correctness regression)

`Document.swift:2067, 2073` did `modifiedParts.insert(headers[i].fileName)` — inserting BASENAME (`"header1.xml"`). `Header.fileName` returns BASENAME per `Header.swift:193`. The writer's overlay-mode dirty-gate at `DocxWriter.swift:141` checks `dirty.contains("word/\(header.fileName)")` — looks for FULL PATH (`"word/header1.xml"`). **The format mismatch meant the writer's overlay-mode SKIPPED re-emitting the modified header — the deletion succeeded in-memory but never persisted to disk.** Same bug for footers. Footnotes/endnotes paths used the correct `"word/footnotes.xml"` / `"word/endnotes.xml"` constants.

Triple-confirmed by R2 + R5 + Codex. R2 grep confirmed every other Document.swift callsite uses the correct `"word/\(headers[i].fileName)"` form (lines 464, 475, 1116, 1125, 1192, 1212, 1228, 1245, 1264, 1280) — A-CONT-2's new code was the lone exception.

Fix: 2-line change to use `"word/\(headers[i].fileName)"` and `"word/\(footers[i].fileName)"`. Test `testDeleteBookmarkInHeaderPersistsToDisk` proves the deletion now reaches disk after roundtrip.

#### A-CONT-3 P0 #2 — `getBookmarks()` skipped paragraph-level container bookmarks (UX regression)

A-CONT-2's `collectBodyLevelBookmarkNamesRecursive` deliberately skipped `.paragraph` cases (its job was body-level markers). For container parts, only that helper was called — paragraph-level bookmarks inside container paragraphs (`Paragraph.bookmarks`) were never surfaced to MCP `list_bookmarks`. Identical UX bug to original #58. Paragraph-level bookmarks in headers are MORE common than body-level — the A-CONT-2 closure delivered the LESS common case.

Fix: new `collectAllBookmarksFromContainer` helper handles both `.paragraph(let para)` (walking `para.bookmarks`) AND `.bookmarkMarker` AND recurses into `.contentControl(_, let inner)`. Replaces `collectBodyLevelBookmarkNamesRecursive` for container paths in `getBookmarks()`. Body iteration still uses the prior structure for paragraph-index semantics. Test `testGetBookmarksSurfacesContainerParagraphLevelBookmarks` proves coverage.

#### A-CONT-3 P0 #3 — `insertBookmark` cross-part duplicate detection

`Document.insertBookmark` at `Document.swift:1977-1994` only walked `body.children` for duplicate detection. After A-CONT-2, a TOC anchor named `_Toc12345` living in a header survived `insertBookmark(name: "_Toc12345")` because the scan missed it — silently produced a duplicate-named bookmark, breaking the global-name uniqueness invariant.

Fix: replace the body-only loop with `Set(getBookmarks().map { $0.name })` lookup. Reuses the now-comprehensive `getBookmarks()` walker from P0 #2 — symmetric scope across getBookmarks/deleteBookmark/insertBookmark. Test `testInsertBookmarkDuplicateNameInContainerThrows` proves the symmetry.

### Tests

- 3 new tests in `Tests/OOXMLSwiftTests/Issue58_60ContentPreservationTests.swift`:
  - `testDeleteBookmarkInHeaderPersistsToDisk` (A-CONT-3 P0 #1 — proves deletion reaches disk)
  - `testGetBookmarksSurfacesContainerParagraphLevelBookmarks` (A-CONT-3 P0 #2)
  - `testInsertBookmarkDuplicateNameInContainerThrows` (A-CONT-3 P0 #3)
- Suite total: 664 tests pass / 1 skipped / 0 failures (661 A-CONT-2 baseline + 3 A-CONT-3 new tests)

### Deferred (out of A-CONT-3 scope)

Per maintainer scope decision:
- A-CONT-2 P1 #4: comments.xml coverage in getBookmarks
- A-CONT-2 P1 #5: matrix-pin negative-arm test (proves no false-pass)
- A-CONT-2 P2 #6: cross-part bookmark span end-marker orphan
- A-CONT-2 P2 #7: `paragraphIndex = -1` sentinel doc note
- A-CONT-2 P2 #9: dedicated tests for new deleteBookmark container paths

### API additions (v0.19.9, additive — no breaking change vs v0.19.8 contract)

- No new public types or signatures. `getBookmarks()` and `insertBookmark()` keep their existing call signatures; the new behavior is strictly additive (returns more, throws more on duplicates).

### Spectra change

This release ships sub-stack A-CONT-3 mini-cycle. Re-numbers planned sub-stack B → v0.19.10 + v3.13.10 (sub-stack C unchanged at v0.20.0 + v3.14.0). Sub-stack A took 4 sub-cycles (A + A-CONT + A-CONT-2 + A-CONT-3) to drain #58 — same trajectory shape as R5 → R5-CONT-4 needing 5 sub-cycles to drain #56. Each round catches what the prior matrix-pin couldn't see — methodology working as designed.

## [0.19.8] - 2026-04-27

### Fixed — Sub-stack A-CONT-2 of #58 (API-layer + SDT-recursion + matrix-pin-fixture mini-mini-cycle from A-CONT verify)

The sub-stack A-CONT 6-AI verify ([report](https://github.com/PsychQuant/che-word-mcp/issues/58#issuecomment-4323658377)) returned BLOCK with 2 P0 + 1 P1 + 1 P2 (3 of 4 reviewers concur — R2 Logic + R5 Devil's Advocate + Codex; R1 Requirements PASS). All three BLOCKs converged on the same 2 findings; R2 alone caught a third (matrix-pin regression-blindness on chosen fixture).

This is sub-cycle 3 for #58 (A → A-CONT → A-CONT-2). R5-CONT-4 took 5 sub-cycles to drain #56; same convergence-cycle pattern at work. Each round catches what the prior matrix-pin couldn't see — the methodology working as designed.

#### A-CONT-2 P0 #1 — `Document.getBookmarks()` walks container `bodyChildren`

Pre-A-CONT-2 `getBookmarks()` ([Document.swift:2122-2153](https://github.com/PsychQuant/ooxml-swift/blob/v0.19.8/Sources/OOXMLSwift/Models/Document.swift#L2122)) iterated only `for child in body.children`. Headers, footers, footnotes, endnotes were never traversed despite the A-CONT CHANGELOG claim of "body + headers + footers + footnotes + endnotes" coverage. A thesis-style document with TOC anchor `<w:bookmarkStart w:name="_Toc12345"/>` at body level inside `header1.xml` round-tripped preserved on disk (A-CONT P0 #1 fix) but was invisible to MCP `list_bookmarks` — same observable symptom as the original #58 P0, just in containers instead of body.

A-CONT-2 extends `getBookmarks()` to walk container `bodyChildren` across headers + footers + footnotes + endnotes. New `collectBodyLevelBookmarkNamesRecursive` helper recurses into block-level `.contentControl(_, let inner)` so SDT-nested markers are also surfaced. Container markers carry `paragraphIndex = -1` sentinel (no paragraph index in body-document sense). Removed the stale comment referencing a `getAllBookmarks()` follow-up helper that didn't exist anywhere in the codebase.

#### A-CONT-2 P0 #2 — `parseContainerChildBodyChildren` SDT recursion

Pre-A-CONT-2 the container parser handled 5 cases (`p`, `tbl`, `sectPr`, `bookmarkStart`, `bookmarkEnd`) + raw default. `parseBodyChildren` had 6 (added `sdt`). The missing `case "sdt":` in the container parser meant block-level SDTs in headers / footers / footnotes / endnotes fell through to `.rawBlockElement` — XML byte-preserved for round-trip ✓ but: (a) bookmarks inside the SDT were NOT surfaced as typed BodyChild entries; (b) `nextBookmarkId` calibration walker explicitly skipped `.rawBlockElement` so SDT-nested bookmark ids were invisible → potential id collision; (c) tables/paragraphs inside the SDT were invisible to typed-model walkers.

A-CONT-2 adds the `case "sdt":` branch mirroring `parseBodyChildren:644-679`: parses SDT metadata via `SDTParser.parseSdtPr`, recursively calls `parseContainerChildBodyChildren` for `<w:sdtContent>` children, appends `.contentControl(metadata, children: sdtChildren)`. The existing `collectBodyLevelBookmarkIds` calibration walker (DocxReader.swift:409-420) already recursed through `.contentControl(_, let inner)` from sub-stack A — so once the parser surfaces the typed `.contentControl`, calibration picks up nested ids correctly.

#### A-CONT-2 P0 #3 — Matrix-pin synthetic-fixture coverage

Pre-A-CONT-2 the `assertContainerBookmarkStartParity` matrix-pin extension was regression-blind on the thesis fixture: R2 Logic verified all 12 container parts (`word/header1.xml` through `header6.xml`, `word/footer1.xml` through `footer4.xml`, `word/footnotes.xml`, `word/endnotes.xml`) have **zero** `<w:bookmarkStart>` elements. The pin asserted `0=0` across every iteration — would PASS even if the A-CONT parser fix were reverted. Same shape as the R5-CONT-4 ternary anti-pattern (`XCTAssertNil(<bool> ? 1000 : nil)`): test framework that LOOKS rigorous but lacks regression sensitivity.

A-CONT-2 adds `testMatrixPinCatchesContainerBookmarkRegression` which builds a synthetic fixture with `<w:hdr>` containing 2 body-level `<w:bookmarkStart>` + matching `<w:bookmarkEnd>`, runs the same matrix-pin assertion path, asserts non-trivial parity (2=2 not 0=0). Catches future parser-asymmetry regressions for real.

#### A-CONT-2 P1 — `deleteBookmark` symmetry with `getBookmarks`

Pre-A-CONT-2 `deleteBookmark(name:)` ([Document.swift:2038-2056](https://github.com/PsychQuant/ooxml-swift/blob/v0.19.8/Sources/OOXMLSwift/Models/Document.swift#L2038)) only matched `.paragraph(...).bookmarks` — couldn't delete body-level `.bookmarkMarker` entries (or container body-level markers). After A-CONT-2 P0 #1, `getBookmarks()` lists names that the prior `deleteBookmark` would throw `BookmarkError.notFound` on — state inconsistency widened by A-CONT.

A-CONT-2 extends `deleteBookmark` with a `tryDeleteBodyLevelBookmark` helper that scans body-level `.bookmarkMarker` entries (matching by name on `.start` markers, removing matching `.end` by id), recurses into `.contentControl`, and applies across body + 4 container types. `modifiedParts` is correctly marked for the owning part (body / specific header / specific footer / footnotes / endnotes). `getBookmarks()` and `deleteBookmark()` are now fully symmetric.

### Tests

- 3 new tests in `Tests/OOXMLSwiftTests/Issue58_60ContentPreservationTests.swift`:
  - `testGetBookmarksSurfacesContainerBodyLevelMarkers` (A-CONT-2 P0 #1)
  - `testParseContainerSDTRecursionPreservesNestedBookmark` (A-CONT-2 P0 #2)
  - `testMatrixPinCatchesContainerBookmarkRegression` (A-CONT-2 P0 #3)
- Suite total: 661 tests pass / 1 skipped / 0 failures (658 A-CONT baseline + 3 A-CONT-2 new tests)

### API additions (v0.19.8, additive — no breaking change vs v0.19.7 contract)

- No new public types or signatures. `getBookmarks()` and `deleteBookmark()` keep their existing call signatures; the new behavior is strictly additive (returns more / accepts more without breaking existing callers).

### Spectra change

This release ships sub-stack A-CONT-2 mini-mini-cycle. Re-numbers planned sub-stack B → v0.19.9 / v3.13.9 (sub-stack C unchanged at v0.20.0 / v3.14.0). Sub-stack A took 3 sub-cycles (A + A-CONT + A-CONT-2) to drain #58 — same trajectory shape as R5 → R5-CONT-4 needing 5 sub-cycles to drain #56.

## [0.19.7] - 2026-04-27

### Fixed — Sub-stack A-CONT of #58 (parser asymmetry mini-cycle from sub-stack A 6-AI verify)

The sub-stack A 6-AI verify ([report](https://github.com/PsychQuant/che-word-mcp/issues/58#issuecomment-4323205184)) returned BLOCK with 2 P0 + 1 P1 + 4 P2/MEDIUM (3 of 6 reviewers PASS, 1 WARN, 1 BLOCK; BLOCK independently confirmed by R2 Logic + R5 Devil's Advocate + direct code read).

Same convergence-cycle pattern as R5-CONT-4 (issue #56): per-task gate caught #58 in `parseBodyChildren` (body.xml entry point); 6-AI verify caught the symmetric sibling in `parseContainerChildBodyChildren` (header / footer / footnote / endnote entry point) that the per-task gate missed. The matrix-pin only exercised body source; container source slipped through.

#### A-CONT P0 #1 — `parseContainerChildBodyChildren` mirrors `parseBodyChildren` branches

`DocxReader.parseContainerChildBodyChildren` ([line 1291-1322](https://github.com/PsychQuant/ooxml-swift/blob/v0.19.7/Sources/OOXMLSwift/IO/DocxReader.swift#L1291-L1322)) had only `case "p"` / `case "tbl"` / `default: continue` after sub-stack A landed. Body-level `<w:bookmarkStart>` / `<w:bookmarkEnd>` inside `<w:hdr>` / `<w:ftr>` / `<w:footnote>` / `<w:endnote>` were still silently dropped on save — same data-loss class #58 was meant to close, just in a different parser entry point.

The dead-code calibration walker added in sub-stack A (`collectBodyLevelBookmarkIds(header.bodyChildren)` at DocxReader.swift:422-432) was the smoking gun: it iterated structures the parser never populated. A-CONT mirrors the exact same fix shape from `parseBodyChildren` (typed branches for bookmarkStart/bookmarkEnd, explicit skip for sectPr, raw-passthrough default) into the container parser entry point.

#### A-CONT P0 #2 — `Document.getBookmarks()` surfaces body-level `.bookmarkMarker` entries

`Document.getBookmarks()` ([line 2122-2136](https://github.com/PsychQuant/ooxml-swift/blob/v0.19.7/Sources/OOXMLSwift/Models/Document.swift#L2122)) iterated only `case .paragraph` reading `para.bookmarks` — never `case .bookmarkMarker`. Pre-fix the marker was dropped on disk (so listing nothing was at least consistent); post-sub-stack-A the marker survives on disk but was invisible to MCP `list_bookmarks` discovery — silent UX regression where users couldn't see, name, jump-to, or delete preserved TOC anchors via MCP.

A-CONT extends `getBookmarks()` to walk body-level `.bookmarkMarker(BookmarkRangeMarker)` entries. Only `.start` markers carry a name (per OOXML — `.end` markers match by id). `paragraphIndex = -1` sentinel indicates "not inside a paragraph" (the marker sits at body level). Tables / content controls / raw block elements remain out of `getBookmarks()` scope — a separate `getAllBookmarks()` walker covering full container coverage is a follow-up if needed.

#### A-CONT P1 — Matrix-pin extension: container-source parity assertion

`testDocumentContentEqualityInvariant` matrix-pin extended with `assertContainerBookmarkStartParity` helper that enumerates all `word/header*.xml` / `word/footer*.xml` / `word/footnotes.xml` / `word/endnotes.xml` parts in source + output, counts `<w:bookmarkStart>` per part, asserts parity. Catches future parser-asymmetry regressions of the same class.

#### A-CONT P2 — Stale comment fix

`parseBodyChildren`'s pre-fix doc comment ("Other elements are skipped") didn't reflect the v0.19.6 `default:` raw-preserve behavior. Updated to document the actual semantics — `<w:bookmarkStart>` / `<w:bookmarkEnd>` produce typed `BodyChild.bookmarkMarker`, `<w:sectPr>` is skipped, all other elements are captured as `BodyChild.rawBlockElement` ("if not typed, preserve as raw" principle).

### Tests

- 2 new tests in `Tests/OOXMLSwiftTests/Issue58_60ContentPreservationTests.swift`: `testHeaderBodyLevelBookmarkRoundTripPreserved`, `testGetBookmarksSurfacesBodyLevelMarkers`
- `testDocumentContentEqualityInvariant` extended with `assertContainerBookmarkStartParity` call (4 container types)
- Suite total: 658 tests pass / 1 skipped / 0 failures (656 v0.19.6 baseline + 2 A-CONT new tests)

### API additions (v0.19.7, additive — no breaking change vs v0.19.6 contract)

- No new public types or signatures. `getBookmarks()` return shape unchanged; the new `paragraphIndex = -1` sentinel for body-level markers is an additive semantic (callers that don't check the sentinel get the body-level bookmarks alongside paragraph-level ones, which is the intended behavior).

### Spectra change

This release ships sub-stack A-CONT mini-cycle. Re-numbers planned sub-stack B → v0.19.8 + v3.13.8 (sub-stack C unchanged at v0.20.0 + v3.14.0). The convergence-cycle methodology is working as intended: per-task gate caught one parser entry point; 6-AI verify caught the symmetric sibling. Sub-stack A took 2 sub-cycles (A + A-CONT) to drain #58 fully — same shape as R5 → R5-CONT-4 needing 5 sub-cycles to drain #56.

## [0.19.6] - 2026-04-27

### Fixed — PsychQuant/che-word-mcp#58 (sub-stack A of document-content-preservation)

Body-level `<w:bookmarkStart>` / `<w:bookmarkEnd>` (e.g., TOC `_Toc<digits>` anchors that wrap multiple paragraphs) were silently dropped on body-mutating save. `DocxReader.parseBodyChildren` switch only handled `<w:p>` / `<w:tbl>` / `<w:sdt>`; the `default: continue` branch silently dropped any other direct child of `<w:body>`. Reproduced on the thesis fixture: 1 of 45 bookmarks lost on round-trip (the TOC anchor matching `_Toc\d+`).

#### Architectural change: BodyChild typed + raw catch-all

`BodyChild` enum gains two cases under the unifying principle "**if not typed, preserve as raw**":

```swift
public enum BodyChild: Equatable {
    case paragraph(Paragraph)
    case table(Table)
    case contentControl(ContentControl, children: [BodyChild])
    case bookmarkMarker(BookmarkRangeMarker)        // ← NEW typed
    case rawBlockElement(RawElement)                 // ← NEW generic catch-all
}
```

- `parseBodyChildren` switch gains explicit `case "bookmarkStart"` and `case "bookmarkEnd"` branches producing `BodyChild.bookmarkMarker(BookmarkRangeMarker(...))`.
- The `default:` branch now captures unrecognized elements as `BodyChild.rawBlockElement(RawElement(name:..., xml:...))` — same architectural pattern as `Run.rawElements` (v0.14.0+, #52). Forward-compatible with other EG_BlockLevelElts members (`<w:moveFromRangeStart>`, body-level `<w:commentRangeStart>`, vendor extensions).
- `<w:sectPr>` gets an explicit `case "sectPr": continue` to preserve pre-fix behavior (it's parsed separately into `WordDocument.sectionProperties`, not into `BodyChild`).

#### `BookmarkRangeMarker.name: String?` field

Added `name: String?` (default nil) to `BookmarkRangeMarker` so body-level marker entries can carry the bookmark's name (paragraph-level markers have the name on `Paragraph.bookmarks`, but body-level markers have no enclosing paragraph). Existing initializer call sites unaffected (new param has default).

#### `nextBookmarkId` calibration extension

The `nextBookmarkId` calibration walker now ALSO walks body-level `BookmarkRangeMarker` entries across body / headers / footers / footnotes / endnotes — not just paragraph-level `paragraph.bookmarkMarkers`. This prevents future API-built bookmarks from colliding with existing body-level ids.

#### Cross-cutting matrix-pin (initial version)

New `testDocumentContentEqualityInvariant` (in `Tests/OOXMLSwiftTests/Issue58_60ContentPreservationTests.swift`) asserts content-equality round-trip on the thesis fixture across preservation classes:

- **Sub-stack A (this version)**: `<w:bookmarkStart>` count parity (catches #58 class). Test passes with 45=45 on thesis fixture.
- **Sub-stack B (lands with v0.19.7 / #59)**: `<w:t>` total character content parity.
- **Sub-stack C (lands with v0.20.0 / #60)**: `<w:rFonts>` / `<w:noProof>` / `<w:lang>` / `<w:kern>` / `w14:*` count parity.

The pin asserts CONTENT equality (counts and joined-strings), not BYTE equality — Word's own canonicalization (e.g., adjacent run consolidation) is allowed to differ. Same architectural pattern as R5-CONT-4's `testRevisionTypeMatrixAcceptRejectCompleteness` structural-symmetry pin.

### Tests

- 4 new tests in `Tests/OOXMLSwiftTests/Issue58_60ContentPreservationTests.swift`: `testBodyLevelBookmarkRoundTripPreserved`, `testBodyLevelUnknownElementPreservedAsRaw`, `testNextBookmarkIdReflectsBodyLevelBookmarksAfterRead`, `testDocumentContentEqualityInvariant` (initial version)
- Suite total: 656 tests pass / 1 skipped / 0 failures (652 v0.19.5 baseline + 4 v0.19.6 new tests)

### API additions (v0.19.6, additive — no breaking change vs v0.19.5 contract)

- `BodyChild.bookmarkMarker(BookmarkRangeMarker)` — typed body-level bookmark marker
- `BodyChild.rawBlockElement(RawElement)` — generic catch-all for unrecognized direct children of `<w:body>`
- `BookmarkRangeMarker.name: String?` — new optional field; default nil; existing callers unaffected

### Spectra change

This release implements sub-stack A of `che-word-mcp-issue-58-59-60-document-content-preservation` (sub-stack B → v0.19.7 / #59 whitespace overlay; sub-stack C → v0.20.0 / #60 RunProperties audit + final matrix-pin extension).

## [0.19.5] - 2026-04-26

### Fixed — 6 P0 + 5 P1 + R5-CONT 7 P0 + 5 P1 + R5-CONT-2 5 P0 + 4 P1 + R5-CONT-3 1 P0 + 4 P1 + R5-CONT-4 1 P0 + 3 P1 from PsychQuant/che-word-mcp#56 rounds 4 + 5 + 6 + 7 + 8 verify

The R3 stack landed in commits dated 2026-04-26 but the round-4 6-AI verify (Agent Team × 5 + Codex) returned BLOCK with 6 P0 + 7 P1 findings spanning walker asymmetry, sentinel collision, attribute-escape gaps, block-level SDT propagation, container-symmetric `replaceText`, and container `<w:tbl>` capture. The R5 stack closed those (see "R5 stack" sub-block below). The round-5 6-AI verify (https://github.com/PsychQuant/che-word-mcp/issues/56#issuecomment-4321866434) then returned BLOCK with 7 NEW P0 findings rooted in a single structural pattern: R5 P0 #6 promoted `bodyChildren` to canonical container storage with `paragraphs` as a flat backward-compat computed view, but several call sites still iterated `.paragraphs`, silently dropping anything inside container tables / contentControls. The R5-CONTINUATION sub-block (§11) closes those + 5 adjacent P1 findings via the same per-task verify gate discipline, in the same release. v0.19.5 ships both stacks as a single coordinated tag.

### R5-CONTINUATION sub-block — 7 P0 + 5 P1 from R5 verify (round-5 stack-completion)

#### R5-CONT P0 #1 — handleMixedContentWrapperRevision walks container bodyChildren

The four container loops in `Document.handleMixedContentWrapperRevision` were iterating `headers[hi].paragraphs` (flat computed view), missing wrappers inside container tables / SDTs. `transformInBodyChildren` is now parameterized over `partKey` and the four container loops route through it on `bodyChildren`. Body branch unchanged. Closes verify R5 P0 #1 + Logic L2 + DA C1.

#### R5-CONT P0 #2 — DocxReader per-container revision propagation walks bodyChildren

`propagateRevisionsFromBodyChildren` parameterized over `source: RevisionSource = .body`. The four per-container revision propagation loops in `DocxReader.read` collapse to single calls of the helper with the correct source label, walking each container's `bodyChildren` (not `.paragraphs` flat view). Typed Revisions inside container tables / nested tables / contentControls now reach `document.revisions.revisions`. Also closes DA-N H1 (hardcoded `.body` source label).

#### R5-CONT P0 #3 — replaceText(.all) recurses into container bodyChildren

The four container loops in `Document.replaceText(scope: .all)` now route through the existing `replaceTextInBodyChildren` recursion, walking `bodyChildren` (incl. tables, nested tables, contentControl). Local-var copy pattern (`var children = container.bodyChildren` → mutate → write back) avoids Swift exclusivity violations from `mutating self` recursive calls. Closes verify R5 P0 #3 + Codex P1 + Regression F1 + DA C1.

#### R5-CONT P0 #4 — partKey unification between DocumentWalker and Header/Footer.fileName

`DocumentWalker.headerPartKey(for:)` and `footerPartKey(for:)` now delegate to `header.fileName` / `footer.fileName` — the same accessor `DocxWriter` uses for dirty-gate checks. Pre-fix the walker's private `defaultHeaderFileName` switch (returned `header2.xml` for `.even`) disagreed with the model's `headerEven.xml` for every (HeaderFooterType, originalFileName=nil) combination, producing silent loss-on-save for API-built containers. Closes verify R5 P0 #4 + Logic L1.

#### R5-CONT P0 #5 — acceptRevision typed .deletion routes by revision.source

New `sourceToPartKey(_ source: RevisionSource) -> String` and `applyToParagraph(at:in:mutate:) -> String?` helpers. The typed `.deletion` branch now consults `revision.source` to find the right paragraph slot (across body, headers, footers, footnotes, endnotes — incl. nested tables / contentControl). Throws `RevisionError.notFound` on miss instead of silent no-op. `modifiedParts` marked with the actual mutated part. Closes verify R5 P0 #5 + DA C2 (silent corruption: container .deletion silently no-op'd OR deleted the wrong body paragraph) + DA H2 (block-level SDT internal .deletion).

#### R5-CONT P0 #6 — toXMLSortedByPosition filters API-built runs/hyperlinks symmetric with contentControls

The four positioned-list builder loops for runs / hyperlinks / fieldSimples / alternateContents now apply the `where position > 0` filter (matching what contentControls already had). The `position == 0` API-built entries emit in the legacy post-content section so they land at end-of-paragraph rather than sorting BEFORE source-loaded children. Closes verify R5 P0 #6 + DA C3 (asymmetric sentinel handling: source-loaded paragraph + `insertText` previously placed text at paragraph head).

#### R5-CONT P0 #7 — getHyperlinks walks all parts

Public `Document.getHyperlinks()` routes through `DocumentWalker.walkAllParagraphs(in: self)` so the returned id/text/url/anchor/type tuple list covers every part (incl. tables / SDT children inside body / headers / footers / footnotes / endnotes). Pre-fix only body top-level paragraphs were listed → the listed-id set was a strict subset of what `updateHyperlink` / `deleteHyperlink` could find. Closes verify R5 P0 #7 + DA C5.

### R5-CONTINUATION P1 follow-ups

- **R5-CONT P1 #8** — `updateHyperlink(url:)` URL sync targets the OWNING part's rels file (`word/_rels/header*.xml.rels`, `footer*.xml.rels`, `footnotes.xml.rels`, `endnotes.xml.rels`) instead of always document-scope. New per-container `relationships: RelationshipsCollection` fields on `Header` / `Footer` / `FootnotesCollection` / `EndnotesCollection`. `Relationship` now Equatable with mutable `target` + optional `targetMode`. New `parseRelationshipsFile(at:)` generic parser; new `writeRelationshipsCollection(_:to:)` writer helper. Codex caught a per-part rId scoping edge case during scoped verify: the merged-rels lookup must search container rels FIRST so colliding ids resolve against the correct part — fixed via `mergedRels = containerRels + documentRels` order with a dedicated regression test. Closes verify R5 P1 #8 + Logic L4 + Codex P1 #4.
- **R5-CONT P1 #9** — Container `toXML()` for Header/Footer/Footnote/Endnote routes through `DocxWriter.xmlForBodyChild` (promoted from private to internal). The `.contentControl` arm now emits `<w:sdt>...</w:sdt>` instead of being silently dropped. Closes verify R5 P1 #9 + Logic L6 + Codex P2.
- **R5-CONT P1 #10** — XML escape sweep audit table tightened to reflect byte-equivalence reality. Investigation found all 20+ remaining local `escapeXML(_:)` helpers across `Hyperlink.swift`, `Footer.swift`, `Comment.swift`, `Image.swift`, `Revision.swift`, and the 10 `Field.swift` instances escape ALL FIVE attribute-significant chars (`& < > " '`) — the prior R3-stack note that "only `'` is missed" was stale. DocxWriter's `escapeAttr` is intentional 4-char (single-quote allowed unescaped inside double-quoted attributes per XML spec). True consolidation onto a single helper is a code-hygiene follow-up; no security impact remains. Closes verify R5 P1 #10 + DA H3.
- **R5-CONT P1 #11** — Roundtrip variant added for `testBlockLevelSDTWrappedRevisionAcceptPersistsThroughRoundtrip` (the highest-risk in-memory mutation test DA H5 explicitly called out). Pure-emit and API-throw tests intentionally remain in-memory because the writer-side check is on the toXML/parser path itself; adding mechanical variants would not catch additional regressions per category. Closes verify R5 P1 #11 + DA H5.
- **R5-CONT P1 #12** — `DocxReader.walkAllParagraphs` private duplicate removed. `nextBookmarkId` calibration now routes through the shared `DocumentWalker.walkAllParagraphs`. Remaining `for header in document.headers` loops in production code are intentional (DocumentWalker primitive itself + per-container revision propagation that needs per-iteration source labels). Closes verify R5 P2 #13 + DA C4 (the "single walker = no walker asymmetry" promise).

### Tests — R5-CONTINUATION

- 11 new tests in `Tests/OOXMLSwiftTests/Issue56R4StackTests.swift` (one per P0 + P1 finding, plus the Codex-caught rId collision regression and the §11.11 SDT-revision roundtrip variant)
- 1 pre-existing test (testUpdateHyperlinkInsideHeaderTableSucceeds, §8.3) updated to use the new `header.relationships` field instead of the pre-R5-CONT `document.hyperlinkReferences` workaround
- Suite total: 640 tests pass / 1 skipped / 0 failures (628 R5 baseline + 12 R5-CONTINUATION new tests)

### API additions (R5-CONTINUATION, additive — no breaking change vs. R5 stack contract)

- `Header.relationships`, `Footer.relationships`, `FootnotesCollection.relationships`, `EndnotesCollection.relationships` — new `var relationships: RelationshipsCollection` fields for per-container rels storage. Empty for API-built containers; populated by DocxReader; emitted by DocxWriter when non-empty.
- `Relationship` now `Equatable`. `target` is mutable; new optional `targetMode: String?` for hyperlink rels (`TargetMode="External"` round-trip).
- `RelationshipsCollection` now `Equatable`.
- `DocxReader.parseRelationshipsFile(at:)` — new internal generic per-file rels parser used for per-container rels.
- `DocxWriter.xmlForBodyChild(_:)` — promoted from `private static` to `internal static` so container `toXML()` can reuse the SDT serialization.

### R5-CONT-2 sub-block — 5 P0 + 4 P1 from R5-CONT 6-AI verify (round-6 cross-cutting completion)

The R5-CONT 6-AI verify (https://github.com/PsychQuant/che-word-mcp/issues/56#issuecomment-4322314964) returned BLOCK with 5 NEW P0 silent-corruption surfaces. Per-task gates closed each NARROW R5 verify finding but missed cross-fix asymmetries: accept↔reject mirror, update↔delete mirror, partial filter coverage, and cross-helper invariants (writer's `paragraphIndex` semantic vs lookup helper's flat-counter semantic). R5-CONT-2 closes the 5 P0 + 4 of 5 P1 (1 P1 deferred — see Caveats below) via the same per-task gate.

#### R5-CONT-2 P0 #1 + #5 — paragraphIndex per-paragraph counter

`propagateRevisionsFromBodyChildren` previously took an external `paragraphIndex` parameter. Container call sites passed `0` for ALL revisions; body case `.contentControl` passed body-children enum index. Both diverged from `applyToParagraph`'s flat-paragraph counter lookup. Fix: helper now uses an internal `var counter = 0` that increments per visited paragraph (recursing into tables / nested tables / SDT inner). Body propagation collapses to a single helper call (the per-case body switch is removed). Container call sites drop the `paragraphIndex: 0` argument. Single source of truth for paragraph-position semantics.

#### R5-CONT-2 P0 #2 — `deleteHyperlink` targets owning part rels

`Document.deleteHyperlink` only updated `document.hyperlinkReferences` and unconditionally marked `word/_rels/document.xml.rels` dirty. Container hyperlinks deleted left orphan rels in `header*.xml.rels` AND wrongly dirtied document rels. Mirror of R5-CONT P1 #8 `updateHyperlink(url:)` fix — new `removeHyperlinkRelTarget(rId:partKey:)` routes via the owning part's relationships; correct rels file marked dirty.

#### R5-CONT-2 P0 #3 — `rejectRevision` typed `.insertion` routes by `revision.source`

`rejectRevision`'s typed `.insertion` branch was body-only — same class as the R5-CONT P0 #5 `acceptRevision` typed `.deletion` bug, but for the reject side and never mirrored. A container-source `.insertion` rejected via `rejectRevision` would silently no-op OR (worse) DELETE BODY TEXT matching `newText` substring. Fix: typed `.insertion` now routes by `revision.source` via the same `applyToParagraph` + `sourceToPartKey` helpers `acceptRevision` already uses. Throws `RevisionError.notFound` on miss instead of silent corruption.

#### R5-CONT-2 P0 #4 — `toXMLSortedByPosition` filter sweep covers all 12 positioned collections

R5-CONT P0 #6 added the `where position > 0` filter to runs / hyperlinks / fieldSimples / alternateContents (4 of 12 positioned collections). The 8 remaining (`bookmarkMarkers`, `commentRangeMarkers`, `permissionRangeMarkers`, `proofErrorMarkers`, `smartTags`, `customXmlBlocks`, `bidiOverrides`, `unrecognizedChildren`) still went into the sort list unconditionally → API-built marker (constructed with explicit `position == 0`) sorted BEFORE every source-loaded child and landed at paragraph head. Fix: filter applied to all 8 remaining collections; symmetric post-content emit added for the position-0 entries.

### R5-CONT-2 P1 follow-ups

- **R5-CONT-2 P1 #6** — `Relationship.rawType: String` preserves the literal source `Type` attribute string. `parseRelationshipsFile` populates it from the source attribute regardless of typed-enum recognition. `writeRelationshipsCollection` prefers `rel.rawType` over `rel.type.rawValue`. Unknown vendor extension types (VML / OLE / Word extension rels) round-trip byte-equivalent instead of being downgraded to `Type=""` (which is invalid OOXML rels).
- **R5-CONT-2 P1 #7** — Hyperlink rels lookup in `parseHyperlink` adds `&& $0.type == .hyperlink` filter to `first(where:)`. Pre-fix the id-only match could resolve a header hyperlink's rId1 to a document-scope rels entry of type `header` (Type=header Target=header1.xml) — wrong-type silent resolution to a part path string. Fix combines with R5-CONT P1 #8's container-first merge order to fully close the cross-part rId resolution surface.
- **R5-CONT-2 P1 #8** — Hyperlink id format includes part scope. Body hyperlinks keep `<rId-or-anchor-or-hl>@<position>`; container hyperlinks (header / footer / footnote / endnote) get prepended with the container part fileName (e.g., `header1.xml:rId1@0`). New `rewriteHyperlinkIdsInBodyChildren` post-processes parsed container bodyChildren to part-scope every hyperlink id (idempotent; only prefixes ids that don't already contain `:`). After R5-CONT P0 #7 made `getHyperlinks` cross-part, two parts producing same `rId@position` were indistinguishable to MCP callers — now disambiguated.
- **R5-CONT-2 P1 #10** — Stale rels file removal. `writeHeader` / `writeFooter` / `writeFootnotes` / `writeEndnotes` add `else if FileManager.default.fileExists(atPath: relsURL.path) { try? FileManager.default.removeItem(at: relsURL) }` so emptying a container's relationships collection (e.g., via `Document.deleteHyperlink`) actually removes the stale `word/_rels/<container>.xml.rels` file from disk on save. Pre-fix overlay-mode preserved the stale file — Word and validators warn about unused relationships.

### Caveats — R5-CONT-2 P1 #9 deferred

R5-CONT verify DA C5 flagged a fileName-collision risk for two API-built `.default` headers without `originalFileName` set (both produce `header1.xml`, leading to `updateHyperlinkRelTarget` partKey-loop matching both). Investigation showed:

1. The public `Document.addHeader(text:type:)` API already routes through `allocateHeaderFileName(for:)` which auto-suffixes — collision only arises when callers SKIP this API and directly construct `Header(id:type:)` then append to `document.headers` raw.
2. A complete fix requires either auto-allocation in `Header` init referencing parent Document state (invasive — Header doesn't know its parent), OR writer-side collision detection + rename (introduces in-memory != on-disk non-determinism).

R5-CONT-2 documents the limitation: callers building multiple same-type containers SHALL use the public `addHeader` / `addFooter` API rather than direct construction. A follow-up issue is tracked for full auto-allocation.

### Tests — R5-CONT-2 stack

- 5 new tests in `Tests/OOXMLSwiftTests/Issue56R4StackTests.swift` (one per P0 #1+#5, P0 #2, P0 #3, P0 #4 — the P1s are validated by the per-fix Codex scoped review and the running suite)
- 1 pre-existing test (`testUpdateHyperlinkUrlInsideHeaderTargetsHeaderRels` §11.8) updated to use `id.contains("rId99")` instead of `id.hasPrefix("rId99")` to accommodate the new container-id format
- Suite total: 645 tests pass / 1 skipped / 0 failures (640 R5-CONT baseline + 5 R5-CONT-2 new tests)

### API additions (R5-CONT-2, additive — no breaking change vs. R5-CONT contract)

- `Relationship.rawType: String` (new field with default = `type.rawValue`); `Relationship.init(...)` gains optional `rawType: String? = nil` parameter
- Container hyperlink id format change is technically observable but follows the same backward-compat path as the R3-stack `<rId>@<position>` change: callers who cached pre-R5-CONT-2 ids and look them up after upgrade will get nil; mitigation is to re-parse documents

### R5-CONT-3 sub-block — 1 P0 + 4 P1 + matrix-completeness pin from R5-CONT-2 6-AI verify (round-7 cycle convergence)

The R5-CONT-2 6-AI verify (https://github.com/PsychQuant/che-word-mcp/issues/56#issuecomment-4322505227) returned BLOCK with 1 P0 + 4 P1 from devil's advocate (other 4 reviewers PASS). The P0: `rejectRevision` typed `.deletion` was a silent no-op at the file level — comment claimed "just clear the marker" but ONLY removed `document.revisions[id]`. `paragraph.revisions` still contained the Revision id; `run.revisionId` still referenced it; `Paragraph.toXML()` still wrapped the runs in `<w:del>` on save. Same class as the R5/R5-CONT/R5-CONT-2 cycle's repeated finding: per-task gate covered SPECIFIC cases (insertion accept/reject mirror) but missed SYMMETRIC siblings (the rest of the typed Revision matrix). R5-CONT-3 closes the P0 + 4 P1 + adds an explicit cross-cutting symmetry test pin to break the convergence cycle.

#### R5-CONT-3 P0 #1 + P1 #2 — `rejectRevision` typed cases clear paragraph + run revision state

`rejectRevision` typed `.deletion` branch (and `.formatting` / `.paragraphChange` / `.formatChange` / `.moveFrom` / `.moveTo`) now route through `applyToParagraph(in: revision.source)` with a `clearMarker` closure that:
- removes the typed Revision from `paragraph.revisions[id]`
- clears `run.revisionId` for any matching run
- clears `paragraph.paragraphFormatChangeRevisionId` (for pPrChange)
- clears `run.formatChangeRevisionId` (for rPrChange)

§15.6's matrix test caught a related gap on first run: §13.3's `rejectRevision` typed `.insertion` (R5-CONT-2) only ran `removeText` but didn't clear paragraph/run state. Same closure pattern extended there (`removeAndClear` replaces the prior `removeText`-only closure).

#### R5-CONT-3 P1 #3 — `sourceToPartKey` throws on orphan container source

Pre-fix the helper silently fell back to `"word/document.xml"` when source `.header(id:X)` / `.footer(id:X)` named a non-existent container. Wrong-part dirty masked orphan-revision logic bugs. Now: throws `RevisionError.notFound(revisionId)`. Signature changes from `(_ source: RevisionSource) -> String` to `(_ source: RevisionSource, revisionId: Int) throws -> String`. All 3 call sites updated with `try` + revisionId argument.

#### R5-CONT-3 P1 #4 — `deleteHyperlink` sweeps legacy doc-scope rels

`removeHyperlinkRelTarget` now defensively sweeps `document.hyperlinkReferences` for the rId when partKey is non-body. R5-CONT P1 #8 introduced per-container rels but documents migrated from the older single-rels model could still carry the same rId in document-scope (legitimate when caller historically used document-scope before the migration). Without this sweep, container deletes leave a doc-scope orphan that never cleans up.

#### R5-CONT-3 P1 #5 — public collision detection + repair for multi-instance same-type containers

R5-CONT-2 §13.8 deferred this — the auto-allocation in `Header` init referencing parent Document state was invasive (Header doesn't know its parent), and writer-side rename introduced in-memory != on-disk non-determinism. R5-CONT-3 closes the deferral with a public diagnostic + opt-in repair helper:

- `Document.containerFileNameCollisions: [(scope: String, fileName: String, indices: [Int])]` — empty when clean, surfaces all collisions for MCP / diagnostic tools to warn before save
- `Document.repairContainerFileNames()` — auto-reassigns `originalFileName` on the SECOND+ instances using the same `allocateHeaderFileName` / `allocateFooterFileName` helper the public `addHeader` / `addFooter` API uses; first instance keeps its existing fileName; marks every reassigned container's part dirty; idempotent

Caller pattern: call `repairContainerFileNames()` just before save when constructing containers via direct `headers.append(...)` rather than `addHeader`. The public API path (addHeader/addFooter) auto-handles, so the helper is only needed for direct-construction flows.

#### R5-CONT-3 cross-cutting symmetry pin — revision type matrix completeness

`testRevisionTypeMatrixAcceptRejectCompleteness` exercises 14 cases (7 typed Revision types × accept + reject). Each case asserts: (1) operation succeeds or throws documented error; (2) `document.revisions` cleared; (3) `paragraph.revisions` cleared (file-state convergence); (4) no silent partial state on revision-id refs. This pin closes the convergence cycle: per-task gate alone discovers SPECIFIC bugs; matrix-pin catches SYMMETRIC siblings before the next 6-AI verify round flags them.

### Tests — R5-CONT-3 stack

- 5 new tests in `Tests/OOXMLSwiftTests/Issue56R4StackTests.swift` (one per §15.1+§15.2 / §15.3 / §15.4 / §15.5 / §15.6)
- Suite total: 650 tests pass / 1 skipped / 0 failures (645 R5-CONT-2 baseline + 5 R5-CONT-3 new)

### API additions (R5-CONT-3, additive — no breaking change vs. R5-CONT-2 contract)

- `Document.containerFileNameCollisions: [(scope: String, fileName: String, indices: [Int])]` — public diagnostic
- `Document.repairContainerFileNames()` — public mutator (idempotent)
- `sourceToPartKey` is private — internal change, no API impact

### R5-CONT-4 sub-block — 1 P0 + 3 P1 from R5-CONT-3 6-AI verify (round-8 acceptRevision symmetry + matrix-pin tightening)

The R5-CONT-3 6-AI verify (https://github.com/PsychQuant/che-word-mcp/issues/56#issuecomment-4322571860 + Codex confirmation 4322576289) returned BLOCK with 1 P0 + 4 P1. The P0: `acceptRevision` typed cases (`.insertion` / `.deletion` / `.formatting` / `.paragraphChange` / `.formatChange` / `.moveFrom` / `.moveTo`) all left paragraph/run revision markers in place — `paragraph.revisions[id]` / `run.revisionId` / `run.formatChangeRevisionId` / `paragraphFormatChangeRevisionId` were never cleared, so `Paragraph.toXML()` re-emitted `<w:ins>` / `<w:del>` / `<w:rPrChange>` / `<w:pPrChange>` wrappers on save. API state said "accepted" but file persistence still had the wrapper. Same class as R5-CONT-2 P0 #1 + R5-CONT-3 P0 #1 — but on the ACCEPT side and across ALL 7 typed branches. R5-CONT-3's §15.6 matrix-pin had an `if operation == "reject"` guard that documented the bug as expected behavior. R5-CONT-4 closes the P0 by mirroring R5-CONT-3's clearMarker pattern onto the accept side, removes the asymmetry guard so the matrix-pin asserts both sides, and replaces a ternary `XCTAssertNil` anti-pattern that false-passed on regression. §17.4 closes the related Logic HIGH `repairContainerFileNames` rels-dirty-marking gap.

#### R5-CONT-4 P0 #1 — `acceptRevision` typed cases clear paragraph + run revision state

`acceptRevision` typed branches now route through `applyToParagraph(in: revision.source, mutate: clearAllMarkers)` with a `clearAllMarkers` closure that:
- removes the typed Revision id from `paragraph.revisions`
- clears `run.revisionId` for any matching run
- clears `paragraph.paragraphFormatChangeRevisionId` (for pPrChange)
- clears `run.formatChangeRevisionId` (for rPrChange)

Mirror of R5-CONT-3 P0 #1 + P1 #2's reject-side fix, applied to all 7 accept-side typed branches. `.deletion` keeps the `removeText` behavior AND adds `clearAllMarkers`. Throws `RevisionError.notFound` on miss instead of silent no-op. `modifiedParts` marked with the actual mutated part. Closes verify R5-CONT-3 P0 #1 + DA R6-NEW-1 (4 adversarial tests DA added all failed pre-fix: `.insertion` → `<w:ins>` remains; `.deletion` → `<w:del>` remains with empty `<w:delText>` — worse corruption; `.formatChange` → `<w:rPrChange>` remains; `.paragraphChange` → `<w:pPrChange>` remains).

#### R5-CONT-4 P1 #2 — Matrix-pin asserts both accept AND reject — removes asymmetry guard

R5-CONT-3's `testRevisionTypeMatrixAcceptRejectCompleteness` had `if operation == "reject"` guarding the paragraph-state cleanup assertions, with comment "consistent with their CURRENT contracts". R5-CONT-3 verify proved that documented "current contract" WAS the bug: the guard hid the §17.1 P0. R5-CONT-4 removes the guard. Both accept and reject SHALL satisfy the same paragraph-state cleanup invariants. The matrix now exercises 14 cases (7 typed Revision types × accept + reject) and asserts file-state convergence on EVERY case. Closes verify R5-CONT-3 P1 §15.6.

#### R5-CONT-4 P1 #3 — Replace ternary `XCTAssertNil` anti-pattern with `XCTAssertNotEqual`

The matrix-pin had `XCTAssertNil(<bool> ? 1000 : nil)` — if a regression set `revisionId = 2000` instead of nil, the ternary would still evaluate to nil (because the `<bool>` would be false) and the assertion would PASS, silently masking the bug. Replaced with `XCTAssertNotEqual(value, 1000)`, which fails for both nil-but-wrong-value AND non-cleared-marker regressions. Closes verify R5-CONT-3 P1 §15.6 / DA R6-NEW-3 / Logic L7.

#### R5-CONT-4 §17.4 — `repairContainerFileNames` marks document rels + content-types dirty

Pre-fix `repairContainerFileNames` reassigned `originalFileName` and marked the renamed container's part dirty (e.g., `word/header2.xml`), but `word/_rels/document.xml.rels` still referenced the OLD path (`header1.xml`) and `[Content_Types].xml` still listed it. `hasNewTypedRelationships` returned false (header IDs unchanged) so the writer's overlay-mode skipped re-emitting `document.xml.rels`. After save, rels pointed to `header1.xml` but the actual file lived at `header2.xml` — Word couldn't open the result. Fix: introduces a `renamed` flag; when ANY rename occurs, marks BOTH `word/_rels/document.xml.rels` AND `[Content_Types].xml` dirty so the writer's overlay-mode re-emits both with the new container fileName references. New test `testRepairContainerFileNamesDirtiesDocumentRels` verifies the dirty-set membership. Closes verify R5-CONT-3 Logic HIGH §15.4.

### Tests — R5-CONT-4 stack

- 2 new tests in `Tests/OOXMLSwiftTests/Issue56R4StackTests.swift` (one per §17.1 P0 + §17.4 Logic HIGH); §17.2 + §17.3 tighten the existing matrix-pin from §15.6
- Suite total: 652 tests pass / 1 skipped / 0 failures (650 R5-CONT-3 baseline + 2 R5-CONT-4 new)

### Caveats — R5-CONT-4

- **§15.5 deferred (contested finding)**: R5-CONT-3 verify Logic HIGH flagged `removeHyperlinkRelTarget`'s legacy doc-scope sweep as over-aggressive (could delete legitimate body rels when a header rel uses the same rId). Codex independently assessed §15.5 as appropriate for its narrow scope (sweep only fires for non-body partKey, the colliding-rId-as-legitimate-body-rel scenario requires migrated docs that already had rId scope ambiguity). Deferred pending agreement; the contested finding is documented here so consumers know the conservative behavior is per-design under at least one reviewer's interpretation.

### API additions (R5-CONT-4, additive — no breaking change vs. R5-CONT-3 contract)

- No new public API. R5-CONT-4 is internal (acceptRevision branches) + test-tightening (matrix-pin) + writer dirty-set repair (`repairContainerFileNames`).

### R5 stack — original 6 P0 + 5 P1 (preserved verbatim from initial v0.19.5 draft)

#### R5 P0 #1 — Mixed-content revision wrapper walker SHALL find wrappers in every part

`Document.handleMixedContentWrapperRevision` no longer body-only. New `DocumentWalker.walkAllParagraphs(in:visit:)` enumerates every paragraph across body (recursing into tables / nested tables / contentControl children), each header (`word/header*.xml`), each footer (`word/footer*.xml`), each footnote, and each endnote — with the originating part key passed to the visit callback. Helper now returns `(paragraph, indexInParagraph, partKey)` and throws `RevisionError.notFound(id)` on miss instead of silent return; caller updates `modifiedParts.insert(partKey)` (not blanket `word/document.xml`) on success and propagates the throw on miss. `acceptRevision` / `rejectRevision` / `acceptAllRevisions` / `rejectAllRevisions` now correctly handle wrappers in headers, footers, footnotes, and endnotes.

#### R5 P0 #2 — Reader assigns source-paragraph child positions starting at 1

`DocxReader.parseParagraph` now initializes `var childPosition = 1` (was 0). `Paragraph.toXMLSortedByPosition` includes ALL contentControls in the positioned-emit list (drops the `> 0` filter); legacy emit path includes only `contentControls.filter { $0.position == 0 }` (the API-built sentinel). `Paragraph.hasSourcePositionedChildren` keeps the `> 0` check (semantics now consistent with positions starting at 1). Eliminates the `position == 0` sentinel collision where a first-child source SDT round-tripped at the same logical position as an API-built one.

#### R5 P0 #3 — Single shared `escapeXMLAttribute` helper across all attribute emit sites

New `internal func escapeXMLAttribute(_:)` in `Sources/OOXMLSwift/IO/XMLAttributeEscape.swift` mapping `& < > " '` → `&amp; &lt; &gt; &quot; &apos;` (Decision 4: `&apos;` not `&#39;` for byte-equivalence with Word). Sweep deletes the 15+ fileprivate duplicates across `Run.swift`, `Revision.swift`, `Paragraph.swift`, `Style.swift`, `Numbering.swift`, `Table.swift`, `Field.swift`, `MathComponent.swift`, `Image.swift`, `Section.swift`, `Comment.swift`, `DocxWriter.swift`. R3-NEW-6's `&#39;` is upgraded to `&apos;` for byte-equivalence. Audit table comment in `Issue56R4StackTests.swift` migrates from R3's deny-list ("all sites covered") to an explicit allow-list naming every emit site that bypasses the helper with rationale (numeric interpolations, pre-validated rIds, verbatim XML, named site-specific exemptions, alternate escape helpers).

#### R5 P0 #4 — Block-level SDT typed Revisions propagate into `document.revisions.revisions`

`DocxReader.read` post-process loop's `case .contentControl` branch now recurses into `contentControl.children` via new `propagateRevisionsFromBodyChildren(_:paragraphIndex:into:)`, propagating any typed `Revision` (with `isMixedContentWrapper`) into `document.revisions.revisions`. Pre-fix `<w:sdt><w:sdtContent><w:p><w:ins w:id="N">...</w:ins></w:p></w:sdtContent></w:sdt>` parsed the typed Revision onto the inner paragraph but the document-level revisions list never saw it — `acceptRevision(id: N)` threw notFound.

#### R5 P0 #5 — `Document.replaceText` symmetric across body and container parts

Headers / footers / footnotes / endnotes branches in `replaceText` (`Document.swift:429-485`) now route through `replaceInParagraphSurfaces(_:find:with:options:)` — the same helper the body path uses. Pre-fix the container loops walked only `para.runs`, silently dropping edits to text inside hyperlinks, fieldSimples, and alternateContents living in headers/footers/footnotes/endnotes. P0 #5's commit also bundled R5 P1 #2 (Footnote.toXML / Endnote.toXML emit from `paragraphs` when populated) because the test path needed both fixes to GREEN.

#### R5 P0 #6 — Container parser captures `<w:tbl>` direct children of header / footer / footnote / endnote roots

`Header`, `Footer`, `Footnote`, `Endnote` gain `public var bodyChildren: [BodyChild] = []` as canonical storage. `paragraphs: [Paragraph]` is now a backward-compatible computed view (get + set; setter preserves table / contentControl positions). `DocxReader.parseContainerBody` and `parseContainerChildBodyChildren` capture both `<w:p>` and `<w:tbl>` direct children. Container `toXML()` emits from `bodyChildren` (Footnote / Endnote keep the legacy single-text-run fallback for API-built notes). `DocumentWalker.walkAllParagraphs` and `DocxReader.walkAllParagraphs` recurse into container `bodyChildren` so paragraphs nested inside container tables (and nested tables) are visible to all walker callers (calibration, mixed-content revision wrapper search, hyperlink ops). `DocxWriter.writeFooter` empty-body sentinel switches from `paragraphs.isEmpty` (would fire for table-only footers) to `bodyChildren.isEmpty`.

### Fixed — P1 follow-ups

- **R5 P1 #1** — `Hyperlink.toXML()` mutation detection upgrades from joined-text comparison to deep `[Run]` equality (`runs == childrenRuns` where `childrenRuns` is `compactMap` of `.run(_)` cases out of `children`). Synthesized `Run.Equatable` covers text + properties via `RunProperties.Equatable`. Closes property-only mutations (e.g., `runs[0].properties.bold = true` with same text) and equal-length text swaps that pre-fix silently dropped on save. Trade-off preserved (non-run order may be lost on hyperlinks containing non-run children, by design).
- **R5 P1 #2** — `Footnote.toXML` and `Endnote.toXML` emit from `bodyChildren` when populated (P0 #5 commit + §6 + §7 cover this). Legacy single-text-run template only fires for API-built notes constructed via the `Footnote(id:text:paragraphIndex:)` initializer without further mutation. A dedicated regression test (`testFootnoteMultiParagraphMutationSurvivesRoundtrip`) pins the contract.
- **R5 P1 #3** — `Document.updateHyperlink` and `Document.deleteHyperlink` walk every part instead of only `body.children[i].paragraph`. New `applyToHyperlink(id:apply:) -> String?` and `removeHyperlink(id:captureRelationshipId:) -> String?` helpers visit body (incl. nested tables / SDT children), headers via `header.bodyChildren`, footers via `footer.bodyChildren`, footnotes, endnotes. `modifiedParts` now picks up the actual owning part. Static dispatch + `Self.` recursion avoids Swift exclusivity errors that arise from `mutating self` recursive calls through `inout` bindings. Fixes the silent "Hyperlink ... not found" mode for any hyperlink living anywhere other than direct body paragraphs.
- **R5 P1 #4** — `SDTParser.parseSDT` recursive call inside `<w:sdtContent>` now passes a positive sibling-counter starting at 1. Nested SDTs receive distinct positions matching their source-document order — no longer collide with the API-built `position == 0` sentinel. Closes DA-N8 (also added a sibling test pinning the one-based counter contract).
- **R5 P1 #5** — Additive `tryAcceptAllRevisions() throws` / `tryRejectAllRevisions() throws` surface aggregate failure as `RevisionError.partialFailure([Int])` listing failing revision ids. Successful sibling revisions are still applied (partial-success semantics). Legacy non-throwing `acceptAllRevisions()` / `rejectAllRevisions()` preserved (delegating via `try?`) so che-word-mcp `Server.swift` compiles unchanged per the R5 design's zero-MCP-source-change discipline.

### Tests — R5 stack

- 14 new tests in `Tests/OOXMLSwiftTests/Issue56R4StackTests.swift` (one per P0 + P1 finding, plus a Codex-added sibling-counter test for §8.4)
- 11 roundtrip variants in `Tests/OOXMLSwiftTests/Issue56R3StackTests.swift` exercising the full DocxWriter→DocxReader cycle on every R3 stack assertion (closes DA-N5 — the all-in-memory R3 pattern was the proven blind spot of R2→R3→R4)
- New helpers: `Helpers/RoundtripHelper.swift` (`roundtrip(_:)`), `Sources/OOXMLSwift/IO/DocumentWalker.swift` (centralized walker abstraction)
- Suite total: 628 tests pass / 1 skipped / 0 failures (582 v0.19.3 baseline + 12 R3 + 14 R4 + 11 roundtrip variants + 9 helper / walker / escape tests)
- Per-task verify gate: scoped Codex CLI run after every P0 / P1 fix; flagged additions fixed inline before commit (e.g., §4 sweep additions for `MathAccent`, `Image`, `Section`, `Comment`)

### API additions (R5 stack, additive — no breaking change vs. v0.19.4 contract)

- `Header.bodyChildren: [BodyChild]`, `Footer.bodyChildren: [BodyChild]`, `Footnote.bodyChildren: [BodyChild]`, `Endnote.bodyChildren: [BodyChild]` — canonical storage promoted from the prior `paragraphs` field. Existing `paragraphs` accessors are now backward-compatible computed views (get + set).
- `RevisionError.partialFailure([Int])` — new error case raised by `tryAcceptAllRevisions` / `tryRejectAllRevisions`.
- `Document.tryAcceptAllRevisions() throws` / `Document.tryRejectAllRevisions() throws` — new throwing variants of the legacy non-throwing accept-all / reject-all methods.
- `internal func escapeXMLAttribute(_:)` (file `XMLAttributeEscape.swift`) — single shared XML attribute escape helper.
- `internal enum DocumentWalker` (file `DocumentWalker.swift`) — `walkAllParagraphs(in:visit:)` and `findUnrecognizedChild(in:name:idMarker:)` cross-part walker.

## [0.19.4] - 2026-04-26 (rolled into v0.19.5; never tagged — see "Skipped versions" above)

### Fixed — 6 P0 + 2 P1 from PsychQuant/che-word-mcp#56 round 3 verify

The v3.13.3 release shipped on top of v0.19.3 and went through a third 6-AI cross-verification round (https://github.com/PsychQuant/che-word-mcp/issues/56#issuecomment-4321007538). Five of six reviewers (logic / regression / security / codex / devil's advocate) returned BLOCK — the R2 fixes themselves introduced 6 new P0 regressions in 4 of 4 batches (anti-pattern: "fixes that save absence but break preserve-order / sync mutation paths"). v0.19.4 closes those 6 P0 plus 2 P1 follow-ups via the spectra change `che-word-mcp-issue-56-r3-stack-completion`. Each fix shipped as an independent commit with its own failing-test → fix → scoped Codex verify gate, breaking the bundle-and-regress cycle.

#### R3-NEW-1 — Hyperlink mutation API round-trips on source-loaded hyperlinks

`Hyperlink.toXML()` now compares `children`-derived run text against `runs` text. Equal → walk `children` (preserves R2 P0-3 source-order between runs and non-run children). Different → walk `runs` (R3-NEW-1: edits via `replaceText` / `updateHyperlink` / `text` setter become visible). Pre-fix v0.19.3 always preferred `children` so source-loaded hyperlink edits silently no-op'd on save.

#### R3-NEW-2 — Paragraph-level `<w:sdt>` round-trips at source position

New `ContentControl.position: Int = 0` field. `DocxReader.parseParagraph` passes `childPosition` to `SDTParser.parseSDT`. `Paragraph.toXMLSortedByPosition` adds contentControls with `position > 0` to the sorted positioned-entry list; legacy post-content emit only fires for `position == 0` (API-built). `hasSourcePositionedChildren` includes `contentControls.position > 0` so SDT-only paragraphs route to sort path. Pre-fix `<w:r>A</w:r><w:sdt>X</w:sdt><w:r>B</w:r>` round-tripped as A → B → SDT.

#### R3-NEW-3 — `insertComment` emits anchor markers on source paragraphs with existing comments

Per-id gate replaces blanket `if commentRangeMarkers.isEmpty` in `Paragraph.toXMLSortedByPosition`. Computes `Set(commentRangeMarkers.map { $0.id })` once; emits `<w:commentRangeStart>` / `<w:commentRangeEnd>` / `<w:commentReference>` for commentIds NOT covered by a source marker. Pre-fix the blanket gate skipped the entire legacy emit when source had any commentRangeMarker → new commentIds via `insertComment` lost ALL their anchor output (comment side-bar showed comment but no scope highlight).

#### R3-NEW-4 — Mixed-content revision wrappers populate both raw and typed representations + accept/reject support

`Revision` gains `isMixedContentWrapper: Bool = false` field. All 4 hasNonRunChild branches in DocxReader (ins/del/moveFrom/moveTo) now append a typed `Revision` with the flag alongside the raw `unrecognizedChildren` capture. `Document.acceptRevision` / `.rejectRevision` detect the flag and delegate to new private `handleMixedContentWrapperRevision` helper that searches body paragraphs (incl. nested table cells) for the matching entry by name + opening-tag-only id match (codex P1 catch: nested bookmarks/comments with same id no longer false-hit), then either replaces rawXML with extracted inner content (accept on insertion/moveTo, reject on deletion/moveFrom) or removes the entry entirely. Pre-fix the typed Revision was missing → MCP `get_revisions` / `accept_*_revision` / `reject_*_revision` tools couldn't see the wrapper but raw XML still emitted on save.

#### R3-NEW-5 — `nextBookmarkId` calibration recurses into tables, headers, footers, footnotes, endnotes

Replaced the early body-only top-level `.paragraph` scan with a comprehensive post-load calibration after all parts are parsed. New private `walkAllParagraphs(in:visit:)` recursively visits paragraphs across body (recursing into tables, nested tables — codex P1 catch — and content controls), headers, footers, footnotes, and endnotes. Pre-fix calibration ran before headers/footers/notes were even loaded AND only saw body top-level paragraphs → bookmarks in table cells / headers / etc. caused false-success calibration → `insertBookmark` allocated id 1 → silent collision with source ids.

#### R3-NEW-6 — XML attribute escape closes rStyle injection sink + audit

Added `fileprivate func escapeXMLAttribute(_ s: String) -> String` in Run.swift (5 chars: `& < > " '`). Routed `RunProperties.toXML` rStyle / color / fontName emits through it. Codex P1 catch: `RunProperties.toChangeXML` (parallel emit path inside `<w:rPrChange>`) was emitting `color` unescaped while `fontName` was already escaped — fixed parity. Audit table comment block in `Issue56R3StackTests.swift` enumerates every direct-emit site across Run.swift / Hyperlink.swift / Paragraph.swift / Footer.swift / Revision.swift / DocxWriter.swift, marked ESCAPED or SAFE-BY-CONSTRUCTION. Pre-fix a malicious source `<w:rStyle w:val='x"/><inj/><w:dummy w:val="y'/>` round-tripped as 3 sibling elements → Word schema reject (and a confidentiality vector if attacker could inject revision authors etc.). Future cleanup (out of R3 scope): consolidate the 6 parallel escape helpers into a shared `XMLEscape.swift`.

### Fixed — 1 P1 follow-up

- **D-3** — `parseHyperlink` now also captures `XMLElement.namespaces` (the separate xmlns: declaration collection in Foundation XMLElement) into `rawAttributes` with the `xmlns:` prefix prepended. Pre-fix `<w:hyperlink xmlns:vendor="..." vendor:custom="x">` round-tripped with the prefixed attribute but lost its namespace declaration → Word schema rejected the unbound prefix.

### Breaking changes

- **D-8 / Hyperlink.id format change introduced in v0.19.3 (P1-7)** — `Hyperlink.id` now follows the format `<rId-or-anchor-or-hl>@<position>` (e.g. `rId5@7`) instead of the v0.19.2 format `<rId-or-anchor-or-hl>` (e.g. `rId5`). This change shipped in v0.19.3 to give two hyperlinks sharing one `r:id` distinct ids — a correctness fix for MCP tools that find / edit / delete hyperlinks by id. **Callers that stored pre-v0.19.3 ids and look them up after upgrade will get nil**. Mitigation: re-parse documents under v0.19.4 to refresh the id cache. No alias / backwards-compatibility shim is provided; the v0.19.3 release is < 7 days old at write time so very little production storage exists.

### Tests

- 12 new tests in `Tests/OOXMLSwiftTests/Issue56R3StackTests.swift` covering each P0 / P1 fix
- Suite total: 582 tests pass / 1 skipped / 0 failures (570 v0.19.3 baseline + 12 new R3 tests, zero regressions)
- Codex CLI scoped verify ran after each P0 fix; flagged P1s fixed inline before commit (R3-NEW-4 nested w:id substring match, R3-NEW-5 nested-table walker, R3-NEW-6 toChangeXML color escape parity)

## [0.19.3] - 2026-04-26

### Fixed — 8 P0 + 3 must-fix P1 from PsychQuant/che-word-mcp#56 round 2 verify

The v3.13.2 release shipped on top of v0.19.2 and went through a second 6-AI cross-verification round (https://github.com/PsychQuant/che-word-mcp/issues/56#issuecomment-4320157395). Five of six reviewers (codex / logic / regression / security / devil's advocate) returned BLOCK; the requirements reviewer's PASS was overturned on every F1–F4 with concrete refutations. v0.19.3 closes the 8 P0 + 3 must-fix P1 in four batches.

#### Batch A — Hyperlink suite

- **P0-1** — `Hyperlink.external` / `.internal` produce hyperlink-styled runs again. v0.19.2 walked `runs` directly without applying the legacy hardcoded `<w:rStyle Hyperlink>` / `0563C1` color / single underline → all 5 MCP `insert_*hyperlink` tools rendered without visual styling. New `RunProperties.rStyle` field carries the style reference; `Hyperlink.makeStyledRun(text:)` builds runs with the Hyperlink character style + blue + underline.
- **P0-2** — `parseHyperlink` no longer lists `w:tgtFrame` / `w:docLocation` as recognized attributes. They had no typed `Hyperlink` field and the writer never emitted them, so v0.19.2 silently dropped vendor / browser-target attributes on round-trip. They now flow into `rawAttributes` and emit via the alphabetical loop.
- **P0-3** — New `HyperlinkChild` enum + `Hyperlink.children: [HyperlinkChild]` preserve source-document order between `<w:r>` and non-run children. `<w:hyperlink><w:r>A</w:r><w:sdt>X</w:sdt><w:r>B</w:r></w:hyperlink>` now round-trips A → SDT → B (was A → B → SDT). Reader populates `children` while keeping `runs` / `rawChildren` for backward-compat reads.
- **P1-7** — `Hyperlink.id` is now `<rId-or-anchor-or-hl>@<position>` so two hyperlinks sharing a single relationship id (legitimate when two anchors target the same URL) parse with distinct ids. MCP tools that find / edit / delete hyperlinks by id again hit the right hyperlink.

#### Batch B — Sort path completeness

- **P0-4** — `Paragraph.toXMLSortedByPosition` now emits `contentControls` after the position-indexed children. Source paragraphs with `<w:sdt>` + any positioned child no longer drop the SDT on save.
- **P0-5** — Sort path also emits the legacy `commentIds` / `footnoteIds` / `endnoteIds` / `hasPageBreak` / legacy `bookmarks` collections. The pre-fix doc-comment claimed they would emit AFTER but the code dropped them entirely — `insert_comment` / `insert_footnote` on a bookmarked source paragraph silently lost the comment / footnote. Each legacy collection is skipped only when its positioned variant is non-empty (Reader keeps both populated; emitting both would double the markers).
- **P0-8** — `hasSourcePositionedChildren` now also treats any run or hyperlink with `position > 0` as a source-loaded signal. Pre-fix a source paragraph with `<w:r>A</w:r><w:hyperlink>L</w:hyperlink><w:r>B</w:r>` (no other markers) routed to legacy → "A B L" output. Now routes to sort.

#### Batch C — Revision wrapper coverage

- **P0-6** — Reader always appends a `Revision` entry on `<w:ins>` / `<w:del>` / `<w:moveFrom>` / `<w:moveTo>` regardless of whether the inner concatenated text is empty. Pre-fix the `if !insertedText.isEmpty` guard meant insertions of pure non-text content (`<w:tab/>`, `<w:br/>`, `<w:drawing>`, `<w:fldChar>`) yielded no revision, the sort-path grouping fell back to a naked `<w:r>`, and the wrapper silently disappeared (regression vs the v3.12.0 #45 Track Changes feature).
- **P0-7** — When a revision wrapper contains any non-`<w:r>` direct child (`<w:hyperlink>`, `<w:sdt>`, `<w:fldSimple>`, `<mc:AlternateContent>`), Reader now captures the whole wrapper verbatim into `unrecognizedChildren` at the wrapper's position. The sort path emits it byte-for-byte. Track Changes flow "user inserted a hyperlink while review mode was on" now round-trips with the hyperlink intact. Trade-off: wrappers with mixed content lose the per-run typed editable surface; pure-run wrappers retain full typed editing as before. Helper: new private `DocxReader.hasNonRunChild(_:)`.

#### Batch D — Bookmark hardening

- **P1-1** — `DocxReader.read(from:)` scans `paragraph.bookmarks` and `paragraph.bookmarkMarkers` after parsing the body, computes the max source bookmark id, and bumps `WordDocument.nextBookmarkId` past it. Pre-fix the counter started at 1 regardless of source content; F2's marker sync turned the previously-latent collision (silent drop) into an active bug (silent overwrite, possible Word schema-reject). `nextBookmarkId` is now `internal`.
- **P1-4** — `appendBookmarkSyncingMarkers` only appends to `bookmarkMarkers` when the paragraph already routes to sort path (`hasSourcePositionedChildren == true`). Pure API-built paragraphs keep `bookmarks`-only emit, restoring the v3.12.0 wrap-around semantic where `addBookmark("foo")` spans the existing run text. F2 had blindly added markers everywhere, downgrading API-path bookmarks to zero-width point bookmarks at paragraph end. `Paragraph.hasSourcePositionedChildren` is now `internal`.

### Test coverage

570 tests pass (557 from v0.19.2 + 13 new in `Issue56RoundtripCompletenessTests`), 1 skipped, 0 failures. New tests are end-to-end Reader → Writer round-trip cases (vs v0.19.2's API-only construction) so the Reader-side filter bugs (P0-2 / P0-7) are now exercised.

### No breaking changes for downstream

All new fields default to empty (`children: []`, `rStyle: nil`, etc.). API-built objects produce byte-equivalent pre-fix output for round-trip-safe paths; the only intentional behavior changes (P0-1 styling, P1-4 bookmark semantics) restore v3.12.0 contracts that v0.19.2 had silently broken. Existing 218+ MCP tools in che-word-mcp are unchanged.

### Follow-up items deferred

The round 2 verify also surfaced 5 non-must-fix P1, 9 P2, and 8 P3 items (devil's advocate NEW-A through NEW-G plus security defense-in-depth and pre-existing non-blocking items). These will be filed as separate follow-up issues for staged remediation; v0.19.3 ships only the must-fix subset to land #56's lossless round-trip contract on the v3.13.x release line.

## [0.19.2] - 2026-04-26

### Fixed — 4 blocking findings from PsychQuant/che-word-mcp#56 verification (F1–F4)

The v3.13.1 release of che-word-mcp shipped on top of v0.19.1 and went through 6-AI cross-verification. Five of the six reviewers initially marked the four #56 Expected requirements as FULLY addressed; the Devil's Advocate reviewer downgraded all four to PARTIAL after surfacing 4 blocking sub-issues that the existing smoke tests didn't cover (concat-text SHA256 + element-count parity miss run-property loss, marker desync, revision wrapper drop, and per-part namespace strip). v0.19.2 fixes all four.

**F1 — `Hyperlink.toXML()` ignored Reader-collected runs / rawAttributes / rawChildren** (`Sources/OOXMLSwift/Models/Hyperlink.swift:151-187`). v0.19.0 added the hybrid model fields but the writer kept emitting a hardcoded single-run blue-underlined `Hyperlink`-styled `<w:r>` regardless of source. Inner-run formatting (bold/italic/color/font), unmodeled `<w:hyperlink>` attributes (`w:tgtFrame`, `w:docLocation`, vendor extensions), and non-Run direct children (nested SDT) were all silently dropped on every round-trip. Rewritten to iterate `runs` (preserving each `RunProperties` via `Run.toXML()`), emit `rawAttributes` alphabetically (skipping any name colliding with a typed attribute), and append `rawChildren` verbatim. Empty-`runs` (API-built path) falls back to the legacy hardcoded styled-run template.

**F2 — `addBookmark` / `deleteBookmark` did not sync `bookmarkMarkers`** (`Sources/OOXMLSwift/Models/Document.swift:1607-1644`). Source-loaded paragraphs always go through `Paragraph.toXMLSortedByPosition` because their existing markers are non-empty. The mutation API only updated `Paragraph.bookmarks` (the typed list), not `Paragraph.bookmarkMarkers` (the position-indexed list the writer actually consults). Result: new bookmarks added via API silently dropped on save (typed entry created but writer never emitted them); deletes left zombie `<w:bookmarkStart w:id="N" w:name=""/>` markers because `emitBookmarkMarker` looked up the deleted bookmark's name via `?? ""` fallthrough. New helper `appendBookmarkSyncingMarkers(to:bookmark:)` is the single insertion entry point, computing `position = max(existing positions) + 1` (start) and `+2` (end) so new bookmarks always land at paragraph tail. `deleteBookmark` now also runs `bookmarkMarkers.removeAll { $0.id == removed.id }`.

**F3 — `<w:ins>` / `<w:del>` / `<w:moveFrom>` / `<w:moveTo>` Reader did not assign `position` or `revisionId` to inner runs** (`Sources/OOXMLSwift/IO/DocxReader.swift:597-684`). Pre-fix, every wrapper-internal run was appended to `paragraph.runs` with `position` defaulting to 0, so source-loaded paragraphs with revision tracking sorted all inserted/moved runs to paragraph front (NEW-1 in the verify report — devil's advocate caught this by comparing line 590 normal `<w:r>` handling against lines 606/650/672). The wrapper element itself was also dropped because the sort-by-position emit (Paragraph.swift:418) emitted runs individually rather than re-grouping by `revisionId`. Two-part fix: Reader assigns `parsedRun.position = childPosition` AND `parsedRun.revisionId = revId` in all four cases; Writer's sort path now uses a `PositionedEntry` enum (`.run(Run)` vs `.xml(String)`) so a post-sort pass can group consecutive `.run` entries with the same `revisionId` and wrap them in a single `<w:ins>` / `<w:del>` / `<w:moveFrom>` / `<w:moveTo>` block via `Revision.toOpeningXML()` / `toClosingXML()`. Track Changes round-trips intact for source-loaded documents (the v3.12.0 #45 feature now composes correctly with #56's source-load infrastructure).

**F4 — Namespace preservation only covered `word/document.xml`** (NEW-2 in the verify report). v0.19.0's `documentRootAttributes` plumbing solved the unbound-prefix problem for the document body but headers, footers, footnotes, and endnotes each still used hardcoded namespace templates. NTPU thesis-class documents have 6 headers with VML watermarks that frequently declare `mc`/`wp`/`w14`/`w15` beyond the hardcoded 5-namespace template — those declarations were silently dropped on round-trip. New `ContainerRootTag.render(elementName:attributes:)` helper generalizes the document-level `renderDocumentRootOpenTag` pattern over any container root element. Reader's `parseDocumentRootAttributes` is now the thin wrapper over a generalized `parseContainerRootAttributes(from:rootElementOpenPrefix:)`. Each of `Header`, `Footer`, `FootnotesCollection`, `EndnotesCollection` gains a `rootAttributes: [String: String]` field; Reader populates from raw bytes per part; their `toXML()` methods consult it and fall back to element-specific defaults (header/footer = 5-namespace VML template; footnotes/endnotes = 2-namespace minimal) when empty. API-built parts emit byte-identical pre-fix output; source-loaded parts round-trip every declaration verbatim.

### Test coverage

557 tests pass (548 from v0.19.1 + 9 new in `Issue56FollowupTests`), 1 skipped, 0 failures. New tests cover each F1–F4 fix plus their fallback behaviors (empty-runs hyperlink, empty-rootAttributes container).

### No breaking changes

All new fields default to empty (`runs: []`, `rootAttributes: [:]`, `revisionId: nil`). API-built objects produce byte-identical pre-fix output. Existing 218+ MCP tools in che-word-mcp are unchanged.

## [0.19.1] - 2026-04-25

### Fixed — pPr double-emission on Phase 4 sort-by-position round-trip (Refs PsychQuant/che-word-mcp#56 follow-up)

Found while running the v0.19.0 round-trip suite against a 570-paragraph NTPU master's thesis fixture. The new sort-by-position emit added in v0.19.0 silently captured `<w:pPr>` into `Paragraph.unrecognizedChildren` because `parseParagraph` had no explicit `case "pPr": break` branch — pPr was already consumed by the dedicated `parseParagraphProperties(from:)` call above the child walker, but it then fell into `default` in the switch and got captured AGAIN as a verbatim raw-carrier.

Symptom: `<w:pPr>` got written twice on save (once via the legacy pPr block at the top of `Paragraph.toXMLSortedByPosition`, once verbatim from `unrecognizedChildren`). xmllint accepts the duplicate (Word ignores the second pPr per ECMA-376), and text content remained intact, but `unrecognizedChildren` count ballooned every round-trip (NTPU thesis: 799 → 1333 entries, +534 spurious pPr captures across 570 paragraphs). File size grew by ~1 KB per paragraph per round-trip.

Fix: 1-line case branch — `case "pPr": break` — stops pPr falling through to the default raw-capture path. Source data: 799 → 229 entries (only oMath, the legitimate raw-carriers). Round-trip: 229 → 229 ✓.

Regression test: `testParseParagraphSkipsPPrInChildWalker` asserts `parseParagraph` never adds `<w:pPr>` to `unrecognizedChildren`.

### Test coverage

548 tests pass (1 skipped, 0 failures).

## [0.19.0] - 2026-04-25

### Fixed — `document.xml` lossless round-trip (Refs PsychQuant/che-word-mcp#56, P0)

Fixes the critical regression where `save_document` silently corrupts `word/document.xml` on every body-mutating MCP call. A trivial `open → insert_paragraph → save` on a typical Word document used to strip 32 of 34 namespace declarations from the `<w:document>` root, wipe 100% of `<w:bookmarkStart>` bookmarks, and drop 354 `<w:t>` text nodes living inside `<w:hyperlink>` / `<w:fldSimple>` / `<mc:AlternateContent>` wrappers (TOC anchor text, cross-reference placeholders, table caption SEQ fields, math notation). All other 41 OOXML parts byte-equal — only `document.xml` itself became invalid.

Three orthogonal root causes addressed in 5 phases (all bundled — splitting Phase 1 alone would change the failure mode from "XML invalid" to "XML valid but text/bookmarks gone", a worse UX):

**Phase 1 — Document root namespace preservation.**
- New `WordDocument.documentRootAttributes: [String: String]` capturing every `xmlns:*` declaration plus `mc:Ignorable` from the source `<w:document>` root.
- `DocxReader.read(from:)` extracts attributes via raw-bytes parser (bypasses libxml2's silent xmlns drop on unused prefixes).
- `DocxWriter.writeDocument` rebuilds the open tag from the captured map, falling back to `xmlns:w` + `xmlns:r` only when the dictionary is empty (preserves create-from-scratch behavior).

**Phase 2 — Bookmark Reader parsing + range markers.**
- New `BookmarkRangeMarker` (kind: start/end, id, position) on `Paragraph.bookmarkMarkers`.
- `DocxReader` paragraph walker now parses `<w:bookmarkStart w:id w:name/>` and `<w:bookmarkEnd w:id/>` (previously zero hits — the `Bookmark` model existed but was write-only).

**Phase 3 — Wrapper hybrid model (typed editable surface + raw passthrough).**
- `Hyperlink` gains `runs: [Run]`, `rawAttributes: [String: String]`, `rawChildren: [String]`, `position: Int`. Existing `text: String` becomes a computed property `runs.map { $0.text }.joined()` for backward compat with existing call sites (zero breaking changes for downstream consumers reading `hyperlink.text`).
- New `FieldSimple` model: `instr: String` + `runs: [Run]` + `rawAttributes` + `position`. `w:instr` whitespace preserved exactly.
- New `AlternateContent` model: `rawXML: String` (verbatim source for byte-equivalent emit) + `fallbackRuns: [Run]` (typed editable mirror of `<mc:Fallback>` content). Documented Non-Goal: edits to `fallbackRuns` may diverge from `<mc:Choice>` content (Word reconciles per its own rules).
- `DocxReader.parseHyperlink` / `parseFieldSimple` / `parseAlternateContent` helpers.

**Phase 4 — `<w:p>` child schema completeness + Writer sort-by-position emit.**
- 6 new raw-carrier types: `CommentRangeMarker`, `PermissionRangeMarker`, `ProofErrorMarker`, `SmartTagBlock`, `CustomXmlBlock`, `BidiOverrideBlock` (each with `position: Int`).
- New `Paragraph.unrecognizedChildren` fallback collection — any `<w:p>` direct child whose local name does not match any typed parser or registered raw-carrier survives the round-trip with verbatim XML + position. Surfaces ECMA-376 spec gaps without silent drops.
- `Run.position: Int` added so direct-child runs participate in sort-by-position emit.
- `Paragraph.toXML()` refactored: when any source-loaded marker collection is non-empty, dispatches to `toXMLSortedByPosition()` which collects `(position, xml)` tuples from every parallel array, sorts by position, and emits in source order. API-built paragraphs (no source markers) keep the legacy emit path — zero breaking changes for existing tools.
- Reader paragraph walker uses `defer { childPosition += 1 }` to assign source-document order positions to every direct child (typed or raw).

**Phase 5 — Test fixture dual-track + tool-mediated edit safety.**
- New `LosslessRoundTripFixtureBuilder` synthesizes a 50–100 KB `.docx` exercising every code path the new Reader / Writer pair must preserve (5+ bookmarks, 3 hyperlinks, 2 fldSimple, 1 AlternateContent, 12 xmlns + mc:Ignorable on root, mixed runs/wrappers across 6 paragraphs).
- New `DocumentXmlLosslessRoundTripTests` (8 tests) covering namespace preservation, bookmark round-trip, hyperlink runs + raw passthrough, FieldSimple SEQ caption, AlternateContent math block, comment range markers, interleaved-children sort-by-position emit, and the builder fixture as a CI regression.
- `WordDocument.replaceText` extended to walk `Hyperlink.runs`, `FieldSimple.runs`, `AlternateContent.fallbackRuns` so tool-mediated edits inside structural wrappers SHALL apply (no silent failure — the v3.12.0 `replace_text` regression where edits inside hyperlinks / SEQ Table captions / math fallbacks returned success but produced no change).

**Test coverage:** 546 ooxml-swift tests pass with 0 failures (8 new tests added by this change).

**Breaking changes:** None. `Hyperlink.text` is now a computed property but observationally equivalent for read access; the setter collapses to single-Run (matching pre-fix multi-run-overwrite behavior).

## [0.18.0] - 2026-04-25

### Added — Track Changes write-side: 5 revision generators + writer extensions (Refs PsychQuant/che-word-mcp#45)

Closes the WRITE-side gap for tracked revisions. Reader infrastructure already populated `paragraph.revisions` from `<w:ins>` / `<w:del>` / `<w:moveFrom>` / `<w:moveTo>` / `<w:rPrChange>` / `<w:pPrChange>` markup, but the writer ignored `paragraph.revisions` entirely — meaning programmatically-added revisions never reached the saved `.docx`. v0.18.0 fills the gap with 6 new `WordDocument` methods and a writer that emits proper revision wrappers.

**New `WordDocument` methods:**

- `allocateRevisionId() -> Int` — scans `revisions.revisions` for max id; returns max+1 (or 1 when empty). Mirrors v0.15.0 `allocateSdtId()` deterministic max+1 pattern.
- `insertTextAsRevision(text:atParagraph:position:author:date:) throws -> Int` — splits the run at `position` (preserves prior + post text + formatting), inserts a new `<w:ins>`-wrapped run, returns allocated revision id.
- `deleteTextAsRevision(atParagraph:start:end:author:date:) throws -> Int` — splits straddling runs at boundaries; tags middle runs with the revision id; writer wraps them with `<w:del>` and substitutes `<w:t>` → `<w:delText>`.
- `moveTextAsRevision(fromParagraph:fromStart:fromEnd:toParagraph:toPosition:author:date:) throws -> (fromId: Int, toId: Int)` — allocates two adjacent ids (`N` and `N+1`); emits paired `<w:moveFrom>` (source) and `<w:moveTo>` (destination). Single-paragraph moves rejected as out of scope.
- `applyRunPropertiesAsRevision(atParagraph:atRunIndex:newProperties:author:date:) throws -> Int` — replaces run formatting; captures previous `RunProperties` for the revision; writer emits `<w:rPrChange>` inside `<w:rPr>`.
- `applyParagraphPropertiesAsRevision(atParagraph:newProperties:author:date:) throws -> Int` — replaces paragraph formatting; captures previous `ParagraphProperties`; writer emits `<w:pPrChange>` inside `<w:pPr>`.

All 5 generators guard `isTrackChangesEnabled()` and throw new `WordError.trackChangesNotEnabled` when off — no auto-enable side effect (per design decision: explicit `enable_track_changes` required to avoid hidden state mutation).

All 5 generators resolve author via 3-tier fallback: explicit non-empty arg → `revisions.settings.author` → `"Unknown"`. They mark `word/document.xml` dirty.

**New typed Run/Paragraph fields linking runs to revisions:**

- `Run.revisionId: Int?` — id of the wrapping `<w:ins>` / `<w:del>` / `<w:moveFrom>` / `<w:moveTo>` revision.
- `Run.formatChangeRevisionId: Int?` — id of the format-change revision whose `previousFormat` describes this run's pre-mutation state. Orthogonal to `revisionId`.
- `Paragraph.paragraphFormatChangeRevisionId: Int?` + `Paragraph.previousProperties: ParagraphProperties?` — pair carrying paragraph-level format change metadata.

**Writer extensions in `Paragraph.toXML()`:**

- Groups consecutive runs sharing the same `revisionId` and emits a single `<w:ins>` / `<w:del>` / `<w:moveFrom>` / `<w:moveTo>` wrapper around the group (instead of one wrapper per run). Multi-run wrapping produces `<w:ins ...><w:r>A</w:r><w:r>B</w:r><w:r>C</w:r></w:ins>`.
- Substitutes `<w:t>` with `<w:delText xml:space="preserve">` when wrapping deletion-typed revisions.
- Emits `<w:rPrChange>` inside a run's `<w:rPr>` when the run carries `formatChangeRevisionId` matching a `.formatChange` revision in `paragraph.revisions`.
- Emits `<w:pPrChange>` inside a paragraph's `<w:pPr>` when the paragraph carries `paragraphFormatChangeRevisionId` matching a `.paragraphChange` revision.

**WordError additions (additive):**

- `case trackChangesNotEnabled` — guard violation when `as_revision: true` is passed but track changes is off.

### Tests

- 525 baseline + 13 net new = **538/538 tests pass**, 1 skipped:
  - 24 `RevisionGenerationTests` covering all 5 generators (insertion run-splitting, deletion boundary splits, move adjacent-id allocation, format change rPrChange/pPrChange emission, error guards, author fallback chain)
  - Multi-run wrapping verified produces single `<w:ins>` containing 3 `<w:r>` siblings
  - `<w:t>` → `<w:delText>` substitution verified

### Migration

Additive release — no API changes to existing methods. `Paragraph.toXML()` behavior for runs without `revisionId` set is unchanged. New `Run` fields default to `nil` so programmatic `Run(text:)` constructions remain Equatable-equal to previous releases.

## [0.14.0] - 2026-04-24

### Added — Run rawElements carrier for unknown OOXML elements (Refs PsychQuant/che-word-mcp#52)

`Run` typed model gains `public var rawElements: [RawElement]?` field carrying verbatim XML for unknown direct children of `<w:r>` (e.g., `<w:pict>` VML watermarks, `<w:object>` OLE embeds, `<w:ruby>` annotations). New `public struct RawElement: Equatable` with `name: String` + `xml: String` fields.

`DocxReader.parseRun` now collects unknown children into `rawElements` (recognized typed kinds — `rPr`, `t`, `drawing`, `oMath`, `oMathPara` — are skipped because they're already captured into typed fields). When no unknown children, `rawElements` stays `nil` (NOT empty array) so programmatic Run construction without rawElements remains Equatable-equal to reader-loaded Runs.

`Run.toXML()` emits typed children in fixed order, then appends rawElements verbatim before `</w:r>`. Empty-text Runs with rawElements (typical NTPU watermark structure: `<w:r>` → `<w:rPr>` → `<w:pict>` with no `<w:t>`) suppress the synthetic empty `<w:t>` to avoid spurious empty text nodes in Word output.

### Added — Header/Footer namespace declarations for VML preservation

`Header.toXML()` and `Footer.toXML()` now declare `xmlns:v` (VML), `xmlns:o` (Office), `xmlns:w10` (Word) at the `<w:hdr>` / `<w:ftr>` root so descendant `<v:shape>` / `<o:lock>` / `<w10:wrap>` resolve when the saved `header*.xml` is re-read. Required for round-trip of preserved VML watermarks.

### Added — `updateAllFields(isolatePerContainer:)` opt-in flag (Refs #52, deferred from #54)

`WordDocument.updateAllFields` gains `isolatePerContainer: Bool = false` parameter. Default `false` preserves prior global-counter-sharing behavior across all container families. When `true`, each container family (body / each header / each footer / footnotes collection / endnotes collection) maintains independent SEQ counter dicts — body's `Figure 3` does NOT increment a header's `Figure` counter.

The returned `[String: Int]` reflects body's final counter state. Per-container final values are reflected in the SEQ runs' rawXML (callers needing per-container introspection can inspect the cached `<w:t>` values directly).

### Tests

- 408 baseline + 7 net new = **451/451 tests pass** across 3 phases:
  - 3 `RunRawElementPreservationTests` (Phase A: VML round-trip, multiple unknowns, Equatable nil-equivalence)
  - 2 `HeaderFooterByteEqualityWithVMLTests` (Phase B: updateAllFields preservation, updateHeader documented limitation)
  - 2 `UpdateAllFieldsCounterIsolationTests` (Phase C: default sharing, isolation flag)
- 6 XCTSkip (pre-existing fixture-gated tests + 1 documented updateHeader API design boundary)

### Compatibility

- **Public API additions** — all opt-in; no removed APIs:
  - `RawElement` struct (new)
  - `Run.rawElements` field (default nil)
  - `updateAllFields(isolatePerContainer:)` parameter (default false)
- **Behavior changes**:
  - DocxReader: previously-dropped unknown Run children now preserved in `Run.rawElements`. Round-trip now byte-preserves VML watermarks / OLE objects in headers/footers. Programmatic callers comparing `Run` instances post-parse will see populated `rawElements` where previously the data was silently lost
  - Header/Footer XML root tags now declare additional namespaces — observable in saved `word/header*.xml` / `word/footer*.xml` byte content
- `DocxReader.parseRun` access changed from `private` to `internal` for `@testable` consumers

### Refs

- PsychQuant/che-word-mcp#52 — Header.toXML raw-XML preservation (closes the v3.7.1 known-limitation paragraph)

## [0.13.5] - 2026-04-24

### Added — Path traversal security baseline (closes che-word-mcp#55)

`isSafeRelativeOOXMLPath()` validator at `Sources/OOXMLSwift/IO/PathValidator.swift`. Defense-in-depth: applied at parse boundary (DocxReader header/footer rel loops) AND at property setters (`Header.originalFileName` / `Footer.originalFileName` `didSet` observers).

Pre-fix, `_rels/document.xml.rels` `Target` attribute flowed unsanitized into `URL.appendingPathComponent` (does NOT normalize `..`) AND into `Header.originalFileName` used at write time. Malicious .docx could read OR write outside `word/` directory at user UID.

Validator rejects: empty / >256 chars (DoS guard), absolute paths, parent traversal (including URL-encoded `%2e%2e` `%2f` `%5c`), control chars (NUL, newlines, < 0x20, 0x7F). Accepts non-ASCII Unicode in printable range.

10 new `PathTraversalSecurityTests` scenarios.

### Added — Multi-instance Header/Footer auto-suffix (closes che-word-mcp#53)

`addHeader()` / `addHeaderWithPageNumber()` / `addFooter()` / `addFooterWithPageNumber()` now call new private `allocateHeaderFileName(for:)` / `allocateFooterFileName(for:)` helpers that auto-suffix the fileName. Multi-instance `.default`-type adds now produce `header1.xml`, `header2.xml`, `header3.xml` instead of all collapsing to `header1.xml`.

Pre-fix: latent bug where `addHeader()` × 2 with default type both produced `Header.fileName == "header1.xml"`. On disk: h2 overwrote h1; in #42 dirty-bit Sets they collapsed to one entry.

Reader-loaded path unchanged (`originalFileName` already populated from `rel.target`). 7 new `MultiInstanceHeaderFooterTests` scenarios.

### Changed — `updateAllFields` coverage extensions (closes che-word-mcp#54)

Bundles 4 sub-findings from #42 verification:

1. **Regex schema-drift detection**: `rewriteCachedResult` now returns `(rewritten: String, didMatch: Bool)`. When `didMatch == false` and a SEQ field with `cachedResultRunIdx` was present, emit stderr warning that cached value may be stale.
2. **Counter-scope documentation**: `updateAllFields()` doc-comment explains SEQ counters are global across body / headers / footers / notes (differs from Word F9 per-section isolation). `isolatePerContainer` flag deferred.
3. **Header-SEQ no-op test**: snapshot-delta assertion confirms updateAllFields adds nothing to modifiedParts when cached value already matches.
4. **Footnote/endnote round-trip tests**: byte-equality verification for note-parts mirrors v0.13.4's header round-trip.

3 new `UpdateAllFieldsCoverageTests` scenarios.

### Tests

- 408 baseline + 36 net new = **444/444 tests pass** across 3 issues:
  - 10 `PathTraversalSecurityTests` (#55)
  - 7 `MultiInstanceHeaderFooterTests` (#53)
  - 3 `UpdateAllFieldsCoverageTests` (#54)
  - Plus dirty-bit verify + earlier sessions
- 5 XCTSkip (pre-existing fixture-gated tests)

### Compatibility

- **No public API changes**. `isSafeRelativeOOXMLPath` is the only new public symbol; defaults preserve all prior behavior for existing callers.
- **Behavior changes**:
  - DocxReader silently drops headers/footers with unsafe rel.target (with stderr warning) — was previously vulnerable
  - `addHeader()` × N with default type now produces sequential fileNames — was silently colliding
  - `updateAllFields` emits stderr warnings on regex schema drift — was silent

### Refs

- PsychQuant/che-word-mcp#53, #54, #55 (all opened during #42 verification on 2026-04-24)

## [0.13.4] - 2026-04-24

### Fixed — `updateAllFields` honest dirty-bit propagation (closes che-word-mcp#42)

Pre-v0.13.4 `WordDocument.updateAllFields` (introduced v0.10.0 for SEQ counter recomputation) **unconditionally** marked every header/footer/footnote/endnote path into `modifiedParts` regardless of whether any SEQ field was actually found there:

```swift
// Pre-v0.13.4 (BROKEN):
modifiedParts.insert("word/document.xml")
for header in headers { modifiedParts.insert("word/\(header.fileName)") }
for footer in footers { modifiedParts.insert("word/\(footer.fileName)") }
// ...always, even when no SEQ in any header
```

Once a header path is in `modifiedParts`, overlay-mode `DocxWriter` re-emits it via `Header.toXML()` — which only knows about typed `paragraphs[]` and silently drops VML watermarks, drawings, and any non-paragraph raw XML. Result on NTPU thesis: 3923-byte VML watermark header → 318-byte `<w:p/>` stub. **P0 silent data loss** on every academic template workflow that called `update_all_fields`.

### Architecture

`processParagraph` now returns `Bool` indicating whether any SEQ field's cached result was actually rewritten. Each container (body / headers / footers / footnotes / endnotes) tracks its own dirty bit during the scan. Only containers with a confirmed SEQ rewrite get inserted into `modifiedParts`:

```swift
// v0.13.4+ (CORRECT):
var bodyDirty = false
for i in 0..<body.children.count {
    if processParagraph(&para, ...) { bodyDirty = true }
}
var dirtyHeaderFiles: Set<String> = []
for i in 0..<headers.count {
    var headerDirty = false
    for j in 0..<headers[i].paragraphs.count {
        if processParagraph(&para, ...) { headerDirty = true }
    }
    if headerDirty { dirtyHeaderFiles.insert(headers[i].fileName) }
}
// ... same for footers/footnotes/endnotes ...
if bodyDirty { modifiedParts.insert("word/document.xml") }
for fileName in dirtyHeaderFiles { modifiedParts.insert("word/\(fileName)") }
```

Additionally, `rewriteCachedResult` is now compared with the original — if the rewritten string equals the input (e.g., counter value didn't actually change), no rewrite is recorded.

### Tests

- `WordDocumentUpdateAllFieldsHeaderPreservationTests.swift` (NEW) — 4 scenarios:
  - `testHeaderWithoutSEQNotMarkedDirty` — body has SEQ, header has only paragraphs → header NOT in modifiedPartsView
  - `testHeaderWithSEQIsMarkedDirty` — header contains SEQ → header IS in modifiedPartsView
  - `testFooterWithoutSEQNotMarkedDirty` — same logic mirrors footers
  - `testUpdateAllFieldsNoSEQAnywhereDoesNotAddToModifiedParts` — true no-op snapshot test
- **407/407 tests pass** (was 403 → +4).

### Known limitation (out of scope for this fix)

When a header DOES legitimately contain a SEQ field (rare — e.g., chapter caption in running header), it still re-emits via `Header.toXML()` which strips co-located VML watermarks/drawings. This requires `Header.toXML()` itself to gain raw-XML preservation, which is a separate architectural change. Current behavior degrades gracefully: the dirty-bit fix eliminates the strip in the common case (no SEQ in header), and the rare edge case is logged for follow-up.

### Compatibility

- **Public API unchanged** — `updateAllFields()` signature identical; semantic guarantee strictly stronger.
- **Behavior change**: `modifiedPartsView` after `updateAllFields` is now a strict subset of the pre-v0.13.4 behavior. No existing test relied on the over-eager dirty marking; no consumer should break.

### Refs

- PsychQuant/che-word-mcp#42 — incident report and root-cause audit

## [0.13.3] - 2026-04-24

### Changed — Serial-only OOXML IO + allocator-based image rId assignment (Refs PsychQuant/che-word-mcp#41)

Two coordinated hardening changes for the `che-word-mcp-insert-crash-autosave-fix` SDD:

#### 1. `DocxReader.read` is now fully serial

Pre-v0.13.3 `DocxReader.swift:438-499` used `DispatchQueue.concurrentPerform` for parallel chunk parsing on bodies with `count >= 256`. Worker threads called `parseParagraph`/`parseTable` against shared libxml2-backed `XMLElement` nodes — libxml2 documents are NOT thread-safe at the document level. The comment "shared data 為唯讀" misjudged lazy-property-access risk on `XMLElement` child collections / attribute dicts.

More importantly, `recover_from_autosave` (che-word-mcp v3.6.0) requires re-parsing the same source bytes to produce identical in-memory state. Parallel chunk parsing introduces non-determinism, undermining the entire save-durability stack.

v0.13.3 removes the parallel block. Parsing is now a single serial loop. New regression test `SerialOnlyOOXMLTests.testNoParallelPrimitivesInOOXMLIO` greps `Sources/OOXMLSwift/IO/` for forbidden symbols (`concurrentPerform`, `withTaskGroup`, `DispatchQueue.global`, `DispatchQueue.async`, `Task.detached`) and asserts zero matches — prevents future regressions.

**Trade-off**: `open_document` on large theses (1000+ paragraphs) sees a 200-800ms regression vs v0.13.2. Acceptable for determinism guarantee. New `DocxReaderDeterminismTests` confirms `body.children.count` and paragraph text are identical across 5 repeated reads.

#### 2. `nextImageRelationshipId` delegates to allocator

`Document.swift:1023` `nextImageRelationshipId` was a naïve counter `4 + headers.count + footers.count + images.count`. Defensive hardening: now delegates to `nextRelationshipId` which already consults original rels via `RelationshipIdAllocator` in overlay mode (introduced v0.12.0).

The naïve counter happened to track the typed model in lockstep for most cases (because all 3 collections grow with assignments), but is fragile against any mismatched assignment — e.g., reader-loaded doc with hyperlinks/comments rels not counted by the formula. New tests in `RelationshipIdAllocatorMutationTests` cover reader-loaded doc + sequential insert + initializer-built doc baseline.

### Tests

- 397 baseline + 6 new = **403/403 tests pass**:
  - `RelationshipIdAllocatorMutationTests` (4 scenarios): reader-loaded non-collision, header+image collision regression, sequential inserts, initializer-built rId4 baseline
  - `SerialOnlyOOXMLTests` (1 scenario): grep-based regression test for parallel primitives
  - `DocxReaderDeterminismTests` (1 scenario): 300-paragraph fixture × 5 reads = identical output

### Compatibility

- **No public API change** — both changes are internal refactors. `DocxReader.read` signature unchanged; `nextImageRelationshipId` is internal.
- **Behavior change**: `open_document` perf regression on large bodies (200-800ms one-time cost per session). `nextImageRelationshipId` may now return higher rIds in edge cases (still collision-free, just not the lowest available).

### Refs

- PsychQuant/che-word-mcp#41 — sequential 3rd insert crash investigation
- Phase B of `che-word-mcp-insert-crash-autosave-fix` Spectra change

## [0.13.2] - 2026-04-23

### Fixed — Atomic-rename save (closes che-word-mcp#36)

Pre-v0.13.2 `DocxWriter.write(_:to:)` deleted the target file BEFORE computing the new bytes:

```swift
// Pre-v0.13.2 (BROKEN):
if FileManager.default.fileExists(atPath: url.path) {
    try FileManager.default.removeItem(at: url)        // ← STEP A: delete original
}
let data = try writeData(document)                      // ← STEP B: any throw here = data loss
try data.write(to: url)                                 // ← STEP C: non-atomic write
```

Three failure modes:
1. **Throw at STEP B** → original deleted, no recovery (the bug behind che-word-mcp#36 incident).
2. **Throw at STEP C** → file is partial / zero-byte.
3. **SIGKILL between A and C** → file gone, no `.bak`, no rollback.

### Architecture

`write(_:to:)` now follows the atomic-rename pattern used by every durable file system writer:

```swift
// v0.13.2+ (CORRECT):
let data = try (compute new bytes — overlay or scratch mode)
let tempURL = url.appendingPathExtension("tmp.\(UUID().uuidString)")
defer { try? FileManager.default.removeItem(at: tempURL) }     // cleanup on throw
try data.write(to: tempURL)
let handle = try FileHandle(forWritingTo: tempURL)
try handle.synchronize()                                       // fsync
try handle.close()
_ = try FileManager.default.replaceItemAt(url, withItemAt: tempURL,
                                          backupItemName: nil, options: [])
```

Properties:
- **Atomicity** — `replaceItemAt` uses POSIX `rename(2)` on same volume (kernel-atomic), copy+delete on cross-volume (Foundation fallback). External observers see either full original or full new bytes.
- **Throw-safe** — any throw at any step leaves `url` byte-preserved. Temp file cleaned up via `defer`.
- **fsync** — bytes flushed to disk before rename, so power loss after rename guarantees the new bytes are durable.

### Tests

- `AtomicSaveTests.swift` (NEW) — 6 tests:
  - `testSuccessfulSaveReplacesTargetAtomically` — happy path, SHA256 transitions cleanly.
  - `testThrowMidWritePreservesOriginalAndNoOrphanTempRemains` — read-only parent dir → write throws → original intact + no orphan tmp.
  - `testProcessKilledMidWritePreservesOriginal` — planted orphan tmp survives next write; original SHA256 invariant preserved across simulated SIGKILL.
  - `testFreshWriteToNonExistentPath` — fresh write produces only the target (no orphans).
  - `testTargetIsAlwaysObservableDuringSuccessfulWrite` — concurrent observer polling `fileExists(atPath:)` NEVER sees target absent during write (RED on pre-v0.13.2; GREEN with atomic-rename).
  - `testNoTempOrphanRemainsAfterSuccessfulOverwrite` — orphan cleanup invariant via `defer`.

**397/397 tests pass** (was 391 → +6 AtomicSaveTests).

### Compatibility

- **Public API unchanged** — `DocxWriter.write(_:to:)` signature identical; semantic guarantee strictly stronger.
- **Behavior change**: target file is no longer deleted as a separate step before write. Callers that observed the deletion gap (none known) would now see continuous file presence.
- **Cross-volume save**: `replaceItemAt` automatically falls back to copy+delete when temp and target are on different mount points. No-data-loss invariant preserved at copy granularity.

### Refs

- PsychQuant/che-word-mcp#36 — incident report and root-cause audit.

## [0.13.1] - 2026-04-23

### Fixed — rels overlay merge + relationship-driven image extraction (closes che-word-mcp#35)

v0.13.0 shipped `WordDocument.modifiedParts: Set<String>` + `DocxWriter` overlay-mode skip-when-not-dirty for typed parts (`document.xml`, `styles.xml`, `fontTable.xml`, `header*.xml`, `footer*.xml`, etc.) — but the **rels layer** still had two regressions on no-op round-trip of NTPU-style fixtures:

**Root cause A** — `DocxReader.extractImages` was directory-driven: it walked `word/media/` and used `targetToId[targetPath] ?? "rId_\(fileName)"` as fallback when the lookup missed. The fallback produced ids like `rId_image1.png` which:
1. Violate the OOXML `rId[0-9]+` convention.
2. Made `hasNewTypedRelationships` return true on no-op load (the typed model's image.id wasn't in originalRels), forcing rels regeneration.

**Root cause B** — `writeDocumentRelationships` built rels **from the typed-model parts list only**. Original rels for parts the typed model doesn't manage (theme / webSettings / customXml / commentsExtensible / commentsIds / people) were silently dropped — which broke theme-font inheritance, comment author identity, watermark VML rendering toggle, and Word 2013+ comment thread metadata after any legitimate rels-changing edit (e.g., `addHeader`).

### Architecture

1. **`Sources/OOXMLSwift/IO/RelationshipsOverlay.swift`** (NEW) — Parallel to `ContentTypesOverlay` from v0.12.0. Parses original `word/_rels/document.xml.rels`; merges typed-model rels with preservation of unknown rel types:
   - Original rel of managed type AND id in typed → emit (typed authoritative on target).
   - Original rel of managed type AND id NOT in typed → drop (deletion).
   - Original rel of any other type → preserve verbatim.
   - Typed rel whose id NOT in original → append as new.

2. **`DocxWriter.writeDocumentRelationships`** — Refactored. Overlay mode dispatches through `RelationshipsOverlay.merge`; scratch mode (no source archive) preserves the pre-v0.13.1 output via new `serializeScratchRels` helper. Adds `typedManagedRelationshipTypes` constant listing the 12 type URLs the model owns.

3. **`DocxReader.extractImages`** — Rewritten **relationship-driven** (was directory-driven). Iterates `relationships.imageRelationships` (source of truth); tries multiple path normalizations (`media/X` / `../media/X` / `word/media/X`); skips orphan rels rather than forge ids. Removed the `"rId_\(fileName)"` fallback entirely.

### Tests

- `RelationshipsOverlayTests.testNoOpRoundTripPreservesDocumentRelsByteEqual` — proves rels file is byte-equal after no-op load+save on multi-header fixture (theme + people rels survive).
- `testAddHeaderPreservesUnknownRelsTypes` — proves addHeader-triggered legitimate rewrite still preserves theme + people rels via overlay merge.
- `testRelsNeverProducesNonNumericIds` — regex-based regression guard against `rId_xxx`-style forged ids in either path.

**391/391 tests pass** (was 388 → +3 RelationshipsOverlay coverage).

### Compatibility

- **Additive for typed callers**: `RelationshipsOverlay` is `internal`; no public API change.
- **Behaviour change**: in overlay mode `writeDocumentRelationships` no longer drops rels for unknown types. Callers that previously relied on the lossy regenerate behavior (e.g., wanted theme rel stripped) — there are no known such callers.
- **Scratch mode unchanged**: `create_document` paths emit the same rels as before.
- `extractImages` orphan media files (in `word/media/` but not referenced by any rel) are now skipped instead of being assigned forged ids.

## [0.13.0] - 2026-04-23

### Added — True byte-preservation via dirty tracking (closes che-word-mcp#23 round-2, #32, #33 contributing fixes)

This release completes the round-trip fidelity work started in v0.12.0. The
PreservedArchive infrastructure preserved unknown parts (theme, customXml,
glossary, etc.) but the writer still **unconditionally re-emitted** every
typed-managed part on every save — so a Reader-loaded NTPU thesis lost its
13 custom font declarations, 6 distinct headers (collapsed to "header1.xml"),
and 4 footers after a single no-op `save_document` round-trip even though no
typed mutation had occurred.

v0.13.0 introduces three architectural changes that make typed-managed parts
behave like unknown parts: skip-when-not-dirty.

1. **`WordDocument.modifiedParts: Set<String>`** — every mutating method
   inserts the corresponding OOXML part path (`"word/document.xml"`,
   `"word/header4.xml"`, `"word/styles.xml"`, etc.). `DocxReader.read()`
   clears the set as the final step, so freshly loaded documents start with
   `modifiedParts.isEmpty == true`. Public `markPartDirty(_:)` lets external
   consumers (e.g., `che-word-mcp` writing to `archiveTempDir/word/theme/theme1.xml`)
   join the dirty-tracking contract.

2. **`Header.originalFileName` / `Footer.originalFileName`** — the pre-v0.13.0
   `fileName` computed property collapsed all `.default` headers to
   `"header1.xml"` regardless of source archive paths, so 6-section NTPU theses
   with `header1.xml`–`header6.xml` had every typed-model lookup hit the same
   file. Reader now populates `originalFileName` from each relationship's
   `Target` attribute; `fileName` returns `originalFileName ?? type-based-default`.

3. **`DocxWriter` overlay-mode skip-when-not-dirty** — every typed-part writer
   in overlay mode is gated by `modifiedParts.contains(<part path>)`. Scratch
   mode (no `archiveTempDir`) writes everything unconditionally — backward
   compatible with `create_document` callers. New helpers `hasNewTypedParts`
   and `hasNewTypedRelationships` ensure `[Content_Types].xml` and
   `word/_rels/document.xml.rels` are still re-emitted when the typed model
   added parts not declared in the source archive.

### Tests

- 4 `MarkDirtyCoverageTests` for the `Set<String>` foundation
- 38 `MarkDirtyCoverageTests` enumerating every WordDocument mutating method
- 8 `HeaderFooterOriginalFileNameTests` for the fileName preservation
- 3 `ReaderDirtyTrackingTests` for Reader instrumentation
- 2 `OverlaySkipWhenNotDirtyTests` proving no-op round-trip preserves typed
  parts byte-equal AND single-edit triggers selective re-emission only
- 6 `MultiHeaderFooterFixtureTests` building a 22-part .docx with 6 headers,
  4 footers, 13 fontTable entries, and 1 `<w15:person>` with full presenceInfo
  — proving end-to-end that editing one header preserves the other 5 byte-equal
  AND markPartDirty + direct write preserves all 13 fontTable entries

Total: 388 tests pass (was 327; +61 v0.13.0 contract coverage).

### Compatibility

- **Additive for typed callers**: `modifiedParts`, `markPartDirty(_:)`,
  `originalFileName` are new APIs. Existing callers compile unchanged.
- **Behaviour change in overlay mode**: writers SKIP for parts not in
  `modifiedParts`. Callers that previously relied on the writer regenerating
  `fontTable.xml` from a hardcoded 3-entry default on every save (which was
  the round-trip bug) will now see the original 13-entry fontTable preserved.
- **Scratch mode unchanged**: `create_document` paths (no source archive)
  emit every part as before.
- **`Header(id:paragraphs:type:)` and `Footer(id:paragraphs:type:...)` gain
  optional `originalFileName: String? = nil` parameter** — callers using
  positional arguments are unaffected.

## [0.12.2] - 2026-04-23

### Fixed — `WordDocument.nextRelationshipId` is now overlay-aware

`WordDocument.addHeader()`, `addFooter()`, and other typed-model add operations
allocate the new relationship's `rId` via `nextRelationshipId`. Previously this
used a naive counter (`headers.count + footers.count`) that would collide with
preserved original `_rels/document.xml.rels` entries in overlay mode (e.g.,
calling `addHeader` on a document with preserved `rId99` would naively return
`rId4`, but the writer's `RelationshipIdAllocator` would then upgrade it to
`rId100` — creating a typed/written rId mismatch).

`nextRelationshipId` now reads `archiveTempDir`'s original rels XML (when set)
and uses `RelationshipIdAllocator` to compute a collision-free rId. In scratch
mode (no archiveTempDir), behavior is unchanged.

### Compatibility

Behavior change only affects callers using overlay mode (Reader-loaded documents)
with Add CRUD tools. Scratch mode (`create_document` MCP path) returns the
same `rId4`, `rId5`, ... sequence as before.

## [0.12.1] - 2026-04-23

### Changed — Promote `WordDocument.archiveTempDir` to public read-only

Promotes the `archiveTempDir: URL?` accessor on `WordDocument` from
`internal` to `public` (read-only). Required by `che-word-mcp` v3.3.0
Phase 2A theme/header/footer CRUD tools, which need to read original OOXML
parts (`word/theme/theme1.xml`, `word/header*.xml`, `word/footer*.xml`)
directly from the preserved archive tempDir.

### Compatibility

Additive and non-breaking. The setter remains internal to `ooxml-swift`
(only `DocxReader` writes it via `preservedArchive`). External callers can
read the URL but cannot mutate the lifecycle outside of `WordDocument.close()`.

## [0.12.0] - 2026-04-23

### Changed — Preserve-by-default round-trip architecture (Phase 1 of `che-word-mcp-ooxml-roundtrip-fidelity`)

`DocxReader.read()` no longer deletes the source archive's unzip tempDir. The
tempDir is now retained on the returned `WordDocument` and released only when
the caller invokes the new `WordDocument.close()` method. `DocxWriter.write()`
detects the preserved tempDir and switches to **overlay mode**: typed-model
parts are overwritten directly into the preserved tempDir, then `ZipHelper.zip`
produces the destination `.docx`. All OOXML parts the typed model does NOT
manage (`word/theme/`, `word/webSettings.xml`, `word/people.xml`,
`word/commentsExtended.xml`, `word/commentsExtensible.xml`,
`word/commentsIds.xml`, `word/glossary/`, `word/customXml/`, etc.) survive
round-trip byte-for-byte.

Closes the lossy round-trip diagnosed in
[`PsychQuant/che-word-mcp#23`](https://github.com/PsychQuant/che-word-mcp/issues/23).
Unblocks the `OOXML parts CRUD completeness` milestone (#24-#31, 8 enhancement
issues) which all require round-trip fidelity to ship.

### Added — Public API

- **`WordDocument.close()`** (`Sources/OOXMLSwift/Models/Document.swift`) —
  new `public mutating func close()`. Releases the preserved archive tempDir;
  idempotent. Callers SHOULD invoke after the final `DocxWriter.write()` to
  free the tempDir; forgetting leaks the directory until process exit (macOS
  reclaims `/tmp` on reboot).

### Added — Internal helper types

- **`PreservedArchive`** (`Sources/OOXMLSwift/IO/PreservedArchive.swift`) —
  thin wrapper over the unzip tempDir URL with a `cleanup()` method. Used as
  internal storage for `WordDocument`'s preserved-archive lifecycle.
- **`RelationshipIdAllocator`** (`Sources/OOXMLSwift/IO/RelationshipIdAllocator.swift`) —
  scans the source's `_rels/document.xml.rels` plus typed-model rIds, returns
  collision-free `rId<N>` strings via `allocate()`. Replaces the prior naive
  counter (`headers.count + footers.count + ...`) at `DocxWriter.swift:238`
  that would collide with preserved original rIds in overlay mode.
- **`ContentTypesOverlay`** + **`PartDescriptor`** (`Sources/OOXMLSwift/IO/ContentTypesOverlay.swift`) —
  parses the source `[Content_Types].xml`, merges typed-part `<Override>`
  entries with preserved entries via the
  preserve-unknown-overrides + dedupe-typed-overrides + add-new-overrides
  algorithm. Supports explicit "deletion" semantics via `typedManagedPatterns`
  (PartName matches a managed pattern but is absent from typedParts → drop).

### Compatibility

**BREAKING-semantic, additive-API-only**:

- API additions are non-breaking — existing code compiles unchanged.
- **Lifecycle is new**: callers that read documents and discard them previously
  worked because `DocxReader` cleaned its tempDir before returning. Now the
  tempDir lives until `close()`; non-`close`-ing callers leak tempDirs that
  macOS eventually reclaims on reboot.
- **MCP server callers** (`che-word-mcp`) must wire `WordDocument.close()` into
  session lifecycle. See `che-word-mcp` v3.3.0 (Phase 2A of the same Spectra
  change) for the integration.

### Tests

15 new XCTest cases in `Tests/OOXMLSwiftTests/RoundTripFidelityTests.swift`:
- `WordDocument.close()` / `archiveTempDir` lifecycle (5)
- `RelationshipIdAllocator` collision avoidance + non-numeric handling (5)
- `ContentTypesOverlay` preserve / replace / add / drop scenarios (3)
- `DocxWriter` overlay round-trip preservation of unknown parts (theme1.xml +
  customXml) + ZIP entry-list equality + Content_Types Override-set equality
  (2)

Full suite **325/325 green** (was 310/310).

## [0.11.0] - 2026-04-23

### Added — `MathAccent` for accent decorators

Adds the OMML accent element `<m:acc>` (ECMA-376 Part 1 §22.1.2.1) so callers
emitting LaTeX-derived equations (`\hat{x}`, `\bar{x}`, `\tilde{x}`, `\dot{x}`,
`\overline{x}`) produce structurally correct OMML editable in MS Word's
native equation editor. Previously these accent macros had no first-class
`MathComponent` representation.

- **`MathAccent`** (`Sources/OOXMLSwift/Models/MathComponent.swift`) — new
  public struct conforming to `MathComponent`. Stored properties: `base:
  [MathComponent]` (math content under the accent) and `accentChar: String`
  (Unicode combining diacritic — typically `"\u{0302}"` circumflex,
  `"\u{0304}"` macron, `"\u{0303}"` tilde, `"\u{0307}"` dot above).
  `toOMML()` emits `<m:acc><m:accPr><m:chr m:val="<c>"/></m:accPr><m:e><base
  OMML></m:e></m:acc>` with XML escaping applied to `accentChar`.

- **`OMMLParser` accent dispatch** — adds `case "acc"` to the recognized-tag
  switch with a `parseMathAccent(_:)` helper. Previously `<m:acc>` subtrees
  were preserved as `UnknownMath`; now they round-trip as typed
  `MathAccent` values.

### Tests

4 new XCTest cases in `MathComponentTests` cover hat over single run, bar
over Greek letter, accent over composite SubSuperScript base, and accent
character requiring XML escape.

### Compatibility

Additive and non-breaking. Existing `<m:acc>` round-trips that returned
`UnknownMath` will now return `MathAccent` — callers pattern-matching with
`as? UnknownMath` should add a `MathAccent` arm.

## [0.10.0] - 2026-04-22

### Added — read-side parsers for fields and OMML

Closes the "write-side only" gap from v2.0.0 `FieldCode` and `MathComponent`. Three downstream `che-word-mcp` issues (#17 caption CRUD, #19 update_all_fields, #21 equation CRUD) all depend on these primitives.

- **`FieldParser`** (`Sources/OOXMLSwift/Parsing/FieldParser.swift`) — walks a `Paragraph`'s runs looking for `<w:fldChar>` field spans, parses `<w:instrText>` into typed `ParsedFieldValue` (cases: `.sequence`, `.styleRef`, `.reference`, `.unknown(instrText:)`). Each `ParsedField` carries `startRunIdx` / `endRunIdx` / `cachedResultRunIdx` so CRUD tools can locate specific runs to modify.

- **`OMMLParser`** (`Sources/OOXMLSwift/Parsing/OMMLParser.swift`) — parses `<m:oMath>` / `<m:oMathPara>` XML into a `[MathComponent]` tree. Recognizes 5 of the 9 core types (`MathRun`, `MathFraction`, `MathSubSuperScript`, `MathRadical`, `MathNary`); unrecognized subtrees preserved as `UnknownMath(rawXML:)` for round-trip safety. (Parsers for `MathDelimiter` / `MathFunction` / `MathLimit` / `MathMatrix` deferred; they still emit via `toOMML()` and round-trip through `UnknownMath`.)

- **`UnknownMath`** — new opaque `MathComponent` struct that preserves raw XML for round-tripping. Note: callers iterating `[MathComponent]` arrays may encounter this type — handle via `as?` cast.

- **`FieldCode.parse(instrText:)`** static method added via extension on `SequenceField`, `StyleRefField`, `ReferenceField`. Returns `Self?` (nil on non-match). `FieldParser` dispatches by trying each in turn. Unknown field types (e.g., `TIME`, `MERGEFIELD`) captured as `.unknown(instrText:)`.

- **`WordDocument.updateAllFields() -> [String: Int]`** — F9-equivalent SEQ counter recomputation across body + headers + footers + footnotes + endnotes. Non-SEQ fields preserved verbatim. Chapter-reset semantics: when a paragraph has `pStyle == "Heading N"`, SEQ fields with `resetLevel == N` restart their counters. Returns map of identifier → final count.

### Tests

40 new XCTest cases across `FieldCodeParseTests`, `FieldParserTests`, `OMMLParserTests`, `UpdateAllFieldsTests`. Full suite 306/306 green.

### Out of scope (follow-up)

- LaTeX parser for `insert_equation(latex:)` (Phase 3 deferred from word-mcp-insertion-primitives).
- IF / CalculationField / DateTimeField / DocumentInfoField / MergeField `parse(instrText:)` — `.unknown` fallback covers them for round-trip; add per-type parsers when CRUD tools target them.
- `MathDelimiter` / `MathFunction` / `MathLimit` / `MathMatrix` parsing — `UnknownMath` preserves round-trip; full parse added when CRUD tools target those shapes.

## [0.9.0] - 2026-04-22

### Added

- **`InsertLocation.afterText(String, instance: Int)` + `.beforeText(...)` cases** — insert paragraph/image relative to a body paragraph containing the given substring. Match is on flattened run text (cross-run safe). `instance` is 1-based to disambiguate when same phrase appears multiple times. Closes use case in [che-word-mcp#14](https://github.com/PsychQuant/che-word-mcp/issues/14) where every insert previously needed `search_text` + `insert_*` as 2 MCP calls.
- **`InsertLocationError.textNotFound(searchText:instance:)`** — new error case for text-anchor resolution failure.
- **`WordDocument.findBodyChildContainingText(_:nthInstance:)`** (private) — helper iterating body paragraphs and matching flattened text.

### Behavior note

Enum case addition technically changes the public surface. Callers that `switch` on `InsertLocation` exhaustively may emit a warning about missing cases. In practice all in-monorepo consumers use partial switches / pass-through, so no breaking impact observed during batch rebuild.

## [0.8.0] - 2026-04-22

### Breaking

- **`WordDocument.replaceText` signature change** — was `replaceText(find:with:all:) -> Int`; now `replaceText(find:with:options:) throws -> Int`. The old `all: Bool` parameter is removed (behavior now "always replaces all matches"). `ReplaceOptions` exposes `scope: ReplaceScope` (`.bodyAndTables` / `.all`), `regex: Bool`, `matchCase: Bool`. Throws `ReplaceError.invalidRegex` on bad regex pattern. Migration: `doc.replaceText(find:with:all: true)` → `try doc.replaceText(find:with:options: ReplaceOptions())`.
- **`MathEquation` deprecated** — `@available(*, deprecated)` annotation applied. The `toXML()` implementation still runs but produces flat `<m:r><m:t>` (not structured OMML). Replace with `MathComponent` AST; `MathEquation` will be removed in 1.0.

### Added

- **Text replacement engine (`TextReplacementEngine`)** — flatten-then-map algorithm. Cross-run matches now succeed (e.g. `"hello world"` spread across `["hello ", "", "world"]` runs). Replacement text inherits the start run's formatting. Non-text runs (fields, drawings) are preserved across splices. Supports `.all` scope (headers, footers, footnotes, endnotes) and regex mode with `$1..$N` backreferences. Closes cross-run-failure part of PsychQuant/che-word-mcp#7.
- **`Document.replaceText` scope `.all`** — when `options.scope == .all`, traversal covers body, table cells, headers, footers, footnotes, endnotes.
- **Math AST (`MathComponent` protocol + 9 types)** — `MathRun`, `MathFraction`, `MathSubSuperScript`, `MathRadical`, `MathNary` (∑/∫/∏/∬/∮/⋃/⋂), `MathDelimiter`, `MathFunction`, `MathLimit`, `MathMatrix`. Each emits structurally correct OMML via `toOMML()`. Replaces the flat `MathEquation.toXML()` string-substitution path. Refs PsychQuant/che-word-mcp#6.
- **`StyleRefField`** conforming to `FieldCode` — produces `STYLEREF <level>[ \s][ \l]` field XML. For caption chapter-number prefixes. Refs PsychQuant/che-word-mcp#9.
- **`ImageDimensions.detect(path:)`** — reads PNG IHDR and JPEG SOFn headers, returns `(widthPx, heightPx, aspectRatio)`. Throws `ImageDimensionsError.unsupportedFormat` for non-PNG/JPEG extensions. Used for auto-aspect image insertion. Refs PsychQuant/che-word-mcp#8.
- **`InsertLocation` enum** — four cases: `.paragraphIndex(Int)`, `.afterImageId(String)`, `.afterTableIndex(Int)`, `.intoTableCell(tableIndex:row:col:)`. Extends `WordDocument` with `insertParagraph(_:at: InsertLocation)` and `insertImage(path:widthPx:heightPx:at: InsertLocation, ...)` overloads. Throws `InsertLocationError` on invalid anchor. Refs PsychQuant/che-word-mcp#8 #9.
- **`FieldCode.toFieldXML()` is now public** — enables external modules (e.g. che-word-mcp) to emit field XML inline via rawXML-bearing runs.

### Changed

- `nextImageRelationshipId` visibility bumped from `private` to `internal` so the `InsertLocation` extension can reuse the id allocator.

## [0.6.1] - 2026-04-16

### Fixed

- **Table-cell revisions now carry distinct location info** — `Revision` gains 3 optional `Int?` fields: `tableRow`, `tableColumn`, `cellParagraphIndex`. Previously all revisions in a table shared the same `paragraphIndex` (the table's body position), making it impossible to distinguish which cell they belonged to. Closes [PsychQuant/ooxml-swift#2](https://github.com/PsychQuant/ooxml-swift/issues/2).

## [0.6.0] - 2026-04-16

### Added

- **Container reading** — `DocxReader.read(from:)` now parses `word/header*.xml`, `word/footer*.xml`, `word/footnotes.xml`, and `word/endnotes.xml`. These were previously write-only in the model; now `document.headers`, `document.footers`, `document.footnotes`, `document.endnotes` are populated on the read path with full paragraph structure (including revisions and comments).
- **`RevisionSource` enum** — new public type: `.body`, `.header(id:)`, `.footer(id:)`, `.footnote(id:)`, `.endnote(id:)`. Every `Revision` now carries a `source` field (default `.body`).
- **`Revision.previousFormatDescription: String?`** — human-readable summary of prior formatting for `.formatChange` and `.paragraphChange` revisions (e.g., `"bold, italic, 12pt Times New Roman"`). Complements the existing structured `previousFormat: RunProperties?`.
- **`WordDocument.getRevisionsFull() -> [Revision]`** — additive API returning all revisions (body + containers) with `source` and `previousFormatDescription`. Mirrors `getCommentsFull()` pattern.
- **`Footnote.paragraphs` / `Endnote.paragraphs`** — new `[Paragraph]` property on both model types, populated by the reader with rich paragraph structure.

### Fixed

- **Nested `rPrChange` / `pPrChange` revisions** — `parseParagraph` now descends into `<w:rPr>` and `<w:pPr>` to detect `<w:rPrChange>` (run formatting change) and `<w:pPrChange>` (paragraph property change), emitting `Revision(type: .formatChange)` and `Revision(type: .paragraphChange)` respectively. Previously these nested change-tracking elements were invisible. Closes Part B of [PsychQuant/ooxml-swift#1](https://github.com/PsychQuant/ooxml-swift/issues/1).
- **Container revision aggregation** — revision aggregation step now walks headers, footers, footnotes, and endnotes after body paragraphs, assigning the correct `RevisionSource` to each. Closes Part C of [PsychQuant/ooxml-swift#1](https://github.com/PsychQuant/ooxml-swift/issues/1).
- **Footnote/endnote separator filtering** — uses `w:type` attribute (not numeric ID) to skip separator and continuation-separator entries, which is robust against ID numbering variations across Word versions.

### Changed

- **`getRevisions()` tuple API** — now filters to `source == .body` only, preserving backward compatibility. Callers that want container revisions should use `getRevisionsFull()`.

### Notes

- `RevisionType.formatting` (`.rPrChange2`) is retained in the enum but not emitted by the parser — awaiting real-world evidence of `<w:rPrChange2>` in OOXML output.
- Spectra change: [`PsychQuant/macdoc:openspec/changes/docx-reader-nested-revisions-and-containers`](https://github.com/PsychQuant/macdoc/tree/main/openspec/changes/docx-reader-nested-revisions-and-containers)
- **Fully closes [PsychQuant/ooxml-swift#1](https://github.com/PsychQuant/ooxml-swift/issues/1)** (all 4 parts: A/B/C/D across v0.5.7 and v0.6.0).

## [0.5.7] - 2026-04-16

### Fixed

- **`DocxReader.parseParagraph` now parses `w:moveFrom` and `w:moveTo` revisions** — previously these two `RevisionType` cases were silently dropped at the top-level paragraph switch, even though the `Revision` model had always declared them. Any document using Word's tracked move feature reported 0 move revisions, undercounting total revisions accordingly. Closes Part A of [PsychQuant/ooxml-swift#1](https://github.com/PsychQuant/ooxml-swift/issues/1).
  - `w:moveFrom` mirrors `w:del`: extract nested `<w:r>` text, emit a `Revision` with `type == .moveFrom` and the moved-out text in `originalText`.
  - `w:moveTo` mirrors `w:ins`: extract nested `<w:r>` text, emit a `Revision` with `type == .moveTo` and the moved-in text in `newText`.
  - Both preserve the `w:id` attribute, so callers can correlate a moveFrom/moveTo pair sharing the same id.

### Added

- **`DocxReader.debugLoggingEnabled: Bool = false`** — opt-in static flag. When set to `true`, the parser writes one line to stderr for each direct child of `<w:p>` whose local name is not one of the recognized cases. Intended for development and test-time surfacing of parser coverage gaps. Zero runtime cost when `false` (guard evaluated before any string formatting). Closes Part D of [PsychQuant/ooxml-swift#1](https://github.com/PsychQuant/ooxml-swift/issues/1).
- **`DocxReader.parseParagraph` is now `internal static`** (was `private static`) — enables unit tests in the `OOXMLSwiftTests` target to exercise the parser directly with hand-constructed `XMLElement` instances via `@testable import`. No new public API surface.

### Notes

- Parts B and C of `ooxml-swift#1` remain open (nested `rPrChange`/`pPrChange` in property parsers + container iteration for headers/footers/footnotes/endnotes). They will ship as a follow-up change with additional API surface including a new `Revision.source` field.
- Spectra change: [`PsychQuant/macdoc:openspec/changes/docx-reader-top-level-revisions`](https://github.com/PsychQuant/macdoc/tree/main/openspec/changes/docx-reader-top-level-revisions)

## [0.5.6] - 2026-04-15

### Added

- `WordDocument.getCommentsFull() -> [Comment]` — returns the complete `Comment` struct for every comment in the document, exposing `parentId` (reply threading), `paraId`, `done`, and `initials`. Companion to the existing `getComments()` tuple API.

### Notes

- `getCommentsFull` is purely additive. The existing `getComments()` tuple API is unchanged.
- Motivation: the prior tuple-returning `getComments()` dropped `parentId`, forcing downstream consumers (e.g., manuscript review threading tools in che-word-mcp) to either lose reply structure or re-parse `comments.xml` manually. `getCommentsFull` provides the full struct without breaking existing callers.
- Spectra change: [`PsychQuant/macdoc:openspec/changes/manuscript-review-markdown-export`](https://github.com/PsychQuant/macdoc/tree/main/openspec/changes/manuscript-review-markdown-export)
