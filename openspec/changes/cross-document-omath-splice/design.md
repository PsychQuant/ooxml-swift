## Context

OOXMLSwift currently has three OMath storage paths, each with distinct visual semantics:

| Carrier | XML shape | Word renders as | LaTeX equivalent |
|---------|-----------|------------------|------------------|
| `Run.rawXML` | `<w:r>...<m:oMath>...</m:oMath>...</w:r>` | inline (with surrounding text) | `$α$` |
| `Paragraph.unrecognizedChildren[name="oMath"]` | `<w:p>...<m:oMath>...</m:oMath>...</w:p>` (direct child of `<w:p>`) | display (own line) | `$$\alpha$$` or `\[\alpha\]` |
| `MathEquation` (deprecated) | API-built; flat `<m:r><m:t>...</m:t></m:r>` from naive LaTeX simplifier | text-only fallback | n/a |

The 郭嘉員 thesis fixture has 522 inline OMath blocks all stored in `Run.rawXML`. Pandoc-generated docx typically uses `unrecognizedChildren` for display equations. A cross-document copy API must respect this distinction — copying inline OMath into a display carrier (or vice versa) changes the visual output and violates the verbatim-copy contract.

`Paragraph.toXML()` emits all 13 position-indexed carriers (runs / unrecognizedChildren / bookmarkMarkers / commentRangeMarkers / proofErrorMarkers / smartTags / customXmlBlocks / bidiOverrides / contentControls / hyperlinks / fieldSimples / alternateContents / permissionRangeMarkers) interleaved by their `position: Int?` fields, with stable sort (equal positions retain insertion order). DocxReader fills `position` for both `Run` and `UnrecognizedChild` from source-document byte offset.

Existing `WordDocument+ReplaceTextWithBoundaryDetection.swift` already implements anchor-Run splitting — same pattern this design reuses for mid-paragraph OMath insertion.

The design decisions below were converged via [PsychQuant/ooxml-swift#57](https://github.com/PsychQuant/ooxml-swift/issues/57) `/idd-diagnose` (6 open questions identified) → `/spectra-discuss` (each question answered with explicit Pros/Cons trade-off table).

## Goals / Non-Goals

**Goals:**

- Verbatim copy of `<m:oMath>` XML blocks between `WordDocument` paragraphs without LaTeX intermediation
- Preserve carrier shape (inline-Run OMath stays inline-Run in target; direct-child OMath stays direct-child in target)
- Preserve source Run's `rPr` (font/size/lang) by default, with escape hatches for cross-doc style/theme conflicts
- Support mid-paragraph splice via `.afterText(...)` / `.beforeText(...)` anchors mirroring existing `InsertLocation` API
- Provide both single-OMath low-level API (full caller control) and paragraph-level batch API (convenient for cross-doc rescue scenarios)
- Maintain round-trip lossless guarantee — `DocxReader.read()` of the spliced doc returns OMath XML byte-equal to source

**Non-Goals:**

- LaTeX-to-structured-OMML conversion — out of scope; existing `MathComponent` AST handles the API-built path
- Document-level batch API (`spliceAllOMath` across entire WordDocument) — paragraph-matching algorithm is caller-specific (different consumers want different matchers); keep matching in caller layer to prevent scope creep
- Cross-document `paragraph` formatting copy — only OMath + its enclosing Run rPr; surrounding paragraph properties stay target's
- Auto-rewrite of namespace prefix — `.lenient` mode accepts mixed prefixes (ECMA-376 compliant); `.strict` mode throws on any mismatch. No string substitution on source XML
- Bulk splice of all OMath from one paragraph in a single call — caller loops `spliceOMath` with incrementing `omathIndex` (or uses `spliceParagraphOMath` for context-anchor auto-derivation)
- Splicing into headers / footers / footnotes / endnotes — body paragraphs only in v0.1 (target indexed by `toBodyParagraphIndex: Int`)

## Decisions

### Carrier preservation strategy

**Decision**: Preserve source carrier shape — inspect source paragraph for OMath in `Run.rawXML` (inline) and `Paragraph.unrecognizedChildren[name="oMath" or "oMathPara"]` (direct-child), and splice into target using the same carrier kind.

**Alternatives considered**:

- *Always Run.rawXML on target*: simpler one-path implementation, mirrors existing `insertEquation(...displayMode: false)` convention. **Rejected** — turns display OMath into inline visually (lossy semantics), breaks Pandoc-style source documents
- *Always unrecognizedChildren on target*: simpler data model. **Rejected** — turns inline OMath into display visually (catastrophic for thesis use case where 522 OMath are all inline-in-Run)

**Rationale**: Caller invokes a "verbatim copy" operation; turning inline into display (or vice versa) violates the contract. Implementation cost (two code paths) is acceptable; the alternative cost (broken visual semantics) is not.

### Joint document-order index for `omathIndex`

**Decision**: When source paragraph contains OMath in both carriers, sort by `position: Int?` (filled by DocxReader for both `Run` and `UnrecognizedChild` from source byte offset) and treat `omathIndex` as "Nth OMath in source-document order, regardless of carrier."

**Alternatives considered**:

- *Per-carrier separate index*: caller specifies `sourceCarrier: .inRun(index: Int) | .directChild(index: Int)`. **Rejected** — exposes implementation detail; caller usually thinks in source-document order, not carrier order

**Rationale**: Joint sort is robust because DocxReader fills `position` for both carriers (verified at `DocxReader.swift:1005` for runs, `DocxReader.swift:1399-1405` for default-case unrecognized children). API-built paragraphs have nil positions but never appear as splice sources in practice (sources are always source-loaded).

### Mid-paragraph splice via anchor-Run split

**Decision**: For `.afterText(...)` / `.beforeText(...)` anchors that resolve mid-Run, split the anchor Run into 2-3 segments at the anchor boundary, copy `rPr` to each segment, and insert the new OMath Run between them. All segments share the source Run's `position` value; stable sort retains insertion order (same trick used by other position-coupled paragraph operations).

**Alternatives considered**:

- *Renumber whole paragraph's positioned entries*: re-sequence all 13 carrier types. **Rejected** — touches every carrier kind (each is a separate array; missing one causes silent ordering bug; this is the fragile area #56 series fixed). Renumber also loses correspondence between `position` and source byte offset, breaking byte-equal round-trip diagnostics
- *Append-only API (no mid-paragraph)*: only support `.atStart` / `.atEnd`. **Rejected** — thesis use case requires mid-paragraph (e.g., "進行 t 檢定" needs splicing the OMath `t` between "進行 " and " 檢定")

**Rationale**: Run-split isolates blast radius to `runs[]` array; the other 12 carriers are untouched. This is the same anchor-split approach `WordDocument+ReplaceTextWithBoundaryDetection.swift` already implements — proven robust by the existing test suite.

### Default `OMathSpliceRpRMode = .full`

**Decision**: Default to verbatim `rPr` copy from source Run to the new target OMath Run. Provide `.omathOnly` (whitelist: rFonts/sz/lang only) and `.discard` (empty rPr) as escape hatches.

**Alternatives considered**:

- *Default `.discard` (empty rPr)*: simpler, predictable. **Rejected** — Cambria Math font reference would be lost, OMath inherits target paragraph's CJK font and renders broken
- *Default `.omathOnly` (whitelist)*: defensive against cross-doc styleRef / themeColor breakage. **Rejected as default** — too cautious for the common case where source rPr is plain `rFonts` + `lang`; whitelist-skipped fields cause subtle visual differences that surprise callers

**Rationale**: User direction「我想要完整就好」prioritizes visual completeness. Cross-doc rPr risks (rStyle ID not present in target's `word/styles.xml`, themeColor referencing different theme) are documented as caveats and addressable via the two opt-out modes when callers know they have such conflicts.

### Two-tier API: `spliceOMath` (single) + `spliceParagraphOMath` (batch)

**Decision**: Provide both single-OMath low-level entry point and paragraph-level batch entry point. The batch variant auto-derives anchor for each OMath from its source-text-context (~5-10 chars on each side via `flattenedDisplayText` slicing) and routes to `.afterText(prefix, instance: 1)`. Throws `OMathSpliceError.contextAnchorNotFound(omathIndex:, snippet:)` per OMath that fails to resolve.

**Alternatives considered**:

- *Only single-OMath API*: caller writes the loop. **Rejected** — for thesis rescue (510+ splice calls across 43 paragraphs), boilerplate explodes by ~200 LOC of fragile per-OMath context extraction
- *Document-level batch (`spliceAllOMath`)*: single call for whole document. **Rejected** — paragraph-matching algorithm is caller-specific (thesis uses 30-char prefix anchor; other use cases want paragraph IDs or content-hash). Putting matcher policy in API leads to scope creep

**Rationale**: Two-tier mirrors common library design (high-level convenience + low-level escape hatch). Paragraph matching stays in caller — `ooxml-swift` doesn't know about thesis vs Pandoc vs other matching strategies.

### Lenient namespace policy by default

**Decision**: Default `OMathSpliceNamespacePolicy = .lenient` — accept prefix mismatch (e.g., source `mml:` + target `m:` both pointing to standard OMML URI) by splicing source XML verbatim, letting target paragraph carry mixed prefixes. Throw `.namespaceMismatch(sourceURI:, targetURI:)` only when URIs differ. `.strict` mode throws on any prefix or URI mismatch.

**Alternatives considered**:

- *Strict-only (issue body original design)*: throw on any namespace mismatch. **Rejected** — for the rare prefix-mismatch case (source from non-standard generator), the splice would abort instead of producing a working result
- *Auto-rewrite prefix (`mml:` → `m:`)*: string substitution to normalize. **Rejected** — violates verbatim-copy contract; string substitution can mis-rewrite attribute values containing `mml:` literal

**Rationale**: ECMA-376 explicitly allows mixed prefixes within one document (each namespace declaration scopes locally). 99% of real docx files use the standard `m:` prefix anyway, making this almost a no-op safeguard. URI mismatch (rare; would mean the source uses a vendor-extended namespace not in the OMML schema) is a real semantic mismatch and warrants a throw.

## Risks / Trade-offs

[**Risk: round-trip lossy after splice**] → Mitigation: 8 round-trip test fixtures in `Tests/OOXMLSwiftTests/OMathSpliceTests.swift` enforce byte-equal `DocxReader.read()` of spliced output vs original `<m:oMath>` block. PR merge blocked on any byte mismatch.

[**Risk: anchor-Run split with whitespace-sensitive `<w:t xml:space="preserve">`**] → Mitigation: copy `xml:space` attribute when splitting; existing `replaceText` code path tested with similar fixtures. Add explicit test case for whitespace-bearing anchor.

[**Risk: cross-doc rStyle reference broken in target (e.g., `<w:rStyle w:val="MathStyle">` where target lacks that style ID)**] → Mitigation: documented as `.full` mode caveat; provide `.omathOnly` (strips rStyle) as opt-out. Real risk is low — Cambria Math is typically applied via direct `rFonts`, not via style reference, in NTPU/Word output.

[**Risk: position-renumber regression**] → Mitigation: split-point share-position approach explicitly avoids renumbering; existing #56 / #99-103 round-trip suites run on every PR and would catch ordering drift.

[**Risk: mixed-prefix output rejected by strict XML validators**] → Mitigation: default `.lenient` produces ECMA-376-compliant output (mixed prefixes are spec-legal); `.strict` mode available for callers needing single-prefix output for downstream tooling.

[**Risk: MathEquation removal collision**] → Mitigation: this API has zero dependency on `MathEquation`; sister concern #58 tracks the deprecation message inconsistency separately.

[**Risk: future migration to `MathComponent` AST**] → Mitigation: the splice API operates on raw XML, completely orthogonal to the structured AST path. Future API additions on the AST side (LaTeX → structured OMML) do not interact with splice operations.
