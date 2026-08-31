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
- Maintain round-trip semantic fidelity — `DocxReader.read()` of the spliced doc returns an OMath subtree with the same namespace-expanded structure, attributes, and resolved text

**Non-Goals:**

- LaTeX-to-structured-OMML conversion — out of scope; existing `MathComponent` AST handles the API-built path
- Document-level batch API (`spliceAllOMath` across entire WordDocument) — paragraph-matching algorithm is caller-specific (different consumers want different matchers); keep matching in caller layer to prevent scope creep
- Cross-document `paragraph` formatting copy — only OMath + its enclosing Run rPr; surrounding paragraph properties stay target's
- Auto-rewrite of namespace prefix — `.lenient` mode accepts mixed prefixes (ECMA-376 compliant); `.strict` mode throws on any mismatch. No string substitution on source XML
- Splicing into headers / footers / footnotes / endnotes — body paragraphs only in v0.1 (target indexed by `toBodyParagraphIndex: Int`)

## Decisions

### Carrier preservation strategy

**Decision**: Preserve source carrier shape — inspect source paragraph for OMath in `Run.rawXML` (inline) and `Paragraph.unrecognizedChildren[name="oMath" or "oMathPara"]` (direct-child), and splice into target using the same carrier kind.

**Alternatives considered**:

- *Always Run.rawXML on target*: simpler one-path implementation, mirrors existing `insertEquation(...displayMode: false)` convention. **Rejected** — turns display OMath into inline visually (lossy semantics), breaks Pandoc-style source documents
- *Always unrecognizedChildren on target*: simpler data model. **Rejected** — turns inline OMath into display visually (catastrophic for thesis use case where 522 OMath are all inline-in-Run)

**Rationale**: Caller invokes a "verbatim copy" operation; turning inline into display (or vice versa) violates the contract. Implementation cost (two code paths) is acceptable; the alternative cost (broken visual semantics) is not.

### Joint document-order index for `omathIndex`

**Decision**: Treat `omathIndex` as "Nth OMath in actual paragraph serialization order, regardless of carrier." The extractor uses the same four regions as Paragraph: absolute start boundary, positive-position window, non-positive post-content buckets, and absolute end boundary; boundary order/position plus stable Run-before-direct collection sequence resolve ties.

**Alternatives considered**:

- *Per-carrier separate index*: caller specifies `sourceCarrier: .inRun(index: Int) | .directChild(index: Int)`. **Rejected** — exposes implementation detail; caller usually thinks in source-document order, not carrier order

**Rationale**: DocxReader fills positive positions for loaded carriers, while API-built and in-memory boundary paragraphs legitimately carry nil/non-positive positions or absolute-boundary metadata and can immediately become splice sources. Mirroring serializer regions keeps extraction, strict namespace target selection, batch anchors, and visible XML order consistent.

### Mid-paragraph splice via anchor-Run split

**Decision**: For `.afterText(...)` / `.beforeText(...)` anchors that resolve mid-Run, split the anchor Run into 2-3 segments at the anchor boundary, copy `rPr` to each segment, and insert the new OMath Run between them. All segments share the source Run's `position` value; stable sort retains insertion order (same trick used by other position-coupled paragraph operations).

**Alternatives considered**:

- *Renumber whole paragraph's positioned entries*: re-sequence all 13 carrier types. **Rejected for inline anchors** — touches every carrier kind (each is a separate array; missing one causes silent ordering bug; this is the fragile area #56 series fixed) and discards useful source-offset diagnostics. Direct-child mid-text insertion is the narrow exception because cross-collection placement requires a unique slot.
- *Append-only API (no mid-paragraph)*: only support `.atStart` / `.atEnd`. **Rejected** — thesis use case requires mid-paragraph (e.g., "進行 t 檢定" needs splicing the OMath `t` between "進行 " and " 檢定")

**Rationale**: Run-split isolates blast radius to `runs[]` array; the other 12 carriers are untouched. This is the same anchor-split approach `WordDocument+ReplaceTextWithBoundaryDetection.swift` already implements — proven robust by the existing test suite.

### Boundary insertion across mixed carrier modes

**Decision**: `.atStart` and `.atEnd` use an internal absolute-boundary serializer lane on the newly inserted Run or UnrecognizedChild. The lane emits before legacy pre-content or after legacy post-content, respectively, and remains outside the ordinary position-sorted window.

**Rationale**: Paragraph serializes legacy pre-content, positive positions, nil/zero type buckets, and legacy post-content through separate paths. A numeric position cannot represent a point outside all four regions, and projecting legacy typed state into raw carriers breaks later page-break/note/bookmark mutations. The dedicated lane preserves all public typed fields and existing numeric positions while making true boundaries safe, including hostile `Int.max` inputs (#122).

### Batch anchor state and mutation direction

**Decision**: Before deriving any anchor, batch splice uses two global preflight phases: (1) validate every extracted source fragment and the relevant initial target fragment for XML/namespace identity, then (2) enforce URI/prefix policy only after the entire validity scan succeeds. `deriveContextAnchors` then returns optional `(snippet, instance, boundaryPlacement)` anchors: nil is unmatched, an empty string without boundary metadata is a matched leading OMath, a non-empty string includes its source occurrence number, and explicit start/end metadata remains authoritative for an unsaved boundary source. Prose comes only from typed Run text that `Run.toXML()` would serialize; raw run/property overrides and drawings hide typed text and are excluded. Inline XML comparison uses the same self-contained namespace normalization as extraction. Anchor groups run in source order for partial-success semantics; shared start/text boundaries apply right-to-left, while an absolute-end group applies left-to-right because end insertion appends.

**Rationale**: The former empty-string sentinel conflated leading and unmatched equations, while hard-coded `instance: 1` misplaced repeated prose. Global reverse mutation also left later equations behind when an earlier source item failed. XML admission is global and separated from namespace policy so malformed input always fails with `OMathSpliceMalformedXMLError`, independent of source order, before any mutation; partial-success semantics apply only to later context-anchor group failures. Explicit occurrence state plus source-ordered, group-atomic application makes failure loud, partial success inspectable, and final document order stable (#125).

Context derivation consumes a Run/direct-OMath event stream sorted by the same four regions as `Paragraph.toXML()`: absolute start, positive positions, non-positive post-content, and absolute end. Runs precede direct unrecognized children at equal positive/non-positive positions, matching collection emission order. Array order alone is never used as document order.

### Direct-child text anchors

**Decision**: Direct-child OMath uses the same Run-anchor resolution as inline OMath. The target Run is split at the original UTF-16 boundary, surrounding carriers are shifted in the canonical position space, and the direct child receives the unique position between prefix and suffix.

**Rationale**: Silently degrading `.afterText` / `.beforeText` to paragraph end violates both the public position enum and batch failure contract. Unique positions are required because the serializer's cross-collection stable tie order cannot place an `UnrecognizedChild` between two equal-position Runs.

Anchor resolution uses the same serializer-visible typed-text rule as batch derivation. Matching hidden `Run.text` beneath `rawXML`, `RunProperties.rawXML`, or `drawing` would split a carrier whose prefix and suffix both retain the opaque override, duplicating content.

When a visible-text Run also carries post-text `rawElements`, splitting clears them from the prefix and keeps them on the suffix only. An empty-text suffix is retained for this purpose, preserving the opaque child exactly once and after the original text/insertion boundary.

### Direct-child local-name preservation

**Decision**: Carry the source direct-child local name through extraction and use it when creating the target `UnrecognizedChild`.

**Rationale**: Hardcoding `oMath` made typed metadata disagree with an `oMathPara` raw root and broke subsequent carrier-sensitive operations (#124).

Validation also rejects a source or relevant target direct-child carrier when its stored `name` disagrees with the validated raw root local name. Both values must describe the same carrier before mutation.

### Default `OMathSpliceRpRMode = .full`

**Decision**: Default to verbatim `rPr` copy from source Run to the new target OMath Run. Provide `.omathOnly` (whitelist: `rFonts`, legacy `fontName`, `fontSize`, `lang`, `bold`, `italic`) and `.discard` (empty rPr) as escape hatches.

**Alternatives considered**:

- *Default `.discard` (empty rPr)*: simpler, predictable. **Rejected** — Cambria Math font reference would be lost, OMath inherits target paragraph's CJK font and renders broken
- *Default `.omathOnly` (whitelist)*: defensive against cross-doc styleRef / themeColor breakage. **Rejected as default** — too cautious for the common case where source rPr is plain `rFonts` + `lang`; whitelist-skipped fields cause subtle visual differences that surprise callers

**Rationale**: User direction「我想要完整就好」prioritizes visual completeness. Cross-doc rPr risks (rStyle ID not present in target's `word/styles.xml`, themeColor referencing different theme) are documented as caveats and addressable via the two opt-out modes when callers know they have such conflicts.

### Two-tier API: `spliceOMath` (single) + `spliceParagraphOMath` (batch)

**Decision**: Provide both single-OMath low-level entry point and paragraph-level batch entry point. The batch variant preserves explicit absolute-boundary metadata or derives a trailing source-text prefix plus its occurrence instance, groups OMath sharing one boundary, and routes each group to `.atStart`, `.atEnd`, or `.afterText(prefix, instance:)`. Throws `OMathSpliceError.contextAnchorNotFound(omathIndex:, snippet:)` when a prose-derived group cannot resolve.

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

Namespace inspection accepts both XML quote styles and resolves character/entity references through `XmlTreeReader`. Foundation `XMLDocument` is the strict validity gate because the lossless tree parser intentionally does not enforce every namespace/QName/XML 1.0 rule. The candidate is trimmed only by XML 1.0 `S` (space, tab, carriage return, line feed), must begin directly with a root start tag, and is parsed inside a synthetic wrapper as exactly one `oMath`/`oMathPara` element. XML declarations, actual DTDs, U+FEFF/non-XML framing whitespace, document-level comments/PIs/CDATA, extra text, and multiple roots therefore fail structurally, while the same `<!DOCTYPE` text inside an OMath comment or CDATA section remains legal. Extracted prefix-less fragments receive a local default OMML declaration so the copied raw subtree is self-contained; malformed fragments fail before target mutation and an explicitly declared non-standard URI is never defaulted to the standard URI. Strict target comparison uses the first OMath in joint serializer order, including an explicit empty prefix for a default namespace.

Inline carrier discovery uses only the raw fragment's document-element local name (`oMath` or `oMathPara`). Attribute values, character data, descendants, and lookalike root names never opt an ordinary `Run.rawXML` fragment into extraction. The same root classifier is reused by anchor lookup and batch context derivation.

Malformed fragment admission throws the additive `OMathSpliceMalformedXMLError` type rather than adding a case to the already released `OMathSpliceError` enum. This preserves source compatibility for external consumers whose switches exhaustively cover the enum's original six cases.

## Risks / Trade-offs

[**Risk: round-trip semantic loss after splice**] → Mitigation: an end-to-end splice/write/reload test compares extracted source and reloaded OMath through `OMathSemanticXML`. Foundation's parsed semantic tree supplies namespace scope (including `xmlns=""` undeclaration) and XML attribute-whitespace normalization. Canonical coverage equates entity spelling, adjacent text/CDATA segmentation, attribute order/quotes, namespace declaration placement, and element/attribute-prefix variants; it rejects namespace-expanded name, structure, child-order, resolved-text, character-reference whitespace, and semantic-attribute changes without sentinel collisions (#123). Exact lexical bytes are not promised after XML parsing.

[**Risk: anchor-Run split with whitespace-sensitive `<w:t xml:space="preserve">`**] → Mitigation: copy `xml:space` attribute when splitting; existing `replaceText` code path tested with similar fixtures. Add explicit test case for whitespace-bearing anchor.

[**Risk: cross-doc rStyle reference broken in target (e.g., `<w:rStyle w:val="MathStyle">` where target lacks that style ID)**] → Mitigation: documented as `.full` mode caveat; provide `.omathOnly` (strips rStyle) as opt-out. Real risk is low — Cambria Math is typically applied via direct `rFonts`, not via style reference, in NTPU/Word output.

[**Risk: position-renumber regression**] → Mitigation: ordinary boundary insertion does not renumber existing carriers; only direct-child mid-text insertion compacts the position-indexed window needed for a unique cross-collection slot. A sentinel matrix pins all 13 position-indexed collections plus legacy pre/post carriers for inline/direct `.atStart` and `.atEnd`.

[**Risk: mixed-prefix output rejected by strict XML validators**] → Mitigation: default `.lenient` produces ECMA-376-compliant output (mixed prefixes are spec-legal); `.strict` mode available for callers needing single-prefix output for downstream tooling.

[**Risk: MathEquation removal collision**] → Mitigation: this API has zero dependency on `MathEquation`; sister concern #58 tracks the deprecation message inconsistency separately.

[**Risk: future migration to `MathComponent` AST**] → Mitigation: the splice API operates on raw XML, completely orthogonal to the structured AST path. Future API additions on the AST side (LaTeX → structured OMML) do not interact with splice operations.
