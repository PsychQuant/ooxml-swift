## ADDED Requirements

### Requirement: Single-OMath verbatim splice between documents

The system SHALL provide `WordDocument.spliceOMath(from:toBodyParagraphIndex:position:omathIndex:rPrMode:namespacePolicy:)` that copies a single `<m:oMath>` XML block verbatim from a source `Paragraph` to a target body paragraph at a caller-specified position.

#### Scenario: Inline OMath spliced from source Run to target paragraph end

- **WHEN** caller invokes `target.spliceOMath(from: sourceParagraph, toBodyParagraphIndex: 5, position: .atEnd, omathIndex: 0)` with `sourceParagraph` containing one OMath stored in `Run.rawXML`
- **THEN** the system SHALL append a new `Run` carrying the source OMath block verbatim as a `Run.rawElements` child (NOT `Run.rawXML` — see rationale below) to `target.body.children[5].runs`
- **AND** the spliced Run's `properties` SHALL match the source Run's properties when `rPrMode == .full`
- **AND** `Run.toXML()` of the spliced Run SHALL serialize `<w:rPr>` (when `properties` is non-empty) *before* the OMath body, inside one properly-opened `<w:r>...</w:r>` element
- **AND** the call SHALL return `1` indicating one OMath block was spliced

PsychQuant/ooxml-swift#117: the spliced Run intentionally does NOT use
`Run.rawXML` for its own OMath payload, even though the *source* Run being
read from disk uses that convention. `Run.toXML()` treats `rawXML` as a
verbatim replacement for the entire `<w:r>...</w:r>` element — it returns
before ever reaching the `<w:rPr>` emission code — so a Run built with
`rawXML` set can never also emit a copied `<w:rPr>`. `rawElements` is a
Run *child* carrier: `Run.toXML()` emits `<w:rPr>` (from `properties`) first,
then walks `rawElements` in order, so both survive together. This is a
construction-time distinction only; after the target document is saved and
reloaded, `DocxReader` re-parses the resulting `<w:r><w:rPr>...</w:rPr>
<m:oMath>...</m:oMath></w:r>` the same way it parses any other OMath run —
splitting the parsed `<w:rPr>` into typed `Run.properties` and storing just
the `<m:oMath>` fragment back into `Run.rawXML`, per the reader's normal
(unchanged) convention.

##### Example: Greek-letter inline math splice

- **GIVEN** source paragraph with run containing `rawXML = "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>α</m:t></m:r></m:oMath>"` and rPr `{ rFonts: { ascii: "Cambria Math" }, sz: 24 }`
- **WHEN** caller invokes `target.spliceOMath(from: source, toBodyParagraphIndex: 5, position: .atEnd, omathIndex: 0, rPrMode: .full)`
- **THEN** target.body.children[5].runs gains one Run whose `rawElements` contains one entry (`name: "oMath"`) whose `xml` equals the source OMath XML byte-for-byte
- **AND** that Run's `properties` equals `{ rFonts: { ascii: "Cambria Math" }, sz: 24 }`
- **AND** that Run's `toXML()` output contains `<w:rPr>` before `<m:oMath`

#### Scenario: Direct-child OMath spliced preserving carrier

- **WHEN** caller invokes `target.spliceOMath(...)` with `sourceParagraph` containing OMath stored in `unrecognizedChildren` (i.e., `<m:oMath>` is direct child of `<w:p>`, not inside a Run)
- **THEN** the system SHALL append an `UnrecognizedChild(name: <source root local name>, rawXML: <source OMath XML>, position: <appropriate>)` to `target.body.children[targetIdx].unrecognizedChildren`
- **AND** `name` SHALL be `"oMath"` when the source root is `<m:oMath>` and `"oMathPara"` when the source root is `<m:oMathPara>` (PsychQuant/ooxml-swift#124 — `name` MUST NOT be hardcoded independently of the payload's actual root element)
- **AND** the spliced OMath SHALL NOT be placed inside a `Run` (carrier shape is preserved)

#### Scenario: Source paragraph has no OMath

- **WHEN** caller invokes `spliceOMath` with a `sourceParagraph` whose runs contain no `Run.rawXML` matching `<m:oMath` AND whose `unrecognizedChildren` contain no entry with `name == "oMath" || "oMathPara"`
- **THEN** the system SHALL throw `OMathSpliceError.sourceHasNoOMath`

#### Scenario: omathIndex out of range

- **WHEN** caller invokes `spliceOMath` with `omathIndex: 5` and the source paragraph contains only 3 OMath blocks
- **THEN** the system SHALL throw `OMathSpliceError.omathIndexOutOfRange(requested: 5, available: 3)`

#### Scenario: Target paragraph index out of range

- **WHEN** caller invokes `spliceOMath` with `toBodyParagraphIndex: 9999` and `target.body.children` has fewer than 9999 paragraph entries
- **THEN** the system SHALL throw `OMathSpliceError.targetParagraphOutOfRange(9999)`

### Requirement: Mid-paragraph splice via anchor-text matching

The system SHALL support `OMathSplicePosition.afterText(_, instance:, options:)` and `.beforeText(_, instance:, options:)` to position the spliced OMath relative to a text anchor within the target paragraph.

#### Scenario: Anchor falls in middle of a run, run is split into segments

- **WHEN** caller invokes `target.spliceOMath(from: source, toBodyParagraphIndex: 5, position: .afterText("進行 ", instance: 1), omathIndex: 0)` and the target paragraph at index 5 contains a single run with text `"所得出的參數進行 檢定："`
- **THEN** the system SHALL split that run into two segments: prefix `"所得出的參數進行 "` and suffix `"檢定："`
- **AND** the spliced OMath Run SHALL be inserted between the two segments
- **AND** all three runs (prefix / OMath / suffix) SHALL share the original run's `position` value
- **AND** the original run's `rPr` SHALL be copied to both prefix and suffix segments
- **AND** the relative emit order in `Paragraph.toXML()` SHALL be: prefix run → OMath run → suffix run (stable sort retains array insertion order for equal positions)

##### Example: Mid-prose splice with whitespace anchor

- **GIVEN** target paragraph with one run `text = "所得出的參數進行 檢定：", properties = { rFonts: { eastAsia: "DFKai-SB" } }, position = 3`
- **WHEN** caller calls `spliceOMath(from: source, toBodyParagraphIndex: 5, position: .afterText("進行 "), omathIndex: 0)` with source providing inline OMath `<m:oMath>...t...</m:oMath>`
- **THEN** target paragraph runs after splice:
  | Index | text | OMath payload (`rawElements`) | position |
  |-------|------|--------------------------------|----------|
  | 0 | "所得出的參數進行 " | nil | 3 |
  | 1 | "" | `[RawElement(name: "oMath", xml: "<m:oMath>...t...</m:oMath>")]` | 3 |
  | 2 | "檢定：" | nil | 3 |
- **AND** rPr `{ rFonts: { eastAsia: "DFKai-SB" } }` is copied to runs at indices 0 and 2

#### Scenario: Anchor not found in target paragraph

- **WHEN** caller invokes `spliceOMath` with `position: .afterText("nonexistent text", instance: 1)` and the target paragraph's `flattenedDisplayText()` does not contain `"nonexistent text"`
- **THEN** the system SHALL throw `OMathSpliceError.anchorNotFound("nonexistent text", instance: 1)`

#### Scenario: Anchor instance > 1 resolves to Nth occurrence

- **WHEN** caller invokes `spliceOMath` with `position: .afterText("檢定", instance: 2)` and the target paragraph contains "檢定" three times at character offsets 10, 30, 50
- **THEN** the system SHALL splice the OMath at offset 32 (immediately after the second "檢定" occurrence)

### Requirement: rPr propagation modes

The system SHALL provide `OMathSpliceRpRMode` with three modes controlling how the source Run's `rPr` is copied to the spliced OMath Run.

#### Scenario: .full mode copies rPr verbatim

- **WHEN** caller invokes `spliceOMath(..., rPrMode: .full)` (the default)
- **THEN** the new OMath Run's `properties` SHALL equal the source Run's `properties` (deep copy)

#### Scenario: .omathOnly mode copies whitelisted fields

- **WHEN** caller invokes `spliceOMath(..., rPrMode: .omathOnly)`
- **THEN** the new OMath Run's `properties` SHALL contain ONLY `rFonts`, `sz`, `szCs`, `lang`, `bold`, `italic` from the source
- **AND** all other fields (`rStyle`, `color`, `highlight`, `verticalAlign`, etc.) SHALL be `nil` / default

#### Scenario: .discard mode resets to default rPr

- **WHEN** caller invokes `spliceOMath(..., rPrMode: .discard)`
- **THEN** the new OMath Run's `properties` SHALL equal `RunProperties()` (default-initialized)

### Requirement: Namespace policy controls prefix/URI mismatch handling

The system SHALL provide `OMathSpliceNamespacePolicy` with `.lenient` (default) and `.strict` modes.

#### Scenario: .lenient mode accepts prefix mismatch with same URI

- **WHEN** source OMath uses prefix `mml:` with URI `http://schemas.openxmlformats.org/officeDocument/2006/math`
- **AND** target document standard prefix is `m:` with the same URI
- **AND** caller invokes `spliceOMath(..., namespacePolicy: .lenient)`
- **THEN** the system SHALL splice the source OMath verbatim WITHOUT rewriting prefixes
- **AND** the splice SHALL NOT throw

#### Scenario: .strict mode throws on prefix mismatch

- **WHEN** source uses `mml:` prefix and target uses `m:` prefix (same URI)
- **AND** caller invokes `spliceOMath(..., namespacePolicy: .strict)`
- **THEN** the system SHALL throw `OMathSpliceError.namespaceMismatch(sourceURI: "...math", targetURI: "...math")` (URIs equal but prefixes differ)

#### Scenario: Both modes throw on URI mismatch

- **WHEN** source URI is `http://example.com/vendor/math` and target standard URI is `http://schemas.openxmlformats.org/officeDocument/2006/math`
- **AND** caller invokes `spliceOMath(...)` with either policy
- **THEN** the system SHALL throw `OMathSpliceError.namespaceMismatch(sourceURI:, targetURI:)`

### Requirement: Paragraph-level batch splice with auto-anchor derivation

The system SHALL provide `WordDocument.spliceParagraphOMath(from:toBodyParagraphIndex:rPrMode:namespacePolicy:)` that copies all OMath blocks from one source paragraph to a corresponding target paragraph in source-document order, auto-deriving the splice anchor for each OMath from its source-text-context (5-10 chars on each side).

#### Scenario: All OMath blocks spliced in source order

- **WHEN** source paragraph contains 3 OMath blocks at positions {after "進行 ", after "α=", after "β="} mixed with surrounding prose
- **AND** target paragraph contains the same prose anchors (e.g., from a related document version)
- **AND** caller invokes `target.spliceParagraphOMath(from: sourcePara, toBodyParagraphIndex: 5)`
- **THEN** the system SHALL splice all 3 OMath blocks at the corresponding target locations
- **AND** the call SHALL return `3`

#### Scenario: Context anchor not found for one OMath

- **WHEN** source paragraph has OMath at position whose preceding context is "大小效果"
- **AND** target paragraph's prose says "規模效果" (e.g., advisor changed wording)
- **AND** caller invokes `spliceParagraphOMath(...)`
- **THEN** the system SHALL throw `OMathSpliceError.contextAnchorNotFound(omathIndex: <N>, snippet: "大小效果")`
- **AND** any OMath blocks that were already spliced before this failure SHALL remain in target (partial-success state; caller can inspect target paragraph)

#### Scenario: Source-leading OMath maps to the start of the target paragraph

PsychQuant/ooxml-swift#125: an extracted OMath's derived anchor is one of
three distinct states, which SHALL NOT be collapsed into each other:

1. **Matched, non-empty prefix** — the common case, routes to `.afterText(prefix, ...)`.
2. **Matched, empty prefix** — the source paragraph literally starts with this OMath (no preceding prose). Only meaningful for the first extracted OMath (index 0).
3. **Unmatched** — the OMath could not be paired back to its source Run at all (an extraction bug, e.g. a namespace-normalization mismatch). This SHALL throw, per the "Context anchor not found" scenario above, NOT silently degrade to any boundary.

- **WHEN** source paragraph's first OMath (index 0) has no preceding prose (state 2 above)
- **AND** caller invokes `spliceParagraphOMath(...)`
- **THEN** the system SHALL splice that OMath via `.atStart` of the target paragraph
- **AND** this SHALL NOT be conflated with state 3 (unmatched) merely because both states share an empty derived string internally — the implementation SHALL distinguish them (e.g. via an optional anchor type) before choosing a position

#### Scenario: Namespace self-containment does not prevent matching an extracted OMath back to its source Run

- **GIVEN** a source Run's `rawXML` lacks an explicit `xmlns:m` declaration (inherited from an ancestor `<w:p>` instead) and `OMathExtractor.extract` therefore injects a self-contained `xmlns:m="..."` onto the corresponding `ExtractedOMath.xml`
- **WHEN** `spliceParagraphOMath` derives that OMath's context anchor
- **THEN** the matching step SHALL apply the same namespace-normalization the extractor itself applied before comparing, so this OMath is correctly recognized as matched (not misclassified as unmatched due to a byte-for-byte comparison against the un-normalized original)

#### Scenario: Batch splice order is asserted, not just count

- **WHEN** a batch splice test verifies `spliceParagraphOMath`'s result
- **THEN** it SHALL assert the serialized prefix/OMath/suffix ORDER for every equation (e.g. via `Paragraph.toXML()` substring-index comparison), not merely the returned equation count — a wrong-boundary bug can still return the correct count while visually misplacing every equation

### Requirement: Round-trip lossless guarantee

PsychQuant/ooxml-swift#123: reload fidelity is **canonical semantic XML
equivalence**, not lexical byte identity. Foundation's `XMLDocument` parser
normalizes legal lexical variants during `DocxReader.read` — for example, a
character reference such as `&#x3B1;` reloads as the literal character `α`.
The reloaded XML is semantically equivalent to the source but not
byte-for-byte identical, so a byte-equality requirement is unsatisfiable by
the typed reader architecture (achieving it would require a separate
raw-byte overlay/provenance channel disproportionate to this splice API).
This requirement was previously worded as a byte-equality guarantee; that
wording was the stale artifact, not the implementation.

The OMath XML written into the target paragraph by `spliceOMath` or
`spliceParagraphOMath` SHALL be semantically equivalent to the source OMath
XML when the target document is subsequently saved with `DocxWriter.write`
and reloaded with `DocxReader.read`. Two XML fragments are semantically
equivalent when, after namespace resolution:

- every element and attribute name resolves to the same namespace URI + local
  name (prefix spelling and namespace-declaration placement are NOT
  significant — `<m:oMath>` and `<mml:oMath>` bound to the same URI are the
  same element);
- every attribute has the same resolved value, regardless of attribute order
  or quote style (`"` vs `'`);
- child elements appear in the same order with the same recursively-resolved
  structure;
- text content is equal after entity/character-reference resolution (e.g.
  `&#x3B1;` and `α` are the same text).

#### Scenario: Saved target reloads with spliced OMath semantically equal to source

- **WHEN** caller splices OMath block X from source, calls `DocxWriter.write(target, to: tempURL)`, and reads `let reloaded = try DocxReader.read(from: tempURL)`
- **THEN** the corresponding `Run.rawXML` (PsychQuant/ooxml-swift#117: after a save/reload round trip, the reloaded Run always uses the reader's normal `rawXML` convention, regardless of whether the in-memory pre-save Run used `rawXML` or `rawElements`) in `reloaded` SHALL be semantically equivalent to source's OMath XML per the definition above
- **AND** the comparison SHALL tolerate differing entity spelling, attribute order, attribute quote style, namespace-declaration placement, and namespace-prefix spelling (same URI)

##### Example: Character reference resolves to literal glyph

- **GIVEN** source OMath text node `<m:t>&#x3B1;</m:t>`
- **WHEN** the target document is saved and reloaded
- **THEN** the reloaded OMath text node is `<m:t>α</m:t>`
- **AND** this SHALL be treated as a passing round trip (semantically equal), not a byte-equality failure

##### Example: Attribute order is not significant

- **GIVEN** two source OMath fragments that are semantically identical except for the order two attributes appear on the root element (e.g. `data-a="1" data-b="2"` vs `data-b="2" data-a="1"`)
- **WHEN** each is spliced, saved, and reloaded independently
- **THEN** both reloads SHALL be treated as semantically equal to their respective sources (each attribute's resolved value present, regardless of position)

##### Example: Namespace declaration placement is not significant

- **GIVEN** one source OMath fragment declares `xmlns:m` on the `<m:oMath>` element itself, and another relies entirely on an ancestor `<w:p>`'s `xmlns:m` declaration being in scope (the ordinary shape `DocxReader.parseRun`'s `.xmlString` extraction produces before `ensureXmlnsDeclared` normalizes it)
- **WHEN** each is spliced, saved, and reloaded independently
- **THEN** both reloads SHALL preserve the same glyph content — where the `xmlns:m` declaration appeared in the source is not part of the equivalence comparison

##### Example: Namespace prefix variant is not significant

- **GIVEN** source OMath using the `mml:` prefix (same OMML URI as the target's `m:` convention)
- **WHEN** spliced with `namespacePolicy: .lenient`, saved, and reloaded
- **THEN** the reloaded content preserves the glyph regardless of which prefix spelling (`m:` or `mml:`) the reload happens to use

### Requirement: No regression on existing OMath round-trip behavior

The introduction of `spliceOMath` / `spliceParagraphOMath` APIs SHALL NOT alter the round-trip behavior of OMath blocks that were not spliced.

#### Scenario: Pre-existing OMath in target paragraph preserved during splice

- **WHEN** target paragraph already contains 2 OMath blocks before splice
- **AND** caller splices a 3rd OMath block via `spliceOMath(..., position: .atEnd, ...)`
- **THEN** the original 2 OMath blocks SHALL remain in target with original `rawXML` and `position`
- **AND** only the new 3rd OMath block SHALL be added at the end

#### Scenario: Existing #85 / #92 / #99-103 fixture suites pass unchanged

- **WHEN** the `Issue85InlineMathFlattenTests`, `Issue92OMMLWalkSurfaceCoverageTests`, and `Issue99FlattenReplaceOMMLBilateralTests` test suites run after this change merges
- **THEN** all tests SHALL pass without modification (additive feature does not affect existing OMath behavior)

### Requirement: `.atStart` / `.atEnd` serialize relative to every paragraph carrier, not just Runs

PsychQuant/ooxml-swift#122: `Paragraph.toXMLSortedByPosition()` interleaves
13 position-indexed carriers (runs, hyperlinks, fieldSimples,
alternateContents, bookmarkMarkers, commentRangeMarkers,
permissionRangeMarkers, proofErrorMarkers, smartTags, customXmlBlocks,
bidiOverrides, unrecognizedChildren, contentControls) by `position`, and
additionally emits any nil/zero-position ("API-built") entry in a fixed
post-content section that always renders AFTER the entire positioned
section — regardless of how large a positive position a competing entry
has. `.atStart` / `.atEnd` SHALL account for all of this, not just
`Run.position` values.

#### Scenario: `.atEnd` serializes after an API-built (nil-position) Run

- **GIVEN** target paragraph containing exactly one API-built `Run` (`position == nil`, e.g. `Run(text: "existing")` never touched by a positioning API)
- **WHEN** caller invokes `spliceOMath(..., position: .atEnd, ...)`
- **THEN** the spliced OMath Run SHALL serialize after that existing Run's text in `Paragraph.toXML()` output

#### Scenario: `.atEnd` serializes after a high-position direct-child carrier

- **GIVEN** target paragraph containing a direct-child OMath (`unrecognizedChildren`) at `position: 100`
- **WHEN** caller invokes `spliceOMath(from: source, ..., position: .atEnd, ...)` with an inline (Run-carried) source OMath
- **THEN** the spliced inline OMath Run SHALL serialize after that direct-child carrier, not before it (boundary computation SHALL consider all 13 carriers' positions, not `targetPara.runs` alone)

#### Scenario: `.atStart` precedes a lower-positioned non-Run carrier

- **GIVEN** target paragraph containing a direct-child OMath at `position: 1` and a Run at `position: 5`
- **WHEN** caller invokes `spliceOMath(from: source, ..., position: .atStart, ...)` with an inline source OMath
- **THEN** the spliced inline OMath Run SHALL serialize before the direct-child carrier (boundary computation's minimum SHALL span all 13 carriers, not `targetPara.runs` alone)

#### Scenario (known limitation, NOT guaranteed): `.atEnd` relative to legacy `footnoteIds`/`endnoteIds`

R2 (independent review of this change, Finding #1): `Paragraph.footnoteIds` /
`endnoteIds` (populated by the public `WordDocument.insertFootnote(text:paragraphIndex:)`
/ `insertEndnote(text:paragraphIndex:)` APIs, which can target any existing
paragraph) are NOT one of the 13 position-indexed carriers this requirement
covers. `Paragraph.sortByPositionPostContentXML()` emits them via a
dedicated, unconditional loop at the very end of the function — structurally
after every positioned entry and every other post-content entry, with no
`position` field on `footnoteIds`/`endnoteIds` (a bare `[Int]`) to
materialize them into the way `materializeZeroPositionCarriers` promotes
every other zero-position carrier.

- **GIVEN** target paragraph containing body text and a footnote reference added via `insertFootnote(text:paragraphIndex:)`
- **WHEN** caller invokes `spliceOMath(..., position: .atEnd, ...)`
- **THEN** the spliced OMath SHALL serialize after the body text (guaranteed)
- **BUT** the spliced OMath is NOT guaranteed to serialize after the footnote reference — it currently does not, and this requirement's "after every existing paragraph carrier" wording does NOT extend to `footnoteIds`/`endnoteIds` (or, by the same unconditional-tail-loop structure, `endnoteIds`)

##### Scope note

This requirement covers exactly the reproductions above (API-built Runs;
mixed Run / direct-child carriers; the 13 position-indexed carriers listed in
this Requirement's opening paragraph). It does NOT cover:

- negative positions or `Int.max` collisions;
- paragraphs in tree-backed mode (`Paragraph.xmlNode != nil`) — none of
  `DocxReader.parseParagraph`'s output is tree-backed, so tree-backed splice
  target support remains a pre-existing, unaffected gap rather than something
  this requirement guarantees;
- `Paragraph.footnoteIds` / `endnoteIds` (see the scenario immediately above)
  or, symmetrically, `hasPageBreak` (a page-break `<w:br>` is likewise emitted
  by a dedicated unconditional loop, in `sortByPositionPreContentXML()`,
  structurally BEFORE the entire positioned section — so `.atStart` is
  correspondingly not guaranteed to precede it either; not separately
  reproduced with a test, but the same code-structure argument applies by
  symmetry).

Making `.atEnd`/`.atStart` genuinely account for these legacy, position-less
fields would require changing how `Paragraph`'s core serializer represents
footnote/endnote references and page breaks (e.g. giving them a position
field the way bookmarks have both legacy `bookmarks` and positioned
`bookmarkMarkers`) — a change to the paragraph model used throughout the
library, not something scoped to the OMath splice API. Deliberately not
attempted in this change.
