## ADDED Requirements

### Requirement: Single-OMath verbatim splice between documents

The system SHALL provide `WordDocument.spliceOMath(from:toBodyParagraphIndex:position:omathIndex:rPrMode:namespacePolicy:)` that copies a single `<m:oMath>` XML block verbatim from a source `Paragraph` to a target body paragraph at a caller-specified position.

#### Scenario: Inline OMath spliced from source Run.rawXML to target paragraph end

- **WHEN** caller invokes `target.spliceOMath(from: sourceParagraph, toBodyParagraphIndex: 5, position: .atEnd, omathIndex: 0)` with `sourceParagraph` containing one OMath stored in `Run.rawXML`
- **THEN** the system SHALL append a new `Run` whose `rawXML` equals the extractor's self-contained OMath fragment to `target.body.children[5].runs`
- **AND** the spliced Run's `properties` SHALL match the source Run's properties when `rPrMode == .full`
- **AND** the call SHALL return `1` indicating one OMath block was spliced

##### Example: Greek-letter inline math splice

- **GIVEN** source paragraph with run containing `rawXML = "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>α</m:t></m:r></m:oMath>"` and rPr `{ rFonts: { ascii: "Cambria Math" }, sz: 24 }`
- **WHEN** caller invokes `target.spliceOMath(from: source, toBodyParagraphIndex: 5, position: .atEnd, omathIndex: 0, rPrMode: .full)`
- **THEN** target.body.children[5].runs gains one Run whose `rawXML` equals the extractor's self-contained OMath XML
- **AND** that Run's `properties` equals `{ rFonts: { ascii: "Cambria Math" }, sz: 24 }`

#### Scenario: Direct-child OMath spliced preserving carrier

- **WHEN** caller invokes `target.spliceOMath(...)` with `sourceParagraph` containing OMath stored in `unrecognizedChildren` (i.e., `<m:oMath>` is direct child of `<w:p>`, not inside a Run)
- **THEN** the system SHALL append an `UnrecognizedChild(name: "oMath", rawXML: <source OMath XML>, position: <appropriate>)` to `target.body.children[targetIdx].unrecognizedChildren`
- **AND** the spliced OMath SHALL NOT be placed inside a `Run` (carrier shape is preserved)

#### Scenario: Direct-child OMathPara preserves local-name metadata

- **WHEN** the source direct child root is `<m:oMathPara>`
- **THEN** the target `UnrecognizedChild.name` SHALL be `oMathPara`
- **AND** save/reload and subsequent extraction SHALL preserve matching metadata and raw root

#### Scenario: At-end follows every paragraph carrier

- **GIVEN** a target paragraph containing positive-position carriers, nil/zero-position API-built carriers, and legacy pre/post carriers such as page breaks, comments, footnotes, endnotes, or bookmarks
- **WHEN** caller splices inline or direct-child OMath at `.atEnd`
- **THEN** every existing carrier SHALL retain its relative serialized order
- **AND** the new OMath SHALL serialize after all of them
- **AND** hostile or API-built positions such as `Int.max` SHALL neither overflow nor be rewritten for ordinary boundary insertion
- **AND** negative-position carriers SHALL remain visible in the non-positive post-content bucket rather than being dropped

#### Scenario: At-start precedes every paragraph carrier

- **GIVEN** a target paragraph containing position-indexed and legacy pre/post carriers
- **WHEN** caller splices inline or direct-child OMath at `.atStart`
- **THEN** the new OMath SHALL serialize before every existing carrier

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
  | Index | text | rawXML | position |
  |-------|------|--------|----------|
  | 0 | "所得出的參數進行 " | nil | 3 |
  | 1 | "" | "<m:oMath>...t...</m:oMath>" | 3 |
  | 2 | "檢定：" | nil | 3 |
- **AND** rPr `{ rFonts: { eastAsia: "DFKai-SB" } }` is copied to runs at indices 0 and 2

#### Scenario: Anchor not found in target paragraph

- **WHEN** caller invokes `spliceOMath` with `position: .afterText("nonexistent text", instance: 1)` and the target paragraph's `flattenedDisplayText()` does not contain `"nonexistent text"`
- **THEN** the system SHALL throw `OMathSpliceError.anchorNotFound("nonexistent text", instance: 1)`

#### Scenario: Anchor instance > 1 resolves to Nth occurrence

- **WHEN** caller invokes `spliceOMath` with `position: .afterText("檢定", instance: 2)` and the target paragraph contains "檢定" three times at character offsets 10, 30, 50
- **THEN** the system SHALL splice the OMath at offset 32 (immediately after the second "檢定" occurrence)

#### Scenario: Direct-child carrier honors text anchor

- **WHEN** the source OMath is a paragraph direct child and caller uses `.afterText` or `.beforeText`
- **THEN** the direct child SHALL serialize at the resolved boundary between split target Runs
- **AND** a missing anchor SHALL throw `anchorNotFound` instead of appending at paragraph end

#### Scenario: Math-script-insensitive anchor retains original split offsets

- **WHEN** caller enables `AnchorLookupOptions(mathScriptInsensitive: true)` and the target contains a Unicode math-script variant such as `H₀`
- **THEN** an ASCII anchor such as `H0` SHALL resolve
- **AND** the Run SHALL split on valid original UTF-16 boundaries

### Requirement: rPr propagation modes

The system SHALL provide `OMathSpliceRpRMode` with three modes controlling how the source Run's `rPr` is copied to the spliced OMath Run.

#### Scenario: .full mode copies rPr verbatim

- **WHEN** caller invokes `spliceOMath(..., rPrMode: .full)` (the default)
- **THEN** the new OMath Run's `properties` SHALL equal the source Run's `properties` (deep copy)

#### Scenario: .omathOnly mode copies whitelisted fields

- **WHEN** caller invokes `spliceOMath(..., rPrMode: .omathOnly)`
- **THEN** the new OMath Run's `properties` SHALL contain ONLY `rFonts`, legacy `fontName`, `fontSize`, `lang`, `bold`, and `italic` from the source
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
- **AND** namespace declarations using either single or double quotes SHALL be parsed equivalently
- **AND** a prefix-less extracted OMath fragment SHALL receive a self-contained default OMML namespace declaration

#### Scenario: Malformed source OMath fails before mutation

- **WHEN** the source OMath root or namespace attributes are not well-formed XML
- **THEN** splice SHALL throw `OMathSpliceMalformedXMLError`
- **AND** the target paragraph SHALL remain unchanged
- **AND** duplicate namespace attributes, unbound prefixes, invalid QNames, forbidden XML 1.0 character references, and DTD-bearing fragments SHALL be rejected

### Requirement: Paragraph-level batch splice with auto-anchor derivation

The system SHALL provide `WordDocument.spliceParagraphOMath(from:toBodyParagraphIndex:rPrMode:namespacePolicy:)` that copies all OMath blocks from one source paragraph to a corresponding target paragraph in source-document order, auto-deriving each anchor from up to 10 trailing prose characters plus its occurrence instance.

#### Scenario: All OMath blocks spliced in source order

- **WHEN** source paragraph contains 3 OMath blocks at positions {after "進行 ", after "α=", after "β="} mixed with surrounding prose
- **AND** target paragraph contains the same prose anchors (e.g., from a related document version)
- **AND** caller invokes `target.spliceParagraphOMath(from: sourcePara, toBodyParagraphIndex: 5)`
- **THEN** the system SHALL splice all 3 OMath blocks at the corresponding target locations
- **AND** the call SHALL return `3`

#### Scenario: Leading OMath maps to target start

- **WHEN** a source OMath has a successfully derived empty prefix because it leads the paragraph
- **THEN** batch splice SHALL map it to `.atStart`, not `.atEnd`

#### Scenario: Namespace self-containment does not break source matching

- **WHEN** extraction adds a root namespace declaration to an inline OMath fragment
- **THEN** batch anchor derivation SHALL still match it to the originating Run
- **AND** the OMath SHALL serialize between its corresponding target prefix and suffix

#### Scenario: Unmatched batch anchor fails loudly

- **WHEN** an extracted OMath cannot be matched to source text context
- **THEN** batch splice SHALL throw `contextAnchorNotFound`
- **AND** it SHALL NOT silently append the OMath at a boundary

#### Scenario: Repeated boundary anchors preserve source order

- **WHEN** multiple OMath blocks share the same anchor, including consecutive leading equations
- **THEN** their final serialized order SHALL equal source-document order
- **AND** identical text snippets at different source occurrences SHALL map to the corresponding target occurrence rather than always instance 1

#### Scenario: Re-extraction follows absolute boundary order

- **WHEN** multiple inline and/or direct-child OMath carriers are inserted into the same absolute boundary
- **THEN** `OMathExtractor.extract` and `omathIndex` SHALL expose them in the same order as paragraph serialization
- **AND** an unsaved direct-child absolute-boundary paragraph SHALL be usable immediately as a batch source

#### Scenario: Unsaved absolute-end batch source remains at end

- **WHEN** one or more OMath carriers with explicit absolute-end metadata are reused as a batch source before save/reload
- **THEN** batch splice SHALL map the group to `.atEnd` even when the source has no preceding prose
- **AND** the target SHALL serialize the group after all prose and legacy post-content carriers
- **AND** multiple members of the end group SHALL retain source-document order

#### Scenario: Context anchor not found for one OMath

- **WHEN** source paragraph has OMath at position whose preceding context is "大小效果"
- **AND** target paragraph's prose says "規模效果" (e.g., advisor changed wording)
- **AND** caller invokes `spliceParagraphOMath(...)`
- **THEN** the system SHALL throw `OMathSpliceError.contextAnchorNotFound(omathIndex: <N>, snippet: "大小效果")`
- **AND** any earlier source anchor groups that were already spliced before this failure SHALL remain in target
- **AND** the failing shared-anchor group SHALL be preflighted as a unit so it does not partially mutate the target

### Requirement: Round-trip semantic XML equivalence

The OMath XML written into the target paragraph by `spliceOMath` or `spliceParagraphOMath` SHALL be semantically XML-equivalent to the extracted source OMath when the target document is subsequently saved with `DocxWriter.write` and reloaded with `DocxReader.read`. Equivalence SHALL compare namespace-expanded element identity, semantic attributes independent of order/quote style, ordered child structure, and resolved text values. Namespace prefix spelling/declaration placement and character/entity spelling SHALL NOT be treated as semantic differences.

#### Scenario: Saved target reloads with semantically equivalent OMath

- **WHEN** caller splices OMath block X from source, calls `DocxWriter.write(target, to: tempURL)`, and reads `let reloaded = try DocxReader.read(from: tempURL)`
- **THEN** the reloaded OMath subtree SHALL have the same `OMathSemanticXML` canonical representation as the extracted source subtree
- **AND** entity versus literal spelling, attribute order/quotes, namespace declaration placement, and equivalent element-prefix variants SHALL compare equal
- **AND** a changed element, semantic attribute value, child order, or resolved text value SHALL compare unequal
- **AND** adjacent ordinary text and CDATA segments with the same resolved text SHALL compare equal
- **AND** `xmlns=""` SHALL remove an inherited default namespace for the element and its descendants
- **AND** literal attribute tabs/newlines/carriage returns SHALL compare according to XML whitespace normalization, while a character-referenced tab SHALL remain semantically distinct from a normalized space

#### Scenario: Admission requires one embeddable OMath fragment

- **WHEN** extracted raw XML contains an XML declaration, an actual DTD, document-level comment/PI, non-whitespace framing text, multiple roots, or a non-OMath root
- **THEN** splice SHALL throw `OMathSpliceMalformedXMLError` before target mutation
- **AND** only XML 1.0 `S` characters (space, tab, carriage return, line feed) MAY surround the root; U+FEFF, non-XML whitespace, or document-level CDATA SHALL be rejected
- **AND** a valid OMath whose internal comment or CDATA contains the ordinary text `<!DOCTYPE` SHALL remain admissible
- **AND** paragraph-level batch splice SHALL validate every extracted source fragment and the initial target before anchor derivation or any group mutation
- **AND** the complete XML-validity scan SHALL precede namespace URI/prefix policy checks, so malformed-input error precedence is independent of source order
- **AND** malformed XML in any batch item SHALL leave the entire target unchanged, independent of source group order or anchor availability

### Requirement: No regression on existing OMath round-trip behavior

The introduction of `spliceOMath` / `spliceParagraphOMath` APIs SHALL NOT alter the round-trip behavior of OMath blocks that were not spliced.

#### Scenario: Pre-existing OMath in target paragraph preserved during splice

- **WHEN** target paragraph already contains 2 OMath blocks before splice
- **AND** caller splices a 3rd OMath block via `spliceOMath(..., position: .atEnd, ...)`
- **THEN** the original 2 OMath blocks SHALL remain in target with original `rawXML`, numeric positions, and relative serialized order
- **AND** only the new 3rd OMath block SHALL be added at the end

#### Scenario: Existing #85 / #92 / #99-103 fixture suites pass unchanged

- **WHEN** the `Issue85InlineMathFlattenTests`, `Issue92OMMLWalkSurfaceCoverageTests`, and `Issue99FlattenReplaceOMMLBilateralTests` test suites run after this change merges
- **THEN** all tests SHALL pass without modification (additive feature does not affect existing OMath behavior)
