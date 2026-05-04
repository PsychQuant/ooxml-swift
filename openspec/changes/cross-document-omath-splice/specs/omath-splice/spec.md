## ADDED Requirements

### Requirement: Single-OMath verbatim splice between documents

The system SHALL provide `WordDocument.spliceOMath(from:toBodyParagraphIndex:position:omathIndex:rPrMode:namespacePolicy:)` that copies a single `<m:oMath>` XML block verbatim from a source `Paragraph` to a target body paragraph at a caller-specified position.

#### Scenario: Inline OMath spliced from source Run.rawXML to target paragraph end

- **WHEN** caller invokes `target.spliceOMath(from: sourceParagraph, toBodyParagraphIndex: 5, position: .atEnd, omathIndex: 0)` with `sourceParagraph` containing one OMath stored in `Run.rawXML`
- **THEN** the system SHALL append a new `Run` with `rawXML` byte-equal to the source OMath block to `target.body.children[5].runs`
- **AND** the spliced Run's `properties` SHALL match the source Run's properties when `rPrMode == .full`
- **AND** the call SHALL return `1` indicating one OMath block was spliced

##### Example: Greek-letter inline math splice

- **GIVEN** source paragraph with run containing `rawXML = "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>α</m:t></m:r></m:oMath>"` and rPr `{ rFonts: { ascii: "Cambria Math" }, sz: 24 }`
- **WHEN** caller invokes `target.spliceOMath(from: source, toBodyParagraphIndex: 5, position: .atEnd, omathIndex: 0, rPrMode: .full)`
- **THEN** target.body.children[5].runs gains one Run whose `rawXML` equals the source OMath XML byte-for-byte
- **AND** that Run's `properties` equals `{ rFonts: { ascii: "Cambria Math" }, sz: 24 }`

#### Scenario: Direct-child OMath spliced preserving carrier

- **WHEN** caller invokes `target.spliceOMath(...)` with `sourceParagraph` containing OMath stored in `unrecognizedChildren` (i.e., `<m:oMath>` is direct child of `<w:p>`, not inside a Run)
- **THEN** the system SHALL append an `UnrecognizedChild(name: "oMath", rawXML: <source OMath XML>, position: <appropriate>)` to `target.body.children[targetIdx].unrecognizedChildren`
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

### Requirement: Round-trip lossless guarantee

The OMath XML written into the target paragraph by `spliceOMath` or `spliceParagraphOMath` SHALL be byte-equal to the source OMath XML when the target document is subsequently saved with `DocxWriter.write` and reloaded with `DocxReader.read`.

#### Scenario: Saved target reloads with spliced OMath byte-equal to source

- **WHEN** caller splices OMath block X from source, calls `DocxWriter.write(target, to: tempURL)`, and reads `let reloaded = try DocxReader.read(from: tempURL)`
- **THEN** the corresponding `Run.rawXML` (or `unrecognizedChildren[].rawXML`) in `reloaded` SHALL equal source's OMath XML byte-for-byte

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
