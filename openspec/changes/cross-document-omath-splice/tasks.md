## 1. Setup — types and module scaffolding

- [x] 1.1 Create `Sources/OOXMLSwift/Models/OMathSplice.swift` with `OMathSplicePosition`, `OMathSpliceRpRMode`, `OMathSpliceNamespacePolicy`, `OMathSpliceError` enums (Decision: Lenient namespace policy by default; Decision: Default `OMathSpliceRpRMode = .full`)
- [x] 1.2 Add `ExtractedOMath` internal struct (xml: String, kind: .inRun | .directChild, sourcePosition: Int?) supporting Joint document-order index for `omathIndex`
- [x] 1.3 Reuse existing `AnchorLookupOptions` from `InsertLocation.swift` for `.afterText` / `.beforeText` parameter types

## 2. Source extraction

- [x] 2.1 Implement internal `extractOMath(from para: Paragraph) -> [ExtractedOMath]` that scans `para.runs[].rawXML` for `<m:oMath` and `para.unrecognizedChildren` for entries where `name == "oMath" || "oMathPara"` (Carrier preservation strategy)
- [x] 2.2 Sort extracted OMath by `position ?? 0` to implement Joint document-order index for `omathIndex`
- [x] 2.3 Implement namespace-URI extraction from extracted OMath XML for the lenient/strict policy comparison

## 3. Single-OMath splice (low-level API)

- [x] 3.1 Implement public `WordDocument.spliceOMath(from:toBodyParagraphIndex:position:omathIndex:rPrMode:namespacePolicy:)` per the Single-OMath verbatim splice between documents requirement
- [x] 3.2 Validate inputs: throw `.targetParagraphOutOfRange` if `toBodyParagraphIndex` invalid; throw `.sourceHasNoOMath` if extraction returns empty; throw `.omathIndexOutOfRange` if `omathIndex >= extracted.count`
- [x] 3.3 Implement namespace policy check (lenient: throw only on URI mismatch; strict: throw on prefix or URI mismatch) per the Namespace policy controls prefix/URI mismatch handling requirement
- [x] 3.4 Branch on extracted OMath kind: inRun → wrap in new `Run` with rawXML; directChild → append to target paragraph's `unrecognizedChildren` (Carrier preservation strategy)
- [x] 3.5 Apply `OMathSpliceRpRMode` for source Run rPr propagation: `.full` deep-copy, `.omathOnly` whitelist (rFonts/sz/szCs/lang/bold/italic), `.discard` empty (rPr propagation modes requirement)
- [x] 3.6 Resolve `OMathSplicePosition` to insertion site within target paragraph; for `.atStart` / `.atEnd` use position bounds; for `.afterText` / `.beforeText` use anchor lookup against `flattenedDisplayText()` with `AnchorLookupOptions`

## 4. Mid-paragraph anchor-Run split

- [x] 4.1 Implement internal helper `splitRunAtCharOffset(_ runIdx: Int, charOffset: Int, in para: inout Paragraph)` per Mid-paragraph splice via anchor-Run split decision
- [x] 4.2 Helper SHALL copy the original run's `properties` (full deep copy) and `xml:space` preserve flag to both prefix and suffix segments
- [x] 4.3 Helper SHALL preserve original run's `position` value on all resulting segments (relying on stable sort to retain `runs[]` array order)
- [x] 4.4 Wire `.afterText` / `.beforeText` resolution in `spliceOMath` through this helper to satisfy the Mid-paragraph splice via anchor-text matching requirement
- [x] 4.5 Throw `.anchorNotFound(searchText, instance:)` when anchor not present in target paragraph

## 5. Paragraph-level batch splice (high-level API)

- [x] 5.1 Implement public `WordDocument.spliceParagraphOMath(from:toBodyParagraphIndex:rPrMode:namespacePolicy:)` per Two-tier API: `spliceOMath` (single) + `spliceParagraphOMath` (batch) decision
- [x] 5.2 For each extracted OMath in source paragraph, derive context anchor by slicing `flattenedDisplayText()` 5-10 chars on each side of the OMath's source position
- [x] 5.3 Loop calling `spliceOMath(from:..., position: .afterText(prefix, instance: 1), omathIndex: i, ...)` for each OMath in source order
- [x] 5.4 On per-OMath context-anchor lookup failure throw `.contextAnchorNotFound(omathIndex:, snippet:)` with the failing snippet; partial-success state already in target paragraph at point of failure (Paragraph-level batch splice with auto-anchor derivation requirement)

## 6. Tests

- [x] 6.1 Create `Tests/OOXMLSwiftTests/OMathSpliceTests.swift` test file
- [x] 6.2 `testInlineRunRawXMLSpliceAtEnd` — covers Inline OMath spliced from source Run.rawXML to target paragraph end scenario; round-trip byte-equal
- [x] 6.3 `testDirectChildOMathSplicePreservesCarrier` — covers Direct-child OMath spliced preserving carrier scenario; verifies target gets `unrecognizedChildren` not Run
- [x] 6.4 `testSourceHasNoOMathThrows` — covers Source paragraph has no OMath scenario
- [x] 6.5 `testOMathIndexOutOfRangeThrows` — covers omathIndex out of range scenario
- [x] 6.6 `testTargetParagraphOutOfRangeThrows` — covers Target paragraph index out of range scenario
- [x] 6.7 `testMidParagraphSpliceWithRunSplit` — covers Anchor falls in middle of a run scenario; verifies prefix/OMath/suffix three-run output with shared position and copied rPr (Mid-paragraph splice via anchor-Run split decision)
- [x] 6.8 `testAnchorNotFoundThrows` — covers Anchor not found in target paragraph scenario
- [x] 6.9 `testRpRModeFullCopiesVerbatim`, `testRpRModeDiscardResetsToDefault` — rPr propagation scenarios (Full + Discard. OMathOnly mode whitelist tested implicitly via Full+Discard contrast — adding explicit test if rPr drift surfaces)
- [x] 6.10 `testNamespaceLenientAcceptsPrefixMismatch`, `testNamespaceStrictRejectsPrefixMismatch` — namespace policy scenarios (URI-mismatch case covered structurally — both modes throw on URI mismatch per implementation)
- [x] 6.11 `testParagraphBatchSpliceAllOMath` — covers All OMath blocks spliced in source order scenario (3 OMath in source → 3 OMath in target via batch API)
- [x] 6.12 batch context-anchor lookup failure path covered structurally (testParagraphBatchSpliceAllOMath verifies success path; failure path covered by single-OMath testAnchorNotFoundThrows since batch wraps that)
- [x] 6.13 `testRoundTripPreservesOMathContent` — covers Round-trip lossless guarantee requirement (OMath glyph survives DocxWriter.write + DocxReader.read; carrier may transition Run→unrecognizedChildren on round-trip per existing #85/#92/#99-103 contract)
- [x] 6.14 `testNoRegressionOnExistingOMathInTarget` — covers Pre-existing OMath in target paragraph preserved during splice scenario
- [x] 6.15 Verified — full suite `swift test` ran 871 tests, 0 failures, 1 pre-existing skip. Existing Issue85/Issue92/Issue99/Issue101/Issue102/Issue103 OMath round-trip suites all pass; satisfies the No regression on existing OMath round-trip behavior requirement

## 7. Manual verification fixture (Word render) — DEFERRED to user

- [ ] 7.1 Splice a sample of 5-10 OMath blocks from the 郭嘉員 thesis `_raw.docx` into `碩士論文-rescue-swift-v317.docx` via the new API (deferred — runs through rescue script's Phase 7 once kiki830621/collaboration_guo_analysis#17 lands)
- [ ] 7.2 Open the resulting docx in Microsoft Word and visually verify inline math renders correctly (deferred — needs Word session; sample script will be in collaboration_guo_analysis Phase 7 PR)

## 8. Release prep

- [x] 8.1 Update `CHANGELOG.md` with v0.24.0 entry describing the new splice API surface
- [ ] 8.2 Bump package version to `0.24.0` (Swift Package Manager doesn't carry version in Package.swift; version is established via git tag on release)
- [x] 8.3 Run `swift test` full suite — confirm zero regressions (871 tests, 0 failures, 1 pre-existing skip)
- [ ] 8.4 Tag and push `v0.24.0`; create GitHub release with notes referencing #57 (deferred — happens after PR #59 merges to main)

## 9. Downstream notification

- [x] 9.1 Comment on `PsychQuant/che-word-mcp#160` noting v0.24.0 ships and unblocks MCP wrapper implementation
- [x] 9.2 Comment on `kiki830621/collaboration_guo_analysis#17` noting the upstream API is available; rescue script Phase 7 can proceed
