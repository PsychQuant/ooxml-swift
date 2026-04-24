# Changelog

All notable changes to ooxml-swift will be documented in this file.

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
