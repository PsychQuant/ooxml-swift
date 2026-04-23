# Changelog

All notable changes to ooxml-swift will be documented in this file.

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
