# Changelog

All notable changes to ooxml-swift will be documented in this file.

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
