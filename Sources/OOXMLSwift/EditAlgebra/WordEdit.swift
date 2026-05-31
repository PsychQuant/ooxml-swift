// WordEdit.swift
// EditAlgebra — Phase 2 of ooxml-edit-isomorphism-foundation (PsychQuant/macdoc#99)
//
// WordEdit is the semantic-layer Edit enum — case names follow Word UI verbs
// (foundation ADR-005 + ADR-006: Word UI behavior as ground truth).
//
// 3 canonical cases in this Phase 2 ship:
//   - applyBold(range:)              ↔ Word UI: select + Cmd-B
//   - applyLink(range:url:)          ↔ Word UI: select + Insert→Hyperlink
//   - applyInsertParagraph(after:content:) ↔ Word UI: position cursor + Enter + type
//
// `lower()` translates each WordEdit case to a `[OOXMLEdit]` decomposition.
// Range-crossing-paragraph cases produce N-element lists (one OOXMLEdit per
// affected paragraph), per foundation ADR-002 Worked Example 4.
//
// Naturality invariant (tested in NaturalityTests):
//   (WordEdit.a ∘ WordEdit.b).lower() == WordEdit.a.lower() ∘ WordEdit.b.lower()

import Foundation

// MARK: - WordRange

/// Identifies a contiguous range of text within a WordDocument by start/end
/// Run ElementIDs and character offsets within those Runs.
///
/// Range validity is scoped to the WordDocument instance the range was
/// resolved against — IDs may become stale if the document is mutated
/// between range creation and `lower()`. `lower()` throws
/// `EditError.pathNotFound` if either ID doesn't resolve.
public struct WordRange: Equatable, Sendable {
    /// ElementID of the first Run containing the start of the range.
    public let startRun: ElementID

    /// Character offset within `startRun`'s text. 0 means "before the first
    /// character"; `startRun.text.count` means "after the last character".
    public let startOffset: Int

    /// ElementID of the last Run containing the end of the range. May equal
    /// `startRun` for single-Run ranges.
    public let endRun: ElementID

    /// Character offset within `endRun`'s text. Semantics same as `startOffset`.
    public let endOffset: Int

    public init(startRun: ElementID, startOffset: Int, endRun: ElementID, endOffset: Int) {
        self.startRun = startRun
        self.startOffset = startOffset
        self.endRun = endRun
        self.endOffset = endOffset
    }
}

// MARK: - ParagraphRef

/// Stable identifier for a paragraph within a WordDocument. Resolves to an
/// `ElementID` at `lower()` time.
///
/// Phase 2 implementation: thin wrapper over `ElementID`; future versions may
/// support paragraph-by-text-search or paragraph-by-style-and-position
/// addressing.
public struct ParagraphRef: Equatable, Sendable {
    public let elementID: ElementID

    public init(_ elementID: ElementID) {
        self.elementID = elementID
    }
}

// MARK: - WordEdit

/// Semantic-layer Edit — addresses Word UI verbs and translates to
/// `OOXMLEdit` via `lower()`.
///
/// Case names follow Word UI conventions (foundation ADR-005: verb prefix
/// `apply*` disambiguates Edit-mutation from property-accessor style).
public enum WordEdit: Edit, Equatable, Sendable {
    /// Apply bold formatting to the selected range — equivalent to Word UI
    /// Cmd-B on a selection.
    ///
    /// Within-single-paragraph: lower() returns a 1-element list with one
    /// `OOXMLEdit.setBold`.
    /// Range crosses paragraph boundary: lower() returns N-element list (one
    /// `OOXMLEdit.setBold` per affected paragraph), per foundation ADR-002
    /// Worked Example 4.
    case applyBold(range: WordRange)

    /// Apply a hyperlink to the selected range — equivalent to Word UI
    /// Insert → Hyperlink (or Cmd-K).
    ///
    /// Lowers to `OOXMLEdit.insertHyperlink`. Atomicity (document.xml +
    /// rels-part) is handled at OOXMLEdit level.
    case applyLink(range: WordRange, url: URL)

    /// Insert a new paragraph after the specified paragraph — equivalent to
    /// Word UI positioning the cursor at end of paragraph + Enter + typing.
    ///
    /// Lowers to `OOXMLEdit.insertParagraph`.
    case applyInsertParagraph(after: ParagraphRef, content: String)

    // MARK: - Edit protocol conformance

    /// Apply this WordEdit by calling `lower()` then applying each
    /// `OOXMLEdit` in sequence via `WordDocument.apply(_:)`. Stub in §1
    /// scaffold — full wiring lands in §2 (Document.apply public API).
    public func apply(to document: WordDocument) throws -> WordDocument {
        throw EditError.notImplemented("WordEdit.apply requires Document.apply(_:) wiring (§2 of #105 tasks)")
    }

    /// Lower this WordEdit to its `[OOXMLEdit]` translation.
    ///
    /// Per macdoc#105 design.md Decision 2 + spec.md "WordEdit Enum with 3
    /// Canonical Cases". `lower()` is non-throwing and takes no document
    /// context (per the `Edit` protocol). Cases that require doc context to
    /// resolve (e.g., cross-paragraph WordRange in applyBold needs to
    /// enumerate intermediate Runs) return `[]`, which triggers the silent-
    /// noop guard in `WordDocument.apply` → throws `notImplemented` with a
    /// message identifying the unsupported input combination.
    public func lower() -> [OOXMLEdit] {
        switch self {

        case .applyBold(let range):
            // Single-Run case (startRun == endRun): 1:1 mapping to
            // OOXMLEdit.setBold(target: startRun, value: true). The offsets
            // are ignored at this layer — the OOXMLEdit applies bold to the
            // ENTIRE run, not a substring. Partial-Run bold (offsets cover
            // only part of the run's text) requires run-splitting, which is
            // a separate OOXMLEdit case pending design (file ooxml-swift
            // issue if needed).
            if range.startRun == range.endRun {
                return [.setBold(target: range.startRun, value: true)]
            }
            // Multi-Run / cross-paragraph case: lower() can't enumerate the
            // runs between startRun and endRun without document context.
            // Returning [] triggers the silent-noop guard in
            // WordDocument.apply → throws notImplemented. Resolution
            // requires either (a) Edit protocol amendment to pass doc to
            // lower(), or (b) a pre-lower step that resolves the range to
            // a list of (runID) on the document side. Tracked in macdoc#110.
            return []

        case .applyLink(let range, let url):
            // Per macdoc#110 §5 design walkthrough Q1 verdict: WordEdit
            // semantic-layer applyLink matches Word UI Cmd-K (wrap
            // existing run), so it lowers to OOXMLEdit.wrapWithHyperlink
            // (Design Y wrap-existing primitive).
            //
            // Single-Run case (startRun == endRun): wrap that run.
            // wrapWithHyperlink requires whole-Run target (MVP); partial-
            // Run wrap (selection covers part of run's text) needs
            // run-splitting which is out of scope here. lower() can't
            // detect partial-Run from offsets alone — startOffset/endOffset
            // would need to be checked against startRun's text length,
            // which requires doc context (per §7 lower() context constraint).
            // Pre-validation in apply() pipeline catches partial-Run by
            // checking the run's actual text length post-lower (currently
            // PHASED — pre-validation hasn't shipped, so partial-Run will
            // produce a hyperlink wrapping the whole Run regardless of
            // offset).
            if range.startRun == range.endRun {
                return [.wrapWithHyperlink(target: range.startRun, href: url)]
            }
            // Multi-Run case: same constraint as applyBold above —
            // intermediate Runs can't be enumerated without doc context.
            return []

        case .applyInsertParagraph(let after, let content):
            // Simplest case: 1:1 mapping. ParagraphRef wraps an ElementID
            // that's already in the right shape for OOXMLEdit.
            return [.insertParagraph(
                after: after.elementID,
                content: content,
                styleId: nil
            )]
        }
    }
}
