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

    /// Lower this WordEdit to its `[OOXMLEdit]` translation. Stub in §1
    /// scaffold — full implementations land in §7 of #105 tasks.
    public func lower() -> [OOXMLEdit] {
        // §7 will implement per-case logic. Returning empty is intentional
        // scaffold marker — naturality tests in §9 will fail on stubs, which
        // is the correct signal that implementation is pending.
        []
    }
}
