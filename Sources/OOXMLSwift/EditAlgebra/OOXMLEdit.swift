// OOXMLEdit.swift
// EditAlgebra — Phase 2 of ooxml-edit-isomorphism-foundation (PsychQuant/macdoc#99)
//
// OOXMLEdit is the syntactic-layer Edit enum. Each case maps to one or more
// `Operation` enum cases (see OOXMLEdit+Operation.swift for the mapping
// table). Case names follow OOXML schema-element naming (foundation ADR-005).
//
// 5 canonical cases in this Phase 2 ship:
//   - insertParagraph(after:content:styleId:)
//   - insertParagraphBefore(before:content:styleId:)
//   - setBold(target:value:)
//   - insertHyperlink(target:href:displayText:)
//   - removeParagraph(target:)
//
// Apply routes through Operation + OperationLog + OperationReducer (NOT
// the older applyOverlay/markDirty per foundation #105 design.md Decision 3).

import Foundation

// MARK: - OOXMLEdit

/// Syntactic-layer Edit — addresses OOXML elements by `ElementID` and emits
/// `Operation` log entries.
///
/// Each case 1:1 or 1:N maps to `Operation` enum cases per the mapping table
/// in `openspec/changes/ooxml-edit-algebra-implementation/design.md`
/// Decision 1. The mapping is the public contract that makes property tests
/// possible and lets callers switch between OOXMLEdit (Edit-typed) and
/// Operation (log-entry-typed) when convenient.
public enum OOXMLEdit: Edit, Equatable, Sendable {
    /// Insert a new paragraph after the specified anchor paragraph.
    ///
    /// - `after`: ElementID of the existing paragraph to insert AFTER
    /// - `content`: text content of the new paragraph
    /// - `styleId`: optional style ID reference (e.g., "Heading1"); `nil` for default
    ///
    /// Maps to `Operation.insertParagraphAfter(after:paragraph:)` with
    /// `ParagraphPayload(text: content, styleId: styleId)`.
    case insertParagraph(after: ElementID, content: String, styleId: String?)

    /// Insert a new paragraph before the specified anchor paragraph.
    ///
    /// Positional variant of `.insertParagraph(after:)`. Maps to
    /// `Operation.insertParagraphBefore(before:paragraph:)`.
    case insertParagraphBefore(before: ElementID, content: String, styleId: String?)

    /// Toggle bold formatting on the specified Run element.
    ///
    /// - `target`: ElementID of the Run (must resolve to a `<w:r>` element)
    /// - `value`: `true` sets `<w:b/>`, `false` removes it (or sets `<w:b w:val="false"/>`
    ///   per OOXML spec — implementation choice in §4.1)
    ///
    /// Maps to `Operation.setRunFormat(target:format:)` with
    /// `RunFormatPayload(bold: value, ...)`.
    ///
    /// Throws `EditError.unsupportedOperation` if target is not a Run.
    case setBold(target: ElementID, value: Bool)

    /// Insert a NEW hyperlink wrapper after the specified target element.
    ///
    /// **Insert semantics (Design X from macdoc#110 §5 walkthrough)**:
    /// Creates a brand-new `<w:hyperlink>` wrapper containing a new
    /// `<w:r><w:t>displayText</w:t></w:r>` child. The wrapper is inserted
    /// as a sibling after `target` (Run or Paragraph). The target itself
    /// is NOT modified.
    ///
    /// - `target`: ElementID of the element to insert AFTER (Run or Paragraph)
    /// - `href`: URL the hyperlink points to
    /// - `displayText`: optional displayed text; `nil` → `href.absoluteString`
    ///
    /// Composite operation: emits `[Operation.insertNode (hyperlink XML),
    /// Operation.addRelationship (rels entry)]`. Atomicity at Edit level
    /// via pre-validation (target exists + rels-part exists) before
    /// emitting Operations.
    ///
    /// For Cmd-K parity (replace existing run with link), use
    /// `wrapWithHyperlink` instead.
    case insertHyperlink(target: ElementID, href: URL, displayText: String?)

    /// Wrap the existing Run element with a `<w:hyperlink>` so its text
    /// becomes the link's displayed text.
    ///
    /// **Wrap semantics (Design Y from macdoc#110 §5 walkthrough)**:
    /// REPLACES `target` in its parent's children with a new
    /// `<w:hyperlink r:id="...">target</w:hyperlink>` wrapper containing
    /// the original Run unchanged.
    ///
    /// - `target`: ElementID of the Run to wrap (MUST resolve to `<w:r>`)
    /// - `href`: URL the hyperlink points to
    ///
    /// Composite: `[Operation.insertNode (wrapper around target),
    /// Operation.removeNode (target's original position),
    /// Operation.addRelationship (rels entry)]`. (Reducer may collapse
    /// these into a single tree-mutation primitive — implementation
    /// detail.) Atomicity via pre-validation.
    ///
    /// MVP constraint: whole-Run only. Partial-Run wrap (selection covers
    /// part of run's text) requires run-splitting first — currently
    /// unsupported. Throws `EditError.unsupportedOperation` if target
    /// isn't a `<w:r>` element.
    case wrapWithHyperlink(target: ElementID, href: URL)

    /// Remove the specified paragraph from the document body.
    ///
    /// - `target`: ElementID of the paragraph to remove
    ///
    /// Maps to `Operation.removeParagraph(id:)`. Body-children indices shift
    /// after removal; sibling elements remain c14n-equal to their input form
    /// (canonical-identity invariant per foundation Requirement).
    case removeParagraph(target: ElementID)

    // MARK: - Edit protocol conformance

    /// Apply this Edit to the WordDocument, routing through Operation + OperationLog +
    /// OperationReducer. Stub implementation in §1 scaffold — full apply
    /// wiring lands in §2 (Document.apply public API).
    public func apply(to document: WordDocument) throws -> WordDocument {
        throw EditError.notImplemented("OOXMLEdit.apply requires Document.apply(_:) wiring (§2 of #105 tasks)")
    }

    /// Returns `[self]` — OOXMLEdit is already the syntactic layer; lower()
    /// is identity. WordEdit cases override this with semantic→syntactic
    /// translations (see WordEdit.swift).
    public func lower() -> [OOXMLEdit] {
        [self]
    }
}
