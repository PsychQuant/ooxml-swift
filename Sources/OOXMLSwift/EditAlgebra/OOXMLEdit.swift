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

    /// Insert a hyperlink wrapping the specified target element.
    ///
    /// - `target`: ElementID of the element to wrap in `<w:hyperlink>` (typically a Run)
    /// - `href`: URL the hyperlink points to
    /// - `displayText`: optional override for displayed text; `nil` uses target's existing text
    ///
    /// Composite operation: emits BOTH `Operation.insertNode` (for the
    /// `<w:hyperlink>` element in document.xml) AND `Operation.updateAttribute`
    /// (for the new Relationship entry in `_rels/document.xml.rels`).
    /// Atomic at Edit level — throws `EditError.preserveViolation` if either
    /// sub-operation cannot apply, with NO partial state in the result.
    case insertHyperlink(target: ElementID, href: URL, displayText: String?)

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
