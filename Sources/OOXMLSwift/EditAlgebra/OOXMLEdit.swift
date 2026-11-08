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
    /// - `inheritAnchorFormat`: `true` gives the new paragraph the anchor's
    ///   format (PsychQuant/ooxml-swift#234). Default `false` inserts a bare
    ///   paragraph as before.
    ///
    /// Maps to `Operation.insertParagraphAfter(after:paragraph:)` with
    /// `ParagraphPayload(text: content, styleId: styleId)`, plus
    /// `inheritAnchorFormat: true` when the option is set.
    ///
    /// **Anchor format** (the same rules as macdoc
    /// `word form apply --insert-after`, PsychQuant/macdoc#249). The anchor
    /// must be a `<w:p>`. The reducer copies, at apply time:
    /// - the anchor's own `<w:pPr>`, without `w:numPr`, `w:sectPr`,
    ///   `w:pPrChange` and `w:pageBreakBefore`; its paragraph-mark `<w:rPr>`
    ///   loses `w:ins`, `w:del`, `w:moveFrom`, `w:moveTo` and `w:rPrChange`.
    ///   A copy left with no children is omitted.
    /// - the `<w:rPr>` of the anchor's first own run with text — the first
    ///   `<w:r>` whose nearest enclosing `<w:p>` is the anchor (runs of a
    ///   text box's paragraphs do not count; runs inside a hyperlink do) and
    ///   whose direct `<w:t>` children hold non-empty text — without
    ///   `w:rPrChange`.
    /// When the anchor ends a section (its pPr holds `w:sectPr`), the
    /// `w:sectPr` moves from the anchor to the new paragraph, so the new
    /// text stays in the anchor's section. A non-nil `styleId` replaces the
    /// copied `w:pStyle`. Undoing the insert puts the `w:sectPr` back on the
    /// anchor, because undo replays the log without the undone entries; as
    /// for every structural op, the insert is undoable only inside a
    /// `batchBegin`…`batchEnd` envelope (the reducer undoes `setText` and
    /// `setParagraphStyle` on their own, nothing else). Element names are
    /// matched in the WordprocessingML namespace, so inline-equation runs
    /// (`m:r`) are not candidates for the first run with text.
    case insertParagraph(after: ElementID, content: String, styleId: String?,
                         inheritAnchorFormat: Bool = false)

    /// Insert a new paragraph before the specified anchor paragraph.
    ///
    /// Positional variant of `.insertParagraph(after:)`. Maps to
    /// `Operation.insertParagraphBefore(before:paragraph:)`.
    ///
    /// `inheritAnchorFormat` follows the rules on `.insertParagraph(after:)`
    /// except for the section break: a paragraph inserted before the anchor
    /// is already in the anchor's section, so `w:sectPr` is neither copied
    /// nor moved.
    case insertParagraphBefore(before: ElementID, content: String, styleId: String?,
                               inheritAnchorFormat: Bool = false)

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

    /// Toggle bold formatting on a text range within a paragraph, located by
    /// substring.
    ///
    /// - `target`: ElementID of the paragraph containing the text (must
    ///   resolve to a `<w:p>` element).
    /// - `substring`: exact text to locate within the paragraph's flattened
    ///   text (concatenation, in document order, of every run's `<w:t>`
    ///   text).
    /// - `instance`: 1-based occurrence to target when `substring` appears
    ///   more than once in the paragraph (default `1`, matching
    ///   `InsertLocation.afterText`'s `instance` convention).
    /// - `value`: `true` writes the canonical naked `<w:b/>`; `false` writes
    ///   an explicit `<w:b w:val="0"/>` off value (PsychQuant/ooxml-swift#161).
    ///
    /// Maps to `Operation.setRunFormatInRange(target:substring:instance:
    /// format:)` with `RunFormatPayload(bold: value)`. Unlike `.setBold`,
    /// `target` here is a PARAGRAPH, not a Run — the reducer resolves
    /// `substring` to a text range that may span or fall inside one or more
    /// runs, splitting run(s) as needed so only the matched text is
    /// reformatted.
    ///
    /// PsychQuant/macdoc#232: `docx apply`'s `set_bold` manifest step has a
    /// paragraph anchor + a `substring` field (declared but previously
    /// unused) — never a Run ElementID — which made `.setBold` always throw
    /// `malformedOp("setRunFormat target must be <w:r>")`. This case is the
    /// paragraph+substring-addressed entry point manifest-driven callers
    /// (`EditPlanner`/`docx-workflow-swift`) should use for `set_bold`
    /// instead.
    ///
    /// Throws (at reducer/apply time — `operations()` itself never inspects
    /// document content) `ReducerError.malformedOp` when `substring` is
    /// empty, not found, `instance < 1`, or `target` does not resolve to a
    /// `<w:p>`.
    case setBoldInRange(target: ElementID, substring: String, value: Bool, instance: Int = 1)

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
