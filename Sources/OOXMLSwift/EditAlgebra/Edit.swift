// Edit.swift
// EditAlgebra — Phase 2 of ooxml-edit-isomorphism-foundation (PsychQuant/macdoc#99)
//
// Edit protocol is the public surface for type-level edits over WordDocument.
// Backing runtime is the OpLog mechanism shipped in v0.31.x (Operation +
// OperationLog + OperationReducer), NOT the older applyOverlay/markDirty.
//
// Per foundation design.md ADR-002:
//   "PRs that introduce new Edit cases SHALL attach a commutative diagram
//    + commute proof."
// See docs/edit-algebra-cd-discipline.md and openspec/changes/ooxml-edit-isomorphism-foundation/design.md
// § ADR-002 Worked Examples.

import Foundation

// MARK: - Edit protocol

/// The public surface for type-level edits over a WordDocument.
///
/// Conformers are typically `OOXMLEdit` (syntactic layer — addresses OOXML
/// elements by ElementID) and `WordEdit` (semantic layer — addresses
/// Word-UI verbs like Cmd-B). `WordEdit.lower()` returns a `[OOXMLEdit]`
/// translation; `OOXMLEdit.lower()` is identity.
///
/// `apply(to:)` returns a new WordDocument (immutable apply — input unchanged)
/// per foundation `ooxml-edit-algebra` Requirement "Edit Apply Surface on
/// Document".
public protocol Edit {
    /// Applies the edit to a WordDocument, returning a new WordDocument with
    /// the edit's Operations applied. Throws `EditError` on validation failures
    /// (path resolution, preserve-violation defensive check, operation log
    /// failures).
    ///
    /// The input `document` is NEVER mutated — the returned WordDocument is
    /// a fresh value with the edit applied.
    func apply(to document: WordDocument) throws -> WordDocument

    /// Returns the syntactic-layer translation of this edit. For OOXMLEdit
    /// cases this returns `[self]` (identity). For WordEdit cases this
    /// returns the corresponding `[OOXMLEdit]` decomposition, possibly with
    /// multiple elements when the WordEdit semantics cross structural
    /// boundaries (e.g., applyBold on a range spanning two paragraphs lowers
    /// to two `OOXMLEdit.setBold` calls).
    ///
    /// `lower()` is total: every WordEdit case MUST return a non-empty list.
    /// If a WordEdit case has no defined translation, file an ooxml-swift
    /// issue rather than returning `[]`.
    func lower() -> [OOXMLEdit]
}

// MARK: - EditError

/// Errors thrown by `Edit.apply(to:)` and downstream Operation execution.
///
/// `pathNotFound` — Edit's target ElementID does not resolve in the input
/// WordDocument. Caller should re-check the ElementID source (typed view
/// lookup, prior WordRange resolution, etc.).
///
/// `preserveViolation` — Edit's Operations resulted in a c14n-form change
/// to a part OUTSIDE the Edit's declared target. This is a defensive check
/// against buggy Edit implementations that accidentally modify unmodified
/// subtrees (per foundation `ooxml-edit-algebra` Requirement
/// "Canonical-Identity Round-Trip Contract"). The `expected` and `actual`
/// fields carry c14n digests for debugging.
///
/// `unsupportedOperation` — Edit's target element type doesn't support the
/// operation (e.g., setBold on a non-Run element). The String carries a
/// human-readable description for diagnostics.
///
/// `notImplemented` — Stub case for in-progress Edit cases (used during
/// scaffold development; production code should never throw this).
///
/// `operationLogFailure` — Internal OperationReducer.materialize threw an
/// error. The underlying string is the localized description of the
/// reducer error for debugging.
public enum EditError: Error, Equatable, Sendable {
    case pathNotFound(ElementID)
    case preserveViolation(part: String, expected: String, actual: String)
    case unsupportedOperation(String)
    case notImplemented(String)
    case operationLogFailure(underlying: String)
}
