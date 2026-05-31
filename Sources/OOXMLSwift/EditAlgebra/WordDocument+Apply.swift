// WordDocument+Apply.swift
// EditAlgebra — Phase 2 of ooxml-edit-isomorphism-foundation (PsychQuant/macdoc#99)
//
// Public apply API per design.md Decision 3 (Option A: WordDocument owns log).
// Routes Edit → OOXMLEdit → Operation → OperationLog → OperationReducer.materialize.
//
// CURRENT LIMITATION: typed views (body/styles/etc.) are NOT re-synced
// from new xmlTrees after apply. xmlTrees and operationLog ARE correctly
// updated. End-to-end tests inspect xmlTrees directly for canonical-
// identity assertions. Typed-view re-sync tracked as item #8 of macdoc#110
// (separate from the multi-part scoping fix that landed in PR #74).

import Foundation

extension WordDocument {

    /// Apply an Edit to this document, returning a new WordDocument with the
    /// edit's emitted Operations appended to the log + affected XmlTrees
    /// re-materialized.
    ///
    /// Per foundation `ooxml-edit-algebra` Requirement "Edit Apply Surface on
    /// Document" + this change's design.md Decision 3:
    /// - Immutable apply (input `self` never mutated)
    /// - Routes through Operation + OperationLog + OperationReducer (NOT
    ///   applyOverlay/markDirty)
    /// - Throws `EditError.pathNotFound` when target ElementID doesn't resolve
    /// - Throws `EditError.preserveViolation` when defensive check fires
    /// - Wraps OperationReducer errors as `EditError.operationLogFailure`
    ///
    /// **Known limitation**: typed views (body/styles/headers/etc.) are NOT
    /// re-synced from new xmlTrees after apply. xmlTrees + operationLog ARE
    /// correct. For end-to-end tests, inspect xmlTrees directly via
    /// `result.xmlTrees["word/document.xml"]`. Typed-view re-sync tracked
    /// as item #8 of macdoc#110 (NOT the multi-part scoping fix in PR #74
    /// which already shipped).
    public func apply(_ edit: any Edit) throws -> WordDocument {
        // 1. Lower edit → OOXMLEdit chain → Operations
        //    WordEdit.lower() returns [OOXMLEdit]; OOXMLEdit.lower() returns [self].
        //    Each OOXMLEdit emits 1+ Operations via the mapping table in
        //    OOXMLEdit+Operation.swift (per design.md Decision 1).
        let ooxmlEdits = edit.lower()

        // Defensive: detect Edits that silently lower to []. OOXMLEdit's
        // lower() always returns [self] (identity), so empty here means a
        // non-OOXMLEdit (typically WordEdit) returned []. This happens in two
        // scenarios: (1) unimplemented stub case, (2) input combination that
        // lower() can't resolve without document context (e.g., cross-
        // paragraph WordRange in applyBold — see WordEdit.swift). Both
        // surface as notImplemented since the apply call can't proceed.
        if ooxmlEdits.isEmpty && !(edit is OOXMLEdit) {
            throw EditError.notImplemented(
                "Edit of type \(type(of: edit)) returned empty lower(). Either the case is not yet implemented (see macdoc#110 / macdoc#105 §7), or the input combination requires document context that the non-throwing no-arg lower() protocol can't access (e.g., cross-paragraph WordRange)."
            )
        }

        var newOps: [Operation] = []
        for ooxmlEdit in ooxmlEdits {
            // OOXMLEdit.operations() may throw EditError.notImplemented for
            // stub cases (§1 scaffold) or EditError.unsupportedOperation for
            // type-mismatch (e.g., setBold on non-Run target).
            let ops = try ooxmlEdit.operations()
            newOps.append(contentsOf: ops)
        }

        // 2. Generate stable opIDs ONCE — shared between persisted log and
        //    per-op materialization log. Critical for replay determinism:
        //    the Reducer derives new-node libraryUUIDs from entry.opID (per
        //    Phase 2c convention), so if newLog and the materialize log used
        //    DIFFERENT opIDs, re-materializing the persisted log would
        //    produce different IDs than the freshly-applied tree.
        let opIDs: [UUID] = newOps.map { _ in UUID() }

        // 3. Build accumulated log = old log + new ops (with shared opIDs).
        //    OperationLog enforces append-only semantics; we copy + extend.
        var newLog = self.operationLog
        for (op, opID) in zip(newOps, opIDs) {
            newLog.append(op, source: .swift, opID: opID)
        }

        // 4. Materialize ops per-part: each op is replayed only against the
        //    part its target lives in.
        //
        //    Per-op rather than per-part-batched because subsequent ops may
        //    reference nodes created by earlier ops (Phase 2c determinism:
        //    new node's libraryUUID == entry.opID). The chain works because
        //    newTrees is mutated in place after each op, so the next op's
        //    partContaining lookup sees the in-flight state.
        //
        //    macdoc#110 fix: replaces the §2 scaffold's "apply tempLog to
        //    every tree" pattern which threw elementNotFound on parts that
        //    didn't contain the op's target.
        var newTrees = self.xmlTrees
        for (op, opID) in zip(newOps, opIDs) {
            guard let partPath = OperationReducer.partContaining(op: op, in: newTrees) else {
                // No part contains the op's target. Surface as
                // operationLogFailure (PHASED #4 — upfront pathNotFound
                // validation lands later).
                throw EditError.operationLogFailure(
                    underlying: "No XmlTree part contains any ElementID referenced by op: \(op)"
                )
            }

            // Build a single-op log carrying the SHARED opID. The Reducer
            // sees entry.opID == opID, so the new node's libraryUUID derives
            // from the same UUID that's persisted in newLog above.
            var singleOpLog = OperationLog()
            singleOpLog.append(op, source: .swift, opID: opID)

            do {
                let materialized = try OperationReducer.materialize(
                    log: singleOpLog,
                    base: newTrees[partPath]!
                )
                newTrees[partPath] = materialized
            } catch {
                throw EditError.operationLogFailure(
                    underlying: "OperationReducer.materialize failed on part '\(partPath)': \(error.localizedDescription)"
                )
            }
        }

        // 5. Construct new WordDocument with updated log + trees.
        //    Typed views (body/styles/etc.) carried over from self — STALE
        //    relative to new xmlTrees. Resync tracked as item #8 of
        //    macdoc#110 (SEPARATE from the multi-part scoping fix shipped
        //    here — don't conflate them when chasing #110).
        var newDocument = self
        newDocument.operationLog = newLog
        newDocument.xmlTrees = newTrees
        return newDocument
    }

    /// Apply a sequence of Edits in order, folding each result into the next
    /// apply. Equivalent to chaining individual `apply` calls.
    ///
    /// Per spec.md Requirement "Document.apply Public Method" — sequence
    /// variant for callers iterating over an edit script.
    public func apply<S: Sequence>(_ edits: S) throws -> WordDocument where S.Element == any Edit {
        var current = self
        for edit in edits {
            current = try current.apply(edit)
        }
        return current
    }
}
