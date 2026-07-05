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

        var newDocument = self
        try newDocument.appendAndMaterialize(newOps)
        return newDocument
    }

    /// Shared op-application core: appends `newOps` to the log with fresh
    /// (shared) opIDs and materializes each op against the part containing
    /// its target. Extracted from `apply(_ edit:)` so the Phase 2 typed
    /// setters (task 3.15, `WordDocument+TypedSetters.swift`) route through
    /// the exact same log + reducer path instead of duplicating it.
    internal mutating func appendAndMaterialize(
        _ newOps: [Operation], source: OpSource = .swift
    ) throws {

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
            newLog.append(op, source: source, opID: opID)
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

        // Single-part fast path: when the doc has exactly one part, skip the
        // partContaining tree walk. materialize will throw elementNotFound
        // if the target isn't in the tree (we wrap that as
        // operationLogFailure same as the multi-part error path). Saves
        // ~3-5µs per op on synthesized fixtures where partContaining's
        // findNode walk was significant overhead.
        //
        // Most real-world docs are multi-part (document.xml + styles.xml +
        // comments.xml + ...), but synthesized fixtures + simple cases hit
        // this fast path.
        let singlePartPath: String? = newTrees.count == 1 ? newTrees.keys.first : nil

        for (op, opID) in zip(newOps, opIDs) {
            let partPath: String

            // Part-addressed ops (addRelationship) carry their target part
            // path in the payload — route directly without partContaining
            // walk. addRelationship needs the rels part tree to exist; if
            // not, we create it on-demand (rels parts often don't exist
            // in synthesized fixtures).
            // §4b (#128): log-only markers / opaque ops append to the log but
            // have no materialization target — skip the per-part apply.
            switch op {
            case .batchBegin, .batchEnd, .beginComponent, .endComponent, .unknown:
                continue
            default:
                break
            }

            if case .addRelationship(let part, _, _, _, _) = op {
                partPath = part
                if newTrees[part] == nil {
                    newTrees[part] = makeEmptyRelationshipsTree()
                }
            } else if case .defineStyle = op {
                // §4b (#128): part-addressed like addRelationship — styles
                // live in word/styles.xml (created on demand for synthesized
                // fixtures). Must precede the single-part fast path so a
                // document-only doc doesn't misroute the style definition.
                partPath = "word/styles.xml"
                if newTrees[partPath] == nil {
                    newTrees[partPath] = makeEmptyStylesTree()
                }
            } else if let single = singlePartPath {
                partPath = single
            } else {
                guard let found = OperationReducer.partContaining(op: op, in: newTrees) else {
                    // No part contains the op's target. Surface as
                    // operationLogFailure (PHASED #4 — upfront pathNotFound
                    // validation lands later).
                    throw EditError.operationLogFailure(
                        underlying: "No XmlTree part contains any ElementID referenced by op: \(op)"
                    )
                }
                partPath = found
            }

            // Build a single-op log carrying the SHARED opID. The Reducer
            // sees entry.opID == opID, so the new node's libraryUUID derives
            // from the same UUID that's persisted in newLog above.
            var singleOpLog = OperationLog()
            singleOpLog.append(op, source: source, opID: opID)

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

        // 5. Commit updated log + trees onto self.
        //    body.children typed view is NOT auto-resynced — calling
        //    resync would create new Paragraph(xmlNode:) instances whose
        //    xmlNode references are different from any other path's
        //    deep-copied tree, breaking the reference-equality Paragraph
        //    Equatable (which downstream comparisons like Naturality tests
        //    depend on). Callers who need a fresh body.children call
        //    `resyncBodyFromDocumentTree()` explicitly.
        //
        //    Per macdoc#110 item #8: the resync mechanism ships here as
        //    opt-in. Future architectural work (content-based Paragraph
        //    Equatable, or always-tree-backed Paragraph that re-reads on
        //    every access) could enable auto-resync — out of scope here.
        self.operationLog = newLog
        self.xmlTrees = newTrees
    }

    /// Constructs an empty `<Relationships>` XmlTree suitable for
    /// addRelationship operations. Used when apply() encounters an
    /// addRelationship op on a doc whose rels part doesn't exist yet
    /// (common in synthesized fixtures).
    ///
    /// The namespace `http://schemas.openxmlformats.org/package/2006/relationships`
    /// is the standard rels-part namespace per ECMA-376.
    /// Constructs an empty `<w:styles>` XmlTree for defineStyle operations
    /// on documents whose styles part doesn't exist yet (§4b, #128).
    internal func makeEmptyStylesTree() -> XmlTree {
        let root = XmlNode.element(
            prefix: "w",
            localName: "styles",
            namespaceURI: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            attributes: [XmlAttribute(
                prefix: "xmlns", localName: "w",
                value: "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
        )
        return XmlTree.synthesized(root: root)
    }

    internal func makeEmptyRelationshipsTree() -> XmlTree {
        let root = XmlNode.element(
            prefix: nil,
            localName: "Relationships",
            namespaceURI: "http://schemas.openxmlformats.org/package/2006/relationships"
        )
        return XmlTree.synthesized(root: root)
    }

    /// Rebuilds `body.children` from the current `xmlTrees["word/document.xml"]`
    /// tree. Walks the `<w:body>` direct children and constructs tree-backed
    /// `Paragraph(xmlNode:)` / `Table(xmlNode:)` values. Non-paragraph /
    /// non-table body elements are dropped from the typed view (see scope
    /// notes below).
    ///
    /// **Opt-in design**: this method is NOT called automatically by
    /// `apply()` because tree-backed Paragraph uses reference equality
    /// (`Paragraph.==` compares `xmlNode === xmlNode`). Auto-resync after
    /// apply would create new XmlNode instances on every apply call,
    /// breaking downstream equality comparisons (notably NaturalityTests
    /// which assert two apply paths produce equal docs). Callers who want
    /// fresh body.children after apply call this method explicitly.
    ///
    /// **Narrow scope** (documented limitations):
    /// - Only `<w:p>` and `<w:tbl>` become typed body children. Other
    ///   body-level elements (`<w:sdt>`, `<w:bookmarkStart>`/End, vendor
    ///   extensions) are NOT re-typed; they remain in xmlTrees but
    ///   disappear from body.children. If your doc has these and you
    ///   rely on body.children to round-trip them, prefer reading from
    ///   xmlTrees directly.
    /// - Only document.xml's body is resynced. styles, headers, footers,
    ///   numbering, footnotes, endnotes remain stale relative to new
    ///   xmlTrees.
    ///
    /// Safe to call multiple times — each call rebuilds from scratch.
    ///
    /// macdoc#110 item #8 tracker. Full auto-resync would require a
    /// downstream refactor of Paragraph Equatable semantics (out of
    /// scope here).
    public mutating func resyncBodyFromDocumentTree() {
        guard let docTree = self.xmlTrees["word/document.xml"] else { return }
        guard let bodyNode = docTree.root.children.first(where: {
            $0.kind == .element && $0.localName == "body"
        }) else { return }

        var newChildren: [BodyChild] = []
        var newTables: [Table] = []

        for child in bodyNode.children where child.kind == .element {
            switch child.localName {
            case "p":
                newChildren.append(.paragraph(Paragraph(xmlNode: child)))
            case "tbl":
                let t = Table(xmlNode: child)
                newChildren.append(.table(t))
                newTables.append(t)
            case "sectPr":
                // Parsed separately into sectionProperties; skip from body
                continue
            default:
                // Other body-level elements (sdt, bookmarkMarker, vendor
                // extensions) are not currently re-typed by apply(). They
                // remain in xmlTrees for byte-equivalent round-trip but
                // disappear from body.children typed view. See scope notes.
                continue
            }
        }

        self.body.children = newChildren
        self.body.tables = newTables
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
