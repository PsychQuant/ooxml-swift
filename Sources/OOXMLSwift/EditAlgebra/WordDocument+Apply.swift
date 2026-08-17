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
        _ newOps: [Operation], source: OpSource = .swift,
        replayOpIDs: [UUID]? = nil,
        replayTimestamps: [Date]? = nil,
        replaySources: [OpSource]? = nil,
        appendToLog: Bool = true
    ) throws {

        // Legacy typed mutators update the typed model and deliberately mark
        // their corresponding XmlTree stale. Before starting a new operation
        // suffix, bridge those dirty typed parts back into isolated, lossless
        // trees; otherwise the checkpoint below would capture an older tree
        // and the first reducer op would silently erase the typed mutation.
        let staleTypedParts = Set(
            modifiedParts.subtracting(treeFreshParts).filter { part in
                xmlTrees[part] != nil
                    || part == "[Content_Types].xml"
                    || part.hasSuffix(".xml")
                    || part.hasSuffix(".rels")
            })
        if !newOps.isEmpty, !staleTypedParts.isEmpty {
            let refreshed = try DocxWriter.materializeTypedTrees(
                self, parts: staleTypedParts)
            for (part, tree) in refreshed {
                xmlTrees[part] = tree
                treeFreshParts.insert(part)
            }
        }

        // 2. Generate stable opIDs ONCE — shared between persisted log and
        //    per-op materialization log. Critical for replay determinism:
        //    the Reducer derives new-node libraryUUIDs from entry.opID (per
        //    Phase 2c convention), so if newLog and the materialize log used
        //    DIFFERENT opIDs, re-materializing the persisted log would
        //    produce different IDs than the freshly-applied tree.
        let opIDs: [UUID] = replayOpIDs ?? newOps.map { _ in UUID() }
        let timestamps: [Date] = replayTimestamps ?? newOps.map { _ in Date() }
        let sources: [OpSource] = replaySources
            ?? Array(repeating: source, count: newOps.count)
        guard opIDs.count == newOps.count,
              timestamps.count == newOps.count,
              sources.count == newOps.count else {
            throw EditError.operationLogFailure(
                underlying: "replay metadata count does not match operation count")
        }

        if appendToLog, !newOps.isEmpty, operationReplayBase == nil {
            operationReplayBase = OperationReplayBase(
                // Reducer materialization is copy-on-write at the XmlTree
                // level, so retaining the current tree references is a safe,
                // constant-time checkpoint. Legacy direct typed mutations
                // invalidate this checkpoint through `markTypedDirty`.
                trees: xmlTrees,
                comments: comments,
                carriedParts: carriedParts,
                modifiedParts: modifiedParts,
                treeFreshParts: treeFreshParts,
                logStartIndex: operationLog.entries.count)
        }

        // 3. Build accumulated log = old log + new ops (with shared opIDs).
        //    OperationLog enforces append-only semantics; we copy + extend.
        var newLog = self.operationLog
        if appendToLog {
            for index in newOps.indices {
                newLog.append(newOps[index], source: sources[index],
                              opID: opIDs[index], at: timestamps[index])
            }
        }


        // A control changes the active history set; patching only the current
        // value cannot recover a value that came from the pre-log document,
        // and cannot reverse structural batches. Rebuild the locally-owned
        // history suffix from its captured base instead. This also makes
        // controls atomic: validation/replay finishes before `self` changes.
        if appendToLog, newOps.contains(where: {
            if case .undo = $0 { return true }
            if case .redo = $0 { return true }
            return false
        }) {
            guard let replayBase = operationReplayBase else {
                throw EditError.operationLogFailure(
                    underlying: "control operation has no replay base")
            }
            var localLog = OperationLog()
            for entry in newLog.entries.dropFirst(replayBase.logStartIndex) {
                localLog.append(entry.op, source: entry.source,
                                opID: entry.opID, at: entry.timestamp)
            }
            let active = try OperationReducer.activeEntryIndices(in: localLog)
            let entries = localLog.entries.enumerated().compactMap { index, entry in
                active.contains(index) ? entry : nil
            }

            var rebuilt = self
            rebuilt.xmlTrees = replayBase.trees.mapValues { $0.deepCopy() }
            rebuilt.comments = replayBase.comments
            rebuilt.carriedParts = replayBase.carriedParts
            rebuilt.modifiedParts = replayBase.modifiedParts
            rebuilt.treeFreshParts = replayBase.treeFreshParts
            rebuilt.operationLog = newLog
            rebuilt.operationReplayBase = replayBase
            try rebuilt.appendAndMaterialize(
                entries.map(\.op),
                replayOpIDs: entries.map(\.opID),
                replayTimestamps: entries.map(\.timestamp),
                replaySources: entries.map(\.source),
                appendToLog: false)
            rebuilt.operationLog = newLog
            rebuilt.operationReplayBase = replayBase
            self = rebuilt
            return
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
        var newComments = self.comments
        var newCarriedParts = self.carriedParts
        var touchedParts: Set<String> = []
        var freshParts: Set<String> = []

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

        let previousLogCount = self.operationLog.entries.count
        for index in newOps.indices {
            let op = newOps[index]
            let opID = opIDs[index]
            let timestamp = timestamps[index]
            let entrySource = sources[index]
            let partPath: String

            // A control operation changes which prior entries participate in
            // replay. The document already contains the prior history, so
            // compute the resulting text/style values from the accumulated
            // prefix and apply only those corrections to the current trees.
            // This preserves operation identity and avoids the incorrect
            // single-op reducer path where undo was a silent no-op.
            switch op {
            case .undo, .redo:
                var prefix = OperationLog()
                let prefixCount = appendToLog
                    ? previousLogCount + index + 1
                    : newLog.entries.count
                for entry in newLog.entries.prefix(prefixCount) {
                    prefix.append(entry.op, source: entry.source,
                                  opID: entry.opID, at: entry.timestamp)
                }
                do {
                    for replacement in try OperationReducer.controlReplacementOperations(
                        for: op, in: prefix) {
                        let replacementPart: String
                        if let single = singlePartPath {
                            replacementPart = single
                        } else if let found = OperationReducer.partContaining(
                            op: replacement, in: newTrees) {
                            replacementPart = found
                        } else {
                            throw ReducerError.elementNotFound(
                                opID: opID,
                                elementID: OperationReducer.referencedElementIDs(
                                    in: replacement).first ?? ElementID(rawString: "lib:\(opID.uuidString)"))
                        }
                        guard var replacementTree = newTrees[replacementPart] else {
                            throw ReducerError.elementNotFound(
                                opID: opID,
                                elementID: OperationReducer.referencedElementIDs(
                                    in: replacement).first ?? ElementID(rawString: "lib:\(opID.uuidString)"))
                        }
                        try OperationReducer.apply(
                            entry: LogEntry(
                                opID: opID, op: replacement,
                                source: entrySource, timestamp: timestamp),
                            to: &replacementTree)
                        newTrees[replacementPart] = replacementTree
                        touchedParts.insert(replacementPart)
                        freshParts.insert(replacementPart)
                    }
                } catch {
                    throw EditError.operationLogFailure(
                        underlying: "control operation replay failed: \(error)")
                }
                continue
            default:
                break
            }

            // Part-addressed ops (addRelationship) carry their target part
            // path in the payload — route directly without partContaining
            // walk. addRelationship needs the rels part tree to exist; if
            // not, we create it on-demand (rels parts often don't exist
            // in synthesized fixtures).
            // §4b (#128): log-only markers / opaque ops append to the log but
            // have no materialization target — skip the per-part apply.
            switch op {
            case .carryPart(let carriedPath, let xml):
                // Raw part channel (format-alignment-engine Phase A): store the
                // verbatim bytes on `carriedParts`. carryPart materializes to a
                // package part, not to a document tree, so skip the per-part
                // apply below. `writeAuthoringPackage` emits it byte-exact.
                guard isSafeRelativeOOXMLPath(carriedPath) else {
                    throw EditError.operationLogFailure(
                        underlying: "unsafe carryPart path: \(carriedPath)")
                }
                newCarriedParts[carriedPath] = Data(xml.utf8)
                continue
            case .carryBinaryPart(let carriedPath, let base64):
                guard isSafeRelativeOOXMLPath(carriedPath) else {
                    throw EditError.operationLogFailure(
                        underlying: "unsafe carryBinaryPart path: \(carriedPath)")
                }
                guard let bytes = Data(base64Encoded: base64) else {
                    throw EditError.operationLogFailure(
                        underlying: "invalid base64 for binary part: \(carriedPath)")
                }
                newCarriedParts[carriedPath] = bytes
                continue
            case .batchBegin, .batchEnd, .beginComponent, .endComponent, .unknown:
                continue
            default:
                break
            }

            if case .addRelationship(let part, _, _, _, _) = op {
                guard isSafeRelativeOOXMLPath(part) else {
                    throw EditError.operationLogFailure(
                        underlying: "unsafe addRelationship part path: \(part)")
                }
                partPath = part
                if newTrees[part] == nil {
                    newTrees[part] = makeEmptyRelationshipsTree()
                }
            } else if case .appendParagraph(let container, _) = op, container == nil {
                // §4b appendParagraph with nil container targets the main
                // body — route directly to word/document.xml. Multi-part
                // docs would otherwise fail the partContaining walk (the op
                // references no existing ElementID).
                partPath = "word/document.xml"
            } else if case .defineStyle = op {
                // §4b (#128): part-addressed like addRelationship — styles
                // live in word/styles.xml (created on demand for synthesized
                // fixtures). Must precede the single-part fast path so a
                // document-only doc doesn't misroute the style definition.
                partPath = "word/styles.xml"
                if newTrees[partPath] == nil {
                    newTrees[partPath] = makeEmptyStylesTree()
                    // A newly-created styles part also needs package metadata.
                    // These are deliberately dirty but not tree-fresh: the
                    // typed overlay writers merge the new style relationship
                    // and content type into the existing package metadata.
                    touchedParts.formUnion([
                        "word/_rels/document.xml.rels", "[Content_Types].xml",
                    ])
                }
            } else if case .appendTable(let container, _) = op, container == nil {
                // format-alignment-engine Phase B (2.5): same routing as
                // appendParagraph(in: nil) — targets the main body.
                partPath = "word/document.xml"
            } else if case .appendBlockMarker = op {
                partPath = "word/document.xml"
            } else if case .setSectionProperties = op {
                // format-alignment-engine Phase B: sectPr always lives in
                // word/document.xml. The at:nil form references no ElementID,
                // so the partContaining walk would fail on multi-part docs.
                partPath = "word/document.xml"
            } else if case .setDocumentRoot = op {
                // word-canonical-forms task 2.1: the document root lives in
                // word/document.xml; references no ElementID.
                partPath = "word/document.xml"
            } else if case .setDocumentProlog = op {
                // word-canonical-forms task 3.1: prolog lives in word/document.xml.
                partPath = "word/document.xml"
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
            singleOpLog.append(op, source: entrySource, opID: opID, at: timestamp)

            do {
                let materialized = try OperationReducer.materialize(
                    log: singleOpLog,
                    base: newTrees[partPath]!
                )
                newTrees[partPath] = materialized
                touchedParts.insert(partPath)
                freshParts.insert(partPath)
                if case .insertComment(let anchor, let commentID, let text, let author) = op {
                    try materializeCommentDefinition(
                        id: commentID,
                        text: text,
                        author: author,
                        opID: opID,
                        timestamp: timestamp,
                        trees: &newTrees
                    )
                    let paragraphIndex = paragraphOrdinal(
                        for: anchor, in: materialized) ?? -1
                    var comment = Comment(
                        id: commentID,
                        author: author,
                        text: text,
                        paragraphIndex: paragraphIndex,
                        date: timestamp
                    )
                    comment.paraId = deterministicCommentParaID(opID)
                    newComments.comments.removeAll { $0.id == commentID }
                    newComments.comments.append(comment)
                    touchedParts.formUnion([
                        "word/comments.xml", "word/_rels/document.xml.rels",
                        "[Content_Types].xml",
                    ])
                    freshParts.formUnion([
                        "word/comments.xml", "word/_rels/document.xml.rels",
                    ])
                }
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
        self.comments = newComments
        self.carriedParts = newCarriedParts
        // Reducer-applied parts: the tree is now the authoritative content —
        // write serializes it directly (tree-first), and the part must be
        // re-emitted (dirty) in overlay mode.
        // Order matters: marking dirty clears freshness (stale-shadow
        // guard), so freshness must be granted AFTER the dirty mark.
        self.modifiedParts.formUnion(touchedParts)
        // Some operations update package metadata through the typed overlay
        // writers rather than producing a complete replacement tree. If a
        // typed→op bridge refreshed that metadata immediately beforehand,
        // revoke its freshness now so the later typed metadata change wins.
        self.treeFreshParts.subtract(touchedParts.subtracting(freshParts))
        self.treeFreshParts.formUnion(freshParts)
        for part in touchedParts { self.carriedParts.removeValue(forKey: part) }
    }

    private func materializeCommentDefinition(
        id: Int,
        text: String,
        author: String,
        opID: UUID,
        timestamp: Date,
        trees: inout [String: XmlTree]
    ) throws {
        let commentsPath = "word/comments.xml"
        let commentsTree = trees[commentsPath]?.deepCopy() ?? makeEmptyCommentsTree()
        guard commentsTree.root.localName == "comments" else {
            throw EditError.operationLogFailure(
                underlying: "word/comments.xml root must be <w:comments>")
        }
        let duplicate = commentsTree.root.children.contains { child in
            child.kind == .element && child.localName == "comment"
                && child.attributeValue(prefix: "w", localName: "id") == String(id)
        }
        guard !duplicate else {
            throw EditError.operationLogFailure(underlying: "duplicate comment id: \(id)")
        }

        let comment = XmlNode.element(prefix: "w", localName: "comment")
        comment.setAttribute(prefix: "w", localName: "id", value: String(id))
        comment.setAttribute(prefix: "w", localName: "author", value: author)
        comment.setAttribute(
            prefix: "w", localName: "date",
            value: ISO8601DateFormatter().string(from: timestamp))
        comment.setAttribute(
            prefix: "w", localName: "initials",
            value: String(author.prefix(2).uppercased()))

        let paragraph = OperationReducer.makeParagraph(
            payload: ParagraphPayload(text: text))
        paragraph.setAttribute(
            prefix: "w14", localName: "paraId", value: deterministicCommentParaID(opID))
        comment.children = [paragraph]
        commentsTree.root.children.append(comment)
        trees[commentsPath] = commentsTree

        let relsPath = "word/_rels/document.xml.rels"
        let rels = trees[relsPath]?.deepCopy() ?? makeEmptyRelationshipsTree()
        let commentsRelationshipType =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
        if !rels.root.children.contains(where: {
            $0.kind == .element
                && $0.attributeValue(prefix: nil, localName: "Type") == commentsRelationshipType
        }) {
            let existingIDs = Set(rels.root.children.compactMap {
                $0.attributeValue(prefix: nil, localName: "Id")
            })
            var relationshipID = "rIdComments"
            var suffix = 2
            while existingIDs.contains(relationshipID) {
                relationshipID = "rIdComments\(suffix)"
                suffix += 1
            }
            let relationship = XmlNode.element(localName: "Relationship")
            relationship.setAttribute(prefix: nil, localName: "Id", value: relationshipID)
            relationship.setAttribute(
                prefix: nil, localName: "Type", value: commentsRelationshipType)
            relationship.setAttribute(prefix: nil, localName: "Target", value: "comments.xml")
            rels.root.children.append(relationship)
        }
        trees[relsPath] = rels
    }

    private func makeEmptyCommentsTree() -> XmlTree {
        let root = XmlNode.element(
            prefix: "w",
            localName: "comments",
            namespaceURI: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            attributes: [
                XmlAttribute(
                    prefix: "xmlns", localName: "w",
                    value: "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
                XmlAttribute(
                    prefix: "xmlns", localName: "w14",
                    value: "http://schemas.microsoft.com/office/word/2010/wordml"),
            ]
        )
        return .synthesized(root: root)
    }

    private func paragraphOrdinal(for id: ElementID, in tree: XmlTree) -> Int? {
        guard let body = tree.root.children.first(where: {
            $0.kind == .element && $0.localName == "body"
        }) else { return nil }
        var ordinal = 0
        for child in body.children where child.kind == .element && child.localName == "p" {
            if ElementID(node: child) == id { return ordinal }
            ordinal += 1
        }
        return nil
    }

    private func deterministicCommentParaID(_ opID: UUID) -> String {
        String(opID.uuidString.replacingOccurrences(of: "-", with: "").prefix(8)).uppercased()
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
            namespaceURI: "http://schemas.openxmlformats.org/package/2006/relationships",
            attributes: [XmlAttribute(
                prefix: nil, localName: "xmlns",
                value: "http://schemas.openxmlformats.org/package/2006/relationships")]
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
