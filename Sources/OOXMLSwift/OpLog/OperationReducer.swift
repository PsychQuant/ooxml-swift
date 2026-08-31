// OperationReducer — pure-replay reducer for the operation log.
//
// Spectra change: operation-reducer-impl, target ooxml-swift v0.31.4.
// Capability: ooxml-operation-reducer
//
// Phase 2b of word-aligned-state-sync. Consumes the OperationLog data
// structures shipped in v0.31.3 and materializes XmlTree state by replaying
// log entries in source-array order on a deep-clone of `base`.
//
// Pure-function semantics: same (log, base) input always produces the same
// output tree. Caller's `base` is NEVER mutated — every entry point starts
// by calling `base.deepCopy()` (added v0.31.4 in `Tree/XmlTree.swift`) and
// applies mutations to the clone.
//
// Phase 2b established replay/time-travel; Phase 2c completes the declared
// element and tree mutation taxonomy. Unsupported or malformed payloads fail
// through typed ReducerError values rather than being silently ignored.

import Foundation

// MARK: - ReplayPoint

/// Time-travel point passed to `OperationReducer.state(log:base:at:)`.
public enum ReplayPoint: Equatable, Sendable {
    /// Replay every entry in `log.entries` (equivalent to `materialize`).
    case latest
    /// Replay `log.entries[0..<N]` (the first N entries).
    case index(Int)
    /// Replay every entry whose `LogEntry.timestamp <= cutoff`, in source-array order.
    case timestamp(Date)
}

// MARK: - ReducerError

/// Typed error surface for `OperationReducer`. The reducer never swallows
/// failures — every error condition surfaces as one of these cases.
public enum ReducerError: Error, Equatable {
    /// An op references an `ElementID` not present in the tree at replay time.
    case elementNotFound(opID: UUID, elementID: ElementID)
    /// An op's payload is structurally invalid (e.g., out-of-range index,
    /// or a Phase 2b unsupported op kind).
    case malformedOp(opID: UUID, reason: String)
    /// `redo` invoked but no `.undo` entry references the target opID.
    case cannotRedo(targetOpID: UUID)
    /// `undo` invoked but the target op cannot be inverted (unsupported op
    /// kind, opaque `.unknown`, or no matching entry).
    case cannotUndo(targetOpID: UUID)
}

// MARK: - OperationReducer

/// Pure-replay reducer namespace. Holds NO state — every entry point is a
/// pure function of its arguments. See doc comments on each function for
/// semantics.
public enum OperationReducer {

    /// Replays `log` on a deep-clone of `base` and returns the materialized tree.
    /// Caller's `base` is NEVER mutated.
    public static func materialize(log: OperationLog, base: XmlTree) throws -> XmlTree {
        var working = base.deepCopy()
        let active = try activeEntryIndices(in: log)
        for (index, entry) in log.entries.enumerated() where active.contains(index) {
            try apply(entry: entry, to: &working)
        }
        return working
    }

    /// Materializes the tree state at a specific replay point.
    public static func state(log: OperationLog, base: XmlTree, at point: ReplayPoint) throws -> XmlTree {
        switch point {
        case .latest:
            return try materialize(log: log, base: base)
        case .index(let N):
            guard N >= 0, N <= log.entries.count else {
                let opID = log.entries.first?.opID ?? UUID()
                throw ReducerError.malformedOp(opID: opID, reason: "index out of range")
            }
            var prefix = OperationLog()
            for entry in log.entries.prefix(N) {
                prefix.append(entry.op, source: entry.source,
                              opID: entry.opID, at: entry.timestamp)
            }
            return try materialize(log: prefix, base: base)
        case .timestamp(let cutoff):
            // Build the actual replay history first. Undo/redo interpretation
            // must only see entries admitted by the cutoff; timestamps are
            // caller-supplied and need not be monotonic in source order.
            var filtered = OperationLog()
            for (sourceIndex, entry) in log.entries.enumerated()
            where entry.timestamp <= cutoff {
                let priorEntries = log.entries.prefix(sourceIndex)
                switch entry.op {
                case .undo(let target):
                    guard priorEntries.contains(where: { $0.opID == target }) else {
                        throw ReducerError.cannotUndo(targetOpID: target)
                    }
                    guard filtered.entries.contains(where: { $0.opID == target }) else {
                        // The target exists in the full log but was excluded
                        // by the cutoff, so this control is outside the
                        // admitted replay relationship as well.
                        continue
                    }
                case .redo(let target):
                    guard priorEntries.contains(where: { $0.opID == target }) else {
                        throw ReducerError.cannotRedo(targetOpID: target)
                    }
                    guard priorEntries.contains(where: {
                        if case .undo(let undoneTarget) = $0.op {
                            return undoneTarget == target
                        }
                        return false
                    }) else {
                        throw ReducerError.cannotRedo(targetOpID: target)
                    }
                    guard filtered.entries.contains(where: { $0.opID == target }) else {
                        continue
                    }
                    guard filtered.entries.contains(where: {
                        if case .undo(let undoneTarget) = $0.op {
                            return undoneTarget == target
                        }
                        return false
                    }) else {
                        // The redo is well-formed in source order, but its
                        // prerequisite undo was excluded by this cutoff.
                        continue
                    }
                default:
                    break
                }
                filtered.append(
                    entry.op, source: entry.source,
                    opID: entry.opID, at: entry.timestamp)
            }
            return try materialize(log: filtered, base: base)
        }
    }

    /// Materializes the log as if the entry with `opID == targetOpID` had
    /// never been applied. Intervening entries see the world without the
    /// target's effect (per Decision 4: replay the whole history with the
    /// targeted op replaced).
    public static func undo(_ targetOpID: UUID, log: OperationLog, base: XmlTree) throws -> XmlTree {
        _ = try controlledEntryIndices(
            targetOpID: targetOpID, before: log.entries.count, in: log)
        var augmented = log
        augmented.append(.undo(targetOpID: targetOpID), source: .swift)
        return try materialize(log: augmented, base: base)
    }

    /// Materializes the log skipping the `.undo` entry whose `targetOpID`
    /// matches the argument — restores the target's effect.
    public static func redo(_ targetOpID: UUID, log: OperationLog, base: XmlTree) throws -> XmlTree {
        var augmented = log
        augmented.append(.redo(targetOpID: targetOpID), source: .swift)
        return try materialize(log: augmented, base: base)
    }

    /// Resolves undo/redo before replay. A control entry changes whether its
    /// target participates in replay; it is not an in-place mutation at the
    /// point where the control appears. This distinction matters whenever a
    /// later operation touches the same element: `A, B, C, undo(B)` must
    /// replay as `A, C`, not apply B's inverse after C.
    internal static func activeEntryIndices(in log: OperationLog) throws -> Set<Int> {
        var targetIsActive: [UUID: Bool] = [:]
        var sawUndo: Set<UUID> = []

        for (index, entry) in log.entries.enumerated() {
            switch entry.op {
            case .undo(let target):
                guard log.entries[..<index].contains(where: {
                    $0.opID == target
                }) else {
                    throw ReducerError.cannotUndo(targetOpID: target)
                }
                _ = try controlledEntryIndices(
                    targetOpID: target, before: index, in: log)
                targetIsActive[target] = false
                sawUndo.insert(target)
            case .redo(let target):
                guard sawUndo.contains(target) else {
                    throw ReducerError.cannotRedo(targetOpID: target)
                }
                _ = try controlledEntryIndices(
                    targetOpID: target, before: index, in: log)
                targetIsActive[target] = true
            default:
                break
            }
        }

        var inactive: Set<Int> = []
        for (target, isActive) in targetIsActive where !isActive {
            inactive.formUnion(try controlledEntryIndices(
                targetOpID: target, before: log.entries.count, in: log))
        }
        return Set(log.entries.indices.filter { index in
            guard !inactive.contains(index) else { return false }
            switch log.entries[index].op {
            case .undo, .redo, .batchBegin, .batchEnd: return false
            default: return true
            }
        })
    }

    /// Returns the ordinary entries controlled by an undo target. Phase 2b
    /// allows individual setText/style entries; the later canonical taxonomy
    /// also makes a closed batchBegin…batchEnd envelope one logical unit.
    private static func controlledEntryIndices(
        targetOpID: UUID, before controlIndex: Int, in log: OperationLog
    ) throws -> Set<Int> {
        guard let targetIndex = log.entries[..<controlIndex].firstIndex(where: {
            $0.opID == targetOpID
        }) else {
            throw ReducerError.cannotUndo(targetOpID: targetOpID)
        }
        switch log.entries[targetIndex].op {
        case .setText, .setParagraphStyle:
            return [targetIndex]
        case .batchBegin:
            var depth = 0
            for index in targetIndex..<controlIndex {
                switch log.entries[index].op {
                case .batchBegin: depth += 1
                case .batchEnd:
                    depth -= 1
                    if depth == 0 {
                        return Set((targetIndex + 1)..<index)
                    }
                default: break
                }
            }
            throw ReducerError.cannotUndo(targetOpID: targetOpID)
        default:
            throw ReducerError.cannotUndo(targetOpID: targetOpID)
        }
    }

    /// Computes the concrete text/style corrections needed when a document
    /// that already materialized its history appends one control operation.
    /// Structural batch members cannot be reversed from the current tree
    /// without a stored base, so this path fails loudly for those members.
    internal static func controlReplacementOperations(
        for control: Operation, in log: OperationLog
    ) throws -> [Operation] {
        let target: UUID
        switch control {
        case .undo(let id), .redo(let id): target = id
        default: return []
        }
        let controlIndex = log.entries.count - 1
        guard log.entries[..<controlIndex].contains(where: {
            $0.opID == target
        }) else {
            if case .undo = control {
                throw ReducerError.cannotUndo(targetOpID: target)
            }
            throw ReducerError.cannotRedo(targetOpID: target)
        }
        let controlled = try controlledEntryIndices(
            targetOpID: target, before: controlIndex, in: log)
        let active = try activeEntryIndices(in: log)

        var textTargets: [ElementID] = []
        var styleTargets: [ElementID] = []
        for index in controlled.sorted() {
            switch log.entries[index].op {
            case .setText(let element, _):
                if !textTargets.contains(element) { textTargets.append(element) }
            case .setParagraphStyle(let element, _):
                if !styleTargets.contains(element) { styleTargets.append(element) }
            case .batchBegin, .batchEnd:
                break
            default:
                throw ReducerError.cannotUndo(targetOpID: target)
            }
        }

        var replacements: [Operation] = []
        for element in textTargets {
            let value = log.entries.indices.reversed().compactMap { index -> String? in
                guard active.contains(index),
                      case .setText(let candidate, let text) = log.entries[index].op,
                      candidate == element else { return nil }
                return text
            }.first ?? ""
            replacements.append(.setText(target: element, text: value))
        }
        for element in styleTargets {
            var value: String?
            for index in log.entries.indices.reversed() where active.contains(index) {
                if case .setParagraphStyle(let candidate, let style) = log.entries[index].op,
                   candidate == element {
                    value = style
                    break
                }
            }
            replacements.append(.setParagraphStyle(target: element, styleId: value))
        }
        return replacements
    }

    /// Walks `log.entries` in REVERSE order and returns the most recent entry
    /// whose op references the given `ElementID`. Returns `nil` if no entry
    /// touches it. `.unknown` ops are opaque (never count as touching).
    public static func blame(elementID: ElementID, log: OperationLog) -> LogEntry? {
        for entry in log.entries.reversed() {
            if touchesElement(entry.op, elementID: elementID) {
                return entry
            }
        }
        return nil
    }

    // MARK: - Internal apply dispatch (called by cache for tail-replay)

    /// Applies an entry's op to the tree, OR interprets it as a control op
    /// (.undo finds the target in the log so far and applies its inverse;
    /// .batchBegin/.batchEnd are no-ops; .unknown is opaque).
    internal static func applyOrInterpret(
        entry: LogEntry,
        entryIndex: Int,
        log: OperationLog,
        to tree: inout XmlTree
    ) throws {
        switch entry.op {
        case .undo(let targetOpID):
            // Find the target in the log so far (entries before this .undo entry).
            // If found, undo the target on the current tree state.
            let targetIndex = log.entries[..<entryIndex].firstIndex { $0.opID == targetOpID }
            guard let targetIndex else {
                // Target not in prior history — treat as a no-op (forward-compat).
                return
            }
            try applyInverse(
                entry: log.entries[targetIndex],
                entryIndex: targetIndex,
                log: log,
                to: &tree
            )
        case .redo(let targetOpID):
            let priorEntries = log.entries[..<entryIndex]
            guard priorEntries.contains(where: { prior in
                if case .undo(let undoneID) = prior.op { return undoneID == targetOpID }
                return false
            }), let target = priorEntries.first(where: { $0.opID == targetOpID }) else {
                throw ReducerError.cannotRedo(targetOpID: targetOpID)
            }
            try apply(entry: target, to: &tree)
        default:
            try apply(entry: entry, to: &tree)
        }
    }

    /// Applies a single entry's op to the tree. Control and package-addressed
    /// operations are interpreted or routed by the surrounding pipeline.
    internal static func apply(entry: LogEntry, to tree: inout XmlTree) throws {
        switch entry.op {
        case .setText(let target, let text):
            guard let node = findNode(elementID: target, in: tree) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: target)
            }
            let textChild = XmlNode.text(text)
            let wordNS = node.namespaceURI
                ?? "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            let wt = XmlNode.element(
                prefix: "w", localName: "t", namespaceURI: wordNS, children: [textChild])
            if node.kind == .element && node.localName == "p" {
                // Paragraph target (the design's canonical case — setText
                // addresses the paragraph via w14:paraId): full-text-replace
                // semantics. Keep <w:pPr> (paragraph formatting), replace
                // every content child (<w:r>, <w:hyperlink>, stray <w:t>,
                // field wrappers, …) with one fresh <w:r><w:t>text</w:t></w:r>.
                // Pre-fix behavior appended <w:t> directly under <w:p> —
                // schema-invalid and invisible to the runs-based typed view
                // (caught by TypedSetterOpLogTests, task 3.15).
                let wr = XmlNode.element(
                    prefix: "w", localName: "r", namespaceURI: wordNS, children: [wt])
                // 7.3 verify P2: full-text-replace applies to CONTENT
                // children only. Range markers and paragraph properties
                // (<w:pPr>, bookmarkStart/End, commentRangeStart/End,
                // proofErr, permStart/End) are structural siblings under
                // CT_P and MUST survive; dropping them orphaned real
                // bookmarks/comments.
                let replacedContent: Set<String> = [
                    "r", "hyperlink", "fldSimple", "sdt", "smartTag", "t",
                ]
                func isReplaceableContent(_ child: XmlNode) -> Bool {
                    // Non-element paragraph children are not structural
                    // anchors and were part of the old content surface too.
                    guard child.kind == .element else { return true }
                    guard replacedContent.contains(child.localName) else { return false }
                    // A comment reference is a structural anchor wrapped in
                    // a run. Replacing paragraph text must not orphan the
                    // comments part by deleting that reference run.
                    return !(child.localName == "r"
                        && containsDescendant(named: "commentReference", in: child))
                }
                let firstRemovedIndex = node.children.firstIndex(where: isReplaceableContent)
                var keptChildren = node.children.filter { !isReplaceableContent($0) }
                // Insert at the exact location of the first removed content
                // child, mapped into the retained-child array. This preserves
                // both non-empty ranges (`start, oldRun, end`) and empty ranges
                // before following text (`start, end, oldRun`).
                let insertionIndex = firstRemovedIndex.map { removedIndex in
                    node.children[..<removedIndex].filter { !isReplaceableContent($0) }.count
                } ?? keptChildren.firstIndex {
                    $0.localName == "r"
                        && containsDescendant(named: "commentReference", in: $0)
                } ?? keptChildren.count
                keptChildren.insert(wr, at: insertionIndex)
                node.children = keptChildren
            } else {
                // Run-shaped target: replace <w:t> direct children with one
                // new <w:t>X</w:t>. Other children (e.g., <w:rPr>) are
                // preserved.
                let nonTextChildren = node.children.filter {
                    !($0.kind == .element && $0.localName == "t")
                }
                node.children = nonTextChildren + [wt]
            }

        case .setParagraphStyle(let target, let styleId):
            guard let node = findNode(elementID: target, in: tree) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: target)
            }
            try setOrRemoveParagraphStyle(node: node, styleId: styleId)

        case .batchBegin, .batchEnd:
            // Markers only — no tree mutation. Phase 2c reducer-level batch
            // semantics (rollback, group undo) are out of scope here.
            return

        case .appendParagraph(let container, let payload):
            // §4b (#128): append as last block-level child of the container
            // (nil = <w:body>). Explicit paraId stamps w14:paraId; absent
            // paraId keeps the opID-derived libraryUUID convention.
            let parent: XmlNode
            if let containerID = container {
                guard let node = findNode(elementID: containerID, in: tree) else {
                    throw ReducerError.elementNotFound(opID: entry.opID, elementID: containerID)
                }
                parent = node
            } else {
                guard let body = tree.root.children.first(where: {
                    $0.kind == .element && $0.localName == "body"
                }) else {
                    throw ReducerError.malformedOp(
                        opID: entry.opID, reason: "appendParagraph(in: nil) requires a <w:body>")
                }
                parent = body
            }
            let newP = makeParagraph(payload: payload)
            if let paraId = payload.paraId, !paraId.isEmpty {
                newP.setAttribute(prefix: "w14", localName: "paraId", value: paraId)
            } else {
                newP.libraryUUID = entry.opID
            }
            // word-canonical-forms task 2.2: Word-authored paragraph attrs,
            // stamped AFTER paraId in observed order: w14:textId, w:rsidR,
            // w:rsidRPr, w:rsidRDefault, w:rsidP.
            if let v = payload.textId { newP.setAttribute(prefix: "w14", localName: "textId", value: v) }
            if let v = payload.rsidR { newP.setAttribute(prefix: "w", localName: "rsidR", value: v) }
            if let v = payload.rsidRPr { newP.setAttribute(prefix: "w", localName: "rsidRPr", value: v) }
            if let v = payload.rsidRDefault { newP.setAttribute(prefix: "w", localName: "rsidRDefault", value: v) }
            if let v = payload.rsidP { newP.setAttribute(prefix: "w", localName: "rsidP", value: v) }
            // Insert before a trailing body-level <w:sectPr> so the section
            // marker stays last (OOXML body shape).
            if let sectIdx = parent.children.firstIndex(where: {
                $0.kind == .element && $0.localName == "sectPr"
            }) {
                parent.children.insert(newP, at: sectIdx)
            } else {
                parent.children.append(newP)
            }

        case .setRuns(let target, let runs):
            // §4b (#128): replace inline content, keep <w:pPr>.
            guard let node = findNode(elementID: target, in: tree) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: target)
            }
            // 7.4 verify P3: setRuns is paragraph-scoped — applying it to a
            // non-<w:p> would silently replace that node's children.
            guard node.localName == "p" else {
                throw ReducerError.malformedOp(
                    opID: entry.opID,
                    reason: "setRuns target must be a <w:p>, got <\(node.localName)>")
            }
            let kept = node.children.filter {
                $0.kind == .element && $0.localName == "pPr"
            }
            node.children = kept + runs.map { makeRunNode($0) }

        case .setParagraphContent(let target, let items):
            // word-canonical-forms task 2.4: ordered run|marker inline
            // content, keeping <w:pPr>. Markers are self-contained leaf
            // elements stamped verbatim in position.
            guard let node = findNode(elementID: target, in: tree) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: target)
            }
            guard node.localName == "p" else {
                throw ReducerError.malformedOp(
                    opID: entry.opID,
                    reason: "setParagraphContent target must be a <w:p>, got <\(node.localName)>")
            }
            let keptPPr = node.children.filter {
                $0.kind == .element && $0.localName == "pPr"
            }
            let contentNodes = try items.map {
                try makeInlineNode($0, opID: entry.opID)
            }
            node.children = keptPPr + contentNodes

        case .appendTable(let container, let payload):
            // format-alignment-engine Phase B (task 2.5): one-op table
            // authoring. Same container semantics as appendParagraph —
            // nil = <w:body>, inserted before a trailing sectPr.
            let parent: XmlNode
            if let containerID = container {
                guard let node = findNode(elementID: containerID, in: tree) else {
                    throw ReducerError.elementNotFound(opID: entry.opID, elementID: containerID)
                }
                parent = node
            } else {
                guard let body = tree.root.children.first(where: {
                    $0.kind == .element && $0.localName == "body"
                }) else {
                    throw ReducerError.malformedOp(
                        opID: entry.opID, reason: "appendTable(in: nil) requires a <w:body>")
                }
                parent = body
            }
            guard payload.rows > 0, payload.columns > 0 else {
                throw ReducerError.malformedOp(
                    opID: entry.opID, reason: "appendTable requires rows > 0 and columns > 0")
            }
            if let cells = payload.cells {
                guard cells.count == payload.rows,
                      cells.allSatisfy({ $0.count == payload.columns }) else {
                    throw ReducerError.malformedOp(
                        opID: entry.opID,
                        reason: "appendTable cells shape must be rows × columns")
                }
            }
            guard payload.rowIds == nil || payload.rowIds?.count == payload.rows,
                  payload.cellIds == nil || isRectangular(
                    payload.cellIds!, rows: payload.rows, columns: payload.columns) else {
                throw ReducerError.malformedOp(
                    opID: entry.opID,
                    reason: "appendTable stable IDs shape must be rows × columns")
            }
            let newTbl = try makeTable(payload: payload, opID: entry.opID)
            newTbl.libraryUUID = entry.opID
            if let sectIdx = parent.children.firstIndex(where: {
                $0.kind == .element && $0.localName == "sectPr"
            }) {
                parent.children.insert(newTbl, at: sectIdx)
            } else {
                parent.children.append(newTbl)
            }

        case .appendBlockMarker(let marker):
            guard let body = tree.root.children.first(where: {
                $0.kind == .element && $0.localName == "body"
            }) else {
                throw ReducerError.malformedOp(
                    opID: entry.opID,
                    reason: "appendBlockMarker requires a <w:body>")
            }
            let node = XmlNode.element(prefix: "w", localName: marker.localName)
            for attribute in marker.attributes {
                node.setAttribute(prefix: attribute.prefix,
                                  localName: attribute.localName,
                                  value: attribute.value)
            }
            if let sectIndex = body.children.firstIndex(where: {
                $0.kind == .element && $0.localName == "sectPr"
            }) {
                body.children.insert(node, at: sectIndex)
            } else {
                body.children.append(node)
            }

        case .setDocumentProlog(let prolog):
            // word-canonical-forms (task 3.1): carry the synthesized prolog
            // (declaration + separator) so a Word CRLF prolog rebuilds exact.
            tree.synthesizedProlog = Array(prolog.utf8)

        case .setDocumentRoot(let attributes):
            // word-canonical-forms (task 2.1): replace the <w:document> root
            // element's attribute list wholesale, in op order. The root is
            // the tree's root element.
            tree.root.attributes = attributes.map {
                XmlAttribute(prefix: $0.prefix, localName: $0.localName, value: $0.value)
            }

        case .setSectionProperties(let at, let section):
            // format-alignment-engine Phase B (task 2.1): typed sectPr
            // stamping. `at: nil` → trailing body sectPr (replace existing);
            // `at: <paragraph>` → sectPr as the LAST child of that
            // paragraph's pPr (mid-body section break), creating pPr on
            // demand.
            let sectPr = makeSectPr(section)
            if let paraID = at {
                guard let node = findNode(elementID: paraID, in: tree) else {
                    throw ReducerError.elementNotFound(opID: entry.opID, elementID: paraID)
                }
                guard node.localName == "p" else {
                    throw ReducerError.malformedOp(
                        opID: entry.opID,
                        reason: "setSectionProperties(at:) target must be a <w:p>, got <\(node.localName)>")
                }
                let pPr: XmlNode
                if let existing = node.children.first(where: {
                    $0.kind == .element && $0.localName == "pPr"
                }) {
                    pPr = existing
                } else {
                    pPr = XmlNode.element(prefix: "w", localName: "pPr")
                    node.children = [pPr] + node.children
                }
                // Replace any existing sectPr; CT_PPr places sectPr last.
                pPr.children = pPr.children.filter {
                    !($0.kind == .element && $0.localName == "sectPr")
                }
                pPr.children.append(sectPr)
            } else {
                guard let body = tree.root.children.first(where: {
                    $0.kind == .element && $0.localName == "body"
                }) else {
                    throw ReducerError.malformedOp(
                        opID: entry.opID,
                        reason: "setSectionProperties(at: nil) requires a <w:body>")
                }
                body.children = body.children.filter {
                    !($0.kind == .element && $0.localName == "sectPr")
                }
                body.children.append(sectPr)
            }

        case .defineStyle(let payload):
            // §4b (#128): define-on-first-use into <w:styles>; duplicate
            // styleId is an idempotent no-op. The apply() pipeline routes
            // this op to the word/styles.xml part.
            let root = tree.root
            let exists = root.children.contains { child in
                child.kind == .element && child.localName == "style"
                    && child.attributes.contains {
                        $0.prefix == "w" && $0.localName == "styleId" && $0.value == payload.styleId
                    }
            }
            if exists { return }
            var styleChildren: [XmlNode] = []
            let nameEl = XmlNode.element(prefix: "w", localName: "name")
            nameEl.setAttribute(prefix: "w", localName: "val", value: payload.name ?? payload.styleId)
            styleChildren.append(nameEl)
            var rPrChildren: [XmlNode] = []
            if let font = payload.font {
                let rFonts = XmlNode.element(prefix: "w", localName: "rFonts")
                rFonts.setAttribute(prefix: "w", localName: "ascii", value: font)
                rFonts.setAttribute(prefix: "w", localName: "eastAsia", value: font)
                rPrChildren.append(rFonts)
            }
            if payload.bold == true {
                rPrChildren.append(XmlNode.element(prefix: "w", localName: "b"))
            }
            if payload.italic == true {
                rPrChildren.append(XmlNode.element(prefix: "w", localName: "i"))
            }
            if let color = payload.color {
                let c = XmlNode.element(prefix: "w", localName: "color")
                c.setAttribute(prefix: "w", localName: "val", value: color)
                rPrChildren.append(c)
            }
            if let fontSize = payload.fontSize {
                let sz = XmlNode.element(prefix: "w", localName: "sz")
                // w:sz is half-points; StylePayload.fontSize is points.
                sz.setAttribute(prefix: "w", localName: "val", value: String(fontSize * 2))
                rPrChildren.append(sz)
            }
            if !rPrChildren.isEmpty {
                styleChildren.append(XmlNode.element(prefix: "w", localName: "rPr", children: rPrChildren))
            }
            let style = XmlNode.element(prefix: "w", localName: "style", children: styleChildren)
            style.setAttribute(prefix: "w", localName: "type", value: "paragraph")
            style.setAttribute(prefix: "w", localName: "styleId", value: payload.styleId)
            root.children.append(style)

        case .beginComponent, .endComponent:
            // §4b (#128): log-metadata markers (batch-marker pattern) —
            // zero tree mutation, zero OOXML output.
            return

        case .insertTab(let containerID), .insertBreak(let containerID),
             .insertNoBreakHyphen(let containerID):
            // §4b (#128): run-scoped inline atoms. A paragraph target gets a
            // synthesized wrapping <w:r> (atoms are schema-invalid outside runs).
            guard let node = findNode(elementID: containerID, in: tree) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: containerID)
            }
            let atomName: String
            switch entry.op {
            case .insertTab: atomName = "tab"
            case .insertBreak: atomName = "br"
            default: atomName = "noBreakHyphen"
            }
            let atom = XmlNode.element(prefix: "w", localName: atomName)
            if node.kind == .element && node.localName == "r" {
                node.children.append(atom)
            } else if node.kind == .element && node.localName == "p" {
                let wrapper = XmlNode.element(prefix: "w", localName: "r", children: [atom])
                node.children.append(wrapper)
            } else {
                throw ReducerError.malformedOp(
                    opID: entry.opID,
                    reason: "inline atom target must be a <w:r> or <w:p>, got <\(node.localName)>")
            }

        case .unknown:
            // Opaque op — Phase 2b reducer treats as no-op. Phase 2b spec
            // explicitly accepts this: log warning, pass through to next op.
            return

        case .undo:
            // Should not reach here — undo is interpreted in applyOrInterpret.
            return

        case .insertParagraphAfter(let after, let payload):
            // Find target + parent + position.
            guard let (parent, idx) = findParentAndIndex(targetID: after, in: tree.root) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: after)
            }
            let newP = makeParagraph(payload: payload)
            // 7.4 verify P0: a payload paraId (e.g. Word-assigned, carried by
            // the import diff) is the paragraph's real identity — stamp it.
            // Only ID-less payloads fall back to the deterministic
            // libraryUUID == entry.opID replay convention.
            if let paraId = payload.paraId, !paraId.isEmpty {
                newP.setAttribute(prefix: "w14", localName: "paraId", value: paraId)
            } else {
                newP.libraryUUID = entry.opID
            }
            parent.children.insert(newP, at: idx + 1)

        case .insertParagraphBefore(let before, let payload):
            guard let (parent, idx) = findParentAndIndex(targetID: before, in: tree.root) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: before)
            }
            let newP = makeParagraph(payload: payload)
            if let paraId = payload.paraId, !paraId.isEmpty {
                newP.setAttribute(prefix: "w14", localName: "paraId", value: paraId)
            } else {
                newP.libraryUUID = entry.opID
            }
            parent.children.insert(newP, at: idx)

        case .removeParagraph(let target):
            guard let (parent, idx) = findParentAndIndex(targetID: target, in: tree.root) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: target)
            }
            parent.children.remove(at: idx)
            // MVP: removed paragraph's cross-part refs (comments, footnotes,
            // bookmarks pointing TO this paragraph) are left dangling. A
            // follow-up should add orphan-collection — file issue if needed.

        case .insertTable(let anchor, let payload):
            guard payload.rows >= 0, payload.columns >= 0,
                  payload.cells == nil || isRectangular(
                    payload.cells!, rows: payload.rows, columns: payload.columns),
                  payload.richCells == nil || isRectangular(
                    payload.richCells!, rows: payload.rows, columns: payload.columns),
                  payload.rowIds == nil || payload.rowIds?.count == payload.rows,
                  payload.cellIds == nil || isRectangular(
                    payload.cellIds!, rows: payload.rows, columns: payload.columns) else {
                throw ReducerError.malformedOp(
                    opID: entry.opID,
                    reason: "table payload dimensions do not match rows/columns")
            }
            guard let (parent, index) = findParentAndIndex(targetID: anchor, in: tree.root) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: anchor)
            }
            let table = try makeTable(payload: payload, opID: entry.opID)
            table.libraryUUID = entry.opID
            parent.children.insert(table, at: index + 1)

        case .removeTable(let target):
            guard let (parent, index) = findParentAndIndex(targetID: target, in: tree.root) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: target)
            }
            guard parent.children[index].localName == "tbl" else {
                throw ReducerError.malformedOp(
                    opID: entry.opID, reason: "removeTable target must be <w:tbl>")
            }
            parent.children.remove(at: index)

        case .setCellText(let tableID, let row, let column, let text):
            guard let table = findNode(elementID: tableID, in: tree) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: tableID)
            }
            guard table.localName == "tbl", row >= 0, column >= 0 else {
                throw ReducerError.malformedOp(
                    opID: entry.opID, reason: "setCellText requires a table and non-negative indices")
            }
            let rows = table.children.filter { $0.kind == .element && $0.localName == "tr" }
            guard row < rows.count else {
                throw ReducerError.malformedOp(opID: entry.opID, reason: "table row out of range")
            }
            let cells = rows[row].children.filter { $0.kind == .element && $0.localName == "tc" }
            guard column < cells.count else {
                throw ReducerError.malformedOp(opID: entry.opID, reason: "table column out of range")
            }
            let cell = cells[column]
            if let paragraph = cell.children.first(where: {
                $0.kind == .element && $0.localName == "p"
            }) {
                replaceParagraphText(paragraph, with: text)
            } else {
                cell.children.append(makeParagraph(payload: ParagraphPayload(text: text)))
            }

        case .insertRun(let paragraphID, let position, let payload):
            guard let paragraph = findNode(elementID: paragraphID, in: tree) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: paragraphID)
            }
            guard paragraph.localName == "p" else {
                throw ReducerError.malformedOp(
                    opID: entry.opID, reason: "insertRun parent must be <w:p>")
            }
            let contentStart = paragraph.children.firstIndex {
                !($0.kind == .element && $0.localName == "pPr")
            } ?? paragraph.children.count
            let contentCount = paragraph.children.count - contentStart
            guard position >= 0, position <= contentCount else {
                throw ReducerError.malformedOp(opID: entry.opID, reason: "run position out of range")
            }
            let run = makeRunNode(payload)
            run.libraryUUID = entry.opID
            paragraph.children.insert(run, at: contentStart + position)

        case .setRunFormat(let target, let format):
            guard let node = findNode(elementID: target, in: tree) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: target)
            }
            guard node.localName == "r" else {
                throw ReducerError.malformedOp(
                    opID: entry.opID, reason: "setRunFormat target must be <w:r>")
            }
            applyRunFormat(format, to: node)

        case .insertBookmark(let paragraphID, let bookmarkID, let name):
            guard let paragraph = findNode(elementID: paragraphID, in: tree) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: paragraphID)
            }
            guard paragraph.localName == "p", bookmarkID >= 0, !name.isEmpty else {
                throw ReducerError.malformedOp(
                    opID: entry.opID,
                    reason: "insertBookmark requires a paragraph, non-negative id, and non-empty name")
            }
            let start = XmlNode.element(prefix: "w", localName: "bookmarkStart")
            start.setAttribute(prefix: "w", localName: "id", value: String(bookmarkID))
            start.setAttribute(prefix: "w", localName: "name", value: name)
            let end = XmlNode.element(prefix: "w", localName: "bookmarkEnd")
            end.setAttribute(prefix: "w", localName: "id", value: String(bookmarkID))
            let contentStart = paragraph.children.firstIndex {
                !($0.kind == .element && $0.localName == "pPr")
            } ?? paragraph.children.count
            paragraph.children.insert(start, at: contentStart)
            paragraph.children.append(end)

        case .insertComment(let paragraphID, let commentID, _, _):
            guard let paragraph = findNode(elementID: paragraphID, in: tree) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: paragraphID)
            }
            guard paragraph.localName == "p", commentID >= 0 else {
                throw ReducerError.malformedOp(
                    opID: entry.opID,
                    reason: "insertComment requires a paragraph and non-negative id")
            }
            let start = XmlNode.element(prefix: "w", localName: "commentRangeStart")
            start.setAttribute(prefix: "w", localName: "id", value: String(commentID))
            let end = XmlNode.element(prefix: "w", localName: "commentRangeEnd")
            end.setAttribute(prefix: "w", localName: "id", value: String(commentID))
            let reference = XmlNode.element(prefix: "w", localName: "commentReference")
            reference.setAttribute(prefix: "w", localName: "id", value: String(commentID))
            let referenceRun = XmlNode.element(prefix: "w", localName: "r", children: [reference])
            let contentStart = paragraph.children.firstIndex {
                !($0.kind == .element && $0.localName == "pPr")
            } ?? paragraph.children.count
            paragraph.children.insert(start, at: contentStart)
            paragraph.children.append(end)
            paragraph.children.append(referenceRun)

        case .redo:
            // Interpreted by applyOrInterpret; defensive no-op when apply is
            // invoked directly by internal callers.
            return

        case .insertNode(let parentID, let position, let nodeXML):
            guard let parent = findNode(elementID: parentID, in: tree) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: parentID)
            }
            guard position >= 0, position <= parent.children.count else {
                throw ReducerError.malformedOp(opID: entry.opID, reason: "node position out of range")
            }
            do {
                let node = try parseXMLFragment(nodeXML)
                node.libraryUUID = entry.opID
                parent.children.insert(node, at: position)
            } catch let error as ReducerError {
                throw error
            } catch {
                throw ReducerError.malformedOp(
                    opID: entry.opID,
                    reason: "Failed to parse insertNode nodeXML: \(error.localizedDescription)")
            }

        case .removeNode(let target):
            guard let (parent, index) = findParentAndIndex(targetID: target, in: tree.root) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: target)
            }
            parent.children.remove(at: index)

        case .updateAttribute(let target, let prefix, let localName, let value):
            guard let node = findNode(elementID: target, in: tree) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: target)
            }
            guard !localName.isEmpty else {
                throw ReducerError.malformedOp(
                    opID: entry.opID, reason: "attribute localName must not be empty")
            }
            if let value {
                node.setAttribute(prefix: prefix, localName: localName, value: value)
            } else {
                node.attributes.removeAll { $0.prefix == prefix && $0.localName == localName }
                node.markDirty()
            }

        case .moveNode(let sourceID, let destinationParentID, let destinationIndex):
            guard let source = findNode(elementID: sourceID, in: tree) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: sourceID)
            }
            guard let (sourceParent, sourceIndex) = findParentAndIndex(
                targetID: sourceID, in: tree.root) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: sourceID)
            }
            guard let destinationParent = findNode(elementID: destinationParentID, in: tree) else {
                throw ReducerError.elementNotFound(
                    opID: entry.opID, elementID: destinationParentID)
            }
            guard source !== destinationParent,
                  !containsNodeIdentical(to: destinationParent, in: source) else {
                throw ReducerError.malformedOp(
                    opID: entry.opID, reason: "cannot move a node into itself or its descendant")
            }
            guard destinationIndex >= 0,
                  destinationIndex <= destinationParent.children.count else {
                throw ReducerError.malformedOp(
                    opID: entry.opID, reason: "destination index out of range")
            }
            sourceParent.children.remove(at: sourceIndex)
            let adjustedIndex = sourceParent === destinationParent && sourceIndex < destinationIndex
                ? destinationIndex - 1 : destinationIndex
            destinationParent.children.insert(source, at: adjustedIndex)

        case .insertSiblingAfter(let after, let nodeXML):
            // Find target's parent + index, parse XML fragment, insert as
            // next sibling. Used by OOXMLEdit.insertHyperlink to insert
            // <w:hyperlink> wrapper after a Run.
            guard let (parent, idx) = findParentAndIndex(targetID: after, in: tree.root) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: after)
            }
            do {
                let newNode = try parseXMLFragment(nodeXML)
                parent.children.insert(newNode, at: idx + 1)
            } catch {
                throw ReducerError.malformedOp(
                    opID: entry.opID,
                    reason: "Failed to parse insertSiblingAfter nodeXML: \(error.localizedDescription)"
                )
            }

        case .wrapWithHyperlink(let target, let rId):
            // Wrap target with <w:hyperlink r:id="rId">...</w:hyperlink>.
            // Cmd-K semantics: target stays in place but is now nested
            // inside a hyperlink wrapper at target's original position.
            // Used by OOXMLEdit.wrapWithHyperlink (via WordEdit.applyLink
            // single-Run case).
            guard let (parent, idx) = findParentAndIndex(targetID: target, in: tree.root) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: target)
            }
            // Deep-clone target so the wrapper child is structurally
            // separate from the original (we'll remove the original
            // implicitly by replacing it in parent.children).
            let clone = parent.children[idx].deepClone()
            // Build <w:hyperlink r:id="rId">
            let wrapper = XmlNode.element(prefix: "w", localName: "hyperlink")
            wrapper.setAttribute(prefix: "r", localName: "id", value: rId)
            wrapper.children = [clone]
            // Replace target's position with the wrapper. (deepClone
            // preserves target's libraryUUID, so future ops addressing
            // target now resolve to the clone inside the wrapper.)
            parent.children[idx] = wrapper

        case .addRelationship(let part, let id, let type, let target, let targetMode):
            // Append a <Relationship> entry to the specified rels part.
            // Rels-part xml has a fixed structure: <Relationships> root
            // with <Relationship> children. We append to the root's
            // children.
            //
            // The rels part is addressed by path, not ElementID — different
            // from element-addressed ops. Reducer's `apply(entry:to:)`
            // takes a single tree, but addRelationship needs to mutate
            // a SPECIFIC rels tree. The caller (WordDocument.apply) must
            // route this op to the rels part's tree before calling apply.
            //
            // Convention: when addRelationship is applied to a tree whose
            // root is <Relationships>, we append. If the root isn't a
            // relationships container, we throw — this is a routing error
            // (caller passed wrong tree).
            guard tree.root.kind == .element,
                  tree.root.localName == "Relationships" else {
                throw ReducerError.malformedOp(
                    opID: entry.opID,
                    reason: "addRelationship requires tree root to be <Relationships> element, got \(tree.root.localName); did caller route this op to the rels part? (part='\(part)')"
                )
            }
            let rel = XmlNode.element(localName: "Relationship")
            rel.setAttribute(prefix: nil, localName: "Id", value: id)
            rel.setAttribute(prefix: nil, localName: "Type", value: type)
            rel.setAttribute(prefix: nil, localName: "Target", value: target)
            if let targetMode = targetMode {
                rel.setAttribute(prefix: nil, localName: "TargetMode", value: targetMode)
            }
            tree.root.children.append(rel)

        case .carryPart, .carryBinaryPart:
            // Part-addressed raw channel: carryPart never mutates a document
            // tree. WordDocument.appendAndMaterialize intercepts it and stores
            // the verbatim bytes on `carriedParts` before the reducer runs; if
            // it reaches here (defensive), it is a no-op on the tree.
            break
        }
    }

    // MARK: - Phase 2c helpers

    private static func makeInlineNode(_ item: InlineItem, opID: UUID) throws -> XmlNode {
        switch item.kind {
        case .run:
            guard let run = item.run else {
                throw ReducerError.malformedOp(opID: opID, reason: "run item missing run")
            }
            let node = makeRunNode(run)
            node.libraryUUID = item.runID
            return node
        case .marker:
            guard let marker = item.marker else {
                throw ReducerError.malformedOp(opID: opID, reason: "marker item missing marker")
            }
            let element = XmlNode.element(prefix: "w", localName: marker.localName)
            for attribute in marker.attributes {
                element.setAttribute(prefix: attribute.prefix,
                                     localName: attribute.localName,
                                     value: attribute.value)
            }
            return element
        case .hyperlink:
            guard let hyperlink = item.hyperlink else {
                throw ReducerError.malformedOp(
                    opID: opID, reason: "hyperlink item missing payload")
            }
            let element = XmlNode.element(prefix: "w", localName: "hyperlink")
            switch hyperlink.targetKind {
            case .external:
                guard let relationshipId = hyperlink.relationshipId else {
                    throw ReducerError.malformedOp(
                        opID: opID, reason: "external hyperlink missing relationshipId")
                }
                element.setAttribute(prefix: "r", localName: "id", value: relationshipId)
            case .anchor:
                element.setAttribute(prefix: "w", localName: "anchor", value: hyperlink.target)
            }
            if let items = hyperlink.items {
                element.children = try items.map { try makeInlineNode($0, opID: opID) }
            } else {
                element.children = hyperlink.runs.map(makeRunNode)
            }
            return element
        }
    }

    /// Walks the tree from `root` looking for a node whose ElementID matches
    /// `targetID`. Returns the target's parent and its index in
    /// `parent.children`. Returns `nil` if target not found.
    ///
    /// Root cannot be the target (root has no parent) — if `targetID` matches
    /// `root` itself, returns nil. That's a limitation: ops can't target the
    /// document root, only nested elements (which is fine for OOXML — body
    /// children and below are the only meaningful targets).
    internal static func findParentAndIndex(
        targetID: ElementID,
        in root: XmlNode
    ) -> (XmlNode, Int)? {
        for (idx, child) in root.children.enumerated() {
            if let id = ElementID(node: child), id == targetID {
                return (root, idx)
            }
            if let found = findParentAndIndex(targetID: targetID, in: child) {
                return found
            }
        }
        return nil
    }

    /// Parses an XML fragment string into an XmlNode by wrapping it in a
    /// synthetic root with all common OOXML namespace declarations and
    /// then stripping the synthetic root.
    ///
    /// Used by Phase 2c Reducer cases (insertSiblingAfter, insertNode)
    /// to materialize the nodeXML payload into a real tree node.
    ///
    /// Throws if the fragment is malformed XML or doesn't contain exactly
    /// one root element after parsing.
    internal static func parseXMLFragment(_ fragment: String) throws -> XmlNode {
        // Wrap the fragment in a synthetic root carrying common OOXML
        // namespace declarations. XmlTreeReader requires a fully-formed
        // document, so we synthesize one.
        let wrappedXML =
            "<__wrap__ " +
            "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
            "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">" +
            fragment +
            "</__wrap__>"
        guard let data = wrappedXML.data(using: .utf8) else {
            throw NSError(
                domain: "OperationReducer.parseXMLFragment",
                code: 1,
                userInfo: [NSLocalizedDescriptionKey: "Failed to encode fragment as UTF-8"]
            )
        }
        let tree = try XmlTreeReader.parse(data)
        // The synthetic root's first element child IS the parsed fragment.
        guard let parsed = tree.root.children.first(where: { $0.kind == .element }) else {
            throw NSError(
                domain: "OperationReducer.parseXMLFragment",
                code: 2,
                userInfo: [NSLocalizedDescriptionKey: "Fragment contained no element"]
            )
        }
        return parsed
    }

    private static func isRectangular<T>(
        _ values: [[T]], rows expectedRows: Int, columns expectedColumns: Int
    ) -> Bool {
        values.count == expectedRows && values.allSatisfy { $0.count == expectedColumns }
    }

    private static func replaceParagraphText(_ paragraph: XmlNode, with text: String) {
        let properties = paragraph.children.filter {
            $0.kind == .element && $0.localName == "pPr"
        }
        paragraph.children = properties + [makeRunNode(RunPayload(text: text))]
    }

    private static func containsNodeIdentical(to target: XmlNode, in root: XmlNode) -> Bool {
        if root === target { return true }
        return root.children.contains { containsNodeIdentical(to: target, in: $0) }
    }

    private static func applyRunFormat(_ format: RunFormatPayload, to run: XmlNode) {
        if let bold = format.bold {
            setBooleanRunProperty("b", enabled: bold, in: run)
        }
        if let italic = format.italic {
            setBooleanRunProperty("i", enabled: italic, in: run)
        }
        if let underline = format.underline {
            if underline {
                setValuedRunProperty("u", value: "single", in: run)
            } else {
                removeRunProperty("u", from: run)
            }
        }
        if let size = format.fontSizeHalfPoints {
            setValuedRunProperty("sz", value: String(size), in: run)
        }
        if let color = format.color {
            setValuedRunProperty("color", value: color, in: run)
        }
        removeEmptyRunProperties(from: run)
    }

    private static func setBooleanRunProperty(
        _ localName: String, enabled: Bool, in run: XmlNode
    ) {
        if enabled {
            _ = upsertRunProperty(localName, in: run)
        } else {
            removeRunProperty(localName, from: run)
        }
    }

    private static func setValuedRunProperty(
        _ localName: String, value: String, in run: XmlNode
    ) {
        upsertRunProperty(localName, in: run)
            .setAttribute(prefix: "w", localName: "val", value: value)
    }

    private static func upsertRunProperty(_ localName: String, in run: XmlNode) -> XmlNode {
        let rPr: XmlNode
        if let existing = run.children.first(where: {
            $0.kind == .element && $0.localName == "rPr"
        }) {
            rPr = existing
        } else {
            rPr = XmlNode.element(prefix: "w", localName: "rPr")
            run.children.insert(rPr, at: 0)
        }
        if let existing = rPr.children.first(where: {
            $0.kind == .element && $0.localName == localName
        }) {
            return existing
        }

        let node = XmlNode.element(prefix: "w", localName: localName)
        let rank = runPropertyRank(localName)
        let insertionIndex = rPr.children.firstIndex {
            runPropertyRank($0.localName) > rank
        } ?? rPr.children.count
        rPr.children.insert(node, at: insertionIndex)
        return node
    }

    private static func removeRunProperty(_ localName: String, from run: XmlNode) {
        guard let rPr = run.children.first(where: {
            $0.kind == .element && $0.localName == "rPr"
        }) else { return }
        rPr.children.removeAll { $0.kind == .element && $0.localName == localName }
    }

    private static func removeEmptyRunProperties(from run: XmlNode) {
        run.children.removeAll {
            $0.kind == .element && $0.localName == "rPr" && $0.children.isEmpty
        }
    }

    private static func runPropertyRank(_ localName: String) -> Int {
        let order = [
            "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
            "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof",
            "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
            "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd",
            "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
            "specVanish", "oMath"
        ]
        return order.firstIndex(of: localName) ?? Int.max
    }

    /// Constructs a fresh `<w:p>` element with one `<w:r><w:t>text</w:t></w:r>`
    /// child run. Optional `<w:pPr><w:pStyle w:val="styleId"/></w:pPr>` is
    /// prepended when `styleId` is non-nil and non-empty.
    ///
    /// The returned paragraph has NO libraryUUID — caller is responsible for
    /// assigning one (typically `entry.opID` for deterministic-replay).
    /// Child Run is anonymous (no libraryUUID) — future Operations addressing
    /// the run must use path-based addressing (entry.opID + child path index).
    internal static func makeParagraph(text: String, styleId: String?) -> XmlNode {
        makeParagraph(payload: ParagraphPayload(text: text, styleId: styleId))
    }

    /// Builds a `<w:p>` from the full payload. pPr children follow CT_PPr
    /// schema order: pStyle, numPr, spacing, ind, jc (format-alignment-engine
    /// Phase B task 2.1 — reducer stamping of the additive pPr fields).
    internal static func makeParagraph(payload: ParagraphPayload) -> XmlNode {
        let wordNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        let textNode = XmlNode.text(payload.text)
        let wt = XmlNode.element(
            prefix: "w", localName: "t", namespaceURI: wordNS, children: [textNode])
        let wr = XmlNode.element(
            prefix: "w", localName: "r", namespaceURI: wordNS, children: [wt])

        var children: [XmlNode] = []
        let pPrChildren = makePPrChildren(payload: payload)
        if !pPrChildren.isEmpty {
            children.append(XmlNode.element(
                prefix: "w", localName: "pPr", namespaceURI: wordNS, children: pPrChildren))
        }
        children.append(wr)
        return XmlNode.element(
            prefix: "w", localName: "p", namespaceURI: wordNS, children: children)
    }

    /// pPr children in CT_PPr schema order from the payload's optional fields.
    private static func makePPrChildren(payload: ParagraphPayload) -> [XmlNode] {
        var out: [XmlNode] = []
        if let styleId = payload.styleId, !styleId.isEmpty {
            let pStyle = XmlNode.element(prefix: "w", localName: "pStyle")
            pStyle.setAttribute(prefix: "w", localName: "val", value: styleId)
            out.append(pStyle)
        }
        if payload.numId != nil || payload.numLevel != nil {
            var numChildren: [XmlNode] = []
            if let ilvl = payload.numLevel {
                let n = XmlNode.element(prefix: "w", localName: "ilvl")
                n.setAttribute(prefix: "w", localName: "val", value: String(ilvl))
                numChildren.append(n)
            }
            if let numId = payload.numId {
                let n = XmlNode.element(prefix: "w", localName: "numId")
                n.setAttribute(prefix: "w", localName: "val", value: String(numId))
                numChildren.append(n)
            }
            out.append(XmlNode.element(prefix: "w", localName: "numPr", children: numChildren))
        }
        if payload.spacingBefore != nil || payload.spacingAfter != nil
            || payload.spacingLine != nil || payload.spacingLineRule != nil {
            let spacing = XmlNode.element(prefix: "w", localName: "spacing")
            if let v = payload.spacingBefore {
                spacing.setAttribute(prefix: "w", localName: "before", value: String(v))
            }
            if let v = payload.spacingAfter {
                spacing.setAttribute(prefix: "w", localName: "after", value: String(v))
            }
            if let v = payload.spacingLine {
                spacing.setAttribute(prefix: "w", localName: "line", value: String(v))
            }
            if let v = payload.spacingLineRule {
                spacing.setAttribute(prefix: "w", localName: "lineRule", value: v)
            }
            out.append(spacing)
        }
        if payload.indentLeft != nil || payload.indentRight != nil
            || payload.indentFirstLine != nil || payload.indentHanging != nil
            || payload.indentFirstLineChars != nil || payload.indentHangingChars != nil {
            let ind = XmlNode.element(prefix: "w", localName: "ind")
            // Order (task 3.1): left, right, firstLineChars, firstLine,
            // hangingChars, hanging.
            if let v = payload.indentLeft {
                ind.setAttribute(prefix: "w", localName: "left", value: String(v))
            }
            if let v = payload.indentRight {
                ind.setAttribute(prefix: "w", localName: "right", value: String(v))
            }
            if let v = payload.indentFirstLineChars {
                ind.setAttribute(prefix: "w", localName: "firstLineChars", value: String(v))
            }
            if let v = payload.indentFirstLine {
                ind.setAttribute(prefix: "w", localName: "firstLine", value: String(v))
            }
            if let v = payload.indentHangingChars {
                ind.setAttribute(prefix: "w", localName: "hangingChars", value: String(v))
            }
            if let v = payload.indentHanging {
                ind.setAttribute(prefix: "w", localName: "hanging", value: String(v))
            }
            out.append(ind)
        }
        if let jc = payload.alignment {
            let n = XmlNode.element(prefix: "w", localName: "jc")
            n.setAttribute(prefix: "w", localName: "val", value: jc)
            out.append(n)
        }
        // Paragraph-mark run properties as the trailing pPr child (task 3.1).
        if let markRun = payload.paragraphMarkRun, let rPr = makeRPr(markRun) {
            out.append(rPr)
        }
        return out
    }

    /// Builds a `<w:tbl>` from the payload (format-alignment-engine Phase B
    /// task 2.5). Canonical minimal form: `<w:tblGrid>` with one bare
    /// `<w:gridCol/>` per column, then one `<w:tr>` per row whose `<w:tc>`
    /// each hold a single paragraph in `makeParagraph`'s own shape — the
    /// vocabulary the reverse extractor recognizes for the byte-equal
    /// upgrade rule.
    internal static func makeTable(payload: TablePayload, opID: UUID) throws -> XmlNode {
        var tblChildren: [XmlNode] = []
        var gridCols: [XmlNode] = []
        for _ in 0..<payload.columns {
            gridCols.append(XmlNode.element(prefix: "w", localName: "gridCol"))
        }
        tblChildren.append(XmlNode.element(prefix: "w", localName: "tblGrid", children: gridCols))
        for row in 0..<payload.rows {
            var cells: [XmlNode] = []
            for column in 0..<payload.columns {
                let paragraphs: [XmlNode]
                if let rich = payload.richCells?[row][column], !rich.isEmpty {
                    paragraphs = try rich.map { richParagraph in
                        let p = makeParagraph(payload: richParagraph.paragraph)
                        if let paraId = richParagraph.paragraph.paraId, !paraId.isEmpty {
                            p.setAttribute(prefix: "w14", localName: "paraId", value: paraId)
                        }
                        if let items = richParagraph.items {
                            let pPr = p.children.filter {
                                $0.kind == .element && $0.localName == "pPr"
                            }
                            p.children = pPr + (try items.map {
                                try makeInlineNode($0, opID: opID)
                            })
                        } else if let runs = richParagraph.runs {
                            let pPr = p.children.filter {
                                $0.kind == .element && $0.localName == "pPr"
                            }
                            p.children = pPr + runs.map(makeRunNode)
                        }
                        return p
                    }
                } else {
                    let text = payload.cells?[row][column] ?? ""
                    paragraphs = [makeParagraph(payload: ParagraphPayload(text: text))]
                }
                let cell = XmlNode.element(
                    prefix: "w", localName: "tc", children: paragraphs)
                if let id = payload.cellIds?[row][column], !id.isEmpty {
                    cell.setAttribute(prefix: "w14", localName: "textId", value: id)
                }
                cells.append(cell)
            }
            let rowNode = XmlNode.element(prefix: "w", localName: "tr", children: cells)
            if let id = payload.rowIds?[row], !id.isEmpty {
                rowNode.setAttribute(prefix: "w14", localName: "textId", value: id)
            }
            tblChildren.append(rowNode)
        }
        let table = XmlNode.element(prefix: "w", localName: "tbl", children: tblChildren)
        if let id = payload.tableId, !id.isEmpty {
            table.setAttribute(prefix: "w14", localName: "textId", value: id)
        }
        return table
    }

    /// Builds a `<w:r>` from a RunPayload. rPr children follow CT_RPr schema
    /// order (rFonts, b, i, color, sz, u, vertAlign); run-level rsids stamp
    /// on `<w:r>` (rsidR, rsidRPr); xml:space rides `<w:t>`. Shared by
    /// setRuns and setParagraphContent (task 2.4).
    /// Builds a `<w:rPr>` from a RunPayload's formatting fields, or nil when
    /// none are set. Children follow CT_RPr schema order: rFonts, b, bCs, i,
    /// iCs, color, sz, szCs, u, vertAlign. rFonts attributes follow Word's
    /// observed order: ascii, eastAsia, hAnsi, hint, asciiTheme,
    /// eastAsiaTheme, hAnsiTheme. Shared by run rPr and the pPr/rPr
    /// paragraph-mark properties (word-canonical-forms task 3.1).
    private static func makeRPr(_ run: RunPayload) -> XmlNode? {
        var rPrChildren: [XmlNode] = []
        if run.fontAscii != nil || run.fontEastAsia != nil || run.fontHAnsi != nil
            || run.fontHint != nil || run.fontAsciiTheme != nil
            || run.fontEastAsiaTheme != nil || run.fontHAnsiTheme != nil {
            let rFonts = XmlNode.element(prefix: "w", localName: "rFonts")
            if let v = run.fontAscii { rFonts.setAttribute(prefix: "w", localName: "ascii", value: v) }
            if let v = run.fontEastAsia { rFonts.setAttribute(prefix: "w", localName: "eastAsia", value: v) }
            if let v = run.fontHAnsi { rFonts.setAttribute(prefix: "w", localName: "hAnsi", value: v) }
            if let v = run.fontHint { rFonts.setAttribute(prefix: "w", localName: "hint", value: v) }
            if let v = run.fontAsciiTheme { rFonts.setAttribute(prefix: "w", localName: "asciiTheme", value: v) }
            if let v = run.fontEastAsiaTheme { rFonts.setAttribute(prefix: "w", localName: "eastAsiaTheme", value: v) }
            if let v = run.fontHAnsiTheme { rFonts.setAttribute(prefix: "w", localName: "hAnsiTheme", value: v) }
            rPrChildren.append(rFonts)
        }
        func appendOnOff(_ localName: String, _ value: Bool?) {
            guard let value else { return }
            let node = XmlNode.element(prefix: "w", localName: localName)
            if !value {
                node.setAttribute(prefix: "w", localName: "val", value: "0")
            }
            rPrChildren.append(node)
        }
        appendOnOff("b", run.bold)
        appendOnOff("bCs", run.boldCs)
        appendOnOff("i", run.italic)
        appendOnOff("iCs", run.italicCs)
        appendOnOff("strike", run.strikethrough)
        appendOnOff("noProof", run.noProof)
        if let color = run.color {
            let c = XmlNode.element(prefix: "w", localName: "color")
            c.setAttribute(prefix: "w", localName: "val", value: color)
            rPrChildren.append(c)
        }
        if let sz = run.sizeHalfPoints {
            let n = XmlNode.element(prefix: "w", localName: "sz")
            n.setAttribute(prefix: "w", localName: "val", value: String(sz))
            rPrChildren.append(n)
        }
        if let szCs = run.sizeCsHalfPoints {
            let n = XmlNode.element(prefix: "w", localName: "szCs")
            n.setAttribute(prefix: "w", localName: "val", value: String(szCs))
            rPrChildren.append(n)
        }
        if let u = run.underline {
            let n = XmlNode.element(prefix: "w", localName: "u")
            n.setAttribute(prefix: "w", localName: "val", value: u)
            rPrChildren.append(n)
        }
        if let va = run.vertAlign {
            let n = XmlNode.element(prefix: "w", localName: "vertAlign")
            n.setAttribute(prefix: "w", localName: "val", value: va)
            rPrChildren.append(n)
        }
        return rPrChildren.isEmpty
            ? nil : XmlNode.element(prefix: "w", localName: "rPr", children: rPrChildren)
    }

    private static func makeRunNode(_ run: RunPayload) -> XmlNode {
        var rChildren: [XmlNode] = []
        if let rPr = makeRPr(run) { rChildren.append(rPr) }
        let wt = XmlNode.element(prefix: "w", localName: "t", children: [XmlNode.text(run.text)])
        if run.preserveSpace == true {
            wt.setAttribute(prefix: "xml", localName: "space", value: "preserve")
        }
        rChildren.append(wt)
        let wr = XmlNode.element(prefix: "w", localName: "r", children: rChildren)
        if let v = run.rsidR { wr.setAttribute(prefix: "w", localName: "rsidR", value: v) }
        if let v = run.rsidRPr { wr.setAttribute(prefix: "w", localName: "rsidRPr", value: v) }
        return wr
    }

    /// Builds a `<w:sectPr>` from the payload. Children follow CT_SectPr
    /// schema order: headerReference*, footerReference*, pgSz, pgMar, cols.
    private static func makeSectPr(_ section: SectionPayload) -> XmlNode {
        var children: [XmlNode] = []
        for ref in section.headerReferences ?? [] {
            let n = XmlNode.element(prefix: "w", localName: "headerReference")
            n.setAttribute(prefix: "w", localName: "type", value: ref.type)
            n.setAttribute(prefix: "r", localName: "id", value: ref.relationshipId)
            children.append(n)
        }
        for ref in section.footerReferences ?? [] {
            let n = XmlNode.element(prefix: "w", localName: "footerReference")
            n.setAttribute(prefix: "w", localName: "type", value: ref.type)
            n.setAttribute(prefix: "r", localName: "id", value: ref.relationshipId)
            children.append(n)
        }
        // <w:type> before pgSz (CT_SectPr order; task 3.1).
        if let v = section.sectionType {
            let t = XmlNode.element(prefix: "w", localName: "type")
            t.setAttribute(prefix: "w", localName: "val", value: v)
            children.append(t)
        }
        if section.pageWidth != nil || section.pageHeight != nil || section.orientation != nil
            || section.pageSizeCode != nil {
            let pgSz = XmlNode.element(prefix: "w", localName: "pgSz")
            if let v = section.pageWidth {
                pgSz.setAttribute(prefix: "w", localName: "w", value: String(v))
            }
            if let v = section.pageHeight {
                pgSz.setAttribute(prefix: "w", localName: "h", value: String(v))
            }
            if let v = section.orientation {
                pgSz.setAttribute(prefix: "w", localName: "orient", value: v)
            }
            if let v = section.pageSizeCode {
                pgSz.setAttribute(prefix: "w", localName: "code", value: String(v))
            }
            children.append(pgSz)
        }
        if section.marginTop != nil || section.marginRight != nil || section.marginBottom != nil
            || section.marginLeft != nil || section.marginHeader != nil
            || section.marginFooter != nil || section.marginGutter != nil {
            let pgMar = XmlNode.element(prefix: "w", localName: "pgMar")
            if let v = section.marginTop {
                pgMar.setAttribute(prefix: "w", localName: "top", value: String(v))
            }
            if let v = section.marginRight {
                pgMar.setAttribute(prefix: "w", localName: "right", value: String(v))
            }
            if let v = section.marginBottom {
                pgMar.setAttribute(prefix: "w", localName: "bottom", value: String(v))
            }
            if let v = section.marginLeft {
                pgMar.setAttribute(prefix: "w", localName: "left", value: String(v))
            }
            if let v = section.marginHeader {
                pgMar.setAttribute(prefix: "w", localName: "header", value: String(v))
            }
            if let v = section.marginFooter {
                pgMar.setAttribute(prefix: "w", localName: "footer", value: String(v))
            }
            if let v = section.marginGutter {
                pgMar.setAttribute(prefix: "w", localName: "gutter", value: String(v))
            }
            children.append(pgMar)
        }
        if section.columnCount != nil || section.columnSpace != nil {
            let cols = XmlNode.element(prefix: "w", localName: "cols")
            if let v = section.columnCount {
                cols.setAttribute(prefix: "w", localName: "num", value: String(v))
            }
            if let v = section.columnSpace {
                cols.setAttribute(prefix: "w", localName: "space", value: String(v))
            }
            children.append(cols)
        }
        // word-canonical-forms task 3.1: <w:docGrid> after cols (CT_SectPr order).
        if section.docGridType != nil || section.docGridLinePitch != nil {
            let docGrid = XmlNode.element(prefix: "w", localName: "docGrid")
            if let v = section.docGridType {
                docGrid.setAttribute(prefix: "w", localName: "type", value: v)
            }
            if let v = section.docGridLinePitch {
                docGrid.setAttribute(prefix: "w", localName: "linePitch", value: String(v))
            }
            children.append(docGrid)
        }
        let sectPr = XmlNode.element(prefix: "w", localName: "sectPr", children: children)
        // sectPr element rsids, order: rsidR, rsidRPr, rsidSect (tasks 2.2/3.1).
        if let v = section.rsidR { sectPr.setAttribute(prefix: "w", localName: "rsidR", value: v) }
        if let v = section.rsidRPr { sectPr.setAttribute(prefix: "w", localName: "rsidRPr", value: v) }
        if let v = section.rsidSect { sectPr.setAttribute(prefix: "w", localName: "rsidSect", value: v) }
        return sectPr
    }

    /// Applies the INVERSE of an entry's op (used by `undo` and by `.undo`
    /// log entry interpretation). Phase 2b inverse coverage is limited to
    /// `setText` and `setParagraphStyle`. Other op kinds throw `.cannotUndo`.
    internal static func applyInverse(
        entry: LogEntry,
        entryIndex: Int,
        log: OperationLog,
        to tree: inout XmlTree
    ) throws {
        switch entry.op {
        case .setText(let target, _):
            // Find the most recent prior setText for the same target.
            let priorText = previousSetText(forTarget: target, beforeIndex: entryIndex, in: log) ?? ""
            // Apply: replace target's text with priorText.
            try apply(entry: LogEntry(
                opID: entry.opID,  // same opID — caller doesn't see this
                op: .setText(target: target, text: priorText),
                source: entry.source,
                timestamp: entry.timestamp
            ), to: &tree)

        case .setParagraphStyle(let target, _):
            // Find the most recent prior setParagraphStyle for the same target.
            let priorStyle = previousSetParagraphStyle(forTarget: target, beforeIndex: entryIndex, in: log)
            try apply(entry: LogEntry(
                opID: entry.opID,
                op: .setParagraphStyle(target: target, styleId: priorStyle),
                source: entry.source,
                timestamp: entry.timestamp
            ), to: &tree)

        default:
            throw ReducerError.cannotUndo(targetOpID: entry.opID)
        }
    }

    // MARK: - Helpers

    /// Recursive tree walk. Returns the first `XmlNode` whose derived
    /// `ElementID` matches the given target. Linear-time per call; Phase 5+
    /// may add an ID index for performance.
    internal static func findNode(elementID: ElementID, in tree: XmlTree) -> XmlNode? {
        // 7.2 verify P1 (ElementID cross-space collision): `w:id`/`r:id`
        // values are independently numbered per OOXML feature (bookmarks,
        // footnote refs, revisions, …) yet derive identical ElementID raw
        // strings. First-match would silently mutate whichever element
        // appears first — collect ALL matches for the collision-prone
        // forms and refuse loudly on ambiguity. `w14:paraId` (128-bit
        // random) keeps the early-return fast path.
        if elementID.raw.hasPrefix("w14:paraId=") || elementID.raw.hasPrefix("lib:") {
            return findFirstNode(elementID: elementID, in: tree.root)
        }
        var matches: [XmlNode] = []
        collectNodes(elementID: elementID, in: tree.root, into: &matches, cap: 2)
        if matches.count > 1 {
            // Ambiguity is not representable as a return value here (the
            // signature predates this guard); surfacing it loudly matters
            // more than the exact channel — callers treat nil as
            // elementNotFound and throw, which is loud, never silent
            // wrong-element mutation.
            return nil
        }
        return matches.first
    }

    private static func containsDescendant(named localName: String, in node: XmlNode) -> Bool {
        if node.kind == .element && node.localName == localName { return true }
        return node.children.contains { containsDescendant(named: localName, in: $0) }
    }

    private static func findFirstNode(elementID: ElementID, in node: XmlNode) -> XmlNode? {
        if let nodeID = ElementID(node: node), nodeID == elementID {
            return node
        }
        for child in node.children {
            if let found = findFirstNode(elementID: elementID, in: child) {
                return found
            }
        }
        return nil
    }

    private static func collectNodes(
        elementID: ElementID, in node: XmlNode, into matches: inout [XmlNode], cap: Int
    ) {
        if matches.count >= cap { return }
        if let nodeID = ElementID(node: node), nodeID == elementID {
            matches.append(node)
            if matches.count >= cap { return }
        }
        for child in node.children {
            collectNodes(elementID: elementID, in: child, into: &matches, cap: cap)
            if matches.count >= cap { return }
        }
    }

    /// Walks log backwards from `beforeIndex` looking for the most recent
    /// `setText` entry targeting the same ElementID. Returns the text value
    /// or `nil` if no prior entry exists.
    private static func previousSetText(forTarget target: ElementID, beforeIndex: Int, in log: OperationLog) -> String? {
        guard beforeIndex > 0 else { return nil }
        for i in (0..<beforeIndex).reversed() {
            if case .setText(let t, let txt) = log.entries[i].op, t == target {
                return txt
            }
        }
        return nil
    }

    private static func previousSetParagraphStyle(forTarget target: ElementID, beforeIndex: Int, in log: OperationLog) -> String? {
        guard beforeIndex > 0 else { return nil }
        for i in (0..<beforeIndex).reversed() {
            if case .setParagraphStyle(let t, let styleId) = log.entries[i].op, t == target {
                return styleId
            }
        }
        return nil
    }

    /// Extracts the ElementIDs an Operation references in its payload. Used
    /// by WordDocument.apply for per-op part scoping (locate the part whose
    /// tree contains the op's target so we don't replay the op against every
    /// part in the document).
    ///
    /// Control ops (.undo/.redo/.batchBegin/.batchEnd/.unknown) have no
    /// referenced ElementIDs and return [].
    internal static func referencedElementIDs(in op: Operation) -> [ElementID] {
        switch op {
        case .insertParagraphAfter(let after, _): return [after]
        case .insertParagraphBefore(let before, _): return [before]
        case .removeParagraph(let id): return [id]
        case .setText(let target, _): return [target]
        case .setParagraphStyle(let target, _): return [target]
        case .insertTable(let at, _): return [at]
        case .removeTable(let id): return [id]
        case .setCellText(let table, _, _, _): return [table]
        case .insertRun(let parent, _, _): return [parent]
        case .setRunFormat(let target, _): return [target]
        case .insertBookmark(let at, _, _): return [at]
        case .insertComment(let anchor, _, _, _): return [anchor]
        case .insertNode(let parent, _, _): return [parent]
        case .removeNode(let target): return [target]
        case .updateAttribute(let target, _, _, _): return [target]
        case .moveNode(let source, let dest, _): return [source, dest]
        case .insertSiblingAfter(let after, _): return [after]
        case .wrapWithHyperlink(let target, _): return [target]
        case .addRelationship:
            // Rels-part operations don't address an ElementID in any tree —
            // they target a part by path. WordDocument.apply's per-op
            // scoping treats this as a special case: returns nil from
            // partContaining → fallback to direct part lookup by name.
            return []
        case .appendParagraph(let container, _): return container.map { [$0] } ?? []
        case .setRuns(let target, _): return [target]
        case .insertTab(let c): return [c]
        case .insertBreak(let c): return [c]
        case .insertNoBreakHyphen(let c): return [c]
        case .defineStyle:
            // Part-addressed (word/styles.xml) — routed by the apply pipeline,
            // same pattern as addRelationship.
            return []
        case .beginComponent, .endComponent:
            // Log-metadata markers; the component id is not a tree node.
            return []
        case .undo, .redo, .batchBegin, .batchEnd, .unknown: return []
        case .carryPart, .carryBinaryPart: return []  // part-addressed raw channel, no ElementID
        case .setSectionProperties(let at, _): return at.map { [$0] } ?? []
        case .appendTable(let container, _): return container.map { [$0] } ?? []
        case .appendBlockMarker: return []
        case .setDocumentRoot: return []  // root-addressed, no ElementID
        case .setDocumentProlog: return []  // document-addressed, no ElementID
        case .setParagraphContent(let target, _): return [target]
        }
    }

    /// Returns the path of the part in `trees` whose tree contains a node
    /// referenced by `op`. Returns nil if no part contains any referenced
    /// node (caller should treat as elementNotFound — the op cannot be
    /// applied).
    ///
    /// Used by WordDocument.apply to scope per-op materialization: each op
    /// is replayed only against the part its target lives in, not against
    /// every part in the document (which would throw elementNotFound for
    /// non-target parts).
    internal static func partContaining(
        op: Operation,
        in trees: [String: XmlTree]
    ) -> String? {
        let targetIDs = referencedElementIDs(in: op)
        guard !targetIDs.isEmpty else {
            // Control op (batch markers, undo/redo, unknown) — no target part.
            // Caller's responsibility to handle (typically: skip materialize).
            return nil
        }
        for (path, tree) in trees {
            for targetID in targetIDs {
                if findNode(elementID: targetID, in: tree) != nil {
                    return path
                }
            }
        }
        return nil
    }

    /// Returns true if the op references the given ElementID in any of its
    /// associated values. `.unknown` ops are opaque and never count.
    private static func touchesElement(_ op: Operation, elementID: ElementID) -> Bool {
        switch op {
        case .insertParagraphAfter(let after, _): return after == elementID
        case .insertParagraphBefore(let before, _): return before == elementID
        case .removeParagraph(let id): return id == elementID
        case .setText(let target, _): return target == elementID
        case .setParagraphStyle(let target, _): return target == elementID
        case .insertTable(let at, _): return at == elementID
        case .removeTable(let id): return id == elementID
        case .setCellText(let table, _, _, _): return table == elementID
        case .insertRun(let parent, _, _): return parent == elementID
        case .setRunFormat(let target, _): return target == elementID
        case .insertBookmark(let at, _, _): return at == elementID
        case .insertComment(let anchor, _, _, _): return anchor == elementID
        case .appendParagraph(let container, _): return container == elementID
        case .setRuns(let target, _): return target == elementID
        case .insertTab(let c): return c == elementID
        case .insertBreak(let c): return c == elementID
        case .insertNoBreakHyphen(let c): return c == elementID
        case .defineStyle: return false
        case .beginComponent(_, let id): return id == elementID
        case .endComponent(let id): return id == elementID
        case .undo, .redo, .batchBegin, .batchEnd: return false
        case .insertNode(let parent, _, _): return parent == elementID
        case .removeNode(let target): return target == elementID
        case .updateAttribute(let target, _, _, _): return target == elementID
        case .moveNode(let source, let dest, _): return source == elementID || dest == elementID
        case .insertSiblingAfter(let after, _): return after == elementID
        case .wrapWithHyperlink(let target, _): return target == elementID
        case .addRelationship: return false  // rels-part operation, not element-addressed
        case .carryPart, .carryBinaryPart: return false  // part-addressed raw channel, not element-addressed
        case .setSectionProperties(let at, _): return at == elementID
        case .appendTable(let container, _): return container == elementID
        case .appendBlockMarker: return false
        case .setDocumentRoot: return false  // root-addressed, not element-addressed
        case .setDocumentProlog: return false  // document-addressed
        case .setParagraphContent(let target, _): return target == elementID
        case .unknown: return false
        }
    }

    /// Sets or removes the paragraph style on a `<w:p>` node by manipulating
    /// the `<w:pPr>` child's `<w:pStyle>` element.
    private static func setOrRemoveParagraphStyle(node: XmlNode, styleId: String?) throws {
        // Find or create <w:pPr>.
        let pPr: XmlNode
        if let existing = node.children.first(where: { $0.kind == .element && $0.localName == "pPr" }) {
            pPr = existing
        } else if let styleId = styleId, !styleId.isEmpty {
            pPr = XmlNode.element(prefix: "w", localName: "pPr")
            // <w:pPr> goes first in <w:p>'s children per OOXML schema.
            node.children = [pPr] + node.children
        } else {
            // No existing pPr and styleId is nil/empty — nothing to do.
            return
        }
        // Remove existing <w:pStyle>.
        pPr.children = pPr.children.filter { !($0.kind == .element && $0.localName == "pStyle") }
        // Add new <w:pStyle> if styleId non-nil.
        if let styleId = styleId, !styleId.isEmpty {
            let pStyle = XmlNode.element(prefix: "w", localName: "pStyle")
            pStyle.setAttribute(prefix: "w", localName: "val", value: styleId)
            pPr.children = [pStyle] + pPr.children
        }
    }
}
