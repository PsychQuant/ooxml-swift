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
// Phase 2b apply scope is narrow: only `setText`, `setParagraphStyle`,
// `batchBegin/End`, `unknown` (no-op markers). Other ops throw `.malformedOp`
// with reason `"Phase 2c implements this op"`. The spec tests only exercise
// setText / undo / redo / blame paths, so this is sufficient. Phase 2c
// (`operation-log-setter-wiring-impl`) implements the remaining ops when
// its own tests exercise them.

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
        for (index, entry) in log.entries.enumerated() {
            try applyOrInterpret(entry: entry, entryIndex: index, log: log, to: &working)
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
            var working = base.deepCopy()
            for i in 0..<N {
                try applyOrInterpret(entry: log.entries[i], entryIndex: i, log: log, to: &working)
            }
            return working
        case .timestamp(let cutoff):
            var working = base.deepCopy()
            for (i, entry) in log.entries.enumerated() where entry.timestamp <= cutoff {
                try applyOrInterpret(entry: entry, entryIndex: i, log: log, to: &working)
            }
            return working
        }
    }

    /// Materializes the log as if the entry with `opID == targetOpID` had
    /// never been applied. Intervening entries see the world without the
    /// target's effect (per Decision 4: replay the whole history with the
    /// targeted op replaced).
    public static func undo(_ targetOpID: UUID, log: OperationLog, base: XmlTree) throws -> XmlTree {
        // Verify target exists.
        guard log.entries.contains(where: { $0.opID == targetOpID }) else {
            throw ReducerError.cannotUndo(targetOpID: targetOpID)
        }
        var working = base.deepCopy()
        for (i, entry) in log.entries.enumerated() {
            if entry.opID == targetOpID {
                try applyInverse(entry: entry, entryIndex: i, log: log, to: &working)
            } else {
                try applyOrInterpret(entry: entry, entryIndex: i, log: log, to: &working)
            }
        }
        return working
    }

    /// Materializes the log skipping the `.undo` entry whose `targetOpID`
    /// matches the argument — restores the target's effect.
    public static func redo(_ targetOpID: UUID, log: OperationLog, base: XmlTree) throws -> XmlTree {
        // Find a .undo entry matching the target.
        let undoIndex = log.entries.firstIndex { entry in
            if case .undo(let tgt) = entry.op, tgt == targetOpID { return true }
            return false
        }
        guard undoIndex != nil else {
            throw ReducerError.cannotRedo(targetOpID: targetOpID)
        }
        var working = base.deepCopy()
        for (i, entry) in log.entries.enumerated() {
            if i == undoIndex {
                continue  // skip the .undo entry → original op stays in effect
            }
            try applyOrInterpret(entry: entry, entryIndex: i, log: log, to: &working)
        }
        return working
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
            let targetEntry = log.entries[..<entryIndex].first { $0.opID == targetOpID }
            guard let target = targetEntry else {
                // Target not in prior history — treat as a no-op (forward-compat).
                return
            }
            try applyInverse(entry: target, entryIndex: entryIndex, log: log, to: &tree)
        default:
            try apply(entry: entry, to: &tree)
        }
    }

    /// Applies a single entry's op to the tree. Phase 2b apply scope: only
    /// `setText`, `setParagraphStyle`, `batchBegin/End`, `unknown` (markers
    /// + opaque). Other op kinds throw `.malformedOp` with reason
    /// "Phase 2c implements this op".
    internal static func apply(entry: LogEntry, to tree: inout XmlTree) throws {
        switch entry.op {
        case .setText(let target, let text):
            guard let node = findNode(elementID: target, in: tree) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: target)
            }
            // Replace <w:t> direct children with one new <w:t>X</w:t>.
            let textChild = XmlNode.text(text)
            let wt = XmlNode.element(prefix: "w", localName: "t", children: [textChild])
            // Filter out existing <w:t> children, append the new one.
            // Other children (e.g., <w:rPr>) are preserved.
            let nonTextChildren = node.children.filter {
                !($0.kind == .element && $0.localName == "t")
            }
            node.children = nonTextChildren + [wt]

        case .setParagraphStyle(let target, let styleId):
            guard let node = findNode(elementID: target, in: tree) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: target)
            }
            try setOrRemoveParagraphStyle(node: node, styleId: styleId)

        case .batchBegin, .batchEnd:
            // Markers only — no tree mutation. Phase 2c reducer-level batch
            // semantics (rollback, group undo) are out of scope here.
            return

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
            let newP = makeParagraph(text: payload.text, styleId: payload.styleId)
            // Deterministic ID derivation: new paragraph's libraryUUID == entry.opID.
            // This makes log replay produce the same tree every time.
            newP.libraryUUID = entry.opID
            parent.children.insert(newP, at: idx + 1)

        case .insertParagraphBefore(let before, let payload):
            guard let (parent, idx) = findParentAndIndex(targetID: before, in: tree.root) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: before)
            }
            let newP = makeParagraph(text: payload.text, styleId: payload.styleId)
            newP.libraryUUID = entry.opID
            parent.children.insert(newP, at: idx)

        case .removeParagraph(let target):
            guard let (parent, idx) = findParentAndIndex(targetID: target, in: tree.root) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: target)
            }
            parent.children.remove(at: idx)
            // MVP: removed paragraph's cross-part refs (comments, footnotes,
            // bookmarks pointing TO this paragraph) are left dangling. A
            // follow-up should add orphan-collection — file issue if needed.

        case .setRunFormat(let target, let format):
            // MVP: bold only (sufficient for OOXMLEdit.setBold). Other
            // fields throw malformedOp — add italic/underline/fontSize/color
            // when corresponding OOXMLEdit cases appear.
            if format.italic != nil || format.underline != nil
               || format.fontSizeHalfPoints != nil || format.color != nil {
                throw ReducerError.malformedOp(
                    opID: entry.opID,
                    reason: "Phase 2c MVP supports RunFormatPayload.bold only — italic/underline/fontSizeHalfPoints/color pending"
                )
            }
            guard let node = findNode(elementID: target, in: tree) else {
                throw ReducerError.elementNotFound(opID: entry.opID, elementID: target)
            }
            if let bold = format.bold {
                try setOrRemoveBold(runNode: node, value: bold)
            }
            // bold == nil → no-op (leave unchanged)

        // Phase 2c remaining unsupported op kinds — implemented incrementally.
        case .insertTable, .removeTable, .setCellText,
             .insertRun,
             .insertBookmark, .insertComment,
             .redo,
             .insertNode, .removeNode, .updateAttribute, .moveNode:
            throw ReducerError.malformedOp(
                opID: entry.opID,
                reason: "Phase 2c implements this op"
            )

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
        }
    }

    // MARK: - Phase 2c helpers

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

    /// Sets or removes `<w:b/>` inside a `<w:r>`'s `<w:rPr>` to reflect the
    /// boolean bold value. `value == true` ensures `<w:b/>` is present;
    /// `value == false` removes any existing `<w:b/>` element (relies on
    /// style inheritance for the "not bold" default).
    ///
    /// Creates `<w:rPr>` if needed; removes empty `<w:rPr>` if it becomes
    /// childless after the operation (keeps tree minimal).
    private static func setOrRemoveBold(runNode: XmlNode, value: Bool) throws {
        // Find or create <w:rPr>. Per OOXML schema, rPr is the FIRST child of <w:r>.
        let rPr: XmlNode
        if let existing = runNode.children.first(where: { $0.kind == .element && $0.localName == "rPr" }) {
            rPr = existing
        } else if value {
            rPr = XmlNode.element(prefix: "w", localName: "rPr")
            runNode.children = [rPr] + runNode.children
        } else {
            // value is false and there's no rPr — already not-bold via inheritance.
            return
        }
        // Remove existing <w:b>.
        rPr.children = rPr.children.filter { !($0.kind == .element && $0.localName == "b") }
        // Add <w:b/> if value is true.
        if value {
            let b = XmlNode.element(prefix: "w", localName: "b")
            rPr.children.append(b)
        }
        // If rPr is now empty, remove it (tree minimization).
        if rPr.children.isEmpty {
            runNode.children = runNode.children.filter { $0 !== rPr }
        }
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
        let textNode = XmlNode.text(text)
        let wt = XmlNode.element(prefix: "w", localName: "t", children: [textNode])
        let wr = XmlNode.element(prefix: "w", localName: "r", children: [wt])

        var children: [XmlNode] = []
        if let styleId = styleId, !styleId.isEmpty {
            let pStyle = XmlNode.element(prefix: "w", localName: "pStyle")
            pStyle.setAttribute(prefix: "w", localName: "val", value: styleId)
            let pPr = XmlNode.element(prefix: "w", localName: "pPr", children: [pStyle])
            children.append(pPr)
        }
        children.append(wr)
        return XmlNode.element(prefix: "w", localName: "p", children: children)
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
        return findNode(elementID: elementID, in: tree.root)
    }

    private static func findNode(elementID: ElementID, in node: XmlNode) -> XmlNode? {
        if let nodeID = ElementID(node: node), nodeID == elementID {
            return node
        }
        for child in node.children {
            if let found = findNode(elementID: elementID, in: child) {
                return found
            }
        }
        return nil
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
        case .addRelationship:
            // Rels-part operations don't address an ElementID in any tree —
            // they target a part by path. WordDocument.apply's per-op
            // scoping treats this as a special case: returns nil from
            // partContaining → fallback to direct part lookup by name.
            return []
        case .undo, .redo, .batchBegin, .batchEnd, .unknown: return []
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
        case .undo, .redo, .batchBegin, .batchEnd: return false
        case .insertNode(let parent, _, _): return parent == elementID
        case .removeNode(let target): return target == elementID
        case .updateAttribute(let target, _, _, _): return target == elementID
        case .moveNode(let source, let dest, _): return source == elementID || dest == elementID
        case .insertSiblingAfter(let after, _): return after == elementID
        case .addRelationship: return false  // rels-part operation, not element-addressed
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
