// WordImport.swift
// word-aligned-state-sync Phase 3 task 4.2 — Word-import diff via element
// identity matching ("Decision 6: Word-import diff via structural
// element-identity matching"; `ooxml-word-sync` Requirement "Word-import
// diff via element identity matching").
//
// This is deliberately NOT a generic tree-edit-distance algorithm (GumTree /
// Zhang-Shriver were considered and rejected in design.md): OOXML paragraphs
// carry explicit identity (`w14:paraId`) for the structural elements that
// matter, and identity-noise (rsids, attribute-order variants) is already
// normalized by `XmlNode.normalizedFingerprint()`.
//
// MVP scope (the spec's normative scenarios): body-level `<w:p>` matching by
// stable ID → structural fingerprint fallback; text change → SetText;
// appeared → InsertParagraphAfter / InsertParagraphBefore; disappeared →
// RemoveParagraph. A matched paragraph whose fingerprint changed but whose
// text is identical (formatting-only edit) has no representable element-level
// op yet — it is surfaced loudly in `unrepresentedChanges` instead of being
// silently swallowed. Representing formatting diffs via the tree-node
// fallback ops is a later Phase 3 iteration.

import Foundation

/// Result of a Word-import diff.
public struct WordImportDiff: Equatable {
    /// Element-level operations reproducing Word's edits when replayed on
    /// the snapshot tree. Ordered: removals, then per-position inserts and
    /// text updates in current-document order.
    public var operations: [Operation]
    /// Matched elements whose content changed in a way the MVP op taxonomy
    /// cannot represent (formatting-only edits). Callers MUST NOT treat an
    /// empty `operations` array as "no change" without checking this.
    public var unrepresentedChanges: [ElementID]

    public init(operations: [Operation] = [], unrepresentedChanges: [ElementID] = []) {
        self.operations = operations
        self.unrepresentedChanges = unrepresentedChanges
    }
}

public enum WordImport {

    /// Diffs the last-synced `snapshot` tree against the freshly re-read
    /// `current` tree (both `word/document.xml`) and infers the operations
    /// Word performed. Pure function — neither input is mutated.
    public static func diff(snapshot: XmlTree, current: XmlTree) -> WordImportDiff {
        let snapParas = bodyParagraphs(of: snapshot)
        let currParas = bodyParagraphs(of: current)

        // Pass 1 — identity matching. Stable IDs first, then structural
        // fingerprints (identity-noise excluded) for ID-less paragraphs,
        // consumed first-unmatched-wins so duplicate-content paragraphs
        // pair positionally.
        var matches: [ObjectIdentifier: XmlNode] = [:]   // current node → snapshot node
        var usedSnapshot = Set<ObjectIdentifier>()

        var snapByStableID: [String: XmlNode] = [:]
        for p in snapParas {
            if let sid = p.stableID, snapByStableID[sid] == nil { snapByStableID[sid] = p }
        }
        for p in currParas {
            if let sid = p.stableID, let snap = snapByStableID[sid],
               !usedSnapshot.contains(ObjectIdentifier(snap)) {
                matches[ObjectIdentifier(p)] = snap
                usedSnapshot.insert(ObjectIdentifier(snap))
            }
        }
        var snapByFingerprint: [String: [XmlNode]] = [:]
        for p in snapParas where p.stableID == nil {
            snapByFingerprint[p.normalizedFingerprint(), default: []].append(p)
        }
        for p in currParas where p.stableID == nil && matches[ObjectIdentifier(p)] == nil {
            let fp = p.normalizedFingerprint()
            if var candidates = snapByFingerprint[fp] {
                while let candidate = candidates.first {
                    candidates.removeFirst()
                    if !usedSnapshot.contains(ObjectIdentifier(candidate)) {
                        matches[ObjectIdentifier(p)] = candidate
                        usedSnapshot.insert(ObjectIdentifier(candidate))
                        break
                    }
                }
                snapByFingerprint[fp] = candidates
            }
        }

        var operations: [Operation] = []
        var unrepresented: [ElementID] = []

        // Pass 2 — removals: snapshot paragraphs never matched by any
        // current paragraph.
        for p in snapParas where !usedSnapshot.contains(ObjectIdentifier(p)) {
            guard let id = ElementID(node: p) else { continue }
            operations.append(.removeParagraph(id: id))
        }

        // Pass 3 — walk current order: appeared paragraphs become inserts
        // anchored on the nearest preceding matched paragraph; matched
        // paragraphs with changed content become SetText (text change) or
        // an unrepresented-change report (formatting-only).
        var lastMatchedID: ElementID?
        for p in currParas {
            if let snap = matches[ObjectIdentifier(p)] {
                if snap.normalizedFingerprint() != p.normalizedFingerprint() {
                    let snapText = concatenatedText(of: snap)
                    let currText = concatenatedText(of: p)
                    if snapText != currText, let id = ElementID(node: p) {
                        operations.append(.setText(target: id, text: currText))
                    } else if let id = ElementID(node: p) {
                        unrepresented.append(id)
                    }
                }
                lastMatchedID = ElementID(node: p)
            } else {
                // Appeared in Word. Anchor after the preceding matched
                // paragraph, or before the first upcoming matched one when
                // Word inserted at the very top.
                let payload = ParagraphPayload(text: concatenatedText(of: p))
                if let anchor = lastMatchedID {
                    operations.append(.insertParagraphAfter(after: anchor, paragraph: payload))
                } else if let nextMatched = currParas.first(where: {
                    matches[ObjectIdentifier($0)] != nil
                }), let beforeID = ElementID(node: nextMatched) {
                    operations.append(.insertParagraphBefore(before: beforeID, paragraph: payload))
                } else if let id = ElementID(node: p) {
                    // Document had no matchable paragraphs at all — surface
                    // rather than guess an anchor.
                    unrepresented.append(id)
                }
            }
        }

        return WordImportDiff(operations: operations, unrepresentedChanges: unrepresented)
    }

    // MARK: - Helpers

    /// Direct `<w:p>` children of `<w:body>` (the MVP diff surface).
    static func bodyParagraphs(of tree: XmlTree) -> [XmlNode] {
        guard let body = tree.root.children.first(where: {
            $0.kind == .element && $0.localName == "body"
        }) else { return [] }
        return body.children.filter { $0.kind == .element && $0.localName == "p" }
    }

    /// Concatenated `<w:t>` text of every descendant, in document order.
    static func concatenatedText(of node: XmlNode) -> String {
        if node.kind == .element && node.localName == "t" {
            return node.children.filter { $0.kind == .text }.map(\.textContent).joined()
        }
        return node.children.map { concatenatedText(of: $0) }.joined()
    }
}
