// WordDocument+TypedSetters.swift
// word-aligned-state-sync Phase 2 task 3.15 — typed mutations routed through
// the operation log ("Decision 4: Typed APIs as views, not as the model").
//
// Why document-scoped rather than `paragraph.text = ...`: typed views are
// value types holding a shared `XmlNode` reference but NOT a reference to
// the document-owned `OperationLog` (a value type on `WordDocument`). A
// free-standing struct setter therefore cannot append to the log. The
// document-scoped form is the same ownership shape the EditAlgebra
// `apply(_ edit:)` surface established (ooxml-edit-algebra design:
// "WordDocument owns log"). The legacy direct-tree setter
// (`paragraph.text =`) remains as the log-less escape hatch and is slated
// for deprecation in Phase 5 (v1.0 cleanup).

import Foundation

extension WordDocument {

    /// Sets the full text of the paragraph addressed by `id`, routing the
    /// mutation through the operation log:
    ///
    /// 1. `.setText(target:text:)` is appended to `operationLog` with
    ///    `source: .swift` (spec `ooxml-operation-log`, scenario
    ///    "Swift-originated operation").
    /// 2. `OperationReducer` materializes the op into the containing part's
    ///    `XmlTree` — the tree, not the typed model, is the state that
    ///    changes.
    /// 3. The typed body view is resynced from the tree so
    ///    `paragraph.text` reads the new value ("Decision 4" read-back).
    /// 4. `word/document.xml` is marked dirty so the overlay save path
    ///    re-serializes the mutated part.
    ///
    /// Throws when no part contains `id` (wrapped reducer
    /// `elementNotFound`) — a silent no-op would violate the
    /// apply-errors-are-reported requirement.
    public mutating func setParagraphText(id: ElementID, _ text: String) throws {
        try appendAndMaterialize([.setText(target: id, text: text)])
        resyncBodyFromDocumentTree()
    }
}

extension Paragraph {

    /// Stable op-log address of this paragraph, derived from the underlying
    /// tree node (`w14:paraId` → … → library UUID, per the ElementID
    /// derivation rules). `nil` when the paragraph is not tree-backed
    /// (legacy detached mode) or the node carries neither a native OOXML ID
    /// nor an assigned library UUID.
    public var elementID: ElementID? {
        guard let node = xmlNode else { return nil }
        return ElementID(node: node)
    }
}
