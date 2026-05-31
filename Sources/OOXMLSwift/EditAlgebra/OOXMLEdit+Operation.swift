// OOXMLEdit+Operation.swift
// EditAlgebra — Phase 2 of ooxml-edit-isomorphism-foundation (PsychQuant/macdoc#99)
//
// Per design.md Decision 1, each `OOXMLEdit` case maps to one or more
// `Operation` enum cases. This extension is the authoritative location of
// that mapping. §3-§6 of #105 tasks fill it in case-by-case.
//
// Mapping table (canonical, per design.md):
//   OOXMLEdit.insertParagraph(after:content:styleId:) → Operation.insertParagraphAfter
//   OOXMLEdit.insertParagraphBefore(before:content:styleId:) → Operation.insertParagraphBefore
//   OOXMLEdit.setBold(target:value:) → Operation.setRunFormat
//   OOXMLEdit.insertHyperlink(target:href:displayText:) → [Operation.insertNode, Operation.updateAttribute]  (composite atomic)
//   OOXMLEdit.removeParagraph(target:) → Operation.removeParagraph

import Foundation

extension OOXMLEdit {
    /// Emits the `[Operation]` log entries this Edit case translates to.
    ///
    /// 1:1 for simple cases (insertParagraph, setBold, removeParagraph);
    /// composite for `insertHyperlink` (returns 2 Operations atomically).
    ///
    /// Stub in §1 scaffold — per-case implementation lands in §3-§6 of #105
    /// tasks. Throws `EditError.notImplemented` for unimplemented cases.
    public func operations() throws -> [Operation] {
        throw EditError.notImplemented(
            "OOXMLEdit.operations() per-case mapping pending — see #105 tasks §3-§6"
        )
    }
}
