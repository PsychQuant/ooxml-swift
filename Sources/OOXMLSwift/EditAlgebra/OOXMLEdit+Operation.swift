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
    /// §3 (this commit): insertParagraph + insertParagraphBefore implemented.
    /// §4-§6 (pending): setBold, insertHyperlink, removeParagraph.
    public func operations() throws -> [Operation] {
        switch self {
        case .insertParagraph(let after, let content, let styleId):
            return [.insertParagraphAfter(
                after: after,
                paragraph: ParagraphPayload(text: content, styleId: styleId)
            )]

        case .insertParagraphBefore(let before, let content, let styleId):
            return [.insertParagraphBefore(
                before: before,
                paragraph: ParagraphPayload(text: content, styleId: styleId)
            )]

        case .setBold(let target, let value):
            return [.setRunFormat(
                target: target,
                format: RunFormatPayload(bold: value)
            )]

        case .removeParagraph(let target):
            return [.removeParagraph(id: target)]

        case .insertHyperlink:
            throw EditError.notImplemented(
                "OOXMLEdit.operations() for \(self) — see #105 tasks §5 (composite design pending user checkpoint)"
            )
        }
    }
}
