// OOXMLEdit+Operation.swift
// EditAlgebra — Phase 2 of ooxml-edit-isomorphism-foundation (PsychQuant/macdoc#99)
//
// Per design.md Decision 1 + macdoc#110 §5 design walkthrough, each
// `OOXMLEdit` case maps to one or more `Operation` enum cases.
//
// Mapping table (canonical):
//   OOXMLEdit.insertParagraph        → Operation.insertParagraphAfter
//   OOXMLEdit.insertParagraphBefore  → Operation.insertParagraphBefore
//   OOXMLEdit.setBold                → Operation.setRunFormat
//   OOXMLEdit.removeParagraph        → Operation.removeParagraph
//   OOXMLEdit.insertHyperlink        → [Operation.insertNode, Operation.addRelationship]
//                                       (composite, pre-validated)
//   OOXMLEdit.wrapWithHyperlink      → [Operation.insertNode (wrapping target),
//                                       Operation.removeNode (target's original position),
//                                       Operation.addRelationship]
//                                       (composite, pre-validated; whole-Run only)

import Foundation

/// Hyperlink relationship type per OOXML spec.
private let hyperlinkRelationshipType =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

/// Default rels part path for document.xml-anchored relationships.
private let documentRelsPath = "word/_rels/document.xml.rels"

extension OOXMLEdit {
    /// Emits the `[Operation]` log entries this Edit case translates to.
    ///
    /// 1:1 for simple cases (insertParagraph, setBold, removeParagraph);
    /// composite for hyperlink cases (returns 2-3 Operations whose
    /// atomicity is enforced upstream by pre-validation in apply).
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

        case .insertHyperlink(let target, let href, let displayText):
            // Insert semantics: new <w:hyperlink> wrapper appended as
            // sibling after target. Allocate a fresh rId for the rels
            // entry. Per Q4 of §5 walkthrough: nil displayText → href.absoluteString.
            //
            // Updated in ooxml-swift#71 Phase 2c: uses the new typed
            // Operation.insertSiblingAfter primitive (was insertNode with
            // position=-1 sentinel, which was semantically incorrect since
            // target is the previous-sibling reference, not the parent).
            let rId = freshRelationshipId()
            let display = displayText ?? href.absoluteString
            let hyperlinkXML = renderHyperlinkXML(rId: rId, displayText: display)
            return [
                .insertSiblingAfter(after: target, nodeXML: hyperlinkXML),
                .addRelationship(
                    part: documentRelsPath,
                    id: rId,
                    type: hyperlinkRelationshipType,
                    target: href.absoluteString,
                    targetMode: "External"
                )
            ]

        case .wrapWithHyperlink(let target, let href):
            // Wrap semantics: <w:hyperlink> takes target's position in
            // parent. Reducer treats insertNode at position -1 with
            // nodeXML containing a placeholder for the original target as
            // "wrap target with this XML". Then removeNode removes the
            // original position. (Two-op formulation; reducer may fuse.)
            //
            // NOTE: target MUST be a <w:r>. Pre-validation in apply
            // verifies this; if it's a paragraph or other element, apply
            // throws EditError.unsupportedOperation.
            let rId = freshRelationshipId()
            let hyperlinkOpenXML = renderHyperlinkOpenWithPlaceholder(rId: rId)
            return [
                .insertNode(parent: target, position: -1, nodeXML: hyperlinkOpenXML),
                .removeNode(target: target),
                .addRelationship(
                    part: documentRelsPath,
                    id: rId,
                    type: hyperlinkRelationshipType,
                    target: href.absoluteString,
                    targetMode: "External"
                )
            ]
        }
    }
}

// MARK: - Hyperlink XML rendering helpers

extension OOXMLEdit {
    /// Generates a fresh relationship ID. Uses a UUID-derived short form
    /// to minimize collision risk with existing rIds (which are typically
    /// `rId1`, `rId2`, etc.). Deterministic rId allocation (re-using the
    /// next free number from the existing rels part) is a Phase 2c
    /// refinement — for MVP, UUID-based IDs avoid the need for doc
    /// context at the operations() step.
    internal static func freshRelationshipId() -> String {
        let raw = UUID().uuidString.replacingOccurrences(of: "-", with: "")
        return "rIdEdit" + String(raw.prefix(8))
    }

    internal func freshRelationshipId() -> String {
        Self.freshRelationshipId()
    }

    /// Renders the full `<w:hyperlink>` XML for `insertHyperlink` (insert
    /// semantics — wrapper contains a new run with displayText).
    internal func renderHyperlinkXML(rId: String, displayText: String) -> String {
        let escaped = escapeXMLContent(displayText)
        return
            "<w:hyperlink r:id=\"\(rId)\">" +
            "<w:r>" +
            "<w:rPr><w:rStyle w:val=\"Hyperlink\"/></w:rPr>" +
            "<w:t>\(escaped)</w:t>" +
            "</w:r>" +
            "</w:hyperlink>"
    }

    /// Renders the `<w:hyperlink>` wrapper for `wrapWithHyperlink` (wrap
    /// semantics — wrapper has a placeholder marker; reducer fills in the
    /// wrapped target). The "$TARGET" placeholder is a sentinel the
    /// Phase 2c Reducer interprets to mean "the node being wrapped".
    internal func renderHyperlinkOpenWithPlaceholder(rId: String) -> String {
        return "<w:hyperlink r:id=\"\(rId)\">$TARGET</w:hyperlink>"
    }

    /// XML-escapes character content (text inside elements). Only handles
    /// `<`, `>`, `&` per the spec for text content.
    internal func escapeXMLContent(_ s: String) -> String {
        return s
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
    }
}
