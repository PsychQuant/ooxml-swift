// OperationReducer+AnchorFormat.swift
// PsychQuant/ooxml-swift#234 — `ParagraphPayload.inheritAnchorFormat`.
//
// Builds the paragraph that `insertParagraphAfter` / `insertParagraphBefore`
// insert when the payload asks for the anchor's format. The rules are the ones
// macdoc `word form apply --insert-after` uses (PsychQuant/macdoc#249); the
// public description lives on `OOXMLEdit.insertParagraph`.
//
// Element names are matched by local name, as everywhere else in the reducer.

import Foundation

extension OperationReducer {

    /// Anchor `w:pPr` children a copy drops: one-off structure (numbering, the
    /// section break, page break before) and the paragraph-property revision
    /// record. The section break is handled separately (moved, not copied).
    static let anchorPPrDroppedChildren: Set<String> = [
        "numPr", "sectPr", "pPrChange", "pageBreakBefore",
    ]

    /// Revision marks of a paragraph-mark `w:rPr`. A copy must not carry them:
    /// they would duplicate their `w:id`.
    static let paragraphMarkRevisionMarks: Set<String> = [
        "ins", "del", "moveFrom", "moveTo", "rPrChange",
    ]

    /// Builds the new paragraph from `payload.text` and the anchor's format.
    ///
    /// - `movesSectionBreak`: `true` for an insert after the anchor — an anchor
    ///   that ends a section hands its `w:sectPr` to the new paragraph, so the
    ///   new text stays in the anchor's section. `false` for an insert before.
    ///
    /// Validation happens before the anchor is touched, so a refused op leaves
    /// the tree as it was.
    static func makeParagraphInheritingFormat(
        payload: ParagraphPayload,
        anchor: XmlNode,
        movesSectionBreak: Bool,
        opID: UUID
    ) throws -> XmlNode {
        guard anchor.kind == .element, anchor.localName == "p" else {
            throw ReducerError.malformedOp(
                opID: opID,
                reason: "inheritAnchorFormat needs a <w:p> anchor, got <\(anchor.localName)>")
        }
        let conflicts = explicitParagraphFields(of: payload)
        guard conflicts.isEmpty else {
            throw ReducerError.malformedOp(
                opID: opID,
                reason: "inheritAnchorFormat takes the format from the anchor; "
                    + "the payload also sets \(conflicts.joined(separator: ", "))")
        }

        let anchorPPr = elementChild(of: anchor, named: "pPr")
        var pPrChildren = copiedParagraphProperties(anchorPPr)
        if let styleId = payload.styleId, !styleId.isEmpty {
            pPrChildren.removeAll { $0.localName == "pStyle" }
            let pStyle = XmlNode.element(prefix: "w", localName: "pStyle", namespaceURI: wordNamespace)
            pStyle.setAttribute(prefix: "w", localName: "val", value: styleId)
            pPrChildren.insert(pStyle, at: 0)
        }
        if movesSectionBreak, let anchorPPr,
           let sectIndex = anchorPPr.children.firstIndex(where: {
               $0.kind == .element && $0.localName == "sectPr"
           }) {
            // Moved, not copied: the original node keeps its source bytes.
            // CT_PPr puts sectPr after rPr and before pPrChange; the copy has
            // no pPrChange, so it goes last.
            pPrChildren.append(anchorPPr.children.remove(at: sectIndex))
        }

        let paragraph = makeParagraph(payload: ParagraphPayload(text: payload.text))
        if let run = elementChild(of: paragraph, named: "r"),
           let rPr = firstOwnTextRun(of: anchor).flatMap(copiedRunProperties) {
            run.children.insert(rPr, at: 0)
        }
        if !pPrChildren.isEmpty {
            paragraph.children.insert(
                XmlNode.element(prefix: "w", localName: "pPr", namespaceURI: wordNamespace,
                                children: pPrChildren),
                at: 0)
        }
        return paragraph
    }

    // MARK: - Pieces

    private static let wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    /// Payload fields besides `text`, `styleId` and `paraId`. With
    /// `inheritAnchorFormat` the format comes from the anchor, so any of these
    /// being set is a contradiction: refused, never silently ignored.
    static func explicitParagraphFields(of p: ParagraphPayload) -> [String] {
        let fields: [(String, Bool)] = [
            ("alignment", p.alignment != nil),
            ("spacingBefore", p.spacingBefore != nil),
            ("spacingAfter", p.spacingAfter != nil),
            ("spacingLine", p.spacingLine != nil),
            ("spacingLineRule", p.spacingLineRule != nil),
            ("indentLeft", p.indentLeft != nil),
            ("indentRight", p.indentRight != nil),
            ("indentFirstLine", p.indentFirstLine != nil),
            ("indentHanging", p.indentHanging != nil),
            ("indentFirstLineChars", p.indentFirstLineChars != nil),
            ("indentHangingChars", p.indentHangingChars != nil),
            ("numId", p.numId != nil),
            ("numLevel", p.numLevel != nil),
            ("paragraphMarkRun", p.paragraphMarkRun != nil),
            ("textId", p.textId != nil),
            ("rsidR", p.rsidR != nil),
            ("rsidRPr", p.rsidRPr != nil),
            ("rsidRDefault", p.rsidRDefault != nil),
            ("rsidP", p.rsidP != nil),
        ]
        return fields.filter(\.1).map(\.0)
    }

    /// The anchor pPr's element children as detached copies, without the
    /// dropped children; the paragraph-mark rPr loses its revision marks and
    /// is omitted when nothing is left.
    private static func copiedParagraphProperties(_ pPr: XmlNode?) -> [XmlNode] {
        guard let pPr else { return [] }
        return pPr.children.compactMap { child -> XmlNode? in
            guard child.kind == .element,
                  !anchorPPrDroppedChildren.contains(child.localName) else { return nil }
            let copy = detachedCopy(child)
            guard copy.localName == "rPr" else { return copy }
            copy.children.removeAll {
                $0.kind == .element && paragraphMarkRevisionMarks.contains($0.localName)
            }
            return copy.children.contains(where: { $0.kind == .element }) ? copy : nil
        }
    }

    /// A detached copy of the run's own rPr without `w:rPrChange`; nil when the
    /// run has none or nothing is left.
    private static func copiedRunProperties(_ run: XmlNode) -> XmlNode? {
        guard let rPr = elementChild(of: run, named: "rPr") else { return nil }
        let copy = detachedCopy(rPr)
        copy.children.removeAll { $0.kind == .element && $0.localName == "rPrChange" }
        return copy.children.contains(where: { $0.kind == .element }) ? copy : nil
    }

    /// The first run whose nearest enclosing `w:p` is `paragraph` and whose
    /// direct `w:t` children hold non-empty text, in document order. Nested
    /// paragraphs (text boxes) are skipped; runs inside hyperlinks, fields or
    /// content controls count.
    static func firstOwnTextRun(of paragraph: XmlNode) -> XmlNode? {
        for child in paragraph.children where child.kind == .element {
            switch child.localName {
            case "p":
                continue
            case "r":
                if !directText(of: child).isEmpty { return child }
            default:
                if let found = firstOwnTextRun(of: child) { return found }
            }
        }
        return nil
    }

    private static func directText(of run: XmlNode) -> String {
        run.children
            .filter { $0.kind == .element && $0.localName == "t" }
            .flatMap { $0.children.filter { $0.kind == .text } }
            .map(\.textContent)
            .joined()
    }

    private static func elementChild(of node: XmlNode, named localName: String) -> XmlNode? {
        node.children.first { $0.kind == .element && $0.localName == localName }
    }

    /// `deepClone()` keeps `libraryUUID`; a copy placed next to its source
    /// must not share the source's identity, or an `ElementID` would resolve
    /// to two nodes.
    private static func detachedCopy(_ node: XmlNode) -> XmlNode {
        let copy = node.deepClone()
        func clearIdentity(_ n: XmlNode) {
            n.libraryUUID = nil
            n.children.forEach(clearIdentity)
        }
        clearIdentity(copy)
        return copy
    }
}
