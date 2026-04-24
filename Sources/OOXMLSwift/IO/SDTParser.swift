import Foundation

/// Parser for `<w:sdt>` (Structured Document Tag / Content Control) elements.
///
/// Surfaces SDTs as first-class `ContentControl` values so DocxReader can
/// attach them to `Paragraph.contentControls` (paragraph-level SDTs) and
/// to a future block-level container (Task 3.4).
///
/// Created for change `che-word-mcp-content-controls-read-write` task 3.1.
/// Tasks 3.2 (12-type discrimination), 3.3 (nested children), 3.4 (block-level),
/// and 3.5 (round-trip fidelity) extend this file.
internal enum SDTParser {

    /// Parse a `<w:sdt>` element into a `ContentControl`.
    ///
    /// - Parameters:
    ///   - element: The `<w:sdt>` XML element.
    ///   - parentSdtId: Outer SDT id when this SDT is nested (Task 3.3).
    ///                  `nil` for top-level SDTs.
    /// - Returns: A `ContentControl` with metadata populated from `<w:sdtPr>`
    ///   and `content` set to the verbatim XML string of `<w:sdtContent>`'s
    ///   children. Nested SDTs inside `<w:sdtContent>` populate `children`.
    static func parseSDT(from element: XMLElement, parentSdtId: Int? = nil) -> ContentControl {
        let sdt = parseSdtPr(from: element)

        // <w:sdtContent> — the content region.
        // For paragraph-level SDTs, capture verbatim XML of inner children;
        // nested `<w:sdt>` siblings become ContentControl.children.
        // Block-level callers (DocxReader.parseBodyChildren) skip this path
        // and feed children via the BodyChild.contentControl carrier instead.
        var contentXML = ""
        var children: [ContentControl] = []

        if let sdtContent = element.elements(forName: "w:sdtContent").first {
            for child in sdtContent.children ?? [] {
                guard let childEl = child as? XMLElement else { continue }
                if childEl.localName == "sdt" {
                    children.append(parseSDT(from: childEl, parentSdtId: sdt.id))
                } else {
                    contentXML += childEl.xmlString
                }
            }
        }

        return ContentControl(
            sdt: sdt,
            content: contentXML,
            children: children,
            parentSdtId: parentSdtId
        )
    }

    /// Parse only the `<w:sdtPr>` metadata of a `<w:sdt>` element.
    /// Used for block-level SDTs (Task 3.4) where the content region is
    /// re-walked by DocxReader.parseBodyChildren as BodyChild values.
    static func parseSdtPr(from element: XMLElement) -> StructuredDocumentTag {
        var sdt = StructuredDocumentTag()
        guard let sdtPr = element.elements(forName: "w:sdtPr").first else { return sdt }

        if let idEl = sdtPr.elements(forName: "w:id").first,
           let idStr = idEl.attribute(forName: "w:val")?.stringValue,
           let id = Int(idStr) {
            sdt.id = id
        }
        if let tagEl = sdtPr.elements(forName: "w:tag").first,
           let tagVal = tagEl.attribute(forName: "w:val")?.stringValue {
            sdt.tag = tagVal
        }
        if let aliasEl = sdtPr.elements(forName: "w:alias").first,
           let aliasVal = aliasEl.attribute(forName: "w:val")?.stringValue {
            sdt.alias = aliasVal
        }
        if let lockEl = sdtPr.elements(forName: "w:lock").first,
           let lockVal = lockEl.attribute(forName: "w:val")?.stringValue,
           let lockType = SDTLockType(rawValue: lockVal) {
            sdt.lockType = lockType
        }
        if let placeholderEl = sdtPr.elements(forName: "w:placeholder").first,
           let docPart = placeholderEl.elements(forName: "w:docPart").first,
           let placeholderVal = docPart.attribute(forName: "w:val")?.stringValue {
            sdt.placeholder = placeholderVal
        }
        if sdtPr.elements(forName: "w:temporary").first != nil {
            sdt.isTemporary = true
        }
        sdt.type = detectType(from: sdtPr)
        return sdt
    }

    /// Detect SDT type by inspecting `<w:sdtPr>` children in spec-defined
    /// priority order. Falls back to `.richText` when no marker present.
    ///
    /// Task 3.2: this implements the priority sequence from
    /// `specs/ooxml-read-back-parsers/spec.md` Requirement 2.
    private static func detectType(from sdtPr: XMLElement) -> SDTType {
        // Priority order from spec (text → picture → date → ... → repeatingSectionItem).
        // First match wins; no marker => richText default.
        for child in sdtPr.children ?? [] {
            guard let el = child as? XMLElement, let local = el.localName else { continue }
            switch local {
            case "text": return .plainText
            case "picture": return .picture
            case "date": return .date
            case "dropDownList": return .dropDownList
            case "comboBox": return .comboBox
            case "checkbox": return .checkbox  // w14:checkbox — namespace stripped to localName
            case "bibliography": return .bibliography
            case "citation": return .citation
            case "group": return .group
            case "repeatingSection": return .repeatingSection  // w15:repeatingSection
            case "repeatingSectionItem": return .repeatingSectionItem
            default: continue
            }
        }
        return .richText
    }
}
