import Foundation

/// Unified part-spanning walker for `WordDocument`.
///
/// Replaces the ad-hoc walkers that previously existed in
/// `Document.handleMixedContentWrapperRevision` (body-only) and the
/// `nextBookmarkId` calibration helper (its own copy of body+headers+footers).
/// Centralizing the recursion eliminates the walker-asymmetry anti-pattern
/// flagged in Issue #56 R4 verify (R4-NEW-2, DA-R1, DA-R3).
///
/// Every callback receives the part key string the paragraph belongs to
/// (e.g., `"word/document.xml"`, `"word/header3.xml"`, `"word/footnotes.xml"`)
/// so callers can mark the correct part dirty in `WordDocument.modifiedParts`
/// instead of blanket-marking `word/document.xml`.
internal enum DocumentWalker {

    // MARK: - Part-key constants

    static let bodyPartKey = "word/document.xml"
    static let footnotesPartKey = "word/footnotes.xml"
    static let endnotesPartKey = "word/endnotes.xml"

    static func headerPartKey(for header: Header) -> String {
        let fileName = header.originalFileName ?? defaultHeaderFileName(for: header.type)
        return "word/\(fileName)"
    }

    static func footerPartKey(for footer: Footer) -> String {
        let fileName = footer.originalFileName ?? defaultFooterFileName(for: footer.type)
        return "word/\(fileName)"
    }

    private static func defaultHeaderFileName(for type: HeaderFooterType) -> String {
        switch type {
        case .first: return "header1.xml"
        case .even: return "header2.xml"
        case .default: return "header3.xml"
        }
    }

    private static func defaultFooterFileName(for type: HeaderFooterType) -> String {
        switch type {
        case .first: return "footer1.xml"
        case .even: return "footer2.xml"
        case .default: return "footer3.xml"
        }
    }

    // MARK: - walkAllParagraphs

    /// Visits every paragraph in the document, in every part, with the part
    /// key the paragraph lives in. Recurses into table cells, nested tables,
    /// and block-level content-control children so callers see the same
    /// universe regardless of where the paragraph is nested.
    static func walkAllParagraphs(in document: WordDocument, visit: (Paragraph, _ partKey: String) -> Void) {
        // Body
        walkBodyChildren(document.body.children, partKey: bodyPartKey, visit: visit)
        // Headers
        for header in document.headers {
            let key = headerPartKey(for: header)
            for para in header.paragraphs { visit(para, key) }
        }
        // Footers
        for footer in document.footers {
            let key = footerPartKey(for: footer)
            for para in footer.paragraphs { visit(para, key) }
        }
        // Footnotes
        for footnote in document.footnotes.footnotes {
            for para in footnote.paragraphs { visit(para, footnotesPartKey) }
        }
        // Endnotes
        for endnote in document.endnotes.endnotes {
            for para in endnote.paragraphs { visit(para, endnotesPartKey) }
        }
    }

    private static func walkBodyChildren(_ children: [BodyChild], partKey: String, visit: (Paragraph, _ partKey: String) -> Void) {
        for child in children {
            switch child {
            case .paragraph(let p):
                visit(p, partKey)
            case .table(let t):
                walkTable(t, partKey: partKey, visit: visit)
            case .contentControl(_, let inner):
                walkBodyChildren(inner, partKey: partKey, visit: visit)
            }
        }
    }

    private static func walkTable(_ table: Table, partKey: String, visit: (Paragraph, _ partKey: String) -> Void) {
        for row in table.rows {
            for cell in row.cells {
                for para in cell.paragraphs { visit(para, partKey) }
                for nested in cell.nestedTables { walkTable(nested, partKey: partKey, visit: visit) }
            }
        }
    }

    // MARK: - findUnrecognizedChild

    /// Locates a `Paragraph.unrecognizedChildren` entry whose opening tag
    /// matches the given element name AND id marker. Returns the paragraph
    /// path so callers can mutate the original via the appropriate accessor.
    ///
    /// Opening-tag-only matching: the substring search is restricted to the
    /// portion before the first `>` of the entry, preventing nested elements
    /// (e.g., `<w:bookmarkStart w:id="50"/>` inside `<w:ins w:id="5">…</w:ins>`)
    /// from false-matching on the outer wrapper's id substring.
    static func findUnrecognizedChild(in document: WordDocument, name: String, idMarker: String) -> (paragraph: Paragraph, indexInParagraph: Int, partKey: String)? {
        var match: (Paragraph, Int, String)?
        walkAllParagraphs(in: document) { para, partKey in
            if match != nil { return }
            for (idx, child) in para.unrecognizedChildren.enumerated() {
                guard child.name == name else { continue }
                guard let openTagEnd = child.rawXML.firstIndex(of: ">") else { continue }
                let openTag = child.rawXML[child.rawXML.startIndex..<openTagEnd]
                if openTag.contains(idMarker) {
                    match = (para, idx, partKey)
                    return
                }
            }
        }
        return match
    }
}
