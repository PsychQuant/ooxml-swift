import Foundation

/// Where to insert a new paragraph or image inside a `WordDocument`.
///
/// Covers the six anchor kinds used by che-word-mcp tools:
/// - `paragraphIndex` — explicit body-level paragraph index.
/// - `afterImageId` — paragraph following the one that contains the image with
///   the given relationship id (returned by `insertImage`).
/// - `afterTableIndex` — paragraph inserted right after the Nth table in the
///   document body.
/// - `intoTableCell` — paragraph inserted inside the specified table cell.
/// - `afterText(searchText, instance)` — paragraph inserted after the body
///   paragraph whose flattened text contains `searchText`. `instance` is 1-based
///   (1 = first match, 2 = second match, ...) — lets callers disambiguate when
///   the same phrase appears multiple times.
/// - `beforeText(searchText, instance)` — same resolution as `afterText` but
///   the new paragraph is inserted *before* the matching paragraph.
public enum InsertLocation: Equatable {
    case paragraphIndex(Int)
    case afterImageId(String)
    case afterTableIndex(Int)
    case intoTableCell(tableIndex: Int, row: Int, col: Int)
    case afterText(String, instance: Int)
    case beforeText(String, instance: Int)
}

/// Error thrown when an `InsertLocation` cannot be resolved in the target document.
public enum InsertLocationError: Error, Equatable {
    case invalidParagraphIndex(Int)
    case imageIdNotFound(String)
    case tableIndexOutOfRange(Int)
    case tableCellOutOfRange(tableIndex: Int, row: Int, col: Int)
    case textNotFound(searchText: String, instance: Int)
}

// MARK: - Document resolution

extension WordDocument {

    /// Insert an image from a local file path at the given `InsertLocation`.
    ///
    /// Wraps the existing `insertImage(path:widthPx:heightPx:...)` but accepts
    /// the full `InsertLocation` enum, unlocking table-cell insertion and
    /// after-image/after-table anchors.
    ///
    /// - Returns: The relationship id assigned to the image (e.g. `"rId5"`).
    /// - Throws: `ImageReference.from` errors; `InsertLocationError` on anchor
    ///   resolution failure.
    @discardableResult
    public mutating func insertImage(
        path: String,
        widthPx: Int,
        heightPx: Int,
        at location: InsertLocation,
        name: String = "Picture",
        description: String = ""
    ) throws -> String {
        let imageId = nextImageRelationshipId
        let imageRef = try ImageReference.from(path: path, id: imageId)
        images.append(imageRef)

        var drawing = Drawing.from(widthPx: widthPx, heightPx: heightPx, imageId: imageId, name: name)
        drawing.description = description
        let run = Run.withDrawing(drawing)
        let para = Paragraph(runs: [run])

        try insertParagraph(para, at: location)
        // insertParagraph(at: location) marks document.xml. New image bumps media + rels + content_types.
        modifiedParts.insert("word/media/\(imageRef.fileName)")
        modifiedParts.insert("word/_rels/document.xml.rels")
        modifiedParts.insert("[Content_Types].xml")
        return imageId
    }

    /// Insert a paragraph at the given `InsertLocation`. See `InsertLocation`.
    public mutating func insertParagraph(_ paragraph: Paragraph, at location: InsertLocation) throws {
        switch location {
        case .paragraphIndex(let idx):
            guard idx >= 0, idx <= body.children.count else {
                throw InsertLocationError.invalidParagraphIndex(idx)
            }
            body.children.insert(.paragraph(paragraph), at: idx)

        case .afterImageId(let rId):
            guard let bodyIdx = findBodyChildContainingImage(rId: rId) else {
                throw InsertLocationError.imageIdNotFound(rId)
            }
            body.children.insert(.paragraph(paragraph), at: bodyIdx + 1)

        case .afterTableIndex(let tableIdx):
            guard let bodyIdx = findBodyChildAt(tableIndex: tableIdx) else {
                throw InsertLocationError.tableIndexOutOfRange(tableIdx)
            }
            body.children.insert(.paragraph(paragraph), at: bodyIdx + 1)

        case .intoTableCell(let tableIdx, let row, let col):
            guard let bodyIdx = findBodyChildAt(tableIndex: tableIdx),
                  case .table(var table) = body.children[bodyIdx],
                  row >= 0, row < table.rows.count,
                  col >= 0, col < table.rows[row].cells.count
            else {
                throw InsertLocationError.tableCellOutOfRange(tableIndex: tableIdx, row: row, col: col)
            }
            table.rows[row].cells[col].paragraphs.append(paragraph)
            body.children[bodyIdx] = .table(table)

        case .afterText(let searchText, let instance):
            guard let bodyIdx = findBodyChildContainingText(searchText, nthInstance: instance) else {
                throw InsertLocationError.textNotFound(searchText: searchText, instance: instance)
            }
            body.children.insert(.paragraph(paragraph), at: bodyIdx + 1)

        case .beforeText(let searchText, let instance):
            guard let bodyIdx = findBodyChildContainingText(searchText, nthInstance: instance) else {
                throw InsertLocationError.textNotFound(searchText: searchText, instance: instance)
            }
            body.children.insert(.paragraph(paragraph), at: bodyIdx)
        }
        modifiedParts.insert("word/document.xml")
    }

    // MARK: Resolution helpers

    /// Return the index in `body.children` of the paragraph whose runs contain
    /// a drawing with the given relationship id, or `nil` if not found.
    private func findBodyChildContainingImage(rId: String) -> Int? {
        for (i, child) in body.children.enumerated() {
            if case .paragraph(let para) = child {
                for run in para.runs {
                    if run.drawing?.imageId == rId {
                        return i
                    }
                }
            }
        }
        return nil
    }

    /// Return the index in `body.children` of the Nth table (0-based).
    private func findBodyChildAt(tableIndex: Int) -> Int? {
        var seen = 0
        for (i, child) in body.children.enumerated() {
            if case .table = child {
                if seen == tableIndex { return i }
                seen += 1
            }
        }
        return nil
    }

    /// Return the index in `body.children` of the `nthInstance`-th paragraph
    /// whose flattened display text contains `needle`. `nthInstance` is 1-based
    /// (1 = first match). Returns `nil` if fewer than `nthInstance` paragraphs
    /// contain the text.
    ///
    /// PsychQuant/che-word-mcp#63 follow-up (verify F1): pre-fix only inspected
    /// `para.runs`, so `insert_image_from_path(after_text: ...)` against an
    /// anchor wrapped in `<w:sdt>` / `<w:hyperlink>` / `<w:fldSimple>` /
    /// `<mc:AlternateContent>` silently threw `textNotFound`. Now mirrors the
    /// surface coverage of `Document.replaceInParagraphSurfaces` so the LOOKUP
    /// path matches the REPLACE path.
    private func findBodyChildContainingText(_ needle: String, nthInstance: Int) -> Int? {
        guard nthInstance >= 1 else { return nil }
        var seen = 0
        for (i, child) in body.children.enumerated() {
            if case .paragraph(let para) = child {
                if para.flattenedDisplayText().contains(needle) {
                    seen += 1
                    if seen == nthInstance { return i }
                }
            }
        }
        return nil
    }
}

// MARK: - Paragraph display-text extension

extension Paragraph {
    /// Concatenate all displayed text across every editable surface this
    /// paragraph carries, in document order: top-level `runs` + `hyperlinks` +
    /// `fieldSimples` + `alternateContents.fallbackRuns` + `contentControls`
    /// (inline `<w:sdt>` content, walked via `TextReplacementEngine`'s read-only
    /// XML helper). Mirrors the surface coverage of
    /// `Document.replaceInParagraphSurfaces` so reading and writing operate on
    /// the same text universe.
    ///
    /// PsychQuant/che-word-mcp#63 follow-up (verify F1 P1).
    public func flattenedDisplayText() -> String {
        var parts: [String] = []
        parts.append(runs.map { $0.text }.joined())
        for h in hyperlinks {
            parts.append(h.runs.map { $0.text }.joined())
        }
        for f in fieldSimples {
            parts.append(f.runs.map { $0.text }.joined())
        }
        for ac in alternateContents {
            parts.append(ac.fallbackRuns.map { $0.text }.joined())
        }
        for cc in contentControls {
            parts.append(flattenContentControlText(cc))
        }
        return parts.joined()
    }

    /// Recursively walk a ContentControl + its nested children, returning the
    /// concatenated display text of every `<w:t>` descendant inside
    /// `cc.content` plus children. Mirrors `Document.replaceInContentControl`'s
    /// recursion so LOOKUP and REPLACE see identical text.
    private func flattenContentControlText(_ cc: ContentControl) -> String {
        var parts: [String] = []
        if !cc.content.isEmpty {
            parts.append(TextReplacementEngine.flatTextOfContentXML(cc.content))
        }
        for child in cc.children {
            parts.append(flattenContentControlText(child))
        }
        return parts.joined()
    }
}
