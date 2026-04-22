import Foundation

/// Where to insert a new paragraph or image inside a `WordDocument`.
///
/// Covers the four anchor kinds used by che-word-mcp tools:
/// - `paragraphIndex` — explicit body-level paragraph index.
/// - `afterImageId` — paragraph following the one that contains the image with
///   the given relationship id (returned by `insertImage`).
/// - `afterTableIndex` — paragraph inserted right after the Nth table in the
///   document body.
/// - `intoTableCell` — paragraph inserted inside the specified table cell.
public enum InsertLocation: Equatable {
    case paragraphIndex(Int)
    case afterImageId(String)
    case afterTableIndex(Int)
    case intoTableCell(tableIndex: Int, row: Int, col: Int)
}

/// Error thrown when an `InsertLocation` cannot be resolved in the target document.
public enum InsertLocationError: Error, Equatable {
    case invalidParagraphIndex(Int)
    case imageIdNotFound(String)
    case tableIndexOutOfRange(Int)
    case tableCellOutOfRange(tableIndex: Int, row: Int, col: Int)
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
        }
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
}
