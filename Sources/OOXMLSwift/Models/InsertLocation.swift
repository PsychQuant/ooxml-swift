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
    /// Inline-mode `WordDocument.insertEquation(at:latex:displayMode:false)` only
    /// supports `.paragraphIndex` anchors per che-word-mcp#67 F2 (semantic
    /// ambiguity for text/image anchors in inline mode). Pre-#91 this rejection
    /// abused `.invalidParagraphIndex(-1)` as a sentinel — structurally lying
    /// because that case is documented for "out-of-range index", not "wrong
    /// anchor kind". This dedicated case lets callers distinguish the two
    /// failures cleanly.
    case inlineModeRequiresParagraphIndex
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
    /// Returns the **top-level** `body.children` index of the BodyChild whose
    /// any descendant paragraph contains `needle` for the `nthInstance`-th
    /// match, or `nil` if not found / inputs invalid.
    ///
    /// **Counting rule (#68)**: each top-level BodyChild that contains the
    /// needle ANYWHERE inside it counts as ONE `nthInstance` (e.g., a table
    /// with the needle in 5 cells = 1 instance, not 5). This mirrors the
    /// pre-#68 behavior where multi-occurrence within ONE top-level
    /// `.paragraph` also counted as 1.
    ///
    /// **Insert-position consequence (#68 — caller `.afterText` / `.beforeText`)**:
    /// returned idx is top-level, so when needle is inside a `.table` /
    /// `.contentControl`, the new paragraph lands AT BODY LEVEL adjacent to
    /// the entire table/SDT — NOT inside the table cell or SDT child list.
    /// This is intentional: `.intoTableCell` exists for inside-cell inserts.
    ///
    /// Public since v0.21.7 (PsychQuant/che-word-mcp#86) — external Swift SPM
    /// consumers (rescue scripts, dxedit CLI, third-party tooling) previously
    /// had to reimplement this with diverging semantics. Exposing the canonical
    /// implementation eliminates that fragmentation.
    public func findBodyChildContainingText(_ needle: String, nthInstance: Int = 1) -> Int? {
        guard nthInstance >= 1, !needle.isEmpty else { return nil }
        var seen = 0
        for (i, child) in body.children.enumerated() {
            if Self.bodyChildContainsText(child, needle: needle) {
                seen += 1
                if seen == nthInstance { return i }
            }
        }
        return nil
    }

    /// PsychQuant/che-word-mcp#68 — recursive descent for text-anchor lookup.
    /// Returns `true` if any paragraph anywhere inside `child` contains `needle`
    /// (via `Paragraph.flattenedDisplayText()`, which already covers inline SDT
    /// per #63). Caller (`findBodyChildContainingText`) treats one matching
    /// top-level BodyChild as ONE `nthInstance` count regardless of how many
    /// nested paragraphs match — consistent with pre-fix top-level paragraph
    /// behavior where multi-occurrence within one paragraph counts once.
    ///
    /// Surfaces walked:
    /// - `.paragraph` — direct text via `flattenedDisplayText()`
    /// - `.table` — `rows[].cells[].paragraphs[]` + `cells[].nestedTables[]`
    ///   (parser depth-limited to 5; recursion depth bounded by parser)
    /// - `.contentControl(_, children:)` — recurse on each child via this helper
    /// - `.bookmarkMarker`, raw catch-all — no text content; returns false
    ///
    /// Public since v0.21.7 (PsychQuant/che-word-mcp#86) — exposed as a primitive
    /// for callers who want to check a single `BodyChild` without the
    /// `nthInstance` enumeration done by `findBodyChildContainingText`.
    public static func bodyChildContainsText(_ child: BodyChild, needle: String) -> Bool {
        switch child {
        case .paragraph(let para):
            return para.flattenedDisplayText().contains(needle)
        case .table(let table):
            return tableContainsText(table, needle: needle)
        case .contentControl(_, let kids):
            return kids.contains { bodyChildContainsText($0, needle: needle) }
        case .bookmarkMarker:
            return false
        case .rawBlockElement:
            return false
        }
    }

    /// Walks every paragraph in every cell of every row, then recurses into
    /// `nestedTables`. Returns true on first match (short-circuit).
    ///
    /// Public since v0.21.7 (PsychQuant/che-word-mcp#86) — exposed alongside
    /// `bodyChildContainsText` so callers building custom traversal can reuse
    /// the depth-bounded table walk.
    public static func tableContainsText(_ table: Table, needle: String) -> Bool {
        for row in table.rows {
            for cell in row.cells {
                for para in cell.paragraphs {
                    if para.flattenedDisplayText().contains(needle) { return true }
                }
                for nested in cell.nestedTables {
                    if tableContainsText(nested, needle: needle) { return true }
                }
            }
        }
        return false
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
    /// PsychQuant/che-word-mcp#85 — include OMML inline math text in top-level runs.
    /// PsychQuant/che-word-mcp#92 — extend OMML walk to hyperlinks /
    /// fieldSimples / alternateContents.fallbackRuns paths via shared
    /// `flattenRunsWithOMML` helper. Pre-#92, those paths used
    /// `runs.map { $0.text }.joined()` and silently dropped OMML inside
    /// their wrappers (same failure class as #85's primary bug, just shifted
    /// to wrappers — rare in practice but real for cross-ref placeholders
    /// emitted by Pandoc/Quarto/LaTeX→docx with embedded math).
    /// contentControls path stays separate (uses RAW XML walking via
    /// `flattenContentControlText` since #63 — different recursion strategy).
    ///
    /// **Cluster fix `flatten-replace-omml-bilateral-coverage`** (closes
    /// che-word-mcp #99 / #100 / #101 / #102 / #103): direct-child OMML
    /// (`<m:oMath>` / `<m:oMathPara>` not wrapped in `<w:r>`) at all 4
    /// wrapper positions — `<w:p>`, `<w:hyperlink>`, `<mc:Fallback>`, and
    /// nested wrapper combinations — is now included in the flatten output.
    /// Source XML position determines emission order (Decision 6).
    ///
    /// **Mirror invariant** (spec capability `ooxml-paragraph-text-mirror`):
    /// this read-side method walks the same wrapper surfaces as
    /// `WordDocument.replaceTextWithBoundaryDetection` (write-side) and both
    /// detect direct-child OMML at the same 4 positions. The two diverge in
    /// how they handle detected OMML — by design, **asymmetric**:
    ///
    /// - **Reads** include each OMML element's `visibleText` so callers can
    ///   locate paragraphs containing math (anchor lookup universe extends
    ///   to OMML).
    /// - **Writes** treat OMML as opaque structural units — replacements
    ///   crossing OMML boundaries refuse with
    ///   `ReplaceResult.refusedDueToOMMLBoundary(occurrences:)`,
    ///   replacements wholly within `<w:t>` ranges proceed normally.
    ///
    /// The asymmetry is principle-driven (spec capability
    /// `ooxml-library-design-principles`): Correctness primacy + Human-like
    /// operations forbid mutating OMML as a side effect of unrelated text
    /// replacement. `<w:delText>` and `<w:instrText>` follow the same opaque
    /// pattern in `Document.replaceInContentControl` for the same reason.
    public func flattenedDisplayText() -> String {
        var parts: [String] = []
        // PsychQuant/che-word-mcp#99 — Pandoc display math:
        // direct-child `<m:oMath>` / `<m:oMathPara>` of `<w:p>` lands in
        // `unrecognizedChildren` (not in any typed wrapper). Interleave with
        // top-level runs by source XML position (Decision 6) so flattened
        // output matches user-visible reading order:
        //   `<w:r>see eq </w:r><m:oMath>δ</m:oMath><w:r> here</w:r>`
        //   → "see eq δ here"  (NOT "see eq  here" with δ appended).
        parts.append(Self.flattenRunsAndDirectChildOMML(
            runs: runs,
            unrecognizedChildren: unrecognizedChildren
        ))
        // PsychQuant/che-word-mcp#100 / #102 — direct-child OMML inside
        // hyperlink wrapper (and nested wrapper combinations like
        // `<w:hyperlink><w:fldSimple>...<m:oMath>...</m:oMath></w:fldSimple></w:hyperlink>`):
        // walk `h.children` (source-of-truth ordered list) instead of just
        // `h.runs` typed projection. Source XML order preserved naturally.
        for h in hyperlinks {
            parts.append(Self.flattenHyperlinkChildren(h.children))
        }
        for f in fieldSimples {
            parts.append(Self.flattenRunsWithOMML(f.runs))
        }
        // PsychQuant/che-word-mcp#101 — direct-child OMML inside `<mc:Fallback>`:
        // typed `ac.fallbackRuns` only captures `<w:r>`-wrapped content.
        // Direct-child `<m:oMath>` lives in `ac.rawXML` Fallback subtree.
        for ac in alternateContents {
            parts.append(Self.flattenRunsWithOMML(ac.fallbackRuns))
            parts.append(Self.extractDirectChildOMMLFromAlternateContentFallback(ac.rawXML))
        }
        for cc in contentControls {
            parts.append(flattenContentControlText(cc))
        }
        return parts.joined()
    }

    /// Walk a `[Run]` list emitting each run's plain `text` plus, for
    /// OMML-bearing runs (whose `rawXML` contains `oMath`), the parsed
    /// math AST's `visibleText` immediately after. Single source of truth
    /// for the OMML walk pattern shared across 4 surface paths in
    /// `flattenedDisplayText` (top-level runs + hyperlinks + fieldSimples +
    /// AC fallbackRuns).
    ///
    /// PsychQuant/che-word-mcp#85 introduced the per-run OMML walk for
    /// top-level runs only. PsychQuant/che-word-mcp#92 extracted it as a
    /// helper so the wrapper paths get identical coverage.
    ///
    /// Reader stores `<m:oMath>` / `<m:oMathPara>` subtrees on `Run.rawXML`
    /// (not as typed children), so the cheap `raw.contains("oMath")`
    /// short-circuit avoids invoking `OMMLParser` on plain runs.
    private static func flattenRunsWithOMML(_ runs: [Run]) -> String {
        var parts: [String] = []
        for run in runs {
            parts.append(run.text)
            if let raw = run.rawXML, raw.contains("oMath") {
                let components = OMMLParser.parse(xml: raw)
                parts.append(components.visibleText)
            }
        }
        return parts.joined()
    }

    /// Position-merged walker for paragraph-level runs interleaved with
    /// direct-child OMML (`<m:oMath>` / `<m:oMathPara>` as direct child of
    /// `<w:p>`, parsed into `Paragraph.unrecognizedChildren`).
    ///
    /// Decision 6 (Spectra change `flatten-replace-omml-bilateral-coverage`):
    /// source XML position determines emission order. A paragraph
    /// `<w:r position=1>see eq </w:r><m:oMath position=2>δ</m:oMath><w:r position=3> here</w:r>`
    /// flattens to `"see eq δ here"`, NOT `"see eq  here" + δ`.
    ///
    /// Fast path: if the paragraph has no direct-child OMML in
    /// `unrecognizedChildren`, falls through to the existing
    /// `flattenRunsWithOMML` (no sort cost on the common case).
    ///
    /// Decision 4 (raw passthrough): direct-child OMML stays in
    /// `unrecognizedChildren`. Walker reads `child.rawXML` and parses with
    /// `OMMLParser` on demand. No mutation, no typed-field promotion.
    private static func flattenRunsAndDirectChildOMML(
        runs: [Run],
        unrecognizedChildren: [UnrecognizedChild]
    ) -> String {
        let directOMath = unrecognizedChildren.filter { child in
            child.name == "oMath" || child.name == "oMathPara"
        }
        if directOMath.isEmpty {
            // Common case — no direct-child OMML; preserve existing semantics
            // (no positional sort, no overhead).
            return flattenRunsWithOMML(runs)
        }

        // Build positional fragments for runs + direct-child OMML, then sort
        // by source position. Runs and OMML children both carry `position`
        // populated by `DocxReader.parseParagraph`'s `childPosition` counter
        // — apples-to-apples comparison.
        enum Fragment {
            case run(Run)
            case oMath(visibleText: String)
        }
        var fragments: [(position: Int, fragment: Fragment)] = []
        for r in runs {
            fragments.append((r.position ?? 0, .run(r)))
        }
        for child in directOMath {
            let visibleText = OMMLParser.parse(xml: child.rawXML).visibleText
            fragments.append((child.position ?? 0, .oMath(visibleText: visibleText)))
        }
        // Stable sort: equal positions preserve relative input order (rare but
        // possible for API-built paragraphs where everything has position 0).
        fragments.sort { $0.position < $1.position }

        var result = ""
        for (_, frag) in fragments {
            switch frag {
            case .run(let r):
                result += r.text
                if let raw = r.rawXML, raw.contains("oMath") {
                    result += OMMLParser.parse(xml: raw).visibleText
                }
            case .oMath(let text):
                result += text
            }
        }
        return result
    }

    /// Walker for `Hyperlink.children` — the source-of-truth ordered list of
    /// hyperlink contents (parallel to legacy `h.runs` projection).
    ///
    /// Closes PsychQuant/che-word-mcp#100 (`<m:oMath>` direct child of
    /// `<w:hyperlink>`) and #102 (nested wrapper —
    /// `<w:hyperlink><w:fldSimple><w:r><m:oMath>...</m:oMath></w:r></w:fldSimple></w:hyperlink>`).
    ///
    /// `.run(r)` cases mirror `flattenRunsWithOMML`'s per-run text + OMML
    /// extraction. `.rawXML(raw)` cases scan the raw subtree for any nested
    /// OMML (whether direct child or wrapped further) — `OMMLParser.parse`
    /// finds `<m:oMath>` / `<m:oMathPara>` anywhere in the input via its
    /// `<m:` prefix scan, so nested wrapper combinations work without
    /// special-casing each shape (covers DA-2 + DA-4 from #92 verify).
    private static func flattenHyperlinkChildren(_ children: [HyperlinkChild]) -> String {
        var result = ""
        for child in children {
            switch child {
            case .run(let r):
                result += r.text
                if let raw = r.rawXML, raw.contains("oMath") {
                    result += OMMLParser.parse(xml: raw).visibleText
                }
            case .rawXML(let raw):
                if raw.contains("oMath") {
                    result += OMMLParser.parse(xml: raw).visibleText
                }
            }
        }
        return result
    }

    /// Extract visible text from `<m:oMath>` / `<m:oMathPara>` elements
    /// appearing as direct children of `<mc:Fallback>` inside an
    /// `<mc:AlternateContent>` raw XML blob.
    ///
    /// Closes PsychQuant/che-word-mcp#101. Typed `ac.fallbackRuns` only
    /// surfaces `<w:r>`-wrapped content from Fallback; this scans the raw
    /// subtree for unwrapped OMML.
    ///
    /// Scope: only Fallback section is scanned. `<mc:Choice>` is the
    /// preferred-rendering branch typically used by Word when it supports
    /// the choice's namespace; including its OMML in flatten would
    /// double-count text the user actually sees. Fallback mirror is
    /// consistent with `ac.fallbackRuns` walk.
    private static func extractDirectChildOMMLFromAlternateContentFallback(_ rawXML: String) -> String {
        guard rawXML.contains("oMath") else { return "" }
        // Carve out just the `<mc:Fallback>...</mc:Fallback>` inner content.
        // Bare-string scan is sufficient because this rawXML is reader-emitted
        // (well-formed) and the AC schema doesn't permit nested
        // `<mc:Fallback>` so backwards search isn't needed.
        guard let openTagStart = rawXML.range(of: "<mc:Fallback") else { return "" }
        guard let openTagEnd = rawXML.range(of: ">", range: openTagStart.upperBound..<rawXML.endIndex) else { return "" }
        guard let closeRange = rawXML.range(of: "</mc:Fallback>", range: openTagEnd.upperBound..<rawXML.endIndex) else { return "" }
        let fallbackInner = String(rawXML[openTagEnd.upperBound..<closeRange.lowerBound])
        // OMMLParser.parseChildren scans for `<m:` and yields all OMML found
        // — visibleText concatenates text from each.
        return OMMLParser.parse(xml: fallbackInner).visibleText
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
