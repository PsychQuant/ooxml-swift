import Foundation

// PsychQuant/che-word-mcp#276: the rules the text-revision generators
// (`insertTextAsRevision`, `deleteTextAsRevision`, `moveTextAsRevision`)
// share. Offsets are positions in the text the read tools show: inserted
// text counts, deleted and moved-from text does not, an inline equation
// counts as its visible text. Both helpers build a new run list and leave
// the paragraph alone, so every refusal happens before anything changes.

/// Raw children of `<w:r>` that are rendering caches rather than content.
/// They do not stop a run from being split; only the first piece keeps them.
private let renderingHintRunChildren: Set<String> = ["lastRenderedPageBreak"]

/// Stage 2 of PsychQuant/che-word-mcp#276 (nested revisions).
private let nestingIssue = "che-word-mcp#293"

extension Run {
    /// PsychQuant/che-word-mcp#276: the run holds something besides `<w:t>`
    /// text — a drawing, an equation or other raw XML, or raw children such
    /// as tabs, breaks, field characters and note references.
    var hasNonTextContent: Bool {
        drawing != nil || rawXML != nil
            || (rawElements ?? []).contains { !renderingHintRunChildren.contains($0.name) }
    }

    /// The run closes a complex field (`<w:fldChar w:fldCharType="end"/>`).
    var endsComplexField: Bool {
        (rawElements ?? []).contains { $0.name == "fldChar" && $0.xml.contains("\"end\"") }
    }

    /// This run cut to `range` of its text. Pieces after the first drop the
    /// rendering hints, which belong to the run's start.
    func piece(_ range: Range<Int>, isFirst: Bool) -> Run {
        var piece = self
        piece.text = String(text.dropFirst(range.lowerBound).prefix(range.count))
        if !isFirst, let raw = rawElements {
            let kept = raw.filter { !renderingHintRunChildren.contains($0.name) }
            piece.rawElements = kept.isEmpty ? nil : kept
        }
        return piece
    }
}

extension Paragraph {
    /// PsychQuant/che-word-mcp#276: true when `run` is deleted or moved-from
    /// text — absent from the accept-all view the read tools show.
    func isRemovedText(_ run: Run) -> Bool {
        guard let id = run.revisionId,
              let type = revisions.first(where: { $0.id == id })?.type else { return false }
        return type == .deletion || type == .moveFrom
    }

    /// PsychQuant/che-word-mcp#276: the runs of the accept-all view. A
    /// deletion read from a file already carried no text, but one made in
    /// the session still did, so the read tools showed it until a reopen.
    var acceptAllViewRuns: [Run] { runs.filter { !isRemovedText($0) } }

    /// Characters `run` contributes to text-revision offsets — what the read
    /// tools show for it (decision 2: inserted text counts, deleted text
    /// does not).
    func textRevisionLength(of run: Run) -> Int {
        if isRemovedText(run) { return 0 }
        var length = run.text.count
        if let raw = run.rawXML, raw.contains("oMath") {
            length += OMMLParser.parse(xml: raw).visibleText.count
        }
        return length
    }

    /// PsychQuant/che-word-mcp#276: length of the text the text-revision
    /// generators address — the paragraph's own runs in the accept-all view.
    /// Hyperlink, field and content-control text is not part of it.
    public var textRevisionLength: Int {
        runs.reduce(0) { $0 + textRevisionLength(of: $1) }
    }

    /// Why offsets in this paragraph cannot be mapped onto its runs, or nil.
    /// Text the library keeps as raw XML at paragraph level is part of the
    /// paragraph but not of any run, so an offset after it would land in the
    /// wrong place.
    var textRevisionBlocker: String? {
        let names = Set(unrecognizedChildren.map(\.name))
        if !names.isDisjoint(with: ["ins", "del", "moveFrom", "moveTo"]) {
            return "the paragraph holds tracked changes kept as raw XML (around a hyperlink, field or similar), whose text offsets cannot address"
        }
        if !names.isDisjoint(with: ["oMath", "oMathPara"]) {
            return "the paragraph holds an equation outside a run, whose text offsets cannot address"
        }
        return nil
    }

    /// The runs with `[start, end)` marked as revision `revisionId`, and the
    /// marked pieces. Refuses an empty span, a span overlapping a tracked
    /// insertion or move (nesting is stage 2), a span with a zero-length run
    /// strictly inside it (field characters, tabs, references, other
    /// tracked deletions — marking around them would leave them untracked
    /// inside the deletion), and a span touching a run with non-text content.
    func markingSpan(start: Int, end: Int, revisionId: Int) throws -> (runs: [Run], marked: [Run]) {
        if let reason = textRevisionBlocker { throw WordError.invalidParameter("paragraph", reason) }
        guard start >= 0, start <= end else { throw WordError.invalidIndex(start) }
        guard end <= textRevisionLength else { throw WordError.invalidIndex(end) }
        guard start < end else {
            throw WordError.invalidParameter("end", "start == end: a tracked deletion needs a non-empty span")
        }
        let span = "[\(start), \(end))"

        var seen = 0
        for run in runs {
            let length = textRevisionLength(of: run)
            defer { seen += length }
            if length == 0 {
                guard start < seen, seen < end else { continue }
                if run.revisionId != nil {
                    throw WordError.invalidParameter("start", "\(span) contains another tracked change; editing around it is not supported yet (\(nestingIssue))")
                }
                throw WordError.invalidParameter("start", "\(span) contains a field character, tab, break, note reference or other non-text content, which a tracked deletion would leave untracked")
            }
            guard seen < end, start < seen + length else { continue }
            if run.revisionId != nil {
                throw WordError.invalidParameter("start", "\(span) overlaps a tracked insertion or move; editing inside a revision is not supported yet (\(nestingIssue))")
            }
            if run.hasNonTextContent {
                throw WordError.invalidParameter("start", "\(span) touches a run that also holds non-text content (a tab, break, field character, image or equation)")
            }
        }

        var result: [Run] = []
        var marked: [Run] = []
        seen = 0
        for run in runs {
            let length = textRevisionLength(of: run)
            defer { seen += length }
            let lower = max(start - seen, 0), upper = min(end - seen, length)
            guard length > 0, lower < upper else {
                result.append(run)
                continue
            }
            if lower > 0 { result.append(run.piece(0..<lower, isFirst: true)) }
            var middle = run.piece(lower..<upper, isFirst: lower == 0)
            middle.revisionId = revisionId
            result.append(middle)
            marked.append(middle)
            if upper < length { result.append(run.piece(upper..<length, isFirst: false)) }
        }
        return (result, marked)
    }

    /// The runs with `newRuns` placed at `position`. At a boundary the new
    /// runs go after deleted text there and after the end of a field (the end
    /// of a field result is the end of the field). Refuses a position inside
    /// a tracked insertion or move, including a boundary between two runs of
    /// one, and a position that would split a run holding non-text content.
    /// The new runs take their neighbour's emit position, so a paragraph read
    /// from a file writes them in place; appended after the last positioned
    /// child they take a position after it.
    func placing(_ newRuns: [Run], at position: Int) throws -> [Run] {
        if let reason = textRevisionBlocker { throw WordError.invalidParameter("paragraph", reason) }
        guard position >= 0, position <= textRevisionLength else { throw WordError.invalidIndex(position) }
        let inside = "\(position) is inside a tracked insertion or move; editing inside a revision is not supported yet (\(nestingIssue))"

        var slot = runs.count
        var split: (index: Int, offset: Int)?
        var seen = 0
        for (index, run) in runs.enumerated() where !isRemovedText(run) {
            let length = textRevisionLength(of: run)
            if position == seen {
                if length == 0, run.endsComplexField { continue }
                slot = index
                break
            }
            if position < seen + length {
                split = (index, position - seen)
                break
            }
            seen += length
        }

        var result = runs
        var placed = newRuns
        if let split {
            let original = runs[split.index]
            if original.revisionId != nil { throw WordError.invalidParameter("position", inside) }
            if original.hasNonTextContent {
                throw WordError.invalidParameter("position", "\(position) is inside a run that also holds non-text content (a tab, break, field character, image or equation); splitting it would copy that content")
            }
            for i in placed.indices { placed[i].position = original.position }
            let length = original.text.count
            result.replaceSubrange(split.index...split.index, with:
                [original.piece(0..<split.offset, isFirst: true)] + placed
                + [original.piece(split.offset..<length, isFirst: false)])
            return result
        }

        // A boundary between two runs of the same revision is still inside it.
        if slot < runs.count, let next = runs[slot].revisionId,
           let previous = runs[..<slot].last(where: { !isRemovedText($0) && textRevisionLength(of: $0) > 0 }),
           previous.revisionId == next {
            throw WordError.invalidParameter("position", inside)
        }
        let emitPosition: Int?
        if slot < runs.count {
            emitPosition = runs[slot].position
        } else if let last = runs.last, (last.position ?? 0) == 0 {
            emitPosition = last.position   // after an unpositioned run: emitted after the positioned children too
        } else {
            emitPosition = maximumSourcePosition.map { $0 + 1 }
        }
        for i in placed.indices { placed[i].position = emitPosition }
        result.insert(contentsOf: placed, at: slot)
        return result
    }
}
