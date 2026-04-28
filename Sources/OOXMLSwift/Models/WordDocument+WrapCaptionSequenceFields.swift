import Foundation

// MARK: - WordDocument.wrapCaptionSequenceFields (PsychQuant/che-word-mcp#62)

extension WordDocument {

    /// Bulk-convert plain-text caption number portions into SEQ-field-bearing
    /// runs across paragraphs whose joined-runs text matches the supplied
    /// regex. Phase 1 supports `scope: .body` only; `.all` (cross-container)
    /// is reserved for a Phase 1.x patch.
    ///
    /// See `openspec/specs/ooxml-content-insertion-primitives/spec.md`
    /// (Requirement: `Document.wrapCaptionSequenceFields converts plain-text
    /// caption number portions to SEQ-field runs`) for the full contract.
    ///
    /// Idempotency: paragraphs whose runs/fieldSimples already contain a SEQ
    /// field for `sequenceName` are reported in `WrapCaptionResult.skipped`
    /// and never double-wrapped.
    ///
    /// Returned `paragraphsModified` and `SkippedParagraph.paragraphIndex`
    /// values are TOP-LEVEL `body.children` indices. When a matched paragraph
    /// is nested inside a `.table` cell or block-level `.contentControl`, the
    /// reported index points to the containing top-level BodyChild — same
    /// semantic as `findBodyChildContainingText` (#68). One body child can
    /// appear multiple times in `paragraphsModified` if multiple nested
    /// paragraphs inside it matched.
    public mutating func wrapCaptionSequenceFields(
        pattern: NSRegularExpression,
        sequenceName: String,
        format: SequenceField.SequenceFormat = .arabic,
        scope: TextScope = .body,
        insertBookmark: Bool = false,
        bookmarkTemplate: String? = nil
    ) throws -> WrapCaptionResult {
        // Pre-mutation validation — reject malformed callers BEFORE any body mutation.
        guard pattern.numberOfCaptureGroups == 1 else {
            throw WrapCaptionError.patternMissingCaptureGroup(
                actual: pattern.numberOfCaptureGroups
            )
        }
        if insertBookmark {
            guard let template = bookmarkTemplate, template.contains("${number}") else {
                throw WrapCaptionError.bookmarkTemplateMissing
            }
        }
        if scope == .all {
            throw WrapCaptionError.scopeNotImplemented(.all)
        }

        var ctx = WrapContext(
            pattern: pattern,
            sequenceName: sequenceName,
            format: format,
            insertBookmark: insertBookmark,
            bookmarkTemplate: bookmarkTemplate,
            nextBookmarkId: nextBookmarkId
        )

        for i in 0..<body.children.count {
            let updated = walkBodyChild(body.children[i], topLevelBodyIndex: i, ctx: &ctx)
            if let updated = updated {
                body.children[i] = updated
            }
        }

        if !ctx.paragraphsModified.isEmpty {
            modifiedParts.insert("word/document.xml")
            nextBookmarkId = ctx.nextBookmarkId
        }

        return WrapCaptionResult(
            matchedParagraphs: ctx.matchedParagraphs,
            fieldsInserted: ctx.fieldsInserted,
            paragraphsModified: ctx.paragraphsModified,
            skipped: ctx.skipped
        )
    }

    // MARK: - Mutable walker context

    private struct WrapContext {
        let pattern: NSRegularExpression
        let sequenceName: String
        let format: SequenceField.SequenceFormat
        let insertBookmark: Bool
        let bookmarkTemplate: String?
        var nextBookmarkId: Int

        var matchedParagraphs: Int = 0
        var fieldsInserted: Int = 0
        var paragraphsModified: [Int] = []
        var skipped: [SkippedParagraph] = []
    }

    /// Recurse into a BodyChild. Returns the updated BodyChild if mutated, or
    /// `nil` if untouched (caller may keep the original reference).
    private func walkBodyChild(
        _ child: BodyChild,
        topLevelBodyIndex: Int,
        ctx: inout WrapContext
    ) -> BodyChild? {
        switch child {
        case .paragraph(var para):
            if let result = tryWrapParagraph(&para, topLevelBodyIndex: topLevelBodyIndex, ctx: &ctx) {
                return result ? .paragraph(para) : nil
            }
            return nil

        case .table(var table):
            var anyChange = false
            for rowIdx in 0..<table.rows.count {
                for cellIdx in 0..<table.rows[rowIdx].cells.count {
                    for paraIdx in 0..<table.rows[rowIdx].cells[cellIdx].paragraphs.count {
                        var para = table.rows[rowIdx].cells[cellIdx].paragraphs[paraIdx]
                        if let result = tryWrapParagraph(&para, topLevelBodyIndex: topLevelBodyIndex, ctx: &ctx) {
                            if result {
                                table.rows[rowIdx].cells[cellIdx].paragraphs[paraIdx] = para
                                anyChange = true
                            }
                        }
                    }
                    for nestedIdx in 0..<table.rows[rowIdx].cells[cellIdx].nestedTables.count {
                        var nested = table.rows[rowIdx].cells[cellIdx].nestedTables[nestedIdx]
                        if walkNestedTable(&nested, topLevelBodyIndex: topLevelBodyIndex, ctx: &ctx) {
                            table.rows[rowIdx].cells[cellIdx].nestedTables[nestedIdx] = nested
                            anyChange = true
                        }
                    }
                }
            }
            return anyChange ? .table(table) : nil

        case .contentControl(let cc, var kids):
            var anyChange = false
            for i in 0..<kids.count {
                if let updated = walkBodyChild(kids[i], topLevelBodyIndex: topLevelBodyIndex, ctx: &ctx) {
                    kids[i] = updated
                    anyChange = true
                }
            }
            return anyChange ? .contentControl(cc, children: kids) : nil

        case .bookmarkMarker, .rawBlockElement:
            return nil
        }
    }

    private func walkNestedTable(
        _ table: inout Table,
        topLevelBodyIndex: Int,
        ctx: inout WrapContext
    ) -> Bool {
        var anyChange = false
        for rowIdx in 0..<table.rows.count {
            for cellIdx in 0..<table.rows[rowIdx].cells.count {
                for paraIdx in 0..<table.rows[rowIdx].cells[cellIdx].paragraphs.count {
                    var para = table.rows[rowIdx].cells[cellIdx].paragraphs[paraIdx]
                    if let result = tryWrapParagraph(&para, topLevelBodyIndex: topLevelBodyIndex, ctx: &ctx) {
                        if result {
                            table.rows[rowIdx].cells[cellIdx].paragraphs[paraIdx] = para
                            anyChange = true
                        }
                    }
                }
                for nestedIdx in 0..<table.rows[rowIdx].cells[cellIdx].nestedTables.count {
                    var nested = table.rows[rowIdx].cells[cellIdx].nestedTables[nestedIdx]
                    if walkNestedTable(&nested, topLevelBodyIndex: topLevelBodyIndex, ctx: &ctx) {
                        table.rows[rowIdx].cells[cellIdx].nestedTables[nestedIdx] = nested
                        anyChange = true
                    }
                }
            }
        }
        return anyChange
    }

    /// Attempt to wrap one paragraph. Returns:
    /// - `nil` if the paragraph did not match the pattern (no work).
    /// - `false` if matched but skipped (already wraps SEQ; ctx updated).
    /// - `true` if matched and rewritten (paragraph mutated, ctx updated).
    private func tryWrapParagraph(
        _ para: inout Paragraph,
        topLevelBodyIndex: Int,
        ctx: inout WrapContext
    ) -> Bool? {
        // Match against a "rendered" view of the paragraph that inlines any
        // existing SEQ field cachedResult — this lets idempotent re-runs still
        // recognize a caption like "圖 4-1：" even though the digit "1" lives
        // inside `<w:t>1</w:t>` of the SEQ run's rawXML, not in `Run.text`.
        // Captions buried in hyperlinks / contentControls remain out of Phase 1
        // scope (separate enhancement if surfaced).
        let renderedText = renderedTextWithSEQ(para, sequenceName: ctx.sequenceName)
        let nsRendered = renderedText as NSString
        let renderedRange = NSRange(location: 0, length: nsRendered.length)

        guard ctx.pattern.firstMatch(in: renderedText, options: [], range: renderedRange) != nil else {
            return nil
        }
        ctx.matchedParagraphs += 1

        // Idempotency check — match counts toward `matchedParagraphs` but if
        // the paragraph already wraps a SEQ field for this sequence name, we
        // skip without modifying.
        if paragraphAlreadyWrapsSEQ(para, sequenceName: ctx.sequenceName) {
            ctx.skipped.append(SkippedParagraph(
                paragraphIndex: topLevelBodyIndex,
                reason: "already wraps SEQ \(ctx.sequenceName)",
                container: nil
            ))
            return false
        }

        // First-call: plain runs only (no SEQ rawXML present yet), so the
        // joined-runs text matches the rendered text 1:1 and is the right
        // surface for locating the capture range across run boundaries.
        let runText = para.runs.map { $0.text }.joined()
        let nsText = runText as NSString
        let fullRange = NSRange(location: 0, length: nsText.length)

        guard let match = ctx.pattern.firstMatch(in: runText, options: [], range: fullRange) else {
            // Defensive: rendered matched but plain didn't (e.g. caption hidden
            // behind a non-SEQ field rawXML). Leave the paragraph alone.
            return nil
        }

        let captureRange = match.range(at: 1)
        guard captureRange.location != NSNotFound,
              captureRange.length > 0 else {
            return nil
        }
        let captureStart = captureRange.location
        let captureEnd = captureRange.location + captureRange.length
        let capturedNumber = nsText.substring(with: captureRange)

        // Locate which run + local offset holds the capture start and end by
        // walking accumulated NSString lengths (UTF-16 units, matching NSRange).
        var preRunIdx = 0
        var preRunLocalEnd = 0
        var postRunIdx = 0
        var postRunLocalStart = 0
        var cursor = 0
        var foundStart = false
        var foundEnd = false
        for (idx, run) in para.runs.enumerated() {
            let runLen = (run.text as NSString).length
            let runStart = cursor
            let runEnd = cursor + runLen
            if !foundStart, captureStart >= runStart, captureStart <= runEnd {
                preRunIdx = idx
                preRunLocalEnd = captureStart - runStart
                foundStart = true
            }
            if !foundEnd, captureEnd >= runStart, captureEnd <= runEnd {
                postRunIdx = idx
                postRunLocalStart = captureEnd - runStart
                foundEnd = true
            }
            cursor = runEnd
            if foundStart && foundEnd { break }
        }
        guard foundStart, foundEnd else {
            // Should be unreachable because the regex matched against the same
            // joined text — defensive fallback: leave paragraph alone.
            return nil
        }

        // Build the SEQ-field run via the existing FieldCode emission machinery
        // (5-run fldChar block) — same emission style insertCaption uses.
        let seqField = SequenceField(
            identifier: ctx.sequenceName,
            format: ctx.format,
            cachedResult: capturedNumber
        )
        var seqRun = Run(text: "")
        seqRun.rawXML = seqField.toFieldXML()

        // Build replacement runs in document order:
        //   [pre-text portion of preRunIdx (if non-empty),
        //    seqRun,
        //    post-text portion of postRunIdx (if non-empty)]
        // Then keep runs before preRunIdx and runs after postRunIdx untouched.
        var newRuns: [Run] = Array(para.runs[..<preRunIdx])

        // Pre-text portion
        let preRun = para.runs[preRunIdx]
        let preRunNS = preRun.text as NSString
        if preRunLocalEnd > 0 {
            var preText = preRun
            preText.text = preRunNS.substring(to: preRunLocalEnd)
            newRuns.append(preText)
        }

        // Optional bookmark wrap: bookmarkStart immediately before SEQ run,
        // bookmarkEnd immediately after.
        if ctx.insertBookmark, let template = ctx.bookmarkTemplate {
            let bookmarkName = template.replacingOccurrences(of: "${number}", with: capturedNumber)
            let bookmarkId = ctx.nextBookmarkId
            ctx.nextBookmarkId += 1
            var startRun = Run(text: "")
            startRun.rawXML = "<w:bookmarkStart w:id=\"\(bookmarkId)\" w:name=\"\(escapeBookmarkName(bookmarkName))\"/>"
            newRuns.append(startRun)
            newRuns.append(seqRun)
            var endRun = Run(text: "")
            endRun.rawXML = "<w:bookmarkEnd w:id=\"\(bookmarkId)\"/>"
            newRuns.append(endRun)
        } else {
            newRuns.append(seqRun)
        }

        // Post-text portion
        let postRun = para.runs[postRunIdx]
        let postRunNS = postRun.text as NSString
        if postRunLocalStart < postRunNS.length {
            var postText = postRun
            postText.text = postRunNS.substring(from: postRunLocalStart)
            newRuns.append(postText)
        }

        // Tail
        if postRunIdx + 1 < para.runs.count {
            newRuns.append(contentsOf: para.runs[(postRunIdx + 1)...])
        }

        para.runs = newRuns

        ctx.fieldsInserted += 1
        ctx.paragraphsModified.append(topLevelBodyIndex)
        return true
    }

    /// Idempotency check: returns true if the paragraph already contains a SEQ
    /// field for `sequenceName` either as a typed `FieldSimple` or embedded as
    /// rawXML in any of its runs (the emission style insertCaption uses).
    private func paragraphAlreadyWrapsSEQ(_ para: Paragraph, sequenceName: String) -> Bool {
        let needle = "SEQ \(sequenceName)"
        for fs in para.fieldSimples where fs.instr.contains(needle) {
            return true
        }
        for run in para.runs {
            if let raw = run.rawXML, raw.contains(needle) {
                return true
            }
        }
        return false
    }

    /// Reconstruct the user-visible text of a paragraph by inlining the
    /// `cachedResult` of any SEQ field for `sequenceName` carried in run
    /// rawXML. This lets the regex matcher recognize already-wrapped captions
    /// on idempotent re-runs (where digits live in `<w:t>N</w:t>` of the SEQ
    /// fldChar block, not in `Run.text`).
    private func renderedTextWithSEQ(_ para: Paragraph, sequenceName: String) -> String {
        let needle = "SEQ \(sequenceName)"
        var parts: [String] = []
        for run in para.runs {
            if let raw = run.rawXML, raw.contains(needle),
               let cached = extractFieldCachedResult(raw) {
                parts.append(cached)
            } else {
                parts.append(run.text)
            }
        }
        return parts.joined()
    }

    /// Extract the text content between `<w:fldChar w:fldCharType="separate"/>`
    /// and `<w:fldChar w:fldCharType="end"/>` from a SEQ field rawXML — that's
    /// where `cachedResult` lives in the 5-run fldChar block emitted by
    /// `FieldCode.toFieldXML()`.
    private func extractFieldCachedResult(_ rawXML: String) -> String? {
        guard let sepRange = rawXML.range(of: "fldCharType=\"separate\"") else { return nil }
        let after = rawXML[sepRange.upperBound...]
        guard let openRange = after.range(of: "<w:t") else { return nil }
        let afterOpen = after[openRange.upperBound...]
        guard let gtRange = afterOpen.range(of: ">") else { return nil }
        let body = afterOpen[gtRange.upperBound...]
        guard let closeRange = body.range(of: "</w:t>") else { return nil }
        return String(body[..<closeRange.lowerBound])
    }

    private func escapeBookmarkName(_ name: String) -> String {
        return name
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
            .replacingOccurrences(of: "'", with: "&apos;")
    }
}
