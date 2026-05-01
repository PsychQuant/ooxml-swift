import Foundation

// MARK: - WordDocument.updateAllFields (v0.10.0)
//
// F9-equivalent SEQ counter recomputation across body / headers / footers /
// footnotes / endnotes. Walks every paragraph, runs `FieldParser.parse`, then
// for each `.sequence(SequenceField)` field maintains a per-identifier counter
// (reset at heading boundaries when the SEQ's resetLevel matches the heading
// level). Non-SEQ fields are preserved verbatim.
//
// Return value: dictionary mapping SEQ identifier to final counter value.
// Callers learn which identifiers were updated.

extension WordDocument {

    /// Recompute SEQ field cached results across the entire document.
    /// Non-SEQ fields (IF, DATE, PAGE, REF, etc.) are preserved verbatim.
    ///
    /// **#42 fix (v0.13.4+)**: only marks `modifiedParts` for containers
    /// where a SEQ field was actually rewritten. Previously we unconditionally
    /// inserted every header/footer/footnote/endnote path, which triggered
    /// overlay-mode re-emission via the typed `Header.toXML()` / `Footer.toXML()`
    /// — those serializers only know about typed `paragraphs[]` and silently
    /// strip VML watermarks, drawings, and any non-paragraph raw XML.
    /// Honest dirty-bit propagation prevents that data loss.
    ///
    /// **Counter scope (v0.13.5+ documentation, #54 sub-finding #8)**: SEQ
    /// counters are global across body / headers / footers / footnotes /
    /// endnotes. Body containing 3 `SEQ Figure` followed by header with 1
    /// `SEQ Figure` will give the header `Figure 4`, not `Figure 1`. This
    /// differs from Word F9 which isolates counters per section. Acceptable
    /// for current callers (NTPU thesis workflow uses distinct identifiers
    /// per running header). If section-isolated counters are needed later,
    /// add an `isolatePerContainer: Bool = false` parameter — out of scope
    /// for this version.
    /// v0.14.0+ (#52): `isolatePerContainer` opt-in flag. When `false`
    /// (default), preserves prior global-counter-sharing behavior. When `true`,
    /// each container family (body / each header / each footer / footnotes
    /// collection / endnotes collection) maintains independent SEQ counter
    /// dicts — body's `Figure 3` does NOT increment header's `Figure` counter.
    /// Returned dict reflects body's final counter state (per spec scenario);
    /// per-container final values are reflected in the SEQ runs' rawXML.
    @discardableResult
    public mutating func updateAllFields(isolatePerContainer: Bool = false) -> [String: Int] {
        var counters: [String: Int] = [:]
        // heading-reset state: for each level N (1-9), how many times we've
        // seen a heading at that level. When a SEQ with resetLevel=N is
        // encountered, if the recorded heading count differs from the counter
        // set at last reset for that identifier, we reset.
        var lastResetHeadingCount: [String: Int] = [:]  // identifier -> heading count at last reset
        var currentHeadingCount: [Int: Int] = [:]  // heading level -> times seen

        // Body — recursive walker.
        //
        // v0.21.9+ (PsychQuant/che-word-mcp#94): pre-fix only the `.paragraph`
        // case was processed at body top-level; `.table` and `.contentControl`
        // were silently skipped so SEQ fields anchored inside table cells or
        // block-level SDTs never updated. Now we recurse into both, mirroring
        // the recursion pattern established by `findBodyChildContainingText`
        // (#68, v0.20.6).
        //
        // Heading-count semantics: only top-level direct `.paragraph` body
        // children count toward `currentHeadingCount` (chapter-reset).
        // Headings nested inside tables / SDTs do NOT increment. Rationale:
        // thesis workflows put chapter headings at body top level; SDT/table
        // -internal headings are rare and would create false resets.
        let bodyDirty = walkAndProcessBodyChildrenForFields(
            &body.children,
            counters: &counters,
            lastResetHeadingCount: &lastResetHeadingCount,
            currentHeadingCount: &currentHeadingCount,
            isTopLevel: true
        )

        // Capture body's final counter state for return (per spec contract).
        // Per-container counter snapshots are NOT in the return value;
        // callers inspecting per-container values inspect SEQ runs' rawXML.
        let bodyCounters = counters

        // Headers — per-header dirty bit. Under isolation, each header gets
        // a fresh counter dict (independent of body and other headers).
        var dirtyHeaderFiles: Set<String> = []
        for i in 0..<headers.count {
            var headerCounters: [String: Int] = isolatePerContainer ? [:] : counters
            var headerLastReset: [String: Int] = isolatePerContainer ? [:] : lastResetHeadingCount
            var headerHeadingCount = currentHeadingCount
            let headerDirty = walkAndProcessBodyChildrenForFields(
                &headers[i].bodyChildren,
                counters: &headerCounters,
                lastResetHeadingCount: &headerLastReset,
                currentHeadingCount: &headerHeadingCount,
                isTopLevel: false
            )
            if !isolatePerContainer {
                counters = headerCounters
                lastResetHeadingCount = headerLastReset
            }
            if headerDirty { dirtyHeaderFiles.insert(headers[i].fileName) }
        }

        // Footers — same isolation pattern as headers.
        var dirtyFooterFiles: Set<String> = []
        for i in 0..<footers.count {
            var footerCounters: [String: Int] = isolatePerContainer ? [:] : counters
            var footerLastReset: [String: Int] = isolatePerContainer ? [:] : lastResetHeadingCount
            var footerHeadingCount = currentHeadingCount
            let footerDirty = walkAndProcessBodyChildrenForFields(
                &footers[i].bodyChildren,
                counters: &footerCounters,
                lastResetHeadingCount: &footerLastReset,
                currentHeadingCount: &footerHeadingCount,
                isTopLevel: false
            )
            if !isolatePerContainer {
                counters = footerCounters
                lastResetHeadingCount = footerLastReset
            }
            if footerDirty { dirtyFooterFiles.insert(footers[i].fileName) }
        }

        // Footnotes — single container family, isolated as a unit.
        var footnotesCounters: [String: Int] = isolatePerContainer ? [:] : counters
        var footnotesLastReset: [String: Int] = isolatePerContainer ? [:] : lastResetHeadingCount
        var footnotesDirty = false
        for i in 0..<footnotes.footnotes.count {
            var footnoteHeadingCount = currentHeadingCount
            if walkAndProcessBodyChildrenForFields(
                &footnotes.footnotes[i].bodyChildren,
                counters: &footnotesCounters,
                lastResetHeadingCount: &footnotesLastReset,
                currentHeadingCount: &footnoteHeadingCount,
                isTopLevel: false
            ) {
                footnotesDirty = true
            }
        }
        if !isolatePerContainer {
            counters = footnotesCounters
            lastResetHeadingCount = footnotesLastReset
        }

        // Endnotes — single container family, isolated as a unit.
        var endnotesCounters: [String: Int] = isolatePerContainer ? [:] : counters
        var endnotesLastReset: [String: Int] = isolatePerContainer ? [:] : lastResetHeadingCount
        var endnotesDirty = false
        for i in 0..<endnotes.endnotes.count {
            var endnoteHeadingCount = currentHeadingCount
            if walkAndProcessBodyChildrenForFields(
                &endnotes.endnotes[i].bodyChildren,
                counters: &endnotesCounters,
                lastResetHeadingCount: &endnotesLastReset,
                currentHeadingCount: &endnoteHeadingCount,
                isTopLevel: false
            ) {
                endnotesDirty = true
            }
        }
        if !isolatePerContainer {
            counters = endnotesCounters
            lastResetHeadingCount = endnotesLastReset
        }
        // In isolation mode, return body's counter snapshot (per spec).
        if isolatePerContainer {
            counters = bodyCounters
        }

        // Honest dirty-bit propagation — only mark parts that ACTUALLY mutated.
        if bodyDirty { modifiedParts.insert("word/document.xml") }
        for fileName in dirtyHeaderFiles { modifiedParts.insert("word/\(fileName)") }
        for fileName in dirtyFooterFiles { modifiedParts.insert("word/\(fileName)") }
        if footnotesDirty { modifiedParts.insert("word/footnotes.xml") }
        if endnotesDirty { modifiedParts.insert("word/endnotes.xml") }
        return counters
    }

    /// Recursive walker for body children — processes SEQ fields in
    /// `.paragraph`, recurses into `.table` cells (rows × cells × paragraphs +
    /// nestedTables) and `.contentControl(_, children:)`. Mirrors the
    /// recursion pattern from `findBodyChildContainingText` (#68).
    ///
    /// `isTopLevel` controls whether headings increment `currentHeadingCount`.
    /// Only direct top-level body paragraphs trigger the heading counter to
    /// avoid container-nested headings causing false SEQ chapter-resets (#94).
    ///
    /// - Returns: `true` if any descendant paragraph had a SEQ field rewritten.
    private func walkAndProcessBodyChildrenForFields(
        _ children: inout [BodyChild],
        counters: inout [String: Int],
        lastResetHeadingCount: inout [String: Int],
        currentHeadingCount: inout [Int: Int],
        isTopLevel: Bool
    ) -> Bool {
        var anyDirty = false
        for i in 0..<children.count {
            if walkAndProcessBodyChildForFields(
                &children[i],
                counters: &counters,
                lastResetHeadingCount: &lastResetHeadingCount,
                currentHeadingCount: &currentHeadingCount,
                isTopLevel: isTopLevel
            ) {
                anyDirty = true
            }
        }
        return anyDirty
    }

    /// - Returns: `true` if any descendant paragraph had a SEQ field rewritten.
    private func walkAndProcessBodyChildForFields(
        _ child: inout BodyChild,
        counters: inout [String: Int],
        lastResetHeadingCount: inout [String: Int],
        currentHeadingCount: inout [Int: Int],
        isTopLevel: Bool
    ) -> Bool {
        var anyDirty = false
        switch child {
        case .paragraph(var para):
            if isTopLevel, let level = headingLevel(of: para) {
                currentHeadingCount[level, default: 0] += 1
            }
            if processParagraph(&para,
                                counters: &counters,
                                lastResetHeadingCount: &lastResetHeadingCount,
                                currentHeadingCount: currentHeadingCount) {
                anyDirty = true
            }
            child = .paragraph(para)
        case .table(var table):
            // Walk every cell paragraph + nested tables. All marked
            // non-top-level so heading-count is unaffected.
            for r in 0..<table.rows.count {
                for c in 0..<table.rows[r].cells.count {
                    for p in 0..<table.rows[r].cells[c].paragraphs.count {
                        var para = table.rows[r].cells[c].paragraphs[p]
                        if processParagraph(&para,
                                            counters: &counters,
                                            lastResetHeadingCount: &lastResetHeadingCount,
                                            currentHeadingCount: currentHeadingCount) {
                            anyDirty = true
                        }
                        table.rows[r].cells[c].paragraphs[p] = para
                    }
                    for n in 0..<table.rows[r].cells[c].nestedTables.count {
                        var nestedAsChild: BodyChild = .table(table.rows[r].cells[c].nestedTables[n])
                        if walkAndProcessBodyChildForFields(
                            &nestedAsChild,
                            counters: &counters,
                            lastResetHeadingCount: &lastResetHeadingCount,
                            currentHeadingCount: &currentHeadingCount,
                            isTopLevel: false
                        ) {
                            anyDirty = true
                        }
                        if case .table(let updatedNested) = nestedAsChild {
                            table.rows[r].cells[c].nestedTables[n] = updatedNested
                        }
                    }
                }
            }
            child = .table(table)
        case .contentControl(let cc, var children):
            for i in 0..<children.count {
                if walkAndProcessBodyChildForFields(
                    &children[i],
                    counters: &counters,
                    lastResetHeadingCount: &lastResetHeadingCount,
                    currentHeadingCount: &currentHeadingCount,
                    isTopLevel: false
                ) {
                    anyDirty = true
                }
            }
            child = .contentControl(cc, children: children)
        case .bookmarkMarker, .rawBlockElement:
            break
        }
        return anyDirty
    }

    /// Process one paragraph: find SEQ fields, increment counters (with reset
    /// when the SEQ's resetLevel matches a fresh heading), rewrite cached
    /// result in rawXML.
    ///
    /// - Returns: `true` when at least one SEQ field's cached result was
    ///   rewritten in the paragraph's rawXML; `false` otherwise. Callers use
    ///   this to mark only containers whose content actually changed (#42 fix).
    private func processParagraph(
        _ para: inout Paragraph,
        counters: inout [String: Int],
        lastResetHeadingCount: inout [String: Int],
        currentHeadingCount: [Int: Int]
    ) -> Bool {
        let fields = FieldParser.parse(paragraph: para)
        guard !fields.isEmpty else { return false }

        var rewroteSomething = false
        for field in fields {
            guard case .sequence(let seq) = field.field else { continue }
            let id = seq.identifier

            // Check reset: if resetLevel is set and heading count at that level
            // has advanced since last reset for this identifier, reset counter.
            if let resetLevel = seq.resetLevel {
                let currentCount = currentHeadingCount[resetLevel] ?? 0
                let lastCount = lastResetHeadingCount[id] ?? -1
                if currentCount != lastCount {
                    counters[id] = 0
                    lastResetHeadingCount[id] = currentCount
                }
            }

            // Increment
            counters[id, default: 0] += 1
            let newValue = counters[id]!

            if rewriteCachedResult(for: field, in: &para, identifier: id, newValue: newValue) {
                rewroteSomething = true
            }
        }
        return rewroteSomething
    }

    private func rewriteCachedResult(
        for field: ParsedField,
        in para: inout Paragraph,
        identifier: String,
        newValue: Int
    ) -> Bool {
        let newText = "\(newValue)"
        switch field.location {
        case .paragraphRun:
            guard let idx = field.cachedResultRunIdx, idx < para.runs.count else { return false }
            return rewriteCachedResultRun(
                &para.runs[idx],
                field: field,
                identifier: identifier,
                newText: newText
            )

        case .hyperlinkRun(let hyperlinkIndex):
            guard hyperlinkIndex < para.hyperlinks.count,
                  let idx = field.cachedResultRunIdx,
                  idx < para.hyperlinks[hyperlinkIndex].runs.count else { return false }
            return rewriteCachedResultRun(
                &para.hyperlinks[hyperlinkIndex].runs[idx],
                field: field,
                identifier: identifier,
                newText: newText
            )

        case .fieldSimple(let index):
            guard index < para.fieldSimples.count else { return false }
            if para.fieldSimples[index].runs.isEmpty {
                para.fieldSimples[index].runs = [Run(text: newText)]
                return true
            }
            return rewriteCachedResultRun(
                &para.fieldSimples[index].runs[0],
                field: field,
                identifier: identifier,
                newText: newText
            )

        case .alternateContentFallbackRun(let alternateContentIndex):
            guard alternateContentIndex < para.alternateContents.count,
                  let idx = field.cachedResultRunIdx,
                  idx < para.alternateContents[alternateContentIndex].fallbackRuns.count else { return false }
            let oldAC = para.alternateContents[alternateContentIndex]
            let (newRawXML, didMatch) = rewriteCachedResult(
                oldAC.rawXML,
                newValue: newText,
                matchingInstrText: field.instrText
            )
            guard didMatch, newRawXML != oldAC.rawXML else {
                if !didMatch {
                    emitCachedResultWarning(identifier: identifier, instrText: field.instrText)
                }
                return false
            }
            var updatedRuns = oldAC.fallbackRuns
            _ = rewriteCachedResultRun(
                &updatedRuns[idx],
                field: field,
                identifier: identifier,
                newText: newText
            )
            para.alternateContents[alternateContentIndex] = AlternateContent(
                rawXML: newRawXML,
                fallbackRuns: updatedRuns,
                position: oldAC.position
            )
            return true

        case .contentControl(let index):
            guard index < para.contentControls.count else { return false }
            let oldXML = para.contentControls[index].content
            let (newXML, didMatch) = rewriteCachedResult(
                oldXML,
                newValue: newText,
                matchingInstrText: field.instrText
            )
            if newXML != oldXML {
                para.contentControls[index].content = newXML
                return true
            }
            if !didMatch {
                emitCachedResultWarning(identifier: identifier, instrText: field.instrText)
            }
            return false
        }
    }

    private func rewriteCachedResultRun(
        _ run: inout Run,
        field: ParsedField,
        identifier: String,
        newText: String
    ) -> Bool {
        let isBakedForm = (run.rawXML?.contains("fldChar") ?? false)

        if isBakedForm {
            let oldXML = run.rawXML ?? ""
            let (newXML, didMatch) = rewriteCachedResult(
                oldXML,
                newValue: newText,
                matchingInstrText: field.instrText
            )
            if newXML != oldXML {
                run.rawXML = newXML
                return true
            }
            if !didMatch {
                emitCachedResultWarning(identifier: identifier, instrText: field.instrText)
            }
            return false
        }

        var rewrote = false
        if let rawXML = run.rawXML {
            let (newXML, didMatch) = rewriteCanonicalCachedText(rawXML, newValue: newText)
            if didMatch && newXML != rawXML {
                run.rawXML = newXML
                rewrote = true
            }
        }
        if run.text != newText {
            run.text = newText
            rewrote = true
        }
        return rewrote
    }

    private func emitCachedResultWarning(identifier: String, instrText: String) {
        FileHandle.standardError.write(
            Data("Warning: updateAllFields could not rewrite cached value for SEQ '\(identifier)' (instrText: '\(instrText)'). Cached value may be stale; potential XML schema drift. (#54)\n".utf8)
        )
    }

    /// Returns heading level (1-9) if paragraph has `pStyle == "Heading N"`, else nil.
    private func headingLevel(of para: Paragraph) -> Int? {
        guard let style = para.properties.style else { return nil }
        let prefix = "Heading "
        guard style.hasPrefix(prefix) else { return nil }
        return Int(style.dropFirst(prefix.count))
    }

    /// Replace the `<w:t>OLD</w:t>` value between `<w:fldChar separate>` and
    /// `<w:fldChar end>` within a field block matching the given instrText.
    /// Only rewrites the specific field — other field blocks in the same rawXML
    /// (if any) are untouched.
    ///
    /// v0.13.5+ (#54 sub-finding #5): returns `(rewritten: String, didMatch: Bool)`.
    /// `didMatch == false` means the regex couldn't locate a SEQ field with the
    /// expected XML structure — caller can use this to detect schema drift
    /// vs. legitimate "value didn't change" no-op.
    private func rewriteCachedResult(_ rawXML: String, newValue: String, matchingInstrText: String) -> (String, didMatch: Bool) {
        // Locate the instrText block to anchor our replacement.
        let instrPattern = "<w:instrText[^>]*>\(escapeForRegex(matchingInstrText))</w:instrText>"
        guard let instrRegex = try? NSRegularExpression(pattern: instrPattern, options: [.dotMatchesLineSeparators]),
              let instrMatch = instrRegex.firstMatch(in: rawXML, options: [], range: NSRange(rawXML.startIndex..., in: rawXML)) else {
            return (rawXML, false)
        }
        // From end of instrText, find the next <w:t>...</w:t>
        let afterInstr = instrMatch.range.location + instrMatch.range.length
        let searchRange = NSRange(location: afterInstr, length: rawXML.count - afterInstr)
        // Note: toFieldXML wraps each run's fldChar/instrText/t in its own
        // <w:r>...</w:r>, so between the separate-fldChar and the cached <w:t>
        // we have `</w:r><w:r>`. Match accordingly.
        let cachedPattern = #"(<w:fldChar[^/]*fldCharType="separate"[^/]*/>\s*</w:r>\s*<w:r[^>]*>\s*<w:t[^>]*>)[^<]*(</w:t>)"#
        guard let cachedRegex = try? NSRegularExpression(pattern: cachedPattern, options: [.dotMatchesLineSeparators]) else {
            return (rawXML, false)
        }
        // Count matches so we can report didMatch honestly.
        let cachedMatchCount = cachedRegex.numberOfMatches(in: rawXML, options: [], range: searchRange)
        let rewritten = cachedRegex.stringByReplacingMatches(
            in: rawXML,
            options: [],
            range: searchRange,
            withTemplate: "$1\(newValue)$2"
        )
        return (rewritten, cachedMatchCount > 0)
    }

    /// Splice a new cached value into a canonical-form cached run's `rawXML`.
    ///
    /// The cached run in canonical 5-run form holds `<w:t>...</w:t>` (optionally
    /// with `<w:rPr>...</w:rPr>` siblings) inside its `rawXML`. Mutating only
    /// `Run.text` would silently no-op because `Run.toXML()` short-circuits on
    /// non-nil `rawXML` (Run.swift:246-248). This helper does the minimum splice
    /// to keep `<w:rPr>` and the `<w:t>` open-tag attributes (`xml:space="preserve"`)
    /// intact while replacing the inner text.
    ///
    /// Returns `(rewritten, didMatch)`. `didMatch == false` means no `<w:t>` was
    /// found in the rawXML — caller leaves rawXML untouched and falls back to
    /// the typed-text path.
    private func rewriteCanonicalCachedText(_ rawXML: String, newValue: String) -> (String, didMatch: Bool) {
        let pattern = #"(<w:t(?:\s[^>]*)?>)[^<]*(</w:t>)"#
        guard let regex = try? NSRegularExpression(pattern: pattern, options: [.dotMatchesLineSeparators]) else {
            return (rawXML, false)
        }
        let range = NSRange(rawXML.startIndex..., in: rawXML)
        let matchCount = regex.numberOfMatches(in: rawXML, options: [], range: range)
        if matchCount == 0 {
            return (rawXML, false)
        }
        // newValue for SEQ counters is digits-only (Int → String), but escape
        // defensively in case future field kinds reuse this helper.
        let escaped = newValue
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
        let rewritten = regex.stringByReplacingMatches(
            in: rawXML,
            options: [],
            range: range,
            withTemplate: "$1\(escaped)$2"
        )
        return (rewritten, true)
    }

    private func escapeForRegex(_ s: String) -> String {
        let metachars = CharacterSet(charactersIn: #"\.^$*+?()[]{}|"#)
        return String(s.map { c -> String in
            if metachars.contains(c.unicodeScalars.first!) {
                return "\\\(c)"
            }
            return String(c)
        }.joined())
    }
}
