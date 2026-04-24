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
    @discardableResult
    public mutating func updateAllFields() -> [String: Int] {
        var counters: [String: Int] = [:]
        // heading-reset state: for each level N (1-9), how many times we've
        // seen a heading at that level. When a SEQ with resetLevel=N is
        // encountered, if the recorded heading count differs from the counter
        // set at last reset for that identifier, we reset.
        var lastResetHeadingCount: [String: Int] = [:]  // identifier -> heading count at last reset
        var currentHeadingCount: [Int: Int] = [:]  // heading level -> times seen

        // Body
        var bodyDirty = false
        for i in 0..<body.children.count {
            if case .paragraph(var para) = body.children[i] {
                if let level = headingLevel(of: para) {
                    currentHeadingCount[level, default: 0] += 1
                }
                if processParagraph(&para, counters: &counters, lastResetHeadingCount: &lastResetHeadingCount, currentHeadingCount: currentHeadingCount) {
                    bodyDirty = true
                }
                body.children[i] = .paragraph(para)
            }
        }

        // Headers — per-header dirty bit
        var dirtyHeaderFiles: Set<String> = []
        for i in 0..<headers.count {
            var headerDirty = false
            for j in 0..<headers[i].paragraphs.count {
                var para = headers[i].paragraphs[j]
                if processParagraph(&para, counters: &counters, lastResetHeadingCount: &lastResetHeadingCount, currentHeadingCount: currentHeadingCount) {
                    headerDirty = true
                }
                headers[i].paragraphs[j] = para
            }
            if headerDirty { dirtyHeaderFiles.insert(headers[i].fileName) }
        }

        // Footers — per-footer dirty bit
        var dirtyFooterFiles: Set<String> = []
        for i in 0..<footers.count {
            var footerDirty = false
            for j in 0..<footers[i].paragraphs.count {
                var para = footers[i].paragraphs[j]
                if processParagraph(&para, counters: &counters, lastResetHeadingCount: &lastResetHeadingCount, currentHeadingCount: currentHeadingCount) {
                    footerDirty = true
                }
                footers[i].paragraphs[j] = para
            }
            if footerDirty { dirtyFooterFiles.insert(footers[i].fileName) }
        }

        // Footnotes
        var footnotesDirty = false
        for i in 0..<footnotes.footnotes.count {
            for j in 0..<footnotes.footnotes[i].paragraphs.count {
                var para = footnotes.footnotes[i].paragraphs[j]
                if processParagraph(&para, counters: &counters, lastResetHeadingCount: &lastResetHeadingCount, currentHeadingCount: currentHeadingCount) {
                    footnotesDirty = true
                }
                footnotes.footnotes[i].paragraphs[j] = para
            }
        }

        // Endnotes
        var endnotesDirty = false
        for i in 0..<endnotes.endnotes.count {
            for j in 0..<endnotes.endnotes[i].paragraphs.count {
                var para = endnotes.endnotes[i].paragraphs[j]
                if processParagraph(&para, counters: &counters, lastResetHeadingCount: &lastResetHeadingCount, currentHeadingCount: currentHeadingCount) {
                    endnotesDirty = true
                }
                endnotes.endnotes[i].paragraphs[j] = para
            }
        }

        // Honest dirty-bit propagation — only mark parts that ACTUALLY mutated.
        if bodyDirty { modifiedParts.insert("word/document.xml") }
        for fileName in dirtyHeaderFiles { modifiedParts.insert("word/\(fileName)") }
        for fileName in dirtyFooterFiles { modifiedParts.insert("word/\(fileName)") }
        if footnotesDirty { modifiedParts.insert("word/footnotes.xml") }
        if endnotesDirty { modifiedParts.insert("word/endnotes.xml") }
        return counters
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

            // Rewrite cachedResult run in rawXML
            if let idx = field.cachedResultRunIdx, idx < para.runs.count {
                let oldXML = para.runs[idx].rawXML ?? ""
                let (newXML, didMatch) = rewriteCachedResult(oldXML, newValue: "\(newValue)", matchingInstrText: field.instrText)
                if newXML != oldXML {
                    para.runs[idx].rawXML = newXML
                    rewroteSomething = true
                } else if !didMatch {
                    // v0.13.5+ (#54 sub-finding #5): regex schema drift detection.
                    // FieldParser saw a SEQ with a cached-result run, but our
                    // rewrite regex couldn't locate the cached <w:t>...</w:t>
                    // block. The cached value may now be stale.
                    FileHandle.standardError.write(
                        Data("Warning: updateAllFields could not rewrite cached value for SEQ '\(id)' (instrText: '\(field.instrText)'). Cached value may be stale; potential XML schema drift. (#54)\n".utf8)
                    )
                }
            }
        }
        return rewroteSomething
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
