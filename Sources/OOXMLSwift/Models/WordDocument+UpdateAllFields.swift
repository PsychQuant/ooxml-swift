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
        for i in 0..<body.children.count {
            if case .paragraph(var para) = body.children[i] {
                if let level = headingLevel(of: para) {
                    currentHeadingCount[level, default: 0] += 1
                }
                processParagraph(&para, counters: &counters, lastResetHeadingCount: &lastResetHeadingCount, currentHeadingCount: currentHeadingCount)
                body.children[i] = .paragraph(para)
            }
        }

        // Headers
        for i in 0..<headers.count {
            for j in 0..<headers[i].paragraphs.count {
                var para = headers[i].paragraphs[j]
                processParagraph(&para, counters: &counters, lastResetHeadingCount: &lastResetHeadingCount, currentHeadingCount: currentHeadingCount)
                headers[i].paragraphs[j] = para
            }
        }

        // Footers
        for i in 0..<footers.count {
            for j in 0..<footers[i].paragraphs.count {
                var para = footers[i].paragraphs[j]
                processParagraph(&para, counters: &counters, lastResetHeadingCount: &lastResetHeadingCount, currentHeadingCount: currentHeadingCount)
                footers[i].paragraphs[j] = para
            }
        }

        // Footnotes
        for i in 0..<footnotes.footnotes.count {
            for j in 0..<footnotes.footnotes[i].paragraphs.count {
                var para = footnotes.footnotes[i].paragraphs[j]
                processParagraph(&para, counters: &counters, lastResetHeadingCount: &lastResetHeadingCount, currentHeadingCount: currentHeadingCount)
                footnotes.footnotes[i].paragraphs[j] = para
            }
        }

        // Endnotes
        for i in 0..<endnotes.endnotes.count {
            for j in 0..<endnotes.endnotes[i].paragraphs.count {
                var para = endnotes.endnotes[i].paragraphs[j]
                processParagraph(&para, counters: &counters, lastResetHeadingCount: &lastResetHeadingCount, currentHeadingCount: currentHeadingCount)
                endnotes.endnotes[i].paragraphs[j] = para
            }
        }

        // updateAllFields rewrites in-place across body + headers/footers/notes;
        // mark every container family dirty so overlay mode re-emits affected parts.
        modifiedParts.insert("word/document.xml")
        for header in headers { modifiedParts.insert("word/\(header.fileName)") }
        for footer in footers { modifiedParts.insert("word/\(footer.fileName)") }
        if !footnotes.footnotes.isEmpty { modifiedParts.insert("word/footnotes.xml") }
        if !endnotes.endnotes.isEmpty { modifiedParts.insert("word/endnotes.xml") }
        return counters
    }

    /// Process one paragraph: find SEQ fields, increment counters (with reset
    /// when the SEQ's resetLevel matches a fresh heading), rewrite cached
    /// result in rawXML.
    private func processParagraph(
        _ para: inout Paragraph,
        counters: inout [String: Int],
        lastResetHeadingCount: inout [String: Int],
        currentHeadingCount: [Int: Int]
    ) {
        let fields = FieldParser.parse(paragraph: para)
        guard !fields.isEmpty else { return }

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
                let newXML = rewriteCachedResult(oldXML, newValue: "\(newValue)", matchingInstrText: field.instrText)
                para.runs[idx].rawXML = newXML
            }
        }
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
    private func rewriteCachedResult(_ rawXML: String, newValue: String, matchingInstrText: String) -> String {
        // Locate the instrText block to anchor our replacement.
        let instrPattern = "<w:instrText[^>]*>\(escapeForRegex(matchingInstrText))</w:instrText>"
        guard let instrRegex = try? NSRegularExpression(pattern: instrPattern, options: [.dotMatchesLineSeparators]),
              let instrMatch = instrRegex.firstMatch(in: rawXML, options: [], range: NSRange(rawXML.startIndex..., in: rawXML)) else {
            return rawXML
        }
        // From end of instrText, find the next <w:t>...</w:t>
        let afterInstr = instrMatch.range.location + instrMatch.range.length
        let searchRange = NSRange(location: afterInstr, length: rawXML.count - afterInstr)
        // Note: toFieldXML wraps each run's fldChar/instrText/t in its own
        // <w:r>...</w:r>, so between the separate-fldChar and the cached <w:t>
        // we have `</w:r><w:r>`. Match accordingly.
        let cachedPattern = #"(<w:fldChar[^/]*fldCharType="separate"[^/]*/>\s*</w:r>\s*<w:r[^>]*>\s*<w:t[^>]*>)[^<]*(</w:t>)"#
        guard let cachedRegex = try? NSRegularExpression(pattern: cachedPattern, options: [.dotMatchesLineSeparators]) else {
            return rawXML
        }
        return cachedRegex.stringByReplacingMatches(
            in: rawXML,
            options: [],
            range: searchRange,
            withTemplate: "$1\(newValue)$2"
        )
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
