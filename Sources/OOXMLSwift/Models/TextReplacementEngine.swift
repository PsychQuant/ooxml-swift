import Foundation

// MARK: - ReplaceScope

/// Which text containers a replace operation traverses.
public enum ReplaceScope: Equatable {
    /// Body paragraphs + table-cell paragraphs.
    case bodyAndTables
    /// `bodyAndTables` plus headers, footers, footnotes, endnotes.
    case all
}

// MARK: - ReplaceOptions

public struct ReplaceOptions: Equatable {
    public var scope: ReplaceScope
    public var regex: Bool
    public var matchCase: Bool

    public init(
        scope: ReplaceScope = .bodyAndTables,
        regex: Bool = false,
        matchCase: Bool = true
    ) {
        self.scope = scope
        self.regex = regex
        self.matchCase = matchCase
    }
}

// MARK: - ReplaceError

public enum ReplaceError: Error, Equatable {
    case invalidRegex(String)
}

// MARK: - TextReplacementEngine

/// Cross-run-safe text replacement for a `Paragraph.runs` array.
///
/// Core algorithm: flatten runs into a single string + offset map, find matches
/// on the flat string, splice replacements back preserving run boundaries.
/// The replacement text inherits the *starting run's* formatting.
///
/// Runs whose `rawXML` or `drawing` is non-nil (field runs, images) are
/// excluded from flattening — matches will not span across them.
public enum TextReplacementEngine {

    // MARK: Flatten

    /// Flatten text runs into a single string plus an offset map.
    /// `map[i]` is the `(runIdx, charOffsetInRun)` pair for the i-th character
    /// of `flat` (where "character" means grapheme cluster, matching Swift's
    /// String default element).
    public static func flattenRuns(_ runs: [Run]) -> (flat: String, map: [(runIdx: Int, offset: Int)]) {
        var flat = ""
        var map: [(runIdx: Int, offset: Int)] = []
        for (runIdx, run) in runs.enumerated() {
            guard isTextRun(run) else { continue }
            for (charIdx, char) in run.text.enumerated() {
                map.append((runIdx, charIdx))
                flat.append(char)
            }
        }
        return (flat, map)
    }

    // MARK: Replace

    /// Replace occurrences of `find` with `replacement` in `runs`. Returns the
    /// number of replacements made. Replacement text inherits the start run's
    /// formatting. When `options.regex == true`, `replacement` may contain
    /// `$1`, `$2`, etc. capture-group backreferences.
    @discardableResult
    public static func replace(
        runs: inout [Run],
        find: String,
        with replacement: String,
        options: ReplaceOptions = ReplaceOptions()
    ) throws -> Int {
        let (flat, map) = flattenRuns(runs)
        guard !flat.isEmpty, !find.isEmpty else { return 0 }

        if options.regex {
            let regexOptions: NSRegularExpression.Options = options.matchCase ? [] : [.caseInsensitive]
            let re: NSRegularExpression
            do {
                re = try NSRegularExpression(pattern: find, options: regexOptions)
            } catch {
                throw ReplaceError.invalidRegex(find)
            }
            let fullNSRange = NSRange(flat.startIndex..., in: flat)
            let matches = re.matches(in: flat, options: [], range: fullNSRange)
            for match in matches.reversed() {
                guard let swiftRange = Range(match.range, in: flat) else { continue }
                // NSRegularExpression.replacementString expands $1..$N using the match's groups.
                let expanded = re.replacementString(for: match, in: flat, offset: 0, template: replacement)
                applyOneReplacement(runs: &runs, flat: flat, map: map, matchRange: swiftRange, replacement: expanded)
            }
            return matches.count
        } else {
            let cmpOptions: String.CompareOptions = options.matchCase ? [] : [.caseInsensitive]
            var ranges: [Range<String.Index>] = []
            var searchStart = flat.startIndex
            while searchStart < flat.endIndex,
                  let range = flat.range(of: find, options: cmpOptions, range: searchStart..<flat.endIndex) {
                ranges.append(range)
                searchStart = range.upperBound
            }
            for range in ranges.reversed() {
                applyOneReplacement(runs: &runs, flat: flat, map: map, matchRange: range, replacement: replacement)
            }
            return ranges.count
        }
    }

    // MARK: Replace inside ContentControl.content (inline <w:sdt>)

    /// Result of `replaceInContentXML`: the rewritten XML fragment and the
    /// number of replacements applied.
    public struct ContentXMLReplaceResult: Equatable {
        public var xml: String
        public var replacements: Int
    }

    /// Replace `find` with `replacement` in an inline-XML fragment that holds
    /// `<w:r><w:t>...</w:t></w:r>` (possibly mixed with `<w:hyperlink>`,
    /// `<w:fldSimple>`, etc) — i.e. the verbatim `content` blob of an inline
    /// `<w:sdt>` content control parsed by `SDTParser.parseSDT`.
    ///
    /// Algorithm mirrors `replace(runs:find:with:options:)`:
    /// 1. Wrap `contentXML` in a synthetic root with `xmlns:w` so Foundation
    ///    `XMLDocument` can parse mixed `<w:r>`/`<w:hyperlink>` etc fragments.
    /// 2. Walk all `<w:t>` text-element descendants in document order, building
    ///    a flat string + offset map (`(elementIdx, charOffset)`).
    ///    Skip `<w:delText>` (TC deletion text — not displayed),
    ///    `<w:instrText>` (field instruction code — not display),
    ///    and any `<w:t>` nested inside a child `<w:sdt>` subtree (those are
    ///    represented as typed `ContentControl.children` by the caller and
    ///    handled via outer recursion — visiting them here would duplicate).
    /// 3. Find matches on flat string with the same regex/literal/case rules
    ///    as the run-level engine.
    /// 4. For each match (reversed to keep earlier offsets valid), splice the
    ///    replacement into the affected `<w:t>` element strings.
    /// 5. Re-serialize the wrapper's children, omitting the wrapper tag.
    ///
    /// Round-trip discipline: only `<w:t>` element string content is mutated.
    /// `xml:space="preserve"` and any other attributes survive verbatim because
    /// we never touch the elements' attribute set.
    public static func replaceInContentXML(
        _ contentXML: String,
        find: String,
        with replacement: String,
        options: ReplaceOptions = ReplaceOptions()
    ) throws -> ContentXMLReplaceResult {
        guard !find.isEmpty, !contentXML.isEmpty else {
            return ContentXMLReplaceResult(xml: contentXML, replacements: 0)
        }

        // Wrap the fragment so XMLDocument has a single root + namespace decl.
        let wrapped = "<__sdtcontent xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            + contentXML + "</__sdtcontent>"
        let xmlDoc: XMLDocument
        do {
            xmlDoc = try XMLDocument(xmlString: wrapped, options: [.nodePreserveAll])
        } catch {
            // Malformed inner XML — bail out without modification rather than
            // throw, so a broken SDT doesn't break the whole replaceText call.
            return ContentXMLReplaceResult(xml: contentXML, replacements: 0)
        }
        guard let root = xmlDoc.rootElement() else {
            return ContentXMLReplaceResult(xml: contentXML, replacements: 0)
        }

        // Collect flattenable <w:t> elements in document order.
        var textElements: [XMLElement] = []
        collectFlattenableTextElements(in: root, into: &textElements)
        guard !textElements.isEmpty else {
            return ContentXMLReplaceResult(xml: contentXML, replacements: 0)
        }

        // Build flat string + offset map.
        var flat = ""
        var map: [(elemIdx: Int, offset: Int)] = []
        for (idx, el) in textElements.enumerated() {
            let text = el.stringValue ?? ""
            for (charIdx, char) in text.enumerated() {
                map.append((idx, charIdx))
                flat.append(char)
            }
        }
        guard !flat.isEmpty else {
            return ContentXMLReplaceResult(xml: contentXML, replacements: 0)
        }

        // Find matches using same rules as the run-level engine.
        var ranges: [Range<String.Index>] = []
        if options.regex {
            let regexOptions: NSRegularExpression.Options = options.matchCase ? [] : [.caseInsensitive]
            let re: NSRegularExpression
            do {
                re = try NSRegularExpression(pattern: find, options: regexOptions)
            } catch {
                throw ReplaceError.invalidRegex(find)
            }
            let fullNSRange = NSRange(flat.startIndex..., in: flat)
            let matches = re.matches(in: flat, options: [], range: fullNSRange)
            // Note: regex with $1..$N expansion needs the match's groups.
            // We splice using the engine's `replacementString(for:in:offset:template:)`.
            for match in matches.reversed() {
                guard let swiftRange = Range(match.range, in: flat) else { continue }
                let expanded = re.replacementString(
                    for: match, in: flat, offset: 0, template: replacement
                )
                applyOneXMLReplacement(
                    textElements: textElements, flat: flat, map: map,
                    matchRange: swiftRange, replacement: expanded
                )
            }
            ranges = matches.compactMap { Range($0.range, in: flat) }
        } else {
            let cmpOptions: String.CompareOptions = options.matchCase ? [] : [.caseInsensitive]
            var searchStart = flat.startIndex
            while searchStart < flat.endIndex,
                  let range = flat.range(of: find, options: cmpOptions, range: searchStart..<flat.endIndex) {
                ranges.append(range)
                searchStart = range.upperBound
            }
            for range in ranges.reversed() {
                applyOneXMLReplacement(
                    textElements: textElements, flat: flat, map: map,
                    matchRange: range, replacement: replacement
                )
            }
        }

        guard !ranges.isEmpty else {
            return ContentXMLReplaceResult(xml: contentXML, replacements: 0)
        }

        // Re-serialize the wrapper's children, dropping the wrapper element.
        var rebuilt = ""
        for child in root.children ?? [] {
            rebuilt += child.xmlString
        }
        return ContentXMLReplaceResult(xml: rebuilt, replacements: ranges.count)
    }

    /// Read-only flat text of an inline `<w:sdt>` content blob — `<w:t>`
    /// descendants joined in document order, mirroring the same flattening
    /// rules as `replaceInContentXML` (skips `<w:delText>` / `<w:instrText>` /
    /// nested `<w:sdt>` subtrees). Used by `findBodyChildContainingText` so
    /// the LOOKUP path matches the REPLACE path's surface coverage.
    /// PsychQuant/che-word-mcp#63 follow-up (verify F1).
    public static func flatTextOfContentXML(_ contentXML: String) -> String {
        guard !contentXML.isEmpty, contentXML.contains("<w:t") else { return "" }
        let wrapped = "<__sdtcontent xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            + contentXML + "</__sdtcontent>"
        guard let xmlDoc = try? XMLDocument(xmlString: wrapped, options: [.nodePreserveAll]),
              let root = xmlDoc.rootElement() else { return "" }
        var elements: [XMLElement] = []
        collectFlattenableTextElements(in: root, into: &elements)
        return elements.compactMap { $0.stringValue }.joined()
    }

    /// Recursively gather `<w:t>` descendants that count as flattenable text,
    /// preserving document order. Skips `<w:delText>` / `<w:instrText>` and
    /// any subtree rooted at a nested `<w:sdt>` (handled by outer recursion).
    private static func collectFlattenableTextElements(
        in element: XMLElement,
        into out: inout [XMLElement]
    ) {
        for child in element.children ?? [] {
            guard let el = child as? XMLElement else { continue }
            let local = el.localName ?? ""
            if local == "sdt" {
                // Nested SDT — its `<w:t>` descendants are owned by the inner
                // ContentControl and visited by outer `replaceInContentControl`
                // recursion. Skipping prevents double-replacement.
                continue
            }
            if local == "t" {
                out.append(el)
                continue
            }
            // delText / instrText carry non-displayed text — never touch them.
            if local == "delText" || local == "instrText" {
                continue
            }
            collectFlattenableTextElements(in: el, into: &out)
        }
    }

    /// Splice one match into the underlying `<w:t>` element string contents.
    /// Mirrors the multi-element path of `applyOneReplacement` but operates on
    /// XML elements: the start element receives `prefix + replacement`, the
    /// end element gets trimmed to its suffix, and any wholly-consumed text
    /// elements between them are emptied (not removed — removing the element
    /// would bring `<w:r>` parents out of sync). Single-element matches collapse
    /// before+replacement+after into the start element.
    private static func applyOneXMLReplacement(
        textElements: [XMLElement],
        flat: String,
        map: [(elemIdx: Int, offset: Int)],
        matchRange: Range<String.Index>,
        replacement: String
    ) {
        let startCharIdx = flat.distance(from: flat.startIndex, to: matchRange.lowerBound)
        let endCharIdx = flat.distance(from: flat.startIndex, to: matchRange.upperBound)

        guard startCharIdx < map.count else { return }
        let (sIdx, sOffset) = map[startCharIdx]

        let eIdx: Int
        let eOffsetExclusive: Int
        if endCharIdx < map.count {
            let (i, o) = map[endCharIdx]
            eIdx = i
            eOffsetExclusive = o
        } else {
            eIdx = textElements.count - 1
            eOffsetExclusive = (textElements[eIdx].stringValue ?? "").count
        }

        if sIdx == eIdx {
            let text = textElements[sIdx].stringValue ?? ""
            let before = String(text.prefix(sOffset))
            let after = String(text.suffix(text.count - eOffsetExclusive))
            textElements[sIdx].stringValue = before + replacement + after
        } else {
            let startText = textElements[sIdx].stringValue ?? ""
            let beforePrefix = String(startText.prefix(sOffset))
            textElements[sIdx].stringValue = beforePrefix + replacement
            let endText = textElements[eIdx].stringValue ?? ""
            let afterSuffix = String(endText.suffix(endText.count - eOffsetExclusive))
            textElements[eIdx].stringValue = afterSuffix
            // Empty out wholly-consumed `<w:t>` elements between start and end.
            // We don't remove the element — its `<w:r>` parent (and any
            // intervening structural wrappers) survives so source-order is
            // preserved.
            if eIdx - sIdx > 1 {
                for midIdx in (sIdx + 1)..<eIdx {
                    textElements[midIdx].stringValue = ""
                }
            }
        }
    }

    // MARK: Private

    private static func isTextRun(_ run: Run) -> Bool {
        return run.rawXML == nil && run.drawing == nil
    }

    private static func applyOneReplacement(
        runs: inout [Run],
        flat: String,
        map: [(runIdx: Int, offset: Int)],
        matchRange: Range<String.Index>,
        replacement: String
    ) {
        let startCharIdx = flat.distance(from: flat.startIndex, to: matchRange.lowerBound)
        let endCharIdx = flat.distance(from: flat.startIndex, to: matchRange.upperBound)

        guard startCharIdx < map.count else { return }
        let (sRunIdx, sOffset) = map[startCharIdx]

        let eRunIdx: Int
        let eOffsetExclusive: Int
        if endCharIdx < map.count {
            let (r, o) = map[endCharIdx]
            eRunIdx = r
            eOffsetExclusive = o
        } else {
            // Match extends to the last character of flat — end run is the last
            // text run, end offset is its full length.
            let lastTextRunIdx = runs.lastIndex(where: isTextRun) ?? sRunIdx
            eRunIdx = lastTextRunIdx
            eOffsetExclusive = runs[lastTextRunIdx].text.count
        }

        if sRunIdx == eRunIdx {
            // Match entirely within a single run.
            let text = runs[sRunIdx].text
            let before = String(text.prefix(sOffset))
            let after = String(text.suffix(text.count - eOffsetExclusive))
            runs[sRunIdx].text = before + replacement + after
        } else {
            // Multi-run match.
            // 1. Trim start run down to prefix + replacement (inheriting its props).
            let startText = runs[sRunIdx].text
            let beforePrefix = String(startText.prefix(sOffset))
            runs[sRunIdx].text = beforePrefix + replacement
            // 2. Trim end run down to just its suffix.
            let endText = runs[eRunIdx].text
            let afterSuffix = String(endText.suffix(endText.count - eOffsetExclusive))
            runs[eRunIdx].text = afterSuffix
            // 3. Remove TEXT runs strictly between start and end. Preserve
            //    non-text runs (field runs with rawXML, drawing runs) that
            //    may sit between — they contributed zero characters to the
            //    flat string, so removing them would silently drop structure.
            //
            //    v0.27.0 (PsychQuant/ooxml-swift#65,
            //    kiki830621/collaboration_guo_analysis#20): also preserve
            //    runs whose `rawElements` is non-empty. A Run shaped like
            //    `<w:r><w:rPr>…</w:rPr><w:commentReference w:id="23"/></w:r>`
            //    parses with `text == ""`, `rawXML == nil`, `drawing == nil`
            //    but `rawElements == [commentReference(...)]`. The pre-fix
            //    `isTextRun` predicate returned true for it, so the loop
            //    deleted the Run wholesale — dropping the commentReference
            //    payload and breaking the comment marker triplet
            //    (`commentRangeStart` + `commentRangeEnd` + `commentReference`).
            //    Word strict validator rejects the resulting docx.
            if eRunIdx - sRunIdx > 1 {
                var idx = eRunIdx - 1
                while idx > sRunIdx {
                    let r = runs[idx]
                    let hasStructuralPayload = !(r.rawElements?.isEmpty ?? true)
                    if isTextRun(r) && !hasStructuralPayload {
                        runs.remove(at: idx)
                    }
                    idx -= 1
                }
            }
        }
    }
}
