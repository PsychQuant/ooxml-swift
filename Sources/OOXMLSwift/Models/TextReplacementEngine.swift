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
            if eRunIdx - sRunIdx > 1 {
                var idx = eRunIdx - 1
                while idx > sRunIdx {
                    if isTextRun(runs[idx]) {
                        runs.remove(at: idx)
                    }
                    idx -= 1
                }
            }
        }
    }
}
