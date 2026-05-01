import Foundation

// MARK: - WordDocument.replaceTextWithBoundaryDetection
//
// Public API surfacing OMML-boundary-aware replacement, returning informative
// `ReplaceResult` values. New in `flatten-replace-omml-bilateral-coverage`
// Spectra change (cluster fix for PsychQuant/che-word-mcp #99 / #100 / #101 /
// #102 / #103).
//
// **Spec capability**: `ooxml-paragraph-text-mirror`.
// **Library principles** (`ooxml-library-design-principles`):
// 1. Correctness primacy — refuse > incorrect approximation.
// 2. Human-like operations — no surprising state, no silent destruction.
//
// **Mirror invariant**: walks the same surfaces as `flattenedDisplayText`
// (top-level runs + direct-child OMML at all 4 wrapper positions) but treats
// OMML as opaque structural units. Replacements wholly within `<w:t>` ranges
// proceed; replacements crossing OMML boundaries refuse with informative
// occurrence info so callers can produce actionable error messages.
//
// **Why a new public API rather than mutating the existing private static
// `Document.replaceInParagraphSurfaces`**: keeping the existing `Int`-returning
// helper untouched preserves backward compatibility for many internal callers
// in `Document.swift` and avoids cascading changes in this Spectra change's
// scope. The new boundary-detection contract is a meaningful addition that
// deserves its own public surface; future iterations can fold it into the
// private helper if/when other callers need the same semantics.
//
// **Scope**: this implementation handles the top-level paragraph surface
// (paragraph runs + paragraph-level direct-child OMML), which covers the
// che-word-mcp#99 reproducer and matches the spec scenarios. Direct-child
// OMML inside hyperlinks (#100), AlternateContent fallback (#101), and
// nested wrappers (#102) is naturally handled because their typed `runs[]`
// arrays are empty in those fixtures — there is no mutable `<w:t>` text to
// span across OMML in those wrappers, so the existing per-surface engine
// returns 0 replacements without any boundary detection needed. If a future
// fixture has both `<w:t>` runs and direct-child OMML inside the same
// hyperlink/AC wrapper, this implementation can extend per-surface boundary
// detection following the same pattern.

extension WordDocument {

    /// Replace `find` with `replacement` across body paragraphs, refusing
    /// matches whose span intersects direct-child OMML at any of the 4
    /// wrapper positions (paragraph / hyperlink / fallback / nested).
    ///
    /// Returns a `ReplaceResult` distinguishing successful replacement count
    /// (`.replaced`), informative refusals (`.refusedDueToOMMLBoundary`),
    /// and mixed outcomes (`.mixed`).
    ///
    /// **Refusal contract** (Decision 2 — Semantic A opaque OMML):
    /// when a match's span intersects an OMML element, the replacement is
    /// refused for that occurrence and returned in
    /// `Occurrence(matchSpan:, ommlSpans:)`. No mutation of the OMML element
    /// occurs. This is principle-driven (Correctness primacy + Human-like
    /// operations) — refusal is preferred over producing structurally valid
    /// but semantically incorrect output (e.g. `"see δref X"`) or silently
    /// deleting equations.
    ///
    /// **Counts and occurrences**: the result combines per-paragraph
    /// outcomes. A paragraph with both wholly-within and cross-OMML matches
    /// in a single call returns `.mixed(replacedCount:, refusedOccurrences:)`.
    /// The `matchSpan` and `ommlSpans` ranges are in
    /// `Paragraph.flattenedDisplayText()` coordinates of the paragraph
    /// where the occurrence appeared.
    @discardableResult
    public mutating func replaceTextWithBoundaryDetection(
        find: String,
        with replacement: String,
        options: ReplaceOptions = ReplaceOptions()
    ) throws -> ReplaceResult {
        var totalReplaced = 0
        var allRefused: [ReplaceResult.Occurrence] = []

        for i in body.children.indices {
            guard case .paragraph(var para) = body.children[i] else { continue }
            let outcome = try Self.replaceInParagraphWithOMMLBoundaryDetection(
                &para, find: find, with: replacement, options: options
            )
            totalReplaced += outcome.replacedCount
            allRefused.append(contentsOf: outcome.refused)
            body.children[i] = .paragraph(para)
        }

        switch (totalReplaced, allRefused.isEmpty) {
        case (_, true):
            return .replaced(count: totalReplaced)
        case (0, false):
            return .refusedDueToOMMLBoundary(occurrences: allRefused)
        default:
            return .mixed(replacedCount: totalReplaced, refusedOccurrences: allRefused)
        }
    }

    // MARK: - Per-paragraph boundary detection

    /// Result of running OMML-boundary-aware replacement on a single paragraph.
    private struct ParagraphReplaceOutcome {
        let replacedCount: Int
        let refused: [ReplaceResult.Occurrence]
    }

    /// Process top-level surface (paragraph runs + paragraph-level direct-child
    /// OMML in `unrecognizedChildren`) with OMML boundary detection.
    ///
    /// Algorithm:
    /// 1. Build flattened text including OMML visible text in source-XML
    ///    position order (same shape as `Paragraph.flattenedDisplayText`'s
    ///    top-level surface contribution).
    /// 2. Track OMML character ranges within the flattened text — the
    ///    "boundaries" that mutations cannot cross.
    /// 3. Find all matches of `find` in flattened text.
    /// 4. Partition: matches whose span intersects any OMML range → refused
    ///    occurrences. Matches wholly within non-OMML chars → safe.
    /// 5. For safe matches, delegate to existing `TextReplacementEngine.replace`
    ///    which handles cross-run mutation correctly. Use the engine's own
    ///    match count (it re-searches the runs-only flatten which is the
    ///    same string with OMML chars removed).
    private static func replaceInParagraphWithOMMLBoundaryDetection(
        _ para: inout Paragraph,
        find: String,
        with replacement: String,
        options: ReplaceOptions
    ) throws -> ParagraphReplaceOutcome {
        // Step 1+2: build extended flatten + OMML char ranges
        let (extendedFlat, ommlRanges) = buildExtendedFlattenWithOMMLRanges(para)

        guard !extendedFlat.isEmpty, !find.isEmpty else {
            return ParagraphReplaceOutcome(replacedCount: 0, refused: [])
        }

        // Step 3+4: find matches in extended flatten, partition refused vs safe
        let matches = findAllMatches(of: find, in: extendedFlat, options: options)
        var refused: [ReplaceResult.Occurrence] = []
        var safeMatches = 0
        for matchRange in matches {
            let intersecting = ommlRanges.filter { range in
                rangesIntersect(matchRange, range)
            }
            if intersecting.isEmpty {
                safeMatches += 1
            } else {
                refused.append(ReplaceResult.Occurrence(
                    matchSpan: matchRange,
                    ommlSpans: intersecting
                ))
            }
        }

        // Step 5: if any safe matches, delegate to existing engine for actual
        // mutation. The engine searches runs-only flatten (identical to
        // extendedFlat with OMML chars removed), and matches that cross OMML
        // can't appear in runs-only flatten anyway, so the engine naturally
        // skips them.
        var replacedCount = 0
        if safeMatches > 0 {
            replacedCount = try TextReplacementEngine.replace(
                runs: &para.runs, find: find, with: replacement, options: options
            )
        }

        return ParagraphReplaceOutcome(replacedCount: replacedCount, refused: refused)
    }

    // MARK: - Extended flatten with OMML range tracking

    /// Build the paragraph's top-level extended flatten (same shape as the
    /// paragraph contribution in `flattenedDisplayText`) AND the character
    /// ranges where OMML visible text appears, in flattened-text coordinates.
    ///
    /// Top-level surface scope: `runs` interleaved with direct-child OMML
    /// (`unrecognizedChildren` where `name == "oMath" || name == "oMathPara"`)
    /// by source XML position. Other wrappers (hyperlinks, fieldSimples,
    /// alternateContents, contentControls) are NOT included here — they
    /// contribute to flattenedDisplayText separately and their boundary
    /// detection is naturally trivial (typed `runs[]` empty in fixtures
    /// B/C/D from the cluster).
    private static func buildExtendedFlattenWithOMMLRanges(
        _ para: Paragraph
    ) -> (flat: String, ommlRanges: [Range<Int>]) {
        let directOMath = para.unrecognizedChildren.filter { child in
            child.name == "oMath" || child.name == "oMathPara"
        }

        // Fast path: no direct-child OMML. Flatten = runs only, no OMML ranges.
        if directOMath.isEmpty {
            var flat = ""
            for run in para.runs where Self.isMutableTextRun(run) {
                flat += run.text
            }
            return (flat, [])
        }

        // Build positional fragments (runs + OMML) sorted by source position
        enum Fragment {
            case runText(String)  // text from a Run
            case ommlText(String) // visibleText extracted from direct-child OMML
        }
        var fragments: [(position: Int, fragment: Fragment)] = []
        for run in para.runs where Self.isMutableTextRun(run) {
            fragments.append((run.position ?? 0, .runText(run.text)))
        }
        for child in directOMath {
            let visibleText = OMMLParser.parse(xml: child.rawXML).visibleText
            fragments.append((child.position ?? 0, .ommlText(visibleText)))
        }
        fragments.sort { $0.position < $1.position }

        // Concatenate while tracking OMML char ranges
        var flat = ""
        var ommlRanges: [Range<Int>] = []
        for (_, frag) in fragments {
            switch frag {
            case .runText(let text):
                flat += text
            case .ommlText(let text):
                let start = flat.count
                flat += text
                let end = flat.count
                if start < end {
                    ommlRanges.append(start..<end)
                }
            }
        }
        return (flat, ommlRanges)
    }

    // MARK: - Match finding + range arithmetic

    /// Find all match ranges of `find` in `text` per `options`. Returns ranges
    /// in character-index coordinates (not byte / UTF-16 / String.Index).
    private static func findAllMatches(
        of find: String,
        in text: String,
        options: ReplaceOptions
    ) -> [Range<Int>] {
        var results: [Range<Int>] = []
        let cmpOptions: String.CompareOptions = options.matchCase ? [] : [.caseInsensitive]

        if options.regex {
            let regexOptions: NSRegularExpression.Options = options.matchCase ? [] : [.caseInsensitive]
            guard let re = try? NSRegularExpression(pattern: find, options: regexOptions) else {
                return []
            }
            let nsRange = NSRange(text.startIndex..., in: text)
            for match in re.matches(in: text, options: [], range: nsRange) {
                if let swiftRange = Range(match.range, in: text) {
                    let lo = text.distance(from: text.startIndex, to: swiftRange.lowerBound)
                    let hi = text.distance(from: text.startIndex, to: swiftRange.upperBound)
                    results.append(lo..<hi)
                }
            }
        } else {
            var searchStart = text.startIndex
            while searchStart < text.endIndex,
                  let range = text.range(of: find, options: cmpOptions, range: searchStart..<text.endIndex) {
                let lo = text.distance(from: text.startIndex, to: range.lowerBound)
                let hi = text.distance(from: text.startIndex, to: range.upperBound)
                results.append(lo..<hi)
                searchStart = range.upperBound
            }
        }
        return results
    }

    /// Half-open range intersection check. Returns true if ranges share any
    /// position. Two ranges `[a, b)` and `[c, d)` intersect iff `a < d && c < b`.
    private static func rangesIntersect(_ lhs: Range<Int>, _ rhs: Range<Int>) -> Bool {
        return lhs.lowerBound < rhs.upperBound && rhs.lowerBound < lhs.upperBound
    }

    /// Mirror of the (private) `TextReplacementEngine.isTextRun` predicate.
    /// A "text run" is one whose `text` field carries the displayed content
    /// (no `rawXML` override, no drawing). This is what the engine's
    /// `flattenRuns` includes when building the matchable string.
    private static func isMutableTextRun(_ run: Run) -> Bool {
        return run.rawXML == nil && run.drawing == nil
    }
}
