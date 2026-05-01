import Foundation

// MARK: - FieldParser (v0.10.0)
//
// Read-side inverse of `FieldCode.toFieldXML()`. Given a `Paragraph`, walks
// its runs looking for `<w:fldChar>` begin/separate/end spans (inside
// `Run.rawXML`) and parses the instrText into typed `FieldCode` values.
//
// The emit side produces a 5-run structure:
//   <w:r><w:fldChar w:fldCharType="begin"/></w:r>
//   <w:r><w:instrText xml:space="preserve"> SEQ Figure ... </w:instrText></w:r>
//   <w:r><w:fldChar w:fldCharType="separate"/></w:r>
//   <w:r><w:t>cached</w:t></w:r>
//   <w:r><w:fldChar w:fldCharType="end"/></w:r>
//
// Per v2.0.0 convention, che-word-mcp writes this entire 5-run block into a
// single `Run.rawXML`. FieldParser accepts both forms:
//   (a) the 5 runs as 5 separate elements in `paragraph.runs`, OR
//   (b) one run whose `rawXML` contains the full 5-run block.

/// Value layer for a parsed field — one case per recognized `FieldCode` type
/// plus `.unknown` for graceful fallback.
public enum ParsedFieldValue {
    case sequence(SequenceField)
    case styleRef(StyleRefField)
    case reference(ReferenceField)
    /// Fallback for field grammars that no registered parser recognized.
    /// Preserves the raw instrText so round-trip `toFieldXML()` reemits it intact.
    case unknown(instrText: String)
}

extension ParsedFieldValue: Equatable {
    public static func == (lhs: ParsedFieldValue, rhs: ParsedFieldValue) -> Bool {
        switch (lhs, rhs) {
        case let (.sequence(a), .sequence(b)):
            return a.identifier == b.identifier && a.format == b.format
                && a.resetLevel == b.resetLevel && a.hideResult == b.hideResult
        case let (.styleRef(a), .styleRef(b)):
            return a.headingLevel == b.headingLevel
                && a.suppressNonDelimiter == b.suppressNonDelimiter
                && a.insertPositionBeforeRef == b.insertPositionBeforeRef
        case let (.reference(a), .reference(b)):
            return a.type == b.type && a.bookmarkName == b.bookmarkName
                && a.includeAboveBelow == b.includeAboveBelow
                && a.createHyperlink == b.createHyperlink
        case let (.unknown(a), .unknown(b)):
            return a == b
        default:
            return false
        }
    }
}

/// One recognized field span inside a paragraph, with run-index location info
/// so CRUD tools can modify specific runs.
public enum ParsedFieldLocation: Equatable {
    case paragraphRun
    case hyperlinkRun(hyperlinkIndex: Int)
    case fieldSimple(index: Int)
    case alternateContentFallbackRun(alternateContentIndex: Int)
    case contentControl(index: Int)
}

public struct ParsedField: Equatable {
    public static func == (lhs: ParsedField, rhs: ParsedField) -> Bool {
        lhs.startRunIdx == rhs.startRunIdx
            && lhs.endRunIdx == rhs.endRunIdx
            && lhs.cachedResultRunIdx == rhs.cachedResultRunIdx
            && lhs.instrText == rhs.instrText
            && lhs.field == rhs.field
            && lhs.location == rhs.location
    }

    /// Index of the run containing `<w:fldChar fldCharType="begin">`.
    public let startRunIdx: Int
    /// Index of the run containing `<w:fldChar fldCharType="end">`.
    public let endRunIdx: Int
    /// Index of the run containing `<w:t>cached</w:t>` between separate and end.
    /// May be `nil` if the field has no cached result run.
    public let cachedResultRunIdx: Int?
    /// Raw instrText string (without the enclosing `<w:instrText>` tags).
    public let instrText: String
    /// Parsed field value, dispatched from instrText.
    public let field: ParsedFieldValue
    /// Paragraph surface where this field was found.
    public let location: ParsedFieldLocation

    public init(
        startRunIdx: Int,
        endRunIdx: Int,
        cachedResultRunIdx: Int?,
        instrText: String,
        field: ParsedFieldValue,
        location: ParsedFieldLocation = .paragraphRun
    ) {
        self.startRunIdx = startRunIdx
        self.endRunIdx = endRunIdx
        self.cachedResultRunIdx = cachedResultRunIdx
        self.instrText = instrText
        self.field = field
        self.location = location
    }
}

public enum FieldParser {

    /// Parse all field spans in the given paragraph's runs. Returns one
    /// `ParsedField` per recognized field. Unknown field types produce
    /// `.unknown(instrText:)` cases — callers never lose data.
    ///
    /// Two-phase scan (PsychQuant/che-word-mcp#104):
    /// - **Phase 1 (baked form)**: scan runs whose rawXML contains BOTH
    ///   `"fldChar"` AND `"instrText"` — this is the v2.0.0 convention where
    ///   `wrapCaptionSequenceFields` and `Field.toFieldXML()` produce a single
    ///   `Run.rawXML` carrying all 5 fldChar elements as nested `<w:r>` blocks.
    /// - **Phase 2 (canonical 5-run form)**: if Phase 1 finds nothing, walk the
    ///   runs as a state machine looking for the disk/native-Word emission form
    ///   where each fldChar element lives in its own `<w:r>`. DocxReader
    ///   produces this form when re-reading any docx (including our own
    ///   `wrapCaptionSequenceFields` output after Writer→Reader roundtrip).
    public static func parse(paragraph: Paragraph) -> [ParsedField] {
        var result: [ParsedField] = []

        result.append(contentsOf: parseRuns(paragraph.runs) { _ in .paragraphRun })

        for (index, hyperlink) in paragraph.hyperlinks.enumerated() {
            result.append(contentsOf: parseRuns(hyperlink.runs) { _ in
                .hyperlinkRun(hyperlinkIndex: index)
            })
        }

        for (index, fieldSimple) in paragraph.fieldSimples.enumerated() {
            result.append(ParsedField(
                startRunIdx: -1,
                endRunIdx: -1,
                cachedResultRunIdx: nil,
                instrText: fieldSimple.instr,
                field: dispatchParse(instrText: fieldSimple.instr),
                location: .fieldSimple(index: index)
            ))
        }

        for (index, alternateContent) in paragraph.alternateContents.enumerated() {
            result.append(contentsOf: parseRuns(alternateContent.fallbackRuns) { _ in
                .alternateContentFallbackRun(alternateContentIndex: index)
            })
        }

        for (index, control) in paragraph.contentControls.enumerated() where !control.content.isEmpty {
            result.append(contentsOf: parseFieldsInRawXML(
                control.content,
                atRunIdx: -1,
                location: .contentControl(index: index)
            ))
        }

        return result
    }

    private static func parseRuns(
        _ runs: [Run],
        locationForRun: (Int) -> ParsedFieldLocation
    ) -> [ParsedField] {
        var result: [ParsedField] = []

        for (runIdx, run) in runs.enumerated() {
            guard let rawXML = run.rawXML else { continue }
            // Phase 1 marker: BOTH fldChar AND instrText in the same rawXML
            // means this run carries a baked-form 5-run block.
            guard rawXML.contains("fldChar"), rawXML.contains("instrText") else { continue }

            let fields = parseFieldsInRawXML(rawXML, atRunIdx: runIdx, location: locationForRun(runIdx))
            result.append(contentsOf: fields)
        }

        // Phase 2: canonical form fallback when Phase 1 found nothing.
        // (Mixed forms — some baked, some canonical — are not supported because
        // they would require the canonical scanner to skip ranges already
        // claimed by Phase 1. No real-world producer mixes the two in a single
        // paragraph; documenting the assumption is sufficient.)
        if result.isEmpty {
            result.append(contentsOf: parseFiveRunSpan(runs: runs, locationForRun: locationForRun))
        }

        return result
    }

    /// Phase-2 walker: detect SEQ fields emitted as 5 separate `<w:r>` runs
    /// (begin / instrText / separate / cachedValue / end). State machine
    /// resets on out-of-order patterns to be robust against malformed
    /// paragraphs.
    private static func parseFiveRunSpan(
        runs: [Run],
        locationForRun: (Int) -> ParsedFieldLocation
    ) -> [ParsedField] {
        // State of an in-progress field span as we walk the runs.
        struct InProgress {
            var beginRunIdx: Int
            var instrText: String?
            var separateRunIdx: Int?
            var cachedRunIdx: Int?
        }

        var result: [ParsedField] = []
        var current: InProgress?

        // Reusable instrText extractor: matches `<w:instrText[ ...]>...</w:instrText>`
        // anywhere in the run's rawXML and returns the inner content.
        let instrTextRegex = try? NSRegularExpression(
            pattern: #"<w:instrText[^>]*>(.*?)</w:instrText>"#,
            options: [.dotMatchesLineSeparators]
        )

        func extractInstrText(_ rawXML: String) -> String? {
            guard let regex = instrTextRegex else { return nil }
            let nsRange = NSRange(rawXML.startIndex..., in: rawXML)
            guard let match = regex.firstMatch(in: rawXML, options: [], range: nsRange),
                  match.numberOfRanges >= 2,
                  let innerRange = Range(match.range(at: 1), in: rawXML) else { return nil }
            return String(rawXML[innerRange])
        }

        // Probe a single run for fldChar/instrText fragments. DocxReader stores
        // unrecognized run children (including `<w:fldChar>` and
        // `<w:instrText>`) in `Run.rawElements` (NOT `Run.rawXML`). Native-Word
        // 5-run paragraphs constructed by hand may instead embed the fragment
        // directly in `Run.rawXML`. Check both surfaces.
        func runFragments(_ run: Run) -> String {
            var pieces: [String] = []
            if let rawXML = run.rawXML { pieces.append(rawXML) }
            if let elems = run.rawElements {
                for elem in elems {
                    pieces.append(elem.xml)
                }
            }
            return pieces.joined()
        }

        for (idx, run) in runs.enumerated() {
            let fragments = runFragments(run)
            // Caption text runs (only `<w:t>`) carry no fldChar/instrText
            // signal. Treat them as neutral — they neither advance nor reset
            // the in-progress span (caption text interleaved with fldChar runs
            // is normal). But a run with `<w:t>` text right after `separate`
            // IS the cached value, so check that case below.
            let hasFldCharBegin = fragments.contains("fldCharType=\"begin\"")
            let hasFldCharSeparate = fragments.contains("fldCharType=\"separate\"")
            let hasFldCharEnd = fragments.contains("fldCharType=\"end\"")
            let hasInstrText = fragments.contains("<w:instrText")
            let hasText = fragments.contains("<w:t")
                       || (run.rawXML == nil && (run.rawElements?.isEmpty ?? true) && !run.text.isEmpty)

            // fldChar begin: start (or restart) a span
            if hasFldCharBegin {
                current = InProgress(beginRunIdx: idx, instrText: nil,
                                     separateRunIdx: nil, cachedRunIdx: nil)
                continue
            }

            // instrText: capture content into in-progress span
            if hasInstrText, var span = current {
                if let extracted = extractInstrText(fragments) {
                    span.instrText = extracted
                    current = span
                }
                continue
            }

            // fldChar separate: mark separator position
            if hasFldCharSeparate, var span = current {
                span.separateRunIdx = idx
                current = span
                continue
            }

            // First text-bearing run after `separate`: capture cached result
            // run index. Native Word emits the cached value as `<w:t>1</w:t>`
            // in its own run between separate and end. After roundtrip,
            // DocxReader exposes that text via `Run.text` (rawXML is nil for
            // this run because `<w:t>` is in `recognizedRunChildren`).
            if let span = current, span.separateRunIdx != nil, span.cachedRunIdx == nil,
               hasText {
                var updated = span
                updated.cachedRunIdx = idx
                current = updated
                continue
            }

            // fldChar end: emit span and reset
            if hasFldCharEnd, let span = current {
                if let instrText = span.instrText {
                    let parsedValue = dispatchParse(instrText: instrText)
                    result.append(ParsedField(
                        startRunIdx: span.beginRunIdx,
                        endRunIdx: idx,
                        cachedResultRunIdx: span.cachedRunIdx,
                        instrText: instrText,
                        field: parsedValue,
                        location: locationForRun(idx)
                    ))
                }
                current = nil
                continue
            }
        }

        return result
    }

    /// Extract field spans from a single Run.rawXML containing one or more
    /// embedded 5-run field blocks. All parsed fields report the same runIdx
    /// (they're logically inside that one paragraph run).
    private static func parseFieldsInRawXML(
        _ rawXML: String,
        atRunIdx runIdx: Int,
        location: ParsedFieldLocation
    ) -> [ParsedField] {
        var results: [ParsedField] = []

        // Scan for <w:instrText ...>...</w:instrText> occurrences.
        // Each one corresponds to one field span (begin → separate → end).
        let instrTextPattern = #"<w:instrText[^>]*>(.*?)</w:instrText>"#
        guard let regex = try? NSRegularExpression(pattern: instrTextPattern, options: [.dotMatchesLineSeparators]) else {
            return []
        }
        let nsRange = NSRange(rawXML.startIndex..., in: rawXML)
        let matches = regex.matches(in: rawXML, options: [], range: nsRange)

        for match in matches {
            guard match.numberOfRanges >= 2,
                  let innerRange = Range(match.range(at: 1), in: rawXML) else { continue }
            let instrText = String(rawXML[innerRange])
            let parsedValue = dispatchParse(instrText: instrText)

            // Try to extract cached result from the following <w:t>...</w:t> between
            // this instrText and the next <w:fldChar end> or next field block.
            let cachedResult = extractCachedResult(afterInstrTextRange: match.range(at: 0), in: rawXML)
            _ = cachedResult  // Currently unused — preserved for future round-trip fidelity

            results.append(ParsedField(
                startRunIdx: runIdx,
                endRunIdx: runIdx,
                cachedResultRunIdx: runIdx,  // Same run — rawXML-embedded form
                instrText: instrText,
                field: parsedValue,
                location: location
            ))
        }

        return results
    }

    /// Extract `<w:t>...</w:t>` text between given instrText range and next fldChar end or field.
    private static func extractCachedResult(afterInstrTextRange instrRange: NSRange, in rawXML: String) -> String? {
        let searchStart = instrRange.location + instrRange.length
        guard searchStart < rawXML.count else { return nil }
        let searchNSRange = NSRange(location: searchStart, length: rawXML.count - searchStart)
        let pattern = #"<w:fldChar[^/]*fldCharType="separate"[^/]*/>\s*<w:r[^>]*>\s*<w:t[^>]*>(.*?)</w:t>"#
        guard let regex = try? NSRegularExpression(pattern: pattern, options: [.dotMatchesLineSeparators]) else {
            return nil
        }
        if let match = regex.firstMatch(in: rawXML, options: [], range: searchNSRange),
           match.numberOfRanges >= 2,
           let innerRange = Range(match.range(at: 1), in: rawXML) {
            return String(rawXML[innerRange])
        }
        return nil
    }

    /// Cascading dispatch: try each registered parser in turn, fall back to unknown.
    private static func dispatchParse(instrText: String) -> ParsedFieldValue {
        if let f = SequenceField.parse(instrText: instrText) { return .sequence(f) }
        if let f = StyleRefField.parse(instrText: instrText) { return .styleRef(f) }
        if let f = ReferenceField.parse(instrText: instrText) { return .reference(f) }
        return .unknown(instrText: instrText)
    }
}
