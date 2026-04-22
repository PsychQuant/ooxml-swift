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
public struct ParsedField: Equatable {
    public static func == (lhs: ParsedField, rhs: ParsedField) -> Bool {
        lhs.startRunIdx == rhs.startRunIdx
            && lhs.endRunIdx == rhs.endRunIdx
            && lhs.cachedResultRunIdx == rhs.cachedResultRunIdx
            && lhs.instrText == rhs.instrText
            && lhs.field == rhs.field
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

    public init(
        startRunIdx: Int,
        endRunIdx: Int,
        cachedResultRunIdx: Int?,
        instrText: String,
        field: ParsedFieldValue
    ) {
        self.startRunIdx = startRunIdx
        self.endRunIdx = endRunIdx
        self.cachedResultRunIdx = cachedResultRunIdx
        self.instrText = instrText
        self.field = field
    }
}

public enum FieldParser {

    /// Parse all field spans in the given paragraph's runs. Returns one
    /// `ParsedField` per recognized field. Unknown field types produce
    /// `.unknown(instrText:)` cases — callers never lose data.
    public static func parse(paragraph: Paragraph) -> [ParsedField] {
        var result: [ParsedField] = []

        for (runIdx, run) in paragraph.runs.enumerated() {
            guard let rawXML = run.rawXML, rawXML.contains("fldChar") else { continue }

            // Per v2.0.0 convention, the whole 5-run block is baked into a single
            // run's rawXML. Extract instrText + cached result from that string.
            let fields = parseFieldsInRawXML(rawXML, atRunIdx: runIdx)
            result.append(contentsOf: fields)
        }

        return result
    }

    /// Extract field spans from a single Run.rawXML containing one or more
    /// embedded 5-run field blocks. All parsed fields report the same runIdx
    /// (they're logically inside that one paragraph run).
    private static func parseFieldsInRawXML(_ rawXML: String, atRunIdx runIdx: Int) -> [ParsedField] {
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
                field: parsedValue
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
