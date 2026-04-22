import Foundation

// MARK: - FieldCode read-side parsers (v0.10.0)
//
// Each `FieldCode`-conforming type exposes a `static func parse(instrText:) -> Self?`
// that recognizes its own `<w:instrText>` grammar. `FieldParser` dispatches by
// trying each type in turn until one returns non-nil.
//
// Design rationale: per-type parsers keep grammar next to the writer
// (`fieldInstruction` getter on the same struct). Adding a new field type
// requires no central registry — just add conformance + a parse method.

extension SequenceField {
    /// Parse " SEQ <identifier>[ \\* <FORMAT>][ \\s <level>][ \\h] " back into a SequenceField.
    /// Returns nil if the instrText does not start with "SEQ ".
    public static func parse(instrText: String) -> SequenceField? {
        let trimmed = instrText.trimmingCharacters(in: .whitespaces)
        guard trimmed.hasPrefix("SEQ ") else { return nil }

        // Tokenize by whitespace, respecting that "\\* ARABIC" etc. are two-token switches.
        var tokens = trimmed.dropFirst("SEQ ".count)
            .split(separator: " ", omittingEmptySubsequences: true)
            .map(String.init)
        guard let identifier = tokens.first else { return nil }
        tokens.removeFirst()

        var format: SequenceFormat = .arabic
        var resetLevel: Int?
        var hideResult = false

        var i = 0
        while i < tokens.count {
            let tok = tokens[i]
            if tok == "\\*" && i + 1 < tokens.count {
                let fmtTok = tokens[i + 1]
                switch fmtTok {
                case "ARABIC": format = .arabic
                case "ALPHABETIC": format = .alphabetic
                case "alphabetic": format = .lowerAlphabetic
                case "ROMAN": format = .roman
                case "roman": format = .lowerRoman
                default:
                    return nil
                }
                i += 2
            } else if tok == "\\s" && i + 1 < tokens.count {
                guard let level = Int(tokens[i + 1]) else { return nil }
                resetLevel = level
                i += 2
            } else if tok == "\\h" {
                hideResult = true
                i += 1
            } else {
                return nil
            }
        }

        return SequenceField(
            identifier: identifier,
            format: format,
            resetLevel: resetLevel,
            hideResult: hideResult,
            cachedResult: nil
        )
    }
}

extension StyleRefField {
    /// Parse " STYLEREF <int>[ \\s][ \\l] " into a StyleRefField.
    public static func parse(instrText: String) -> StyleRefField? {
        let trimmed = instrText.trimmingCharacters(in: .whitespaces)
        guard trimmed.hasPrefix("STYLEREF ") else { return nil }

        let tokens = trimmed.dropFirst("STYLEREF ".count)
            .split(separator: " ", omittingEmptySubsequences: true)
            .map(String.init)
        guard let levelStr = tokens.first, let level = Int(levelStr) else { return nil }

        var suppressNonDelimiter = false
        var insertPositionBeforeRef = false
        for tok in tokens.dropFirst() {
            switch tok {
            case "\\s": suppressNonDelimiter = true
            case "\\l": insertPositionBeforeRef = true
            default: return nil
            }
        }

        return StyleRefField(
            headingLevel: level,
            suppressNonDelimiter: suppressNonDelimiter,
            insertPositionBeforeRef: insertPositionBeforeRef,
            cachedResult: nil
        )
    }
}

extension ReferenceField {
    /// Parse " (REF|PAGEREF|NOTEREF) <bookmark>[ \\p][ \\h] " into a ReferenceField.
    public static func parse(instrText: String) -> ReferenceField? {
        let trimmed = instrText.trimmingCharacters(in: .whitespaces)
        let type: ReferenceFieldType
        let rest: Substring
        if trimmed.hasPrefix("REF ") {
            type = .ref
            rest = trimmed.dropFirst("REF ".count)
        } else if trimmed.hasPrefix("PAGEREF ") {
            type = .pageRef
            rest = trimmed.dropFirst("PAGEREF ".count)
        } else if trimmed.hasPrefix("NOTEREF ") {
            type = .noteRef
            rest = trimmed.dropFirst("NOTEREF ".count)
        } else {
            return nil
        }

        let tokens = rest.split(separator: " ", omittingEmptySubsequences: true).map(String.init)
        guard let bookmarkName = tokens.first else { return nil }

        var includeAboveBelow = false
        var createHyperlink = false
        for tok in tokens.dropFirst() {
            switch tok {
            case "\\p": includeAboveBelow = true
            case "\\h": createHyperlink = true
            default: return nil
            }
        }

        return ReferenceField(
            type: type,
            bookmarkName: bookmarkName,
            includeAboveBelow: includeAboveBelow,
            createHyperlink: createHyperlink,
            cachedResult: nil
        )
    }
}
