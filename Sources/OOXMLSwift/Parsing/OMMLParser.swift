import Foundation

// MARK: - OMMLParser (v0.10.0)
//
// Read-side inverse of `MathComponent.toOMML()`. Given an `<m:oMath>` XML
// string (optionally wrapped in `<m:oMathPara>`), returns a tree of
// `MathComponent` values matching the concrete structs.
//
// Implementation: balanced-brace tag splitter over the top-level children,
// dispatching each `<m:tag>` by name. Nested content inside each recognized
// tag is re-parsed recursively. Unrecognized tags preserved via `UnknownMath`.

public enum OMMLParser {

    /// Parse an OOXML math XML string into a tree of `MathComponent` values.
    /// Strips `<m:oMath>` and `<m:oMathPara>` wrappers if present.
    public static func parse(xml: String) -> [MathComponent] {
        var input = xml.trimmingCharacters(in: .whitespacesAndNewlines)
        input = stripWrapper(input, tag: "m:oMathPara") ?? input
        input = stripWrapper(input, tag: "m:oMath") ?? input
        return parseChildren(input)
    }

    private static func stripWrapper(_ xml: String, tag: String) -> String? {
        let openPattern = "<\(tag)"
        let closePattern = "</\(tag)>"
        guard let openRange = xml.range(of: openPattern),
              openRange.lowerBound == xml.startIndex,
              let openTagEnd = xml.range(of: ">", range: openRange.upperBound..<xml.endIndex),
              let closeRange = xml.range(of: closePattern, options: .backwards) else {
            return nil
        }
        let inner = xml[openTagEnd.upperBound..<closeRange.lowerBound]
        return String(inner)
    }

    private static func parseChildren(_ xml: String) -> [MathComponent] {
        var results: [MathComponent] = []
        var remaining = xml[...]
        while let open = remaining.range(of: "<m:") {
            remaining = remaining[open.lowerBound...]
            guard let (block, rest) = extractBlock(remaining) else { break }
            if let component = parseBlock(String(block)) {
                results.append(component)
            }
            remaining = rest
        }
        return results
    }

    private static func extractBlock(_ xml: Substring) -> (Substring, Substring)? {
        let afterOpenBracket = xml.index(xml.startIndex, offsetBy: 1)
        let afterPrefix = xml.index(afterOpenBracket, offsetBy: 2)
        var cursor = afterPrefix
        while cursor < xml.endIndex {
            let c = xml[cursor]
            if c == " " || c == ">" || c == "/" { break }
            cursor = xml.index(after: cursor)
        }
        let tagName = String(xml[afterPrefix..<cursor])

        guard let tagCloseIdx = xml.range(of: ">", range: cursor..<xml.endIndex) else {
            return nil
        }
        let beforeCloseIdx = xml.index(before: tagCloseIdx.lowerBound)
        if beforeCloseIdx >= cursor && xml[beforeCloseIdx] == "/" {
            let afterClose = xml.index(after: tagCloseIdx.lowerBound)
            return (xml[xml.startIndex..<afterClose], xml[afterClose...])
        }

        let closePattern = "</m:\(tagName)>"
        var depth = 1
        var searchStart = tagCloseIdx.upperBound
        while searchStart < xml.endIndex {
            guard let nextClose = xml.range(of: closePattern, range: searchStart..<xml.endIndex) else {
                return nil
            }
            // Count exact-tag opens between searchStart and nextClose. Avoid
            // prefix false matches (e.g., "<m:r" matching "<m:rPr") by requiring
            // a word boundary character after the tag name.
            let depthAdd = countExactOpens(in: xml, tagName: tagName, range: searchStart..<nextClose.lowerBound)
            depth += depthAdd
            depth -= 1
            if depth == 0 {
                let blockEnd = nextClose.upperBound
                return (xml[xml.startIndex..<blockEnd], xml[blockEnd...])
            }
            searchStart = nextClose.upperBound
        }
        return nil
    }

    /// Count opens of `<m:tagName>` or `<m:tagName ` or `<m:tagName/>` in range.
    /// Rejects prefix matches like `<m:rPr` when looking for `<m:r`.
    private static func countExactOpens(in xml: Substring, tagName: String, range: Range<Substring.Index>) -> Int {
        let openPattern = "<m:\(tagName)"
        var count = 0
        var cursor = range.lowerBound
        while cursor < range.upperBound,
              let found = xml.range(of: openPattern, range: cursor..<range.upperBound) {
            let afterIdx = found.upperBound
            if afterIdx < xml.endIndex {
                let c = xml[afterIdx]
                if c == " " || c == ">" || c == "/" {
                    count += 1
                }
            }
            cursor = afterIdx
        }
        return count
    }

    private static func parseBlock(_ block: String) -> MathComponent? {
        guard block.hasPrefix("<m:") else { return nil }
        let afterPrefix = block.index(block.startIndex, offsetBy: 3)
        var cursor = afterPrefix
        while cursor < block.endIndex {
            let c = block[cursor]
            if c == " " || c == ">" || c == "/" { break }
            cursor = block.index(after: cursor)
        }
        let tagName = String(block[afterPrefix..<cursor])

        switch tagName {
        case "r":
            return parseMathRun(block)
        case "f":
            return parseMathFraction(block)
        case "sSub", "sSup", "sSubSup":
            return parseMathSubSuperScript(block, kind: tagName)
        case "rad":
            return parseMathRadical(block)
        case "nary":
            return parseMathNary(block)
        case "acc":
            return parseMathAccent(block)
        default:
            return UnknownMath(rawXML: block)
        }
    }

    private static func parseMathRun(_ block: String) -> MathRun {
        let text: String
        if let tRange = block.range(of: "<m:t>"),
           let tCloseRange = block.range(of: "</m:t>", options: .backwards) {
            text = decodeXMLEntities(String(block[tRange.upperBound..<tCloseRange.lowerBound]))
        } else {
            text = ""
        }

        var style: MathStyle?
        if let styRange = block.range(of: #"<m:sty m:val=""#) {
            let afterVal = block[styRange.upperBound...]
            if let quoteEnd = afterVal.firstIndex(of: "\"") {
                let styVal = String(afterVal[..<quoteEnd])
                style = MathStyle(rawValue: styVal)
            }
        }

        return MathRun(text: text, style: style)
    }

    private static func parseMathFraction(_ block: String) -> MathFraction {
        let numXML = extractInner(block, childTag: "m:num") ?? ""
        let denXML = extractInner(block, childTag: "m:den") ?? ""
        return MathFraction(
            numerator: parseChildren(numXML),
            denominator: parseChildren(denXML)
        )
    }

    private static func parseMathSubSuperScript(_ block: String, kind: String) -> MathSubSuperScript {
        let baseXML = extractInner(block, childTag: "m:e") ?? ""
        let subXML = (kind == "sSub" || kind == "sSubSup") ? extractInner(block, childTag: "m:sub") : nil
        let supXML = (kind == "sSup" || kind == "sSubSup") ? extractInner(block, childTag: "m:sup") : nil
        return MathSubSuperScript(
            base: parseChildren(baseXML),
            sub: subXML.map { parseChildren($0) },
            sup: supXML.map { parseChildren($0) }
        )
    }

    private static func parseMathRadical(_ block: String) -> MathRadical {
        let radicandXML = extractInner(block, childTag: "m:e") ?? ""
        let degreeXML = extractInner(block, childTag: "m:deg") ?? ""
        let hasDegree = !degreeXML.isEmpty && !block.contains("<m:degHide m:val=\"1\"/>")
        return MathRadical(
            radicand: parseChildren(radicandXML),
            degree: hasDegree ? parseChildren(degreeXML) : nil
        )
    }

    private static func parseMathAccent(_ block: String) -> MathAccent {
        var accentChar = ""
        if let chrRange = block.range(of: #"<m:chr m:val=""#) {
            let afterVal = block[chrRange.upperBound...]
            if let quoteEnd = afterVal.firstIndex(of: "\"") {
                accentChar = decodeXMLEntities(String(afterVal[..<quoteEnd]))
            }
        }
        let baseXML = extractInner(block, childTag: "m:e") ?? ""
        return MathAccent(
            base: parseChildren(baseXML),
            accentChar: accentChar
        )
    }

    private static func parseMathNary(_ block: String) -> MathNary {
        var op: MathNary.NaryOperator = .sum
        if let chrRange = block.range(of: #"<m:chr m:val=""#) {
            let afterVal = block[chrRange.upperBound...]
            if let quoteEnd = afterVal.firstIndex(of: "\"") {
                let chrVal = String(afterVal[..<quoteEnd])
                op = MathNary.NaryOperator(rawValue: chrVal) ?? .sum
            }
        }
        let subXML = extractInner(block, childTag: "m:sub") ?? ""
        let supXML = extractInner(block, childTag: "m:sup") ?? ""
        let baseXML = extractInner(block, childTag: "m:e") ?? ""
        return MathNary(
            op: op,
            sub: subXML.isEmpty ? nil : parseChildren(subXML),
            sup: supXML.isEmpty ? nil : parseChildren(supXML),
            base: parseChildren(baseXML)
        )
    }

    private static func extractInner(_ block: String, childTag: String) -> String? {
        let openPattern = "<\(childTag)>"
        let closePattern = "</\(childTag)>"
        guard let openRange = block.range(of: openPattern) else { return nil }
        var searchStart = openRange.upperBound
        var depth = 1
        while searchStart < block.endIndex {
            guard let nextClose = block.range(of: closePattern, range: searchStart..<block.endIndex) else {
                return nil
            }
            var tmp = searchStart
            while tmp < nextClose.lowerBound {
                if let nested = block.range(of: openPattern, range: tmp..<nextClose.lowerBound) {
                    depth += 1
                    tmp = nested.upperBound
                } else {
                    break
                }
            }
            depth -= 1
            if depth == 0 {
                return String(block[openRange.upperBound..<nextClose.lowerBound])
            }
            searchStart = nextClose.upperBound
        }
        return nil
    }

    private static func decodeXMLEntities(_ s: String) -> String {
        return s
            .replacingOccurrences(of: "&lt;", with: "<")
            .replacingOccurrences(of: "&gt;", with: ">")
            .replacingOccurrences(of: "&amp;", with: "&")
    }
}
