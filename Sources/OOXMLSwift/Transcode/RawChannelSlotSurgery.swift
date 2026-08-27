// RawChannelSlotSurgery.swift
// raw-channel-slot-support (#171) — run-level text substitution inside a
// carried word/document.xml part.
//
// A document whose main part fails the DSL upgrade rides the raw channel as
// one `.carryPart` op; no paragraph-level ops exist, so DSL and op-level
// slots cannot reach its paragraphs. This surgery locates a paragraph by
// `w14:paraId` inside the carried XML string, extracts its visible text (the
// slot default), and — only when the substituted value differs from the
// current text (identity shortcut) — replaces the paragraph's children with a
// single run carrying the dominant text run's rPr. Everything outside the
// designated paragraph stays byte-identical, so an all-default execution
// reproduces the reference byte-equal and the existing Stage B acceptance
// applies unchanged.
//
// Location is STRUCTURE-AWARE, never bare token search (verify round 1):
// `w14:paraId` legitimately appears on `<w:tr>` elements and inside `<w:t>`
// text content, `<w:p>` nests through `<w:txbxContent>`, and `>` may occur
// inside quoted attribute values. The scanner below walks the XML once —
// quote-aware inside tags, skipping comments/CDATA/PIs, tracking `<w:p>`
// nesting depth — and recognizes the paraId only as an attribute of a
// `<w:p>` opening tag. Import re-applies every designation guard and
// verifies post-surgery well-formedness: a silently unfilled or corrupted
// official form is worse than a loud stop.

import Foundation

enum RawChannelSlotSurgery {

    /// Where a designated paraId lives, per structure-aware scan.
    enum Location {
        case unique(ParagraphSpan)
        case duplicate(count: Int)
        case notOnParagraph(carrier: String)
        case absent
    }

    /// A located `<w:p …>…</w:p>` (or self-closing `<w:p …/>`) element.
    struct ParagraphSpan {
        /// Full element range, opening `<` through closing `>`.
        let range: Range<String.Index>
        /// Index just past the opening tag's `>`.
        let contentStart: String.Index
        let selfClosing: Bool
    }

    /// The carried `word/document.xml` content, if the log rides the raw
    /// channel for its main part.
    static func documentCarryXML(log: OperationLog) -> String? {
        for entry in log.entries {
            if case .carryPart(let path, let xml) = entry.op, path == "word/document.xml" {
                return xml
            }
        }
        return nil
    }

    /// Structure-aware location of the paraId. Exactly one `<w:p>` carrier →
    /// `.unique`; several → `.duplicate`; carried only by non-paragraph
    /// elements → `.notOnParagraph` (first carrier name); otherwise `.absent`
    /// (occurrences inside text content or attribute values of the SEARCHED
    /// token never anchor — only opening-tag attributes are consulted).
    static func locate(paraId: String, in xml: String) -> Location {
        var paragraphSpans: [ParagraphSpan] = []
        var otherCarrier: String? = nil
        // Open `<w:p>` elements awaiting their close: (element start, content start).
        var openParagraphs: [(start: String.Index, contentStart: String.Index, isTarget: Bool)] = []

        var i = xml.startIndex
        while i < xml.endIndex, let lt = xml[i...].firstIndex(of: "<") {
            let after = xml.index(after: lt)
            guard after < xml.endIndex else { break }
            if xml[lt...].hasPrefix("<!--") {
                i = skipPast("-->", from: lt, in: xml); continue
            }
            if xml[lt...].hasPrefix("<![CDATA[") {
                i = skipPast("]]>", from: lt, in: xml); continue
            }
            if xml[lt...].hasPrefix("<?") {
                i = skipPast("?>", from: lt, in: xml); continue
            }
            if xml[lt...].hasPrefix("<!") {   // DOCTYPE etc — not in OOXML parts, skip tag
                i = skipPast(">", from: lt, in: xml); continue
            }
            if xml[after] == "/" {
                // Closing tag.
                let nameStart = xml.index(after: after)
                let name = tagName(in: xml, from: nameStart)
                guard let gt = xml[nameStart...].firstIndex(of: ">") else { break }
                let closeEnd = xml.index(after: gt)
                if name == "w:p", let open = openParagraphs.popLast() {
                    if open.isTarget {
                        paragraphSpans.append(ParagraphSpan(
                            range: open.start..<closeEnd,
                            contentStart: open.contentStart,
                            selfClosing: false))
                    }
                }
                i = closeEnd
                continue
            }
            // Opening tag: read name, then scan attributes quote-aware.
            let name = tagName(in: xml, from: after)
            guard let tag = scanTag(in: xml, from: lt) else { break }
            let carries = attributeValue(named: "w14:paraId",
                                         inTag: xml[tag.attrStart..<tag.gtIndex]) == paraId
            if name == "w:p" {
                if tag.selfClosing {
                    if carries {
                        paragraphSpans.append(ParagraphSpan(
                            range: lt..<tag.end, contentStart: tag.end, selfClosing: true))
                    }
                } else {
                    openParagraphs.append((start: lt, contentStart: tag.end, isTarget: carries))
                }
            } else if carries, otherCarrier == nil {
                otherCarrier = String(name)
            }
            i = tag.end
        }

        switch paragraphSpans.count {
        case 0: return otherCarrier.map { .notOnParagraph(carrier: $0) } ?? .absent
        case 1: return .unique(paragraphSpans[0])
        default: return .duplicate(count: paragraphSpans.count)
        }
    }

    /// The paragraph's visible text: every `<w:t>` / `<w:t …>` element's
    /// content, concatenated and XML-unescaped (named entities AND numeric
    /// character references). `<w:t` is matched strictly (next char `>` or
    /// space) so `<w:tab/>` and friends never match.
    static func paragraphText(in span: ParagraphSpan, xml: String) -> String {
        guard !span.selfClosing else { return "" }
        return collectTextElements(in: String(xml[span.contentStart..<span.range.upperBound]))
    }

    /// Rebuilds the paragraph with its text replaced: the opening tag and the
    /// (depth-aware) `<w:pPr>` block are preserved verbatim, the remaining
    /// children are replaced by a single run carrying the dominant text run's
    /// `<w:rPr>` and one `<w:t xml:space="preserve">` element. Values with
    /// characters forbidden by XML 1.0 are refused.
    static func substituted(span: ParagraphSpan, xml: String,
                            value: String, slotName: String) throws -> String {
        let escaped = try escapedTextValue(value, slotName: slotName)
        let fragment = String(xml[span.range])
        let openingTag: String
        if span.selfClosing {
            // `<w:p …/>` reopens as `<w:p …>`.
            var tag = fragment
            if let slash = tag.range(of: "/>", options: .backwards) {
                tag.replaceSubrange(slash.lowerBound..<tag.endIndex, with: ">")
            }
            openingTag = tag
            return openingTag + runBlock(rPr: "", escaped: escaped) + "</w:p>"
        }
        let contentOffset = xml.distance(from: span.range.lowerBound, to: span.contentStart)
        let tagEnd = fragment.index(fragment.startIndex, offsetBy: contentOffset)
        openingTag = String(fragment[..<tagEnd])
        let content = String(fragment[tagEnd..<fragment.index(fragment.endIndex, offsetBy: -"</w:p>".count)])
        let pPr = leadingPropertiesBlock(name: "w:pPr", in: content)
        let rPr = dominantRunProperties(inContent: content)
        return openingTag + pPr + runBlock(rPr: rPr, escaped: escaped) + "</w:p>"
    }

    /// Rewrites the carried document.xml for each raw-channel slot. Import
    /// re-applies every designation guard (stale directive, duplicate,
    /// non-paragraph carrier, missing binding → loud failure) and, when any
    /// substitution happened, verifies the resulting part is well-formed XML
    /// before letting it into the log. Values equal to the paragraph's
    /// current text leave the XML untouched (identity shortcut).
    static func apply(_ log: OperationLog,
                      rawSlots: [String: String],
                      bindings: [String: String]) throws -> OperationLog {
        var rebuilt = OperationLog()
        for entry in log.entries {
            var op = entry.op
            if case .carryPart(let path, var xml) = op, path == "word/document.xml" {
                var didSubstitute = false
                // Ranges invalidate after each replacement — re-locate per slot.
                for paraId in rawSlots.keys.sorted() {
                    let name = rawSlots[paraId]!
                    guard let value = bindings[name] else {
                        throw TranscodeError.rawSlotExecutionFailure(
                            name: name,
                            reason: "no call-site binding for slot '\(name)' (paraId \(paraId))")
                    }
                    switch locate(paraId: paraId, in: xml) {
                    case .unique(let span):
                        if paragraphText(in: span, xml: xml) == value { continue }
                        let replacement = try substituted(
                            span: span, xml: xml, value: value, slotName: name)
                        xml.replaceSubrange(span.range, with: replacement)
                        didSubstitute = true
                    case .duplicate(let count):
                        throw TranscodeError.rawSlotExecutionFailure(
                            name: name,
                            reason: "paraId \(paraId) is carried by \(count) <w:p> elements — refusing to guess")
                    case .notOnParagraph(let carrier):
                        throw TranscodeError.rawSlotExecutionFailure(
                            name: name,
                            reason: "paraId \(paraId) is carried by <\(carrier)>, not a paragraph")
                    case .absent:
                        throw TranscodeError.rawSlotExecutionFailure(
                            name: name,
                            reason: "paraId \(paraId) not found in the carried document.xml (stale directive?)")
                    }
                }
                if didSubstitute {
                    let parser = XMLParser(data: Data(xml.utf8))
                    guard parser.parse() else {
                        throw TranscodeError.rawSlotExecutionFailure(
                            name: rawSlots.first?.value ?? "?",
                            reason: "post-surgery well-formedness check failed: \(parser.parserError.map(String.init(describing:)) ?? "unknown parser error")")
                    }
                }
                op = .carryPart(partPath: path, xml: xml)
            }
            rebuilt.append(op, source: entry.source, opID: entry.opID, at: entry.timestamp)
        }
        return rebuilt
    }

    // MARK: - tag scanning internals

    private struct TagScan {
        /// Index just past the closing `>`.
        let end: String.Index
        /// Index of the `>` itself.
        let gtIndex: String.Index
        /// Start of the attribute region (just past the tag name).
        let attrStart: String.Index
        let selfClosing: Bool
    }

    /// Reads an element name starting at `start` (first char after `<` or
    /// `</`): letters/digits/`:`/`_`/`-`/`.` per XML Name (ASCII subset —
    /// OOXML names are ASCII).
    private static func tagName(in xml: String, from start: String.Index) -> Substring {
        var j = start
        while j < xml.endIndex {
            let c = xml[j]
            if c == " " || c == "\t" || c == "\n" || c == "\r" || c == ">" || c == "/" { break }
            j = xml.index(after: j)
        }
        return xml[start..<j]
    }

    /// Scans an opening tag starting at its `<`, quote-aware, so `>` inside a
    /// quoted attribute value never terminates the tag.
    private static func scanTag(in xml: String, from lt: String.Index) -> TagScan? {
        let nameStart = xml.index(after: lt)
        let name = tagName(in: xml, from: nameStart)
        var j = xml.index(nameStart, offsetBy: name.count, limitedBy: xml.endIndex) ?? xml.endIndex
        let attrStart = j
        var quote: Character? = nil
        var lastNonQuoteChar: Character = " "
        while j < xml.endIndex {
            let c = xml[j]
            if let q = quote {
                if c == q { quote = nil }
            } else if c == "\"" || c == "'" {
                quote = c
            } else if c == ">" {
                let selfClosing = lastNonQuoteChar == "/"
                return TagScan(end: xml.index(after: j), gtIndex: j,
                               attrStart: attrStart, selfClosing: selfClosing)
            }
            if quote == nil { lastNonQuoteChar = c }
            j = xml.index(after: j)
        }
        return nil
    }

    /// The value of an attribute inside an opening tag's attribute region:
    /// the name must be preceded by whitespace and followed by (optional
    /// whitespace) `=` (optional whitespace) quoted value. Quote-aware so a
    /// VALUE containing the attribute-name text never matches.
    private static func attributeValue(named attrName: String, inTag attrs: Substring) -> String? {
        var j = attrs.startIndex
        while j < attrs.endIndex {
            // Skip whitespace.
            while j < attrs.endIndex, isXMLWhitespace(attrs[j]) { j = attrs.index(after: j) }
            guard j < attrs.endIndex, attrs[j] != "/" else { return nil }
            // Read attribute name.
            let nameStart = j
            while j < attrs.endIndex, !isXMLWhitespace(attrs[j]), attrs[j] != "=" {
                j = attrs.index(after: j)
            }
            let name = attrs[nameStart..<j]
            // Skip whitespace, expect '='.
            while j < attrs.endIndex, isXMLWhitespace(attrs[j]) { j = attrs.index(after: j) }
            guard j < attrs.endIndex, attrs[j] == "=" else { continue }
            j = attrs.index(after: j)
            while j < attrs.endIndex, isXMLWhitespace(attrs[j]) { j = attrs.index(after: j) }
            guard j < attrs.endIndex, attrs[j] == "\"" || attrs[j] == "'" else { return nil }
            let q = attrs[j]
            let valueStart = attrs.index(after: j)
            guard let valueEnd = attrs[valueStart...].firstIndex(of: q) else { return nil }
            if name == attrName[...] {
                return String(attrs[valueStart..<valueEnd])
            }
            j = attrs.index(after: valueEnd)
        }
        return nil
    }

    private static func isXMLWhitespace(_ c: Character) -> Bool {
        c == " " || c == "\t" || c == "\n" || c == "\r"
    }

    private static func skipPast(_ marker: String, from: String.Index, in xml: String) -> String.Index {
        guard let r = xml.range(of: marker, range: from..<xml.endIndex) else { return xml.endIndex }
        return r.upperBound
    }

    // MARK: - fragment internals (operate on a single paragraph's content)

    private static func runBlock(rPr: String, escaped: String) -> String {
        "<w:r>" + rPr + "<w:t xml:space=\"preserve\">" + escaped + "</w:t></w:r>"
    }

    /// The leading properties block (`<w:pPr>…</w:pPr>`) of the paragraph
    /// content, depth-aware: `w:pPrChange` nests another `w:pPr` inside, so
    /// the close is found by depth counting, never first-match.
    private static func leadingPropertiesBlock(name: String, in content: String) -> String {
        var i = content.startIndex
        // Skip whitespace between the opening tag and the first child.
        while i < content.endIndex, isXMLWhitespace(content[i]) { i = content.index(after: i) }
        guard content[i...].hasPrefix("<\(name)") else { return "" }
        let afterName = content.index(i, offsetBy: name.count + 1)
        guard afterName < content.endIndex else { return "" }
        let next = content[afterName]
        guard next == ">" || next == "/" || isXMLWhitespace(next) else { return "" }
        guard let span = elementSpan(name: name, openingAt: i, in: content) else { return "" }
        return String(content[span])
    }

    /// Depth-aware element span: from the opening `<name…>` at `start` to the
    /// matching `</name>` (or the end of a self-closing tag), counting nested
    /// same-name elements.
    private static func elementSpan(name: String, openingAt start: String.Index,
                                    in content: String) -> Range<String.Index>? {
        guard let firstTag = scanTag(in: content, from: start) else { return nil }
        if firstTag.selfClosing { return start..<firstTag.end }
        var depth = 1
        var i = firstTag.end
        while i < content.endIndex, let lt = content[i...].firstIndex(of: "<") {
            if content[lt...].hasPrefix("<!--") { i = skipPast("-->", from: lt, in: content); continue }
            let after = content.index(after: lt)
            guard after < content.endIndex else { return nil }
            if content[after] == "/" {
                let nameStart = content.index(after: after)
                if content[nameStart...].hasPrefix("\(name)>") {
                    depth -= 1
                    let closeEnd = content.index(nameStart, offsetBy: name.count + 1)
                    if depth == 0 { return start..<closeEnd }
                    i = closeEnd
                    continue
                }
                guard let gt = content[after...].firstIndex(of: ">") else { return nil }
                i = content.index(after: gt)
                continue
            }
            let tagNameHere = tagName(in: content, from: after)
            guard let tag = scanTag(in: content, from: lt) else { return nil }
            if tagNameHere == name[...], !tag.selfClosing { depth += 1 }
            i = tag.end
        }
        return nil
    }

    /// The `<w:rPr>…</w:rPr>` block of the DIRECT-child run with the longest
    /// visible text (ties: first). Depth-aware on both the run boundary
    /// (runs inside nested structures are not direct children) and the rPr
    /// close (`w:rPrChange` nests another `w:rPr`).
    private static func dominantRunProperties(inContent content: String) -> String {
        var best = ""
        var bestLength = -1
        var depth = 0
        var i = content.startIndex
        while i < content.endIndex, let lt = content[i...].firstIndex(of: "<") {
            if content[lt...].hasPrefix("<!--") { i = skipPast("-->", from: lt, in: content); continue }
            let after = content.index(after: lt)
            guard after < content.endIndex else { break }
            if content[after] == "/" {
                guard let gt = content[after...].firstIndex(of: ">") else { break }
                depth -= 1
                i = content.index(after: gt)
                continue
            }
            let name = tagName(in: content, from: after)
            guard let tag = scanTag(in: content, from: lt) else { break }
            if depth == 0, name == "w:r", !tag.selfClosing {
                guard let runSpan = elementSpan(name: "w:r", openingAt: lt, in: content) else { break }
                let run = String(content[runSpan])
                let textLength = collectTextElements(in: run).count
                if textLength > bestLength {
                    bestLength = textLength
                    best = ""
                    // First direct <w:rPr> inside the run (depth-aware close).
                    let runContentStart = tag.end
                    var k = runContentStart
                    while k < content.endIndex, k < runSpan.upperBound, isXMLWhitespace(content[k]) {
                        k = content.index(after: k)
                    }
                    if content[k...].hasPrefix("<w:rPr") {
                        if let rPrSpan = elementSpan(name: "w:rPr", openingAt: k, in: content),
                           rPrSpan.upperBound <= runSpan.upperBound {
                            best = String(content[rPrSpan])
                        }
                    }
                }
                i = runSpan.upperBound
                continue
            }
            if !tag.selfClosing { depth += 1 }
            i = tag.end
        }
        return best
    }

    /// Concatenated, unescaped `<w:t>` content of a fragment. `<w:t` matched
    /// strictly (next char `>` or space).
    private static func collectTextElements(in fragment: String) -> String {
        var out = ""
        var search = fragment.startIndex
        while let open = fragment.range(of: "<w:t", range: search..<fragment.endIndex) {
            search = open.upperBound
            guard open.upperBound < fragment.endIndex else { break }
            let next = fragment[open.upperBound]
            guard next == ">" || next == " " else { continue }
            guard let tagEnd = fragment[open.upperBound...].firstIndex(of: ">") else { break }
            if fragment[fragment.index(before: tagEnd)] == "/" {
                search = fragment.index(after: tagEnd); continue
            }
            let contentStart = fragment.index(after: tagEnd)
            guard let close = fragment.range(of: "</w:t>", range: contentStart..<fragment.endIndex) else { break }
            out += xmlUnescape(String(fragment[contentStart..<close.lowerBound]))
            search = close.upperBound
        }
        return out
    }

    // MARK: - escaping

    /// Escapes a slot value for text-node position, refusing characters
    /// forbidden by XML 1.0 (they cannot be represented in a well-formed
    /// part, even escaped).
    private static func escapedTextValue(_ value: String, slotName: String) throws -> String {
        for scalar in value.unicodeScalars {
            let v = scalar.value
            let legal = v == 0x9 || v == 0xA || v == 0xD
                || (v >= 0x20 && v <= 0xD7FF)
                || (v >= 0xE000 && v <= 0xFFFD)
                || (v >= 0x10000 && v <= 0x10FFFF)
            guard legal else {
                throw TranscodeError.rawSlotExecutionFailure(
                    name: slotName,
                    reason: "value contains a character forbidden by XML 1.0 (U+\(String(v, radix: 16, uppercase: true)))")
            }
        }
        return value.replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
    }

    /// Single-pass entity decoding: the five named entities plus numeric
    /// character references (`&#dd;` / `&#xhh;`). A decoded `&` is never
    /// re-scanned, so `&amp;#65;` correctly yields the literal `&#65;`.
    private static func xmlUnescape(_ s: String) -> String {
        guard s.contains("&") else { return s }
        var out = ""
        var i = s.startIndex
        while i < s.endIndex {
            let c = s[i]
            guard c == "&" else { out.append(c); i = s.index(after: i); continue }
            guard let semi = s[i...].firstIndex(of: ";") else { out.append(c); i = s.index(after: i); continue }
            let entity = s[s.index(after: i)..<semi]
            var decoded: String? = nil
            switch entity {
            case "amp": decoded = "&"
            case "lt": decoded = "<"
            case "gt": decoded = ">"
            case "quot": decoded = "\""
            case "apos": decoded = "'"
            default:
                if entity.hasPrefix("#x") || entity.hasPrefix("#X") {
                    if let v = UInt32(entity.dropFirst(2), radix: 16), let sc = Unicode.Scalar(v) {
                        decoded = String(Character(sc))
                    }
                } else if entity.hasPrefix("#") {
                    if let v = UInt32(entity.dropFirst()), let sc = Unicode.Scalar(v) {
                        decoded = String(Character(sc))
                    }
                }
            }
            if let d = decoded {
                out += d
                i = s.index(after: semi)
            } else {
                out.append(c)
                i = s.index(after: i)
            }
        }
        return out
    }
}
