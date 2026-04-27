import Foundation

/// v0.19.10+ (#59 sub-stack B): whitespace preservation overlay for OOXML
/// `<w:t xml:space="preserve">[whitespace]</w:t>` text nodes.
///
/// **Background**: Foundation `XMLDocument` strips whitespace-only text node
/// `stringValue` to "" regardless of the `xml:space="preserve"` attribute AND
/// regardless of `XMLNode.Options.nodePreserveWhitespace` parse option (verified
/// by isolated probe in [#59 diagnosis](https://github.com/PsychQuant/che-word-mcp/issues/59)).
/// This is a structural limitation of Foundation's libxml2-backed parser on
/// macOS — not a configuration bug.
///
/// **Approach**: pre-parse byte-stream scan over raw OOXML XML bytes. For each
/// `<w:t xml:space="preserve">[whitespace]</w:t>` element encountered in DOM
/// document order, record the whitespace content keyed by element sequence
/// index. `parseRun` consults the overlay when `t.stringValue.isEmpty` to
/// recover the lost whitespace bytes.
///
/// **Why not switch parsers**: 1-2 weeks of work + new dependency + affects all
/// 10 `XMLDocument(data:)` call sites in DocxReader.swift. Whitespace overlay
/// is contained, surgical, and follows the same architectural pattern as
/// `WordDocument.modifiedParts` overlay (the v0.13.0 byte-preservation
/// architecture).
///
/// **Limitation**: only handles `<w:t>` text nodes with the `xml:space="preserve"`
/// attribute (the standard OOXML form for whitespace-significant text). Bare
/// `<w:t>...</w:t>` without the attribute is not in scope (Word always emits
/// the attribute when text starts/ends with whitespace).
internal struct WhitespaceOverlay {
    /// Map: element sequence index in DOM document order → recovered whitespace text.
    /// Only populated for `<w:t xml:space="preserve">[whitespace-only]</w:t>` elements.
    /// Non-whitespace `<w:t>` elements are NOT in the map — caller falls back to
    /// `t.stringValue` for those.
    private let whitespaceByIndex: [Int: String]

    /// v0.19.11+ (#59 B-CONT P1, R5 finding): same overlay mechanism for
    /// `<w:delText xml:space="preserve">[whitespace-only]</w:delText>`.
    /// Tracked-deletion of whitespace was permanently lost pre-fix because
    /// (a) overlay only scanned `<w:t`, and (b) the `<w:del>` parser path
    /// at `DocxReader.swift:970` reads `delText.stringValue` directly with
    /// no overlay consult. Indexed independently from `whitespaceByIndex`
    /// since `<w:t>` and `<w:delText>` element counts are independent.
    private let delTextWhitespaceByIndex: [Int: String]

    /// Scan the raw XML byte stream for `<w:t xml:space="preserve">[whitespace]</w:t>`
    /// patterns. Records (sequence-index, whitespace-content) pairs keyed by the
    /// element's position in DOM document order — i.e., the same order that
    /// `XMLDocument`'s `elements(forName: "w:t")` yields when walking the DOM.
    init(scanning data: Data) {
        guard let xml = String(data: data, encoding: .utf8) else {
            self.whitespaceByIndex = [:]
            self.delTextWhitespaceByIndex = [:]
            return
        }
        var map: [Int: String] = [:]
        var delTextMap: [Int: String] = [:]

        // Walk every `<w:t` opening tag in source order. For each one, classify:
        //   - has `xml:space="preserve"`? → potential whitespace candidate
        //   - inner text is whitespace-only? → record in map
        //   - other → skip (caller's `stringValue` works fine)
        // Sequence index is the running count of `<w:t` opening tags encountered.
        //
        // v0.19.11+ (#59 B-CONT P0-A): tag-name boundary check. The bare prefix
        // scan `<w:t` falsely matches every OOXML element whose qualified name
        // starts with `w:t` — `<w:tab>`, `<w:tbl>`, `<w:tc>`, `<w:tr>`,
        // `<w:tblPr>`, `<w:tcPr>`, `<w:trPr>`, `<w:tblGrid>`, `<w:trHeight>`,
        // `<w:tblBorders>`, `<w:tcBorders>`, `<w:tblCellMar>`, `<w:tblLayout>`,
        // `<w:tblLook>`, `<w:tblStyle>`, etc. The DOM walker
        // `element.elements(forName: "w:t")` is exact-match. Without a boundary
        // check, scanner index desyncs immediately in any document with tables
        // or tabs (i.e., every real Word file). Triple-confirmed by sub-stack B
        // 6-AI verify (R2 + R5 + Codex). Fix: peek the character after `<w:t`
        // and only count when it's a valid tag-name terminator.
        var index = 0
        var searchStart = xml.startIndex
        while let openRange = xml.range(of: "<w:t", range: searchStart..<xml.endIndex) {
            let afterToken = openRange.upperBound

            // P0-A (#59 B-CONT): tag-name boundary check.
            guard afterToken < xml.endIndex else { break }
            let boundary = xml[afterToken]
            guard boundary == ">" || boundary == " " || boundary == "\t"
                  || boundary == "\n" || boundary == "\r" || boundary == "/" else {
                // Prefix collision (e.g. `<w:tab>`, `<w:tbl>`). Advance past
                // this match without incrementing index — DOM walker won't see
                // it as a `<w:t>` element either.
                searchStart = afterToken
                continue
            }

            // Find tag close: scan forward for `>`.
            guard let tagClose = xml.range(of: ">", range: afterToken..<xml.endIndex) else {
                searchStart = afterToken
                continue
            }
            let tagAttrs = String(xml[afterToken..<tagClose.lowerBound])

            // Self-close (`<w:t/>`): empty element, no content to recover.
            if tagAttrs.hasSuffix("/") {
                index += 1
                searchStart = tagClose.upperBound
                continue
            }

            // Find closing `</w:t>`.
            guard let closeRange = xml.range(of: "</w:t>", range: tagClose.upperBound..<xml.endIndex) else {
                searchStart = tagClose.upperBound
                continue
            }
            let innerText = String(xml[tagClose.upperBound..<closeRange.lowerBound])

            // Only record if (a) `xml:space="preserve"` is present AND (b) inner
            // text is non-empty AND all-whitespace. The xml:space check rules out
            // typical text content; the all-whitespace check rules out content
            // that Foundation will round-trip correctly anyway.
            //
            // v0.19.11+ (#59 B-CONT P1, R5 finding): XML-decode entity-encoded
            // whitespace before the whitespace check. Pre-fix the check ran on
            // raw bytes — `&#x09;&#x09;` (two tabs) sees `&`, `#`, `x`, `0`,
            // `9` which aren't Character.isWhitespace, so the element wasn't
            // stored. Foundation later decodes and then strips the resulting
            // whitespace stringValue → permanent loss.
            if tagAttrs.contains("xml:space=\"preserve\"") {
                let decoded = Self.decodeXMLEntities(in: innerText)
                if !decoded.isEmpty, decoded.allSatisfy({ $0.isWhitespace }) {
                    // Store the DECODED whitespace text — that's what
                    // parseRun's overlay-consult should hand to the Run model.
                    map[index] = decoded
                }
            }

            index += 1
            searchStart = closeRange.upperBound
        }

        // Second pass: scan `<w:delText` elements (mirror of the `<w:t>` loop
        // above). Same boundary check + same xml:space + decoded-whitespace
        // logic; independent counter because DOM walks them via a separate
        // `elements(forName: "w:delText")` query, not interleaved with `<w:t>`.
        var delIndex = 0
        var delSearchStart = xml.startIndex
        while let openRange = xml.range(of: "<w:delText", range: delSearchStart..<xml.endIndex) {
            let afterToken = openRange.upperBound
            guard afterToken < xml.endIndex else { break }
            let boundary = xml[afterToken]
            guard boundary == ">" || boundary == " " || boundary == "\t"
                  || boundary == "\n" || boundary == "\r" || boundary == "/" else {
                delSearchStart = afterToken
                continue
            }
            guard let tagClose = xml.range(of: ">", range: afterToken..<xml.endIndex) else {
                delSearchStart = afterToken
                continue
            }
            let tagAttrs = String(xml[afterToken..<tagClose.lowerBound])
            if tagAttrs.hasSuffix("/") {
                delIndex += 1
                delSearchStart = tagClose.upperBound
                continue
            }
            guard let closeRange = xml.range(of: "</w:delText>", range: tagClose.upperBound..<xml.endIndex) else {
                delSearchStart = tagClose.upperBound
                continue
            }
            let innerText = String(xml[tagClose.upperBound..<closeRange.lowerBound])
            if tagAttrs.contains("xml:space=\"preserve\"") {
                let decoded = Self.decodeXMLEntities(in: innerText)
                if !decoded.isEmpty, decoded.allSatisfy({ $0.isWhitespace }) {
                    delTextMap[delIndex] = decoded
                }
            }
            delIndex += 1
            delSearchStart = closeRange.upperBound
        }

        self.whitespaceByIndex = map
        self.delTextWhitespaceByIndex = delTextMap
    }

    /// v0.19.11+ (#59 B-CONT P0-B): count `<w:t>` opening tags in a raw XML
    /// substring, applying the same boundary check as the main scanner.
    ///
    /// Used by `DocxReader` raw-capture call sites (`parseInsRevisionWrapper`
    /// non-run-child path, `parseAlternateContent` `<mc:Choice>` skip,
    /// generic `.rawBlockElement` capture) to advance
    /// `WhitespaceParseContext.counter` by the number of `<w:t>` elements
    /// in the subtree the parser is about to skip.
    ///
    /// Without this, scanner counts the skipped `<w:t>` elements during
    /// pre-scan but parser never visits them — counter desyncs and every
    /// subsequent `parseRun` overlay lookup queries the wrong index.
    /// Triple-confirmed by sub-stack B 6-AI verify (R2 + Codex).
    /// v0.19.11+ (#59 B-CONT P1, R5 finding): decode XML character entities
    /// in a small inner-text fragment so the all-whitespace check sees the
    /// actual decoded characters, not raw entity bytes.
    ///
    /// Handles:
    ///   - Numeric decimal: `&#9;`, `&#10;`, `&#13;`, `&#32;` etc.
    ///   - Numeric hex: `&#x09;`, `&#x0A;`, `&#x0D;`, `&#x20;`, `&#xA0;` (NBSP) etc.
    ///   - Named: `&nbsp;` (XML 1.0 spec doesn't define it but some Word emitters use)
    ///
    /// Other entities (`&amp;`, `&lt;`, `&gt;`, `&quot;`, `&apos;`) are NOT
    /// whitespace so leaving them un-decoded doesn't affect the whitespace
    /// check (any `&` that's not part of a recognized whitespace entity makes
    /// the check return false correctly — non-whitespace).
    static func decodeXMLEntities(in xml: String) -> String {
        // Fast path: if there's no `&`, the input is already decoded.
        guard xml.contains("&") else { return xml }

        var result = ""
        result.reserveCapacity(xml.count)
        var i = xml.startIndex

        while i < xml.endIndex {
            if xml[i] == "&", let semicolon = xml.range(of: ";", range: i..<xml.endIndex) {
                let entity = String(xml[xml.index(after: i)..<semicolon.lowerBound])
                if let scalar = Self.scalarForEntity(entity) {
                    result.append(Character(scalar))
                    i = semicolon.upperBound
                    continue
                }
            }
            result.append(xml[i])
            i = xml.index(after: i)
        }
        return result
    }

    /// Resolve an XML entity body (the part between `&` and `;`) to a Unicode
    /// scalar. Returns nil for unrecognized entities (caller leaves them as-is).
    private static func scalarForEntity(_ entity: String) -> Unicode.Scalar? {
        if entity.hasPrefix("#x") || entity.hasPrefix("#X") {
            let hex = entity.dropFirst(2)
            guard let value = UInt32(hex, radix: 16) else { return nil }
            return Unicode.Scalar(value)
        }
        if entity.hasPrefix("#") {
            let dec = entity.dropFirst()
            guard let value = UInt32(dec, radix: 10) else { return nil }
            return Unicode.Scalar(value)
        }
        // Named entities — only handle `nbsp` since it's whitespace-relevant.
        if entity == "nbsp" {
            return Unicode.Scalar(0xA0)
        }
        return nil
    }

    static func countWtElements(in xml: String) -> Int {
        var count = 0
        var searchStart = xml.startIndex
        while let openRange = xml.range(of: "<w:t", range: searchStart..<xml.endIndex) {
            let afterToken = openRange.upperBound
            guard afterToken < xml.endIndex else { break }
            let boundary = xml[afterToken]
            if boundary == ">" || boundary == " " || boundary == "\t"
                || boundary == "\n" || boundary == "\r" || boundary == "/" {
                count += 1
            }
            searchStart = afterToken
        }
        return count
    }

    /// v0.19.11+ (#59 B-CONT P1): mirror of `countWtElements` for
    /// `<w:delText>` elements. Used to advance `WhitespaceParseContext.delTextCounter`
    /// at raw-capture sites that may contain `<w:delText>` elements (e.g.,
    /// `<w:del>` with non-run children gets raw-captured along with its
    /// inner `<w:delText>` elements).
    static func countDelTextElements(in xml: String) -> Int {
        var count = 0
        var searchStart = xml.startIndex
        while let openRange = xml.range(of: "<w:delText", range: searchStart..<xml.endIndex) {
            let afterToken = openRange.upperBound
            guard afterToken < xml.endIndex else { break }
            let boundary = xml[afterToken]
            if boundary == ">" || boundary == " " || boundary == "\t"
                || boundary == "\n" || boundary == "\r" || boundary == "/" {
                count += 1
            }
            searchStart = afterToken
        }
        return count
    }

    /// Recover whitespace text for the `<w:t>` element at the given sequence
    /// index in DOM document order. Returns nil if the element is not in the
    /// overlay map (i.e., it's not a whitespace-only `xml:space="preserve"`
    /// element — caller should fall back to `XMLElement.stringValue`).
    func text(forElementSequenceIndex index: Int) -> String? {
        return whitespaceByIndex[index]
    }

    /// v0.19.11+ (#59 B-CONT P1): recover whitespace text for the
    /// `<w:delText>` element at the given sequence index in DOM document order.
    /// Independent index from `text(forElementSequenceIndex:)` — DOM walks
    /// `<w:delText>` separately via `elements(forName: "w:delText")`.
    func delText(forElementSequenceIndex index: Int) -> String? {
        return delTextWhitespaceByIndex[index]
    }

    /// Diagnostic — returns the count of recovered `<w:t>` whitespace entries.
    var recoveredCount: Int {
        return whitespaceByIndex.count
    }
}
