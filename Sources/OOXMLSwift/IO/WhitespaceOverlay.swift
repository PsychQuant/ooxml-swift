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

    /// Scan the raw XML byte stream for `<w:t xml:space="preserve">[whitespace]</w:t>`
    /// patterns. Records (sequence-index, whitespace-content) pairs keyed by the
    /// element's position in DOM document order — i.e., the same order that
    /// `XMLDocument`'s `elements(forName: "w:t")` yields when walking the DOM.
    init(scanning data: Data) {
        guard let xml = String(data: data, encoding: .utf8) else {
            self.whitespaceByIndex = [:]
            return
        }
        var map: [Int: String] = [:]

        // Walk every `<w:t` opening tag in source order. For each one, classify:
        //   - has `xml:space="preserve"`? → potential whitespace candidate
        //   - inner text is whitespace-only? → record in map
        //   - other → skip (caller's `stringValue` works fine)
        // Sequence index is the running count of `<w:t` opening tags encountered.
        var index = 0
        var searchStart = xml.startIndex
        while let openRange = xml.range(of: "<w:t", range: searchStart..<xml.endIndex) {
            // Advance the cursor past the `<w:t` token; we'll find the closing `>`
            // of the tag (could be `>` for `<w:t>` or `<w:t xml:space="preserve">`,
            // or `/>` for `<w:t/>` self-close — though self-close means no content).
            let afterToken = openRange.upperBound

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
            if tagAttrs.contains("xml:space=\"preserve\""),
               !innerText.isEmpty,
               innerText.allSatisfy({ $0.isWhitespace }) {
                map[index] = innerText
            }

            index += 1
            searchStart = closeRange.upperBound
        }

        self.whitespaceByIndex = map
    }

    /// Recover whitespace text for the `<w:t>` element at the given sequence
    /// index in DOM document order. Returns nil if the element is not in the
    /// overlay map (i.e., it's not a whitespace-only `xml:space="preserve"`
    /// element — caller should fall back to `XMLElement.stringValue`).
    func text(forElementSequenceIndex index: Int) -> String? {
        return whitespaceByIndex[index]
    }

    /// Diagnostic — returns the count of recovered whitespace entries.
    var recoveredCount: Int {
        return whitespaceByIndex.count
    }
}
