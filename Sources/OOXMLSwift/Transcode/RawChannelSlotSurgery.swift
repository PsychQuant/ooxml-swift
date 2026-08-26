// RawChannelSlotSurgery.swift
// raw-channel-slot-support (#171) — run-level text substitution inside a
// carried word/document.xml part.
//
// A document whose main part fails the DSL upgrade rides the raw channel as
// one `.carryPart` op; no paragraph-level ops exist, so DSL and op-level
// slots cannot reach its paragraphs. This surgery locates a paragraph by
// `w14:paraId` inside the carried XML string, extracts its visible text (the
// slot default), and — only when the substituted value differs from the
// current text (identity shortcut) — replaces the paragraph's runs with a
// single run carrying the dominant text run's rPr. Everything outside the
// designated paragraph stays byte-identical, so an all-default execution
// reproduces the reference byte-equal and the existing Stage B acceptance
// applies unchanged.

import Foundation

enum RawChannelSlotSurgery {

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

    /// Number of times the exact `w14:paraId="<id>"` token occurs in the XML.
    static func occurrences(paraId: String, in xml: String) -> Int {
        xml.components(separatedBy: token(paraId)).count - 1
    }

    /// The full `<w:p …>…</w:p>` (or self-closing `<w:p …/>`) fragment range
    /// of the paragraph carrying the paraId, or nil when absent. `w:p`
    /// elements never nest, so the first `</w:p>` after the opening tag closes
    /// this paragraph.
    static func paragraphFragmentRange(paraId: String, in xml: String) -> Range<String.Index>? {
        guard let tokenRange = xml.range(of: token(paraId)) else { return nil }
        guard let open = xml.range(of: "<w:p ", options: .backwards,
                                   range: xml.startIndex..<tokenRange.lowerBound) else { return nil }
        guard let tagEnd = xml[tokenRange.upperBound...].firstIndex(of: ">") else { return nil }
        // Self-closing paragraph: `<w:p …/>` — the opening tag is the fragment.
        if xml.index(before: tagEnd) >= tokenRange.upperBound, xml[xml.index(before: tagEnd)] == "/" {
            return open.lowerBound..<xml.index(after: tagEnd)
        }
        guard let close = xml.range(of: "</w:p>", range: tagEnd..<xml.endIndex) else { return nil }
        return open.lowerBound..<close.upperBound
    }

    /// The paragraph's visible text: every `<w:t>` / `<w:t …>` element's
    /// content, concatenated and XML-unescaped. `<w:t` is matched strictly
    /// (next char `>` or space) so `<w:tab/>` and friends never match.
    static func paragraphText(fragment: String) -> String {
        var out = ""
        var search = fragment.startIndex
        while let open = fragment.range(of: "<w:t", range: search..<fragment.endIndex) {
            search = open.upperBound
            guard open.upperBound < fragment.endIndex else { break }
            let next = fragment[open.upperBound]
            guard next == ">" || next == " " else { continue }
            guard let tagEnd = fragment[open.upperBound...].firstIndex(of: ">") else { break }
            // Self-closing `<w:t/>` carries no text.
            if fragment[fragment.index(before: tagEnd)] == "/" { search = fragment.index(after: tagEnd); continue }
            let contentStart = fragment.index(after: tagEnd)
            guard let close = fragment.range(of: "</w:t>", range: contentStart..<fragment.endIndex) else { break }
            out += xmlUnescape(String(fragment[contentStart..<close.lowerBound]))
            search = close.upperBound
        }
        return out
    }

    /// Rebuilds the paragraph with its text replaced: the opening tag and
    /// `<w:pPr>` block are preserved verbatim, the runs collapse to a single
    /// run carrying the dominant text run's `<w:rPr>` and one
    /// `<w:t xml:space="preserve">` element with the escaped value.
    static func substituted(fragment: String, value: String) -> String {
        let openingTag: String
        if let tagEnd = fragment.firstIndex(of: ">") {
            if fragment[fragment.index(before: tagEnd)] == "/" {
                // Self-closing paragraph re-opens: `<w:p …/>` → `<w:p …>`.
                openingTag = String(fragment[..<fragment.index(before: tagEnd)]) + ">"
            } else {
                openingTag = String(fragment[...tagEnd])
            }
        } else {
            openingTag = fragment
        }
        var pPr = ""
        if let open = fragment.range(of: "<w:pPr>"),
           let close = fragment.range(of: "</w:pPr>", range: open.upperBound..<fragment.endIndex) {
            pPr = String(fragment[open.lowerBound..<close.upperBound])
        } else if let selfClosing = fragment.range(of: "<w:pPr/>") {
            pPr = String(fragment[selfClosing])
        }
        let rPr = dominantRunProperties(fragment: fragment)
        return openingTag + pPr
            + "<w:r>" + rPr
            + "<w:t xml:space=\"preserve\">" + xmlEscape(value) + "</w:t></w:r></w:p>"
    }

    /// Rewrites the carried document.xml for each raw-channel slot whose
    /// call-site value differs from the paragraph's current text. Values equal
    /// to the current text leave the XML untouched (identity shortcut) — the
    /// all-default execution therefore stays byte-equal by construction.
    static func apply(_ log: OperationLog,
                      rawSlots: [String: String],
                      bindings: [String: String]) -> OperationLog {
        var rebuilt = OperationLog()
        for entry in log.entries {
            var op = entry.op
            if case .carryPart(let path, var xml) = op, path == "word/document.xml" {
                // Ranges invalidate after each replacement — re-locate per slot.
                for paraId in rawSlots.keys.sorted() {
                    guard let name = rawSlots[paraId], let value = bindings[name],
                          let range = paragraphFragmentRange(paraId: paraId, in: xml) else { continue }
                    let fragment = String(xml[range])
                    if paragraphText(fragment: fragment) == value { continue }
                    xml.replaceSubrange(range, with: substituted(fragment: fragment, value: value))
                }
                op = .carryPart(partPath: path, xml: xml)
            }
            rebuilt.append(op, source: entry.source, opID: entry.opID, at: entry.timestamp)
        }
        return rebuilt
    }

    // MARK: - internals

    private static func token(_ paraId: String) -> String {
        "w14:paraId=\"\(paraId)\""
    }

    /// The `<w:rPr>…</w:rPr>` block of the run with the longest visible text
    /// (ties: first). Empty when the paragraph has no runs or the dominant run
    /// carries no rPr. `<w:r` is matched strictly (next char `>` or space) so
    /// `<w:rPr>` never matches as a run opening.
    private static func dominantRunProperties(fragment: String) -> String {
        var best = ""
        var bestLength = -1
        var search = fragment.startIndex
        while let open = fragment.range(of: "<w:r", range: search..<fragment.endIndex) {
            search = open.upperBound
            guard open.upperBound < fragment.endIndex else { break }
            let next = fragment[open.upperBound]
            guard next == ">" || next == " " else { continue }
            guard let close = fragment.range(of: "</w:r>", range: open.upperBound..<fragment.endIndex) else { break }
            let run = String(fragment[open.lowerBound..<close.upperBound])
            let textLength = paragraphText(fragment: run).count
            if textLength > bestLength {
                bestLength = textLength
                if let rPrOpen = run.range(of: "<w:rPr>"),
                   let rPrClose = run.range(of: "</w:rPr>", range: rPrOpen.upperBound..<run.endIndex) {
                    best = String(run[rPrOpen.lowerBound..<rPrClose.upperBound])
                } else {
                    best = ""
                }
            }
            search = close.upperBound
        }
        return best
    }

    private static func xmlEscape(_ s: String) -> String {
        s.replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
    }

    private static func xmlUnescape(_ s: String) -> String {
        s.replacingOccurrences(of: "&lt;", with: "<")
            .replacingOccurrences(of: "&gt;", with: ">")
            .replacingOccurrences(of: "&quot;", with: "\"")
            .replacingOccurrences(of: "&apos;", with: "'")
            .replacingOccurrences(of: "&amp;", with: "&")
    }
}
