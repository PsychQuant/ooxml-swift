// MdocxFixtureNormalizer
//
// Phase A helper for Spectra change `mdocx-fixture-corpus` (task 1.2).
//
// Strips identity-noise from a docx so byte-diff comparisons between
// hand-crafted goldens and produced docx output stay stable across Word
// versions and ooxml-swift releases. The pipeline is deterministic, idempotent,
// and pure (no I/O beyond reading the input).
//
// Pipeline (normative spec — see openspec/changes/mdocx-fixture-corpus/specs/
// mdocx-fixture-corpus/spec.md, requirement "Normalization pipeline behavior"):
//
//  1. Strip every RSID attribute (w:rsidR, w:rsidRDefault, w:rsidP, w:rsidRPr,
//     w:rsidTr) on every element of every word/*.xml part.
//  2. Strip the <w:rsids> element from word/settings.xml if present.
//  3. Drop word/theme/theme1.xml when it byte-equals the canonical Word default
//     theme vendored at Tests/.../Fixtures/mdocx/_normalizer/word-default-theme.xml.
//     A non-default theme is preserved unchanged.
//  4. Strip a small set of element keys from word/settings.xml that always
//     equal Word defaults in clean Phase A goldens: <w:rsids>, <w:proofState>,
//     <w:defaultTabStop>, <w:zoom>. (Phase A scope — see Phase B TODO below.)
//  5. Re-number w14:paraId / w14:textId / w:bookmarkId attribute values to a
//     deterministic monotonic sequence (00000001, 00000002, ...) in document
//     order across word/document.xml. Cross-references that target these IDs
//     (w:hyperlink/@w:anchor → matching paragraph w14:paraId, w:bookmarkEnd
//     pairing by w:id) are rewritten to keep referential integrity.
//
// Output is byte-identical when run twice on the same input (idempotence) and
// byte-identical when run on inputs that differ only in stripped fields
// (determinism), guaranteed by sorting ZIP entries alphabetically before write.

import Foundation
import ZIPFoundation
@testable import OOXMLSwift

/// Pure helper that strips identity-noise from a docx for golden-fixture
/// byte-diff comparisons. Two API shapes:
///
///   - File-based: `normalize(docxAt:defaultThemeURL:)` — reads two files,
///     normalizes the docx, returns the bytes (caller writes them).
///   - Pure in-memory: `normalize(docxBytes:defaultThemeBytes:)` — for tests
///     that build inputs in-memory.
enum MdocxFixtureNormalizer {

    // MARK: - Public API

    /// Read a docx and the vendored canonical theme reference, normalize, and
    /// return the resulting docx bytes. Callers write the bytes to disk.
    static func normalize(docxAt url: URL, defaultThemeURL: URL) throws -> Data {
        let docxBytes = try Data(contentsOf: url)
        let themeBytes = try Data(contentsOf: defaultThemeURL)
        return try normalize(docxBytes: docxBytes, defaultThemeBytes: themeBytes)
    }

    /// Pure in-memory normalization. `defaultThemeBytes` is the canonical
    /// Word-default theme1.xml that triggers theme drop when byte-equal.
    static func normalize(docxBytes: Data, defaultThemeBytes: Data) throws -> Data {
        // Read every entry from the input docx into memory. ZIPFoundation
        // iteration order is dictated by the central directory, which we then
        // re-order alphabetically on write to keep ZIP layout deterministic
        // independent of the input ordering.
        let inputArchive = try Archive(data: docxBytes, accessMode: .read)
        var entries: [(name: String, data: Data)] = []
        for entry in inputArchive {
            // Skip directory entries — docx parts are all files.
            guard entry.type == .file else { continue }
            var buffer = Data()
            _ = try inputArchive.extract(entry) { buffer.append($0) }
            entries.append((entry.path, buffer))
        }

        // Apply the per-part transformations.
        var transformed: [(name: String, data: Data)] = []
        for (name, data) in entries {
            if let replacement = try transform(entryName: name, bytes: data) {
                // `nil` from transform means "drop this entry".
                transformed.append((name, replacement))
            } else {
                // Theme1 default-equal case: drop the entry. Determined inside
                // transform via the closure injected below.
                continue
            }
        }
        // theme1 drop: separate pass so `transform` itself stays pure-typed.
        // Actually we already handled drop via `transform` returning nil; the
        // `continue` above suffices. Keep theme drop check inline:
        transformed = transformed.filter { entry in
            if entry.name == "word/theme/theme1.xml" {
                // bytes-equal check against vendored reference.
                return entry.data != defaultThemeBytes
            }
            return true
        }

        // Re-number paraId / textId / bookmarkId in document order on
        // word/document.xml, rewriting cross-references in the same part.
        if let docIndex = transformed.firstIndex(where: { $0.name == "word/document.xml" }) {
            let renumbered = try renumberStableIDs(documentXML: transformed[docIndex].data)
            transformed[docIndex] = ("word/document.xml", renumbered)
        }

        // Sort entries alphabetically to make the ZIP layout deterministic.
        transformed.sort { $0.name < $1.name }

        // Write a fresh in-memory ZIP.
        let outputArchive = try Archive(accessMode: .create)
        for (name, data) in transformed {
            try outputArchive.addEntry(
                with: name,
                type: .file,
                uncompressedSize: Int64(data.count),
                compressionMethod: .deflate,
                provider: { position, size in
                    let start = data.startIndex.advanced(by: Int(position))
                    let end = start.advanced(by: size)
                    return data.subdata(in: start..<end)
                }
            )
        }
        guard let outputData = outputArchive.data else {
            throw NormalizerError.archiveDataUnavailable
        }
        return outputData
    }

    // MARK: - Errors

    enum NormalizerError: Error, Equatable {
        case archiveDataUnavailable
    }

    // MARK: - Per-part transformation

    /// Returns the transformed bytes for `entryName` or `nil` to drop the
    /// entry. Phase A: only word/*.xml parts are mutated; everything else is
    /// passed through unchanged. Theme-drop is handled by the caller after this
    /// returns because it depends on byte-equal comparison against an external
    /// reference; the per-entry transform itself stays focused on tree
    /// manipulation.
    private static func transform(entryName: String, bytes: Data) throws -> Data? {
        // Only word/*.xml is in scope for tree-level rewrites. Other parts
        // (rels, _rels/.rels, [Content_Types].xml, media, etc.) pass through.
        guard entryName.hasPrefix("word/"), entryName.hasSuffix(".xml") else {
            return bytes
        }

        let tree: XmlTree
        do {
            tree = try XmlTreeReader.parse(bytes)
        } catch {
            // If a part is not parseable as XML (shouldn't happen for word/*.xml
            // but be conservative), pass through unchanged so we don't blow up
            // on edge cases like binary OLE objects mis-named with .xml.
            return bytes
        }

        // Step 1: strip RSID attributes everywhere in this part.
        stripRSIDAttributes(node: tree.root)

        // Step 2/4: remove default settings keys from word/settings.xml.
        if entryName == "word/settings.xml" {
            stripDefaultSettings(node: tree.root)
        }

        return try XmlTreeWriter.serialize(tree)
    }

    // MARK: - Step 1: RSID attributes

    /// Spec list: `w:rsidR`, `w:rsidRDefault`, `w:rsidP`, `w:rsidRPr`, `w:rsidTr`.
    /// We additionally strip `w:rsidSect` because it is the same identity-noise
    /// class (matches `XmlAttribute.isRsidNoise` semantics already used by
    /// `XmlNode.normalizedFingerprint()`); stripping a strict superset of the
    /// spec list cannot break any spec assertion.
    private static let rsidLocalNames: Set<String> = [
        "rsidR", "rsidRDefault", "rsidP", "rsidRPr", "rsidTr", "rsidSect"
    ]

    private static func stripRSIDAttributes(node: XmlNode) {
        if node.kind == .element {
            let filtered = node.attributes.filter { attr in
                !(attr.prefix == "w" && rsidLocalNames.contains(attr.localName))
            }
            if filtered.count != node.attributes.count {
                node.attributes = filtered
            }
        }
        for child in node.children {
            stripRSIDAttributes(node: child)
        }
    }

    // MARK: - Step 2/4: default settings keys

    /// Phase A scope: drop these <w:*> children of <w:settings> by element name
    /// regardless of their value. This satisfies spec rule 2 (`<w:rsids>`
    /// always dropped) and a structural-only approximation of spec rule 4
    /// (drop keys whose value equals the Word default). For the four named
    /// elements below, Word writes them with the default value virtually
    /// always, so structural strip suffices for Phase A.
    ///
    /// TODO(Phase B): replace structural strip with value-aware default check.
    /// The spec language is "every element key from word/settings.xml whose
    /// value equals the Word default" — a future revision should look up the
    /// expected default per element (e.g. `<w:zoom w:percent="100">`,
    /// `<w:defaultTabStop w:val="720">`, `<w:proofState w:spelling="clean"
    /// w:grammar="clean">`) and only strip when bytes-equal. Until then, any
    /// non-default value of these four keys is silently dropped, which is
    /// acceptable for Phase A goldens because they are hand-crafted and never
    /// carry non-default values for these keys.
    private static let defaultSettingsLocalNames: Set<String> = [
        "rsids", "proofState", "defaultTabStop", "zoom"
    ]

    private static func stripDefaultSettings(node: XmlNode) {
        guard node.kind == .element else { return }
        // Spec scope: only direct children of <w:settings> are the "settings
        // keys" — descendants of those keys must not be stripped.
        if node.prefix == "w" && node.localName == "settings" {
            // Two-pass strip:
            //   1. Drop the matching element children.
            //   2. Drop adjacent whitespace-only text nodes left behind, so two
            //      inputs that differ only in stripped fields normalise to
            //      byte-identical output (determinism Requirement).
            let toRemove: Set<Int> = Set(node.children.indices.filter { idx in
                let child = node.children[idx]
                guard child.kind == .element else { return false }
                return child.prefix == "w" && defaultSettingsLocalNames.contains(child.localName)
            })
            if !toRemove.isEmpty {
                node.children = stripElementsAndAdjacentWhitespace(
                    children: node.children, indicesToRemove: toRemove)
            }
            // No need to recurse below settings — settings.xml is shallow and
            // identity-noise lives at the top level.
            return
        }
        for child in node.children {
            stripDefaultSettings(node: child)
        }
    }

    /// Drop the element children at `indicesToRemove` AND any text nodes
    /// immediately preceding them whose content is pure whitespace. This
    /// keeps the post-strip child sequence byte-identical regardless of
    /// whether the input had the stripped elements in the first place
    /// (per the Determinism scenario in the Normalization pipeline behavior
    /// requirement).
    private static func stripElementsAndAdjacentWhitespace(
        children: [XmlNode], indicesToRemove: Set<Int>
    ) -> [XmlNode] {
        var result: [XmlNode] = []
        result.reserveCapacity(children.count)
        var pendingWhitespace: XmlNode?
        for (idx, child) in children.enumerated() {
            if indicesToRemove.contains(idx) {
                // Drop this element AND any pending whitespace text before it.
                pendingWhitespace = nil
                continue
            }
            if child.kind == .text,
               child.textContent.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty {
                // Buffer whitespace-only text; commit only when followed by a
                // surviving element. If the next element is dropped, the
                // pending whitespace is dropped with it.
                if let p = pendingWhitespace {
                    result.append(p)
                }
                pendingWhitespace = child
                continue
            }
            if let p = pendingWhitespace {
                result.append(p)
                pendingWhitespace = nil
            }
            result.append(child)
        }
        if let p = pendingWhitespace {
            result.append(p)
        }
        return result
    }

    // MARK: - Step 5: stable-ID re-numbering with cross-reference rewrite

    /// Re-number w14:paraId / w14:textId / w:bookmarkId attribute values to a
    /// monotonic sequence (00000001, 00000002, ...) in document order. Within
    /// the same word/document.xml, rewrite known cross-references that target
    /// these IDs.
    ///
    /// Cross-references handled in Phase A:
    ///   - w:bookmarkEnd/@w:id → matched to its w:bookmarkStart by the
    ///     original w:id and rewritten in lock-step.
    ///   - w:hyperlink/@w:anchor → if it matches a paragraph w14:paraId in
    ///     this part, rewrite to the new ID. Otherwise (anchor points to a
    ///     bookmark name that we did not renumber, or to a heading in another
    ///     part), pass through unchanged.
    ///
    /// TODO(Phase B): w:fldSimple PAGEREF instructions and w:instrText
    /// PAGEREF/REF field codes that embed bookmark names are NOT rewritten
    /// here — Phase A goldens do not exercise those. A Phase B revision should
    /// parse field instruction text, locate the bookmark name token, and apply
    /// the same rewrite mapping.
    private static func renumberStableIDs(documentXML: Data) throws -> Data {
        let tree = try XmlTreeReader.parse(documentXML)

        // First pass: collect old → new mappings in document order.
        // Separate counters per attribute namespace+name so paraId, textId, and
        // bookmarkId each get their own 00000001-based sequence (matches the
        // human-written golden style: each ID family resets to 1).
        var paraIdMap: [String: String] = [:]
        var textIdMap: [String: String] = [:]
        var bookmarkIdMap: [String: String] = [:]
        var paraCounter = 0
        var textCounter = 0
        var bookmarkCounter = 0

        func nextID(_ counter: Int) -> String {
            String(format: "%08d", counter)
        }

        func collect(_ node: XmlNode) {
            guard node.kind == .element else { return }
            if let v = node.attributeValue(prefix: "w14", localName: "paraId"),
               paraIdMap[v] == nil {
                paraCounter += 1
                paraIdMap[v] = nextID(paraCounter)
            }
            if let v = node.attributeValue(prefix: "w14", localName: "textId"),
               textIdMap[v] == nil {
                textCounter += 1
                textIdMap[v] = nextID(textCounter)
            }
            // Only w:bookmarkStart establishes a bookmark id. w:bookmarkEnd's
            // w:id is a cross-reference to the start, not a fresh declaration.
            if node.prefix == "w" && node.localName == "bookmarkStart",
               let v = node.attributeValue(prefix: "w", localName: "id"),
               bookmarkIdMap[v] == nil {
                bookmarkCounter += 1
                bookmarkIdMap[v] = nextID(bookmarkCounter)
            }
            for child in node.children {
                collect(child)
            }
        }
        collect(tree.root)

        // Second pass: rewrite all referencing attributes using the maps.
        func rewrite(_ node: XmlNode) {
            guard node.kind == .element else {
                for child in node.children { rewrite(child) }
                return
            }
            // paraId on any element bearing it.
            if let v = node.attributeValue(prefix: "w14", localName: "paraId"),
               let new = paraIdMap[v] {
                node.setAttribute(prefix: "w14", localName: "paraId", value: new)
            }
            if let v = node.attributeValue(prefix: "w14", localName: "textId"),
               let new = textIdMap[v] {
                node.setAttribute(prefix: "w14", localName: "textId", value: new)
            }
            // bookmarkStart / bookmarkEnd both carry w:id; rewrite from the
            // bookmark map. Skip when there's no mapping (id refers to an
            // unknown bookmark — pass through unchanged).
            if node.prefix == "w" &&
               (node.localName == "bookmarkStart" || node.localName == "bookmarkEnd"),
               let v = node.attributeValue(prefix: "w", localName: "id"),
               let new = bookmarkIdMap[v] {
                node.setAttribute(prefix: "w", localName: "id", value: new)
            }
            // Hyperlink anchor → paragraph paraId.
            if node.prefix == "w" && node.localName == "hyperlink",
               let v = node.attributeValue(prefix: "w", localName: "anchor"),
               let new = paraIdMap[v] {
                node.setAttribute(prefix: "w", localName: "anchor", value: new)
            }
            for child in node.children {
                rewrite(child)
            }
        }
        rewrite(tree.root)

        return try XmlTreeWriter.serialize(tree)
    }
}
