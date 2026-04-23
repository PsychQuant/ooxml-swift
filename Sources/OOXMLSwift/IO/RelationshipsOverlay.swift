import Foundation

/// One `<Relationship>` row from `word/_rels/document.xml.rels`.
internal struct RelationshipDescriptor: Equatable {
    /// Relationship id, e.g., `"rId10"`. MUST match `^rId[0-9]+$` per OOXML.
    let id: String
    /// Type URL, e.g., `"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"`.
    let type: String
    /// Target path relative to the rels file, e.g., `"header1.xml"`,
    /// `"theme/theme1.xml"`, `"media/image1.png"`, `"https://example.com"`.
    let target: String
    /// `"External"` for hyperlinks pointing outside the package, otherwise `nil`.
    let targetMode: String?
}

/// Merges typed-model relationships with preserved original relationships from
/// `archiveTempDir`'s `word/_rels/document.xml.rels` so unknown relationship
/// types (theme / webSettings / people / customXml / commentsExtensible /
/// commentsIds) survive round-trip — even when a legitimate edit (e.g.,
/// `addHeader`) triggers a rels rewrite.
///
/// **Algorithm**:
/// 1. Parse original rels into ordered list.
/// 2. For each original rel:
///    - If `rel.type ∈ typedManagedTypes` AND typed model still claims `rel.id`
///      → emit (typed model authoritative on target; ID preserved).
///    - If `rel.type ∈ typedManagedTypes` AND typed model dropped `rel.id`
///      → drop (deletion).
///    - If `rel.type ∉ typedManagedTypes` → preserve verbatim (theme/webSettings/etc.).
/// 3. For each typed-model rel whose ID is NOT in original → append as new.
///
/// Added in v0.13.1 (closes che-word-mcp#35).
internal struct RelationshipsOverlay {

    private let originalRels: [RelationshipDescriptor]

    init(originalRelsXML: String) {
        self.originalRels = Self.parseRels(originalRelsXML)
    }

    /// Compute the merged `word/_rels/document.xml.rels` body.
    ///
    /// - Parameters:
    ///   - typedRels: rels the writer wants to emit, derived from the typed
    ///     model's current state (headers/footers/images/etc.).
    ///   - typedManagedTypes: relationship type URLs the typed model owns. An
    ///     original rel matching this set BUT not in `typedRels` is interpreted
    ///     as a deletion (e.g., `delete_header` removed this rel).
    /// - Returns: Complete XML document for `word/_rels/document.xml.rels`.
    func merge(
        typedRels: [RelationshipDescriptor],
        typedManagedTypes: Set<String>
    ) -> String {
        var merged: [RelationshipDescriptor] = []
        let typedById: [String: RelationshipDescriptor] = Dictionary(
            uniqueKeysWithValues: typedRels.map { ($0.id, $0) }
        )
        var emittedIds = Set<String>()

        // Pass 1: walk original rels in order. Preserve unknown types verbatim;
        // re-emit managed types only if typed model still has them.
        for rel in originalRels {
            if typedManagedTypes.contains(rel.type) {
                if let typed = typedById[rel.id] {
                    merged.append(typed)
                    emittedIds.insert(rel.id)
                }
                // Typed model deleted this managed rel — drop.
            } else {
                merged.append(rel)
                emittedIds.insert(rel.id)
            }
        }

        // Pass 2: append typed rels not already emitted (newly added parts).
        for rel in typedRels where !emittedIds.contains(rel.id) {
            merged.append(rel)
        }

        return Self.serialize(merged)
    }

    // MARK: - Parsing

    private static func parseRels(_ xml: String) -> [RelationshipDescriptor] {
        var result: [RelationshipDescriptor] = []
        // Match `<Relationship ... />` (self-closing) — the only standard form
        // for rels files. `[^>]*?` matches lazily so the trailing `/` is not
        // captured into the attrs group; we deliberately do NOT exclude `/`
        // because Type URLs like "http://schemas..." contain forward slashes.
        let pattern = #"<Relationship\b([^>]*?)/>"#
        guard let regex = try? NSRegularExpression(pattern: pattern) else { return result }
        let nsString = xml as NSString
        for match in regex.matches(in: xml, range: NSRange(location: 0, length: nsString.length))
        where match.numberOfRanges >= 2 {
            let attrs = nsString.substring(with: match.range(at: 1))
            guard let id = attribute(attrs, name: "Id"),
                  let type = attribute(attrs, name: "Type"),
                  let target = attribute(attrs, name: "Target")
            else { continue }
            let targetMode = attribute(attrs, name: "TargetMode")
            result.append(RelationshipDescriptor(
                id: id, type: type, target: target, targetMode: targetMode
            ))
        }
        return result
    }

    private static func attribute(_ attrs: String, name: String) -> String? {
        let escaped = NSRegularExpression.escapedPattern(for: name)
        let pattern = #"\b\#(escaped)="([^"]*)""#
        guard let regex = try? NSRegularExpression(pattern: pattern) else { return nil }
        let nsString = attrs as NSString
        guard let match = regex.firstMatch(
            in: attrs,
            range: NSRange(location: 0, length: nsString.length)
        ), match.numberOfRanges >= 2 else { return nil }
        return nsString.substring(with: match.range(at: 1))
    }

    // MARK: - Serialization

    private static func serialize(_ rels: [RelationshipDescriptor]) -> String {
        var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        xml += "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        for rel in rels {
            xml += "<Relationship Id=\"\(escape(rel.id))\""
            xml += " Type=\"\(escape(rel.type))\""
            xml += " Target=\"\(escape(rel.target))\""
            if let mode = rel.targetMode {
                xml += " TargetMode=\"\(escape(mode))\""
            }
            xml += "/>"
        }
        xml += "</Relationships>"
        return xml
    }

    private static func escape(_ s: String) -> String {
        return s
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
    }
}
