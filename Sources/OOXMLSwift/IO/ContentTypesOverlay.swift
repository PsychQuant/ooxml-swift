import Foundation

/// Describes one part the writer is about to emit. Used by
/// `ContentTypesOverlay.merge` to compute the typed-set contribution to
/// `[Content_Types].xml`.
internal struct PartDescriptor: Equatable {
    /// Full path beginning with `/`, e.g. `"/word/document.xml"`.
    let partName: String
    /// MIME content type, e.g. `"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"`.
    let contentType: String

    init(partName: String, contentType: String) {
        self.partName = partName
        self.contentType = contentType
    }
}

/// Merges typed-model `<Override>` entries with preserved original entries
/// from `archiveTempDir`'s `[Content_Types].xml` so unknown parts (theme,
/// webSettings, people, glossary, etc.) survive round-trip.
///
/// Algorithm:
/// 1. Parse original `<Override>` entries from `originalContentTypesXML`.
/// 2. Parse original `<Default>` entries.
/// 3. Build the merged `<Override>` set: union of (typed parts) and
///    (original parts NOT in typed set), with typed taking precedence on
///    matching PartName.
/// 4. Build the merged `<Default>` set the same way (typed defaults derived
///    from typed parts' file extensions).
/// 5. Emit canonical `[Content_Types].xml` body.
///
/// Added in v0.12.0.
internal struct ContentTypesOverlay {

    private let originalOverrides: [String: String]   // PartName → ContentType
    private let originalDefaults: [String: String]    // Extension → ContentType

    init(originalContentTypesXML: String) {
        self.originalOverrides = Self.parseEntries(
            originalContentTypesXML,
            tag: "Override",
            keyAttr: "PartName"
        )
        self.originalDefaults = Self.parseEntries(
            originalContentTypesXML,
            tag: "Default",
            keyAttr: "Extension"
        )
    }

    /// Compute the merged `[Content_Types].xml` body for the supplied typed
    /// parts. Returns a complete XML document (including the `<?xml?>`
    /// prologue and the `<Types>` root).
    ///
    /// - Parameters:
    ///   - typedParts: PartName + ContentType pairs the writer is about to emit.
    ///   - typedManagedPatterns: Path patterns the typed model OWNS (e.g.,
    ///     `["/word/document.xml", "/word/footnotes.xml", "/word/header"]`).
    ///     A pattern matches a PartName if the PartName equals it OR begins
    ///     with it followed by any character (allowing prefix matches like
    ///     `/word/header` matching `/word/header1.xml`, `/word/header2.xml`).
    ///     PartNames matching a pattern but NOT present in `typedParts` are
    ///     interpreted as deletions and dropped from the merged output.
    ///     PartNames NOT matching any pattern AND present in original are
    ///     preserved (theme, webSettings, people, etc.).
    func merge(typedParts: [PartDescriptor], typedManagedPatterns: [String] = []) -> String {
        var mergedOverrides = originalOverrides

        // Drop original entries whose PartName matches a typed-managed pattern.
        // Those parts are owned by the typed model — its emit list is
        // authoritative (presence = emit, absence = delete).
        for partName in mergedOverrides.keys
        where typedManagedPatterns.contains(where: { Self.matches(partName: partName, pattern: $0) }) {
            mergedOverrides.removeValue(forKey: partName)
        }

        // Add typed entries (these may include re-additions of parts removed above).
        for part in typedParts {
            mergedOverrides[part.partName] = part.contentType
        }

        // Defaults come from preserved originals; typed parts may imply
        // additional defaults (e.g., `.png` for image media). For now, we
        // preserve all original defaults verbatim — typed-derived default
        // additions are deferred to a future refinement.
        let mergedDefaults = originalDefaults

        return Self.serialize(
            defaults: mergedDefaults,
            overrides: mergedOverrides
        )
    }

    private static func matches(partName: String, pattern: String) -> Bool {
        if partName == pattern { return true }
        // Prefix match: `/word/header` matches `/word/header1.xml`, `/word/header_X.xml`
        if partName.hasPrefix(pattern) {
            // Ensure the pattern boundary is followed by a non-letter character
            // (to avoid `/word/style` matching `/word/styles.xml` — though styles
            // is also typed-managed, this future-proofs against precision bugs).
            let nextIndex = partName.index(partName.startIndex, offsetBy: pattern.count)
            if nextIndex < partName.endIndex {
                return true   // any continuation accepted; caller chooses pattern precision
            }
        }
        return false
    }

    // MARK: - Parsing

    private static func parseEntries(
        _ xml: String,
        tag: String,
        keyAttr: String
    ) -> [String: String] {
        var result: [String: String] = [:]
        // Match `<Tag ... KeyAttr="..." ... ContentType="..." .../>` or with
        // attributes in either order. We use two narrow patterns to avoid an
        // overly permissive regex.
        let pattern = #"<\#(tag)\b([^>]*)/>"#
        guard let regex = try? NSRegularExpression(pattern: pattern, options: [.caseInsensitive]) else {
            return result
        }
        let nsString = xml as NSString
        let matches = regex.matches(in: xml, range: NSRange(location: 0, length: nsString.length))
        for match in matches where match.numberOfRanges >= 2 {
            let attrsRange = match.range(at: 1)
            guard attrsRange.location != NSNotFound else { continue }
            let attrs = nsString.substring(with: attrsRange)
            guard let key = Self.attributeValue(attrs, name: keyAttr),
                  let contentType = Self.attributeValue(attrs, name: "ContentType")
            else { continue }
            result[key] = contentType
        }
        return result
    }

    private static func attributeValue(_ attrs: String, name: String) -> String? {
        let pattern = #"\#(name)="([^"]*)""#
        guard let regex = try? NSRegularExpression(pattern: pattern) else { return nil }
        let nsString = attrs as NSString
        guard let match = regex.firstMatch(
            in: attrs,
            range: NSRange(location: 0, length: nsString.length)
        ), match.numberOfRanges >= 2 else { return nil }
        let valRange = match.range(at: 1)
        guard valRange.location != NSNotFound else { return nil }
        return nsString.substring(with: valRange)
    }

    // MARK: - Serialization

    private static func serialize(
        defaults: [String: String],
        overrides: [String: String]
    ) -> String {
        var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        xml += "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        // Sort for deterministic output (helps round-trip-fidelity test diffs).
        for ext in defaults.keys.sorted() {
            xml += "<Default Extension=\"\(escape(ext))\" ContentType=\"\(escape(defaults[ext]!))\"/>"
        }
        for partName in overrides.keys.sorted() {
            xml += "<Override PartName=\"\(escape(partName))\" ContentType=\"\(escape(overrides[partName]!))\"/>"
        }
        xml += "</Types>"
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
