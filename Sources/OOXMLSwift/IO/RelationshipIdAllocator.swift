import Foundation

/// Generates collision-free `rId` values for OOXML relationships, accounting
/// for both preserved original relationships (from `archiveTempDir`'s
/// `_rels/document.xml.rels`) and typed-model fields.
///
/// Replaces the prior naive counter (`headers.count + footers.count + ...`)
/// at `DocxWriter.swift:238` which collided with preserved original rIds in
/// preserve-by-default mode (see `che-word-mcp-ooxml-roundtrip-fidelity`).
///
/// Added in v0.12.0.
internal final class RelationshipIdAllocator {

    /// Next integer to allocate. Starts at `max(observed) + 1`.
    private var nextId: Int

    /// Already-observed integers (for fast `reserve` collision detection).
    private var observed: Set<Int>

    /// Initialize the allocator by scanning `originalRelsXML` for `Id="rId<N>"`
    /// patterns and merging with `additionalReservedIds`.
    ///
    /// - Parameters:
    ///   - originalRelsXML: Raw XML content of `_rels/document.xml.rels` from
    ///     the source archive. Pass empty string when no source archive exists
    ///     (initializer-built documents).
    ///   - additionalReservedIds: rIds the typed model already uses (e.g.,
    ///     existing header/footer/image rIds). Empty array is fine.
    init(originalRelsXML: String, additionalReservedIds: [String] = []) {
        var observed: Set<Int> = []

        // Scan the original rels XML for `Id="rId<digits>"` patterns.
        let pattern = #"Id="rId(\d+)""#
        if let regex = try? NSRegularExpression(pattern: pattern) {
            let nsString = originalRelsXML as NSString
            let matches = regex.matches(
                in: originalRelsXML,
                range: NSRange(location: 0, length: nsString.length)
            )
            for match in matches where match.numberOfRanges >= 2 {
                let intRange = match.range(at: 1)
                guard intRange.location != NSNotFound else { continue }
                let intStr = nsString.substring(with: intRange)
                if let n = Int(intStr) {
                    observed.insert(n)
                }
            }
        }

        // Merge typed-field rIds. Skip non-numeric suffixes per spec
        // ("Allocator handles non-numeric rId values gracefully").
        for id in additionalReservedIds {
            if let n = Self.numericSuffix(of: id) {
                observed.insert(n)
            }
        }

        self.observed = observed
        self.nextId = (observed.max() ?? 0) + 1
    }

    /// Allocate a new collision-free rId.
    func allocate() -> String {
        // Advance past any reserved IDs added between allocations.
        while observed.contains(nextId) {
            nextId += 1
        }
        let id = "rId\(nextId)"
        observed.insert(nextId)
        nextId += 1
        return id
    }

    /// Reserve a specific rId, marking it as taken so future `allocate()`
    /// calls do not return it. No-op if already reserved or if the ID has a
    /// non-numeric suffix.
    func reserve(_ id: String) {
        if let n = Self.numericSuffix(of: id) {
            observed.insert(n)
            if n >= nextId {
                nextId = n + 1
            }
        }
    }

    /// Extract the integer portion of `"rId<N>"`, returning nil for malformed
    /// IDs or non-numeric suffixes (e.g., `"rIdAbc"`).
    private static func numericSuffix(of id: String) -> Int? {
        guard id.hasPrefix("rId") else { return nil }
        let suffix = id.dropFirst(3)
        return Int(suffix)
    }
}
