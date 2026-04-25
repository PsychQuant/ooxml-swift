import Foundation

// MARK: - AlternateContent (PsychQuant/che-word-mcp#56 Phase 3)

/// `<mc:AlternateContent>` wrapper — Word's "I have multiple ways to render
/// this" markup compatibility envelope. Carries one or more `<mc:Choice>`
/// blocks (preferred renderings, gated by feature requirements like
/// `Requires="wps14"`) plus a single `<mc:Fallback>` block (legacy text
/// renderings for clients that lack the required features).
///
/// Common payloads:
/// - WordprocessingShape (`wps:`) drawings: choice has full vector, fallback
///   has flat text describing the shape
/// - Embedded math (`<m:oMath>`): choice has typed math AST, fallback has
///   ASCII transliteration like "Pearson (Spearman)"
/// - Custom OLE / chart blocks: choice has rich content, fallback has
///   placeholder text
///
/// Hybrid model per design decision "Hybrid model (typed surface + raw
/// passthrough), not raw-only carriers":
/// - `rawXML` — verbatim source XML of the entire `<mc:AlternateContent>`
///   block. Writer emits this byte-equivalent so a no-op round-trip preserves
///   `<mc:Choice>` content (which we do not type model — its surface area is
///   too large).
/// - `fallbackRuns` — `<w:r>` children extracted from `<mc:Fallback>` as
///   typed Runs. MCP tools (`replace_text`, `format_text`) operate on this
///   surface so users can edit fallback text without resorting to raw XML
///   manipulation. Out-of-scope: edits to `fallbackRuns` are NOT
///   automatically re-serialized into `rawXML`. Callers wishing to apply
///   fallback edits to the saved XML must invoke a future
///   `regenerateRawXMLFromFallbackRuns()` method (deferred to a later SDD).
///   In practice this means: the v3.13.0 read-side surface lets tools
///   discover and report on fallback text but does not yet propagate edits
///   to disk. Word reconciles `<mc:Choice>` vs `<mc:Fallback>` per its own
///   rules — out of scope for this requirement.
public struct AlternateContent: Equatable {
    /// Verbatim source XML of the `<mc:AlternateContent>` block, used by the
    /// Writer for byte-equivalent emit.
    public var rawXML: String

    /// Typed Runs extracted from the `<mc:Fallback>` child, surfaced for
    /// tool-mediated read access. Empty when source has no fallback or the
    /// fallback contains no `<w:r>` children.
    public var fallbackRuns: [Run]

    /// Source-document order index for Phase 4 sort-by-position emit.
    public var position: Int

    public init(rawXML: String, fallbackRuns: [Run] = [], position: Int = 0) {
        self.rawXML = rawXML
        self.fallbackRuns = fallbackRuns
        self.position = position
    }
}
