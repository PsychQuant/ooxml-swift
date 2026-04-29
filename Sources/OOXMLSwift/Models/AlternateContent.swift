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
///   manipulation.
///
/// **Dirty-tracking (PsychQuant/ooxml-swift#6, F8)**: Writes to `fallbackRuns`
/// flip `fallbackRunsModified` to `true` via `didSet`. The
/// `Paragraph.toXMLSortedByPosition()` emit path checks this flag and throws
/// `RoundtripError.unserializedFallbackEdit(position:)` rather than silently
/// emitting stale `rawXML`. Initializer assignment does NOT fire `didSet`,
/// so freshly-constructed (Reader-loaded) values start clean. To commit
/// typed edits, the caller must rebuild the `AlternateContent` with a
/// regenerated `rawXML` (a future `regenerateRawXMLFromFallbackRuns()`
/// helper is deferred to a later SDD; for now, callers either accept the
/// throw or manually construct a new value with the regenerated XML).
public struct AlternateContent: Equatable {
    /// Verbatim source XML of the `<mc:AlternateContent>` block, used by the
    /// Writer for byte-equivalent emit.
    public var rawXML: String

    /// Typed Runs extracted from the `<mc:Fallback>` child, surfaced for
    /// tool-mediated read access. Empty when source has no fallback or the
    /// fallback contains no `<w:r>` children. **Mutating this array flips
    /// `fallbackRunsModified` to `true`** — see the type-level doc-comment
    /// for the dirty-tracking contract.
    public var fallbackRuns: [Run] {
        didSet { fallbackRunsModified = true }
    }

    /// Source-document order index for Phase 4 sort-by-position emit.
    public var position: Int

    /// Dirty flag flipped by `fallbackRuns`'s `didSet` observer
    /// (PsychQuant/ooxml-swift#6, F8). `false` for Reader-loaded /
    /// freshly-constructed values; `true` once any caller mutation touches
    /// `fallbackRuns`. The emit path consults this to refuse stale-rawXML
    /// writes.
    public private(set) var fallbackRunsModified: Bool = false

    public init(rawXML: String, fallbackRuns: [Run] = [], position: Int = 0) {
        self.rawXML = rawXML
        self.fallbackRuns = fallbackRuns
        self.position = position
        // didSet does not fire from initializer assignment, so the flag
        // stays false here — this is the load-bearing invariant for F8.
    }
}
