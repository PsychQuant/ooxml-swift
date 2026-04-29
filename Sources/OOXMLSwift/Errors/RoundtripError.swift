import Foundation

/// Round-trip completeness errors thrown by `Paragraph` emit paths when the
/// typed mutation surface and the raw-XML emit surface have drifted apart.
///
/// **Origin**: PsychQuant/ooxml-swift#6 — bundled hardening for findings F8
/// (`AlternateContent.fallbackRuns` typed-edit drift) surfaced by the
/// verification of PsychQuant/che-word-mcp#56.
///
/// **Caller guidance**: this throw fires when caller code mutated a typed
/// surface (e.g., `AlternateContent.fallbackRuns`) but the underlying
/// `rawXML` was not re-serialised. Callers SHOULD propagate the error so
/// the offending paragraph can be surfaced to the user. Silencing via
/// `try?` swallows the corruption signal and writes stale data — this is
/// the silent-failure mode the change exists to prevent.
public enum RoundtripError: Error, LocalizedError, Equatable {
    /// `AlternateContent.fallbackRuns` was mutated since construction but the
    /// `rawXML` was not regenerated; the emit path refuses to write the
    /// stale `rawXML` because doing so would silently discard the typed
    /// edit. The associated `position` is the `AlternateContent.position`
    /// (the source-document order index for sort-by-position emit), NOT the
    /// paragraph index — use it to locate the offending value within
    /// `paragraph.alternateContents`.
    case unserializedFallbackEdit(position: Int)

    public var errorDescription: String? {
        switch self {
        case let .unserializedFallbackEdit(position):
            return "AlternateContent at position \(position) has typed fallbackRuns edits that have not been re-serialised into rawXML; refusing to write stale data."
        }
    }
}
