import Foundation

/// Where a text-mutation operation walks within a `WordDocument`.
///
/// PsychQuant/che-word-mcp#62 (`wrap_caption_seq`) introduced this enum so
/// bulk-mutation tools can opt between body-only (the 90% case for thesis
/// caption rescue) and the full part-container traversal that
/// `WordDocument.updateAllFields(isolatePerContainer:)` already uses for SEQ
/// counter computation.
///
/// - `.body` — only `body.children` (recursing into table cells + block-level
///   SDT children).
/// - `.all` — body PLUS headers, footers, footnotes, endnotes — exactly the
///   set traversed by `updateAllFields(isolatePerContainer: false)` so SEQ
///   counter coverage and SEQ field placement coverage are symmetric.
public enum TextScope: Equatable, Sendable {
    case body
    case all
}
