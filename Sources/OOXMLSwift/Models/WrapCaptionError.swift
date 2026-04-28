import Foundation

/// Errors thrown by `WordDocument.wrapCaptionSequenceFields(...)`.
///
/// PsychQuant/che-word-mcp#62 design Decisions 1 + 2 — pre-mutation validation
/// rejects malformed callers BEFORE the document is touched, so a bad regex
/// or bad bookmark configuration does not partially-mutate the body.
public enum WrapCaptionError: Error, Equatable {
    /// Pattern compiled but does not contain exactly one capture group.
    /// Spec requires exactly one numeric capture group whose match becomes the
    /// SEQ field's `cachedResult`.
    case patternMissingCaptureGroup(actual: Int)

    /// `insertBookmark == true` but `bookmarkTemplate` is nil OR does not
    /// contain the literal `${number}` placeholder. Without `${number}`, every
    /// matched paragraph would receive the same bookmark name — which violates
    /// `<w:bookmarkStart>` name uniqueness within the document.
    case bookmarkTemplateMissing

    /// `scope == .all` was passed but cross-container walking is not yet
    /// implemented. Phase 1 ships `.body` only; `.all` lands alongside the
    /// MCP wrapper integration test in a Phase 1.x patch.
    case scopeNotImplemented(TextScope)
}
