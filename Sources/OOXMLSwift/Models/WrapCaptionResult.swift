import Foundation

/// Per-paragraph result returned by `WordDocument.wrapCaptionSequenceFields(...)`.
///
/// PsychQuant/che-word-mcp#62 design Decision 5 — structured per-paragraph
/// shape so LLM callers can verify "did all N captions get fields?" rather
/// than parsing free-form summary strings. `skipped` surfaces idempotency
/// no-ops (e.g., paragraph already wraps a SEQ field for the same identifier)
/// so the caller doesn't think the call silently failed.
public struct WrapCaptionResult: Equatable, Sendable {
    /// Number of paragraphs whose `flattenedDisplayText()` matched the regex
    /// (sum of `paragraphsModified.count` and `skipped.count`).
    public let matchedParagraphs: Int

    /// Number of SEQ fields actually written (matched minus skipped).
    public let fieldsInserted: Int

    /// Body-level (or part-container-level) paragraph indices, in document
    /// order, of paragraphs that received a new SEQ field.
    public let paragraphsModified: [Int]

    /// Paragraphs that matched the pattern but were intentionally not modified
    /// (e.g., already wrap a SEQ field for the same identifier).
    public let skipped: [SkippedParagraph]

    public init(
        matchedParagraphs: Int,
        fieldsInserted: Int,
        paragraphsModified: [Int],
        skipped: [SkippedParagraph]
    ) {
        self.matchedParagraphs = matchedParagraphs
        self.fieldsInserted = fieldsInserted
        self.paragraphsModified = paragraphsModified
        self.skipped = skipped
    }
}

/// One element of `WrapCaptionResult.skipped`.
///
/// `container` is `nil` for body-scope skips (Phase 1 ships body-only) and
/// reserved for the future `.all` cross-container path (e.g.,
/// `"header:default"` / `"footer:rId7"` / `"footnote:N"`).
public struct SkippedParagraph: Equatable, Sendable {
    public let paragraphIndex: Int
    public let reason: String
    public let container: String?

    public init(paragraphIndex: Int, reason: String, container: String? = nil) {
        self.paragraphIndex = paragraphIndex
        self.reason = reason
        self.container = container
    }
}
