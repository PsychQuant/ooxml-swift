import Foundation

// MARK: - FieldSimple (PsychQuant/che-word-mcp#56 Phase 3)

/// `<w:fldSimple>` wrapper holding a Word field expression and its current
/// rendered runs. Typed editable surface per design decision "Hybrid model
/// (typed surface + raw passthrough), not raw-only carriers" so MCP tools
/// like `replace_text` can locate and modify text inside SEQ Table captions,
/// REF cross-references, TOC entries, and other simple field results.
///
/// Source XML shape:
/// ```
/// <w:fldSimple w:instr=" SEQ Table \* ARABIC ">
///   <w:r><w:t>1</w:t></w:r>
/// </w:fldSimple>
/// ```
///
/// Reader populates `instr` with the leading/trailing whitespace preserved
/// exactly. Writer emits the field with the same `w:instr` value.
public struct FieldSimple: Equatable {
    /// `w:instr` value — the field expression (e.g., ` SEQ Table \* ARABIC `,
    /// ` REF tab:foo \h `, ` PAGE \* MERGEFORMAT `). Whitespace must round-trip
    /// byte-equivalent so existing field-recalc tools (e.g., `update_all_fields`)
    /// can still match the source text.
    public var instr: String

    /// Inner `<w:r>` children — the rendered field result. Editable surface
    /// for tool-mediated `replace_text` / `format_text` so v3.12.0's silent
    /// failure for text inside fldSimple wrappers no longer happens.
    public var runs: [Run]

    /// Unrecognized `<w:fldSimple>` attributes (e.g., `w:fldLock`, `w:dirty`,
    /// vendor extensions). Captured verbatim so a no-op round-trip is
    /// attribute-lossless.
    public var rawAttributes: [String: String]

    /// Source-document order index, used by Phase 4 `Paragraph.toXML()`
    /// sort-by-position emit.
    public var position: Int

    public init(
        instr: String,
        runs: [Run] = [],
        rawAttributes: [String: String] = [:],
        position: Int = 0
    ) {
        self.instr = instr
        self.runs = runs
        self.rawAttributes = rawAttributes
        self.position = position
    }
}
