import Foundation

/// Result of `Document.replaceInParagraphSurfaces` — distinguishes successful
/// replacement count from informative refusals due to OMML boundary intersection.
///
/// **Spec capability**: `ooxml-paragraph-text-mirror`. Introduced as part of the
/// `flatten-replace-omml-bilateral-coverage` change (cluster fix for
/// PsychQuant/che-word-mcp #99 / #100 / #101 / #102 / #103).
///
/// **Library design principles** governing this type
/// (`ooxml-library-design-principles`):
/// 1. **Correctness primacy** — refusal is preferred over an incorrect
///    approximation. When a replacement match span intersects an `<m:oMath>` /
///    `<m:oMathPara>` element, mutating around or through the OMML would either
///    produce structurally valid but semantically incorrect output (e.g.
///    `"see δref X"` from `replace("eq δ here", "ref X")`) or silently delete
///    the equation. Both violate human-user expectations, so we refuse.
/// 2. **Human-like operations** — operations correspond to actions a human
///    Word user would consciously perform. A human editing prose around an
///    equation would either (a) explicitly delete the equation first and then
///    rephrase, or (b) keep the rephrasing strictly within the surrounding
///    `<w:t>` text. The library mirrors this by refusing cross-boundary
///    mutation by default; an explicit opt-in (e.g. `omml_handling: "drop"`)
///    is a deliberate follow-up — out of scope for this change.
///
/// **Asymmetric mirror invariant**: `Paragraph.flattenedDisplayText()` (read)
/// and `Document.replaceInParagraphSurfaces` (write) walk the same wrapper
/// surfaces and detect direct-child OMML at the same 4 positions, but diverge
/// in how they handle detected OMML. Reads include OMML `visibleText` (so
/// anchor lookup can locate paragraphs containing math). Writes treat OMML
/// as opaque structural units — replacements wholly within `<w:t>` ranges
/// proceed, replacements crossing OMML boundaries refuse.
public enum ReplaceResult: Equatable {
    /// Successful replacement. `count == 0` means the find-string did not
    /// appear in any walked surface AND no OMML boundary was intersected.
    /// `count > 0` means N occurrences were mutated.
    ///
    /// Anchor-not-found is **not** a refusal — it returns `.replaced(count: 0)`,
    /// distinct from `.refusedDueToOMMLBoundary(occurrences: [])`. The empty
    /// occurrences list reserved for the refusal case is structurally
    /// impossible (refusal carries at least one occurrence describing the
    /// boundary intersection that triggered it).
    case replaced(count: Int)

    /// Refused because at least one find-string match span intersected an
    /// `<m:oMath>` / `<m:oMathPara>` direct-child element. Each `Occurrence`
    /// carries the match position in flattened-text coordinates plus the
    /// intersecting OMML spans, so callers can produce actionable error
    /// messages (e.g. "Cannot replace 'eq δ here'; equation appears at
    /// character 7. Rephrase find to avoid the equation, or use a
    /// dedicated equation-editing tool to mutate equation content.").
    ///
    /// Future cases MAY be added (e.g.
    /// `.refusedDueToHyperlinkBoundary(...)`,
    /// `.refusedDueToContentControlBoundary(...)`) without removing existing
    /// cases. Each new boundary refusal type SHALL conform to the same
    /// "informative occurrence" pattern.
    case refusedDueToOMMLBoundary(occurrences: [Occurrence])

    /// Mixed outcome — the same find string appeared multiple times in a
    /// paragraph; some occurrences were wholly within `<w:t>` (replaced)
    /// and others crossed an OMML boundary (refused). Carries both signals.
    /// `replacedCount > 0` AND `refusedOccurrences.isEmpty == false` are
    /// the invariants for this case (otherwise use `.replaced` or
    /// `.refusedDueToOMMLBoundary` directly).
    case mixed(replacedCount: Int, refusedOccurrences: [Occurrence])

    /// Single refused match occurrence, with positions expressed as character
    /// ranges in `Paragraph.flattenedDisplayText()` coordinates.
    ///
    /// `matchSpan`: where the find-string matched. Half-open range
    /// `[start, end)`.
    /// `ommlSpans`: one or more OMML element spans intersecting `matchSpan`.
    /// Multiple spans appear when a single match crosses two or more
    /// equation elements.
    public struct Occurrence: Equatable {
        public let matchSpan: Range<Int>
        public let ommlSpans: [Range<Int>]

        public init(matchSpan: Range<Int>, ommlSpans: [Range<Int>]) {
            self.matchSpan = matchSpan
            self.ommlSpans = ommlSpans
        }
    }
}
