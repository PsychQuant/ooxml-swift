import Foundation

// Cross-document OMath splice — public API surface.
//
// Spec: openspec/changes/cross-document-omath-splice/specs/omath-splice/spec.md
// Design: openspec/changes/cross-document-omath-splice/design.md
// Issue: PsychQuant/ooxml-swift#57

/// Position within a target paragraph where a spliced OMath block should be inserted.
///
/// `.afterText` / `.beforeText` mirror `InsertLocation`'s anchor pattern, including
/// the `AnchorLookupOptions` knob for math-script-insensitive matching.
public enum OMathSplicePosition: Equatable {
    case atStart
    case atEnd
    case afterText(_ anchor: String, instance: Int = 1, options: AnchorLookupOptions = AnchorLookupOptions())
    case beforeText(_ anchor: String, instance: Int = 1, options: AnchorLookupOptions = AnchorLookupOptions())
}

/// Controls how the source Run's `RunProperties` (rPr) propagate to the spliced OMath Run.
///
/// `.full` is the default — the Cambria Math font reference, language tag, and other
/// formatting are preserved verbatim. `.omathOnly` strips style-ID / theme references
/// that may not resolve in the target document. `.discard` resets rPr entirely.
public enum OMathSpliceRpRMode: Equatable {
    case full
    case omathOnly
    case discard
}

/// Controls how prefix / URI mismatches between source and target OMath namespaces
/// are handled.
///
/// `.lenient` accepts prefix mismatch with same URI — the spliced XML carries its
/// own `xmlns:` declaration so mixed prefixes within one document are spec-legal
/// (ECMA-376 allows local namespace declarations). Throws only when URIs differ.
/// `.strict` throws on any prefix or URI mismatch — useful for byte-equal
/// round-trip fixtures or callers requiring single-prefix output.
public enum OMathSpliceNamespacePolicy: Equatable {
    case lenient
    case strict
}

/// Errors surfaced by `WordDocument.spliceOMath` and `spliceParagraphOMath`.
public enum OMathSpliceError: Error, Equatable {
    case sourceHasNoOMath
    case omathIndexOutOfRange(requested: Int, available: Int)
    case targetParagraphOutOfRange(Int)
    case anchorNotFound(String, instance: Int)
    case namespaceMismatch(sourceURI: String, targetURI: String)
    case contextAnchorNotFound(omathIndex: Int, snippet: String)
}

// MARK: - Internal: extracted OMath descriptor

/// Internal carrier-agnostic descriptor of one OMath block extracted from a source paragraph.
///
/// Used to implement the joint document-order index for `omathIndex` (Decision Q2):
/// callers index "Nth OMath in source-document order, regardless of carrier."
internal struct ExtractedOMath {
    /// The verbatim `<m:oMath>...</m:oMath>` XML (or `<m:oMathPara>` wrapper) from source.
    let xml: String
    /// Which carrier the OMath was loaded from in the source paragraph.
    let kind: Kind
    /// Source-document byte position (filled by DocxReader for both Run and UnrecognizedChild).
    /// Used for joint sort across the two carriers. May be nil for API-built paragraphs.
    let sourcePosition: Int?
    /// For `.inRun` kind, the source Run's properties (used for `.full` / `.omathOnly` rPr propagation).
    /// For `.directChild`, nil (no enclosing Run).
    let sourceRunProperties: RunProperties?
    /// PsychQuant/ooxml-swift#124: the OMath root element's local name (`"oMath"` or
    /// `"oMathPara"`), read directly off `xml`'s opening tag rather than assumed. Used
    /// by `spliceAsDirectChild` so a direct-child `oMathPara` round-trips as
    /// `oMathPara`, not a hardcoded `"oMath"`.
    let rootName: String

    enum Kind: Equatable {
        case inRun
        case directChild  // OMath as direct child of `<w:p>` via Paragraph.unrecognizedChildren
    }
}

// MARK: - Internal: OMath extraction from source paragraph

internal enum OMathExtractor {
    /// Extracts all OMath blocks from a source paragraph, sorted by source-document order.
    ///
    /// Inspects two carriers:
    /// 1. `Run.rawXML` containing `<m:oMath` (inline OMath embedded in a Run, per #85/#92)
    /// 2. `Paragraph.unrecognizedChildren` where `name == "oMath" || "oMathPara"`
    ///    (direct-child OMath, per #99/#100/#101/#102)
    ///
    /// Joint sort by `position ?? 0` to implement caller-intuitive index semantics:
    /// "Nth OMath in source-document order, regardless of carrier" (Decision Q2).
    ///
    /// Spec: Carrier preservation strategy (Decision Q1)
    static func extract(from paragraph: Paragraph) -> [ExtractedOMath] {
        var collected: [ExtractedOMath] = []

        // Carrier 1: Run.rawXML (inline OMath in Run)
        for run in paragraph.runs {
            guard let raw = run.rawXML, raw.contains("<") else { continue }
            // Substring match for `:oMath` or `<oMath` to handle prefix or default namespace.
            // Specific match handles both `<m:oMath` / `<mml:oMath` / `<oMath`.
            let hasOMath = raw.contains(":oMath")
                || raw.contains(":oMathPara")
                || raw.contains("<oMath")
                || raw.contains("<oMathPara")
            guard hasOMath else { continue }
            let normalized = ensureXmlnsDeclared(in: raw)
            collected.append(ExtractedOMath(
                xml: normalized,
                kind: .inRun,
                sourcePosition: run.position,
                sourceRunProperties: run.properties,
                rootName: rootLocalName(of: normalized)
            ))
        }

        // Carrier 2: Paragraph.unrecognizedChildren (direct-child OMath)
        for child in paragraph.unrecognizedChildren where child.name == "oMath" || child.name == "oMathPara" {
            collected.append(ExtractedOMath(
                xml: ensureXmlnsDeclared(in: child.rawXML),
                kind: .directChild,
                sourcePosition: child.position,
                sourceRunProperties: nil,
                // PsychQuant/ooxml-swift#124: preserve the *source* local name
                // (`child.name`) rather than re-deriving it from `xml` — both
                // agree in the normal case, but `child.name` is the value the
                // reader already trusted, so prefer it as source of truth.
                rootName: child.name
            ))
        }

        // Sort by source-document position (joint document-order index — Decision Q2).
        // Stable sort preserves insertion order on equal positions.
        return collected.sorted { ($0.sourcePosition ?? 0) < ($1.sourcePosition ?? 0) }
    }

    /// If the given OMath rawXML's opening tag lacks an `xmlns:<prefix>="<URI>"`
    /// declaration for the OMath namespace prefix, inject one with the standard URI.
    /// This is required when the source-side parser inherits xmlns from the parent
    /// `<w:p>` but `XMLElement.xmlString` doesn't carry inherited declarations
    /// — the extracted rawXML must be self-contained for round-trip correctness.
    static func ensureXmlnsDeclared(in xml: String) -> String {
        guard let prefix = OMathNamespace.extractPrefix(from: xml), !prefix.isEmpty else {
            // Default-namespace OMath — assume xmlns="..." declared elsewhere or not needed.
            return xml
        }
        // Already declared?
        if xml.contains("xmlns:\(prefix)=") || xml.contains("xmlns:\(prefix) =") {
            return xml
        }
        // Inject after the opening element name. Find the first `<prefix:` and inject
        // ` xmlns:prefix="..."` after the element name (before any other attributes).
        let openTag = "<\(prefix):"
        guard let openIdx = xml.range(of: openTag) else { return xml }
        // Find end of element name (first whitespace or `>` or `/`).
        var nameEnd = openIdx.upperBound
        while nameEnd < xml.endIndex,
              !xml[nameEnd].isWhitespace,
              xml[nameEnd] != ">",
              xml[nameEnd] != "/" {
            nameEnd = xml.index(after: nameEnd)
        }
        let standardURI = "http://schemas.openxmlformats.org/officeDocument/2006/math"
        let injection = " xmlns:\(prefix)=\"\(standardURI)\""
        return String(xml[..<nameEnd]) + injection + String(xml[nameEnd...])
    }

    /// PsychQuant/ooxml-swift#124: reads the OMath root element's local name
    /// (e.g. `"oMath"` or `"oMathPara"`) directly off the opening tag, stripping
    /// any namespace prefix. Falls back to `"oMath"` only when the fragment has
    /// no recognizable opening tag at all (defensive default; extraction already
    /// guarantees `hasOMath` matched before this is called).
    static func rootLocalName(of xml: String) -> String {
        guard let lessThan = xml.firstIndex(of: "<") else { return "oMath" }
        var idx = xml.index(after: lessThan)
        var qname = ""
        while idx < xml.endIndex,
              !xml[idx].isWhitespace,
              xml[idx] != ">",
              xml[idx] != "/" {
            qname.append(xml[idx])
            idx = xml.index(after: idx)
        }
        if let colonIdx = qname.firstIndex(of: ":") {
            return String(qname[qname.index(after: colonIdx)...])
        }
        return qname.isEmpty ? "oMath" : qname
    }
}

// MARK: - Internal: namespace inspection helpers

internal enum OMathNamespace {
    /// Extracts the `xmlns:` URI for the OMath prefix in the given XML fragment.
    /// Returns the URI string, or nil if no `xmlns:` declaration found.
    ///
    /// Heuristic: scans for `xmlns:<prefix>="<URI>"` where the prefix is the same
    /// one used in the opening element name (e.g. `<mml:oMath xmlns:mml="...">` → `mml`).
    /// Falls back to scanning for any `xmlns:` declaration if prefix detection fails.
    static func extractURI(from xml: String) -> String? {
        // Find the prefix from the opening tag (e.g. `<m:oMath` → "m", `<mml:oMath` → "mml").
        guard let lessThan = xml.firstIndex(of: "<") else { return nil }
        let afterLT = xml.index(after: lessThan)
        guard let colonIdx = xml[afterLT...].firstIndex(of: ":") else {
            // Could be default-namespace OMath (no prefix). Scan for any `xmlns="..."`.
            return extractURIByPattern(xml: xml, pattern: #"\bxmlns\s*=\s*"([^"]+)""#)
        }
        let prefix = String(xml[afterLT..<colonIdx])
        // Look for `xmlns:<prefix>="<URI>"`.
        let pattern = #"\bxmlns:"# + NSRegularExpression.escapedPattern(for: prefix) + #"\s*=\s*"([^"]+)""#
        return extractURIByPattern(xml: xml, pattern: pattern)
    }

    /// Extracts the OMath prefix (e.g. "m" or "mml") from the opening element.
    /// Returns nil if no prefix used (default namespace).
    static func extractPrefix(from xml: String) -> String? {
        guard let lessThan = xml.firstIndex(of: "<") else { return nil }
        let afterLT = xml.index(after: lessThan)
        guard let colonIdx = xml[afterLT...].firstIndex(of: ":") else {
            return nil
        }
        // Verify the colon belongs to the element name (no whitespace before it).
        let candidate = String(xml[afterLT..<colonIdx])
        guard !candidate.isEmpty,
              candidate.allSatisfy({ $0.isLetter || $0.isNumber || $0 == "_" || $0 == "-" }) else {
            return nil
        }
        return candidate
    }

    private static func extractURIByPattern(xml: String, pattern: String) -> String? {
        guard let regex = try? NSRegularExpression(pattern: pattern) else { return nil }
        let ns = xml as NSString
        guard let match = regex.firstMatch(in: xml, range: NSRange(location: 0, length: ns.length)),
              match.numberOfRanges > 1 else { return nil }
        return ns.substring(with: match.range(at: 1))
    }
}

// MARK: - Internal: rPr propagation modes

internal extension RunProperties {
    /// Returns a copy of this `RunProperties` filtered by the given splice mode.
    ///
    /// - `.full`: returns self verbatim (deep copy via Equatable struct semantics).
    /// - `.omathOnly`: returns a new `RunProperties` with only OMath-rendering-relevant fields:
    ///   `rFonts`, `fontName`, `fontSize`, `lang`, `bold`, `italic`. Other fields (rStyle /
    ///   color / highlight / verticalAlign / etc.) are dropped.
    /// - `.discard`: returns `RunProperties()` (default-initialized).
    ///
    /// Spec: rPr propagation modes (Decision Q4)
    ///
    /// R2 fix (independent review of PsychQuant/ooxml-swift#117, Finding #2):
    /// this branch never copied `lang`, even though it was always part of the
    /// documented `.omathOnly` whitelist (spec.md's "rPr propagation modes"
    /// Requirement, and design.md, both list `rFonts`/`sz`/`lang`) — a
    /// pre-existing gap between the whitelist's own documentation and its
    /// implementation, predating (and not introduced by) the carrier fix in
    /// this same file. `lang` is copied verbatim (no on/off tri-state concern
    /// — `LanguageProperties` has no boolean fields, so there is no
    /// absent-vs-explicit-off distinction to preserve here, unlike bold/italic).
    func filteredForOMathSplice(mode: OMathSpliceRpRMode) -> RunProperties {
        switch mode {
        case .full:
            return self
        case .omathOnly:
            var out = RunProperties()
            out.rFonts = self.rFonts
            out.fontName = self.fontName
            out.fontSize = self.fontSize
            out.lang = self.lang
            if let bold = self.specifiedBold { out.bold = bold }
            if let italic = self.specifiedItalic { out.italic = italic }
            return out
        case .discard:
            return RunProperties()
        }
    }
}
