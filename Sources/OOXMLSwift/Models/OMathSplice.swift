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
/// `.strict` throws on any prefix or URI mismatch — useful for callers
/// requiring a single-prefix output convention.
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

/// A source OMath payload was not one embeddable, namespace-valid XML
/// fragment. Kept separate from `OMathSpliceError` so adding malformed-input
/// reporting does not break external exhaustive switches over that released
/// public enum's original six cases.
public struct OMathSpliceMalformedXMLError: Error, Equatable {
    public init() {}
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
    /// Original local name for a paragraph direct-child carrier.
    /// Nil for inline Run carriers.
    let directChildName: String?
    /// Absolute serializer lane metadata used by atStart/atEnd insertions.
    let boundaryPlacement: ParagraphBoundaryPlacement?
    let boundaryOrder: Int?
    /// Deterministic tie-breaker matching the serializer's collection order.
    let sourceSequence: Int
    /// Source-document byte position (filled by DocxReader for both Run and UnrecognizedChild).
    /// Used for joint sort across the two carriers. May be nil for API-built paragraphs.
    let sourcePosition: Int?
    /// For `.inRun` kind, the source Run's properties (used for `.full` / `.omathOnly` rPr propagation).
    /// For `.directChild`, nil (no enclosing Run).
    let sourceRunProperties: RunProperties?

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
    /// Joint sort mirrors Paragraph's four serializer regions: absolute start,
    /// positive positions, non-positive post-content, and absolute end.
    /// Stable collection sequence breaks ties across Run/direct-child carriers.
    ///
    /// Spec: Carrier preservation strategy (Decision Q1)
    static func extract(from paragraph: Paragraph) -> [ExtractedOMath] {
        var collected: [ExtractedOMath] = []
        var sourceSequence = 0

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
            collected.append(ExtractedOMath(
                xml: ensureXmlnsDeclared(in: raw),
                kind: .inRun,
                directChildName: nil,
                boundaryPlacement: run.paragraphBoundaryPlacement,
                boundaryOrder: run.paragraphBoundaryOrder,
                sourceSequence: sourceSequence,
                sourcePosition: run.position,
                sourceRunProperties: run.properties
            ))
            sourceSequence += 1
        }

        // Carrier 2: Paragraph.unrecognizedChildren (direct-child OMath)
        for child in paragraph.unrecognizedChildren where child.name == "oMath" || child.name == "oMathPara" {
            collected.append(ExtractedOMath(
                xml: ensureXmlnsDeclared(in: child.rawXML),
                kind: .directChild,
                directChildName: child.name,
                boundaryPlacement: child.paragraphBoundaryPlacement,
                boundaryOrder: child.paragraphBoundaryOrder,
                sourceSequence: sourceSequence,
                sourcePosition: child.position,
                sourceRunProperties: nil
            ))
            sourceSequence += 1
        }

        // Mirror Paragraph's four serializer regions: absolute start, positive
        // position window, nil/zero post-content buckets, absolute end.
        return collected.sorted { lhs, rhs in
            let left = serializerSortKey(lhs)
            let right = serializerSortKey(rhs)
            if left.bucket != right.bucket { return left.bucket < right.bucket }
            if left.order != right.order { return left.order < right.order }
            return lhs.sourceSequence < rhs.sourceSequence
        }
    }

    private static func serializerSortKey(_ omath: ExtractedOMath) -> (bucket: Int, order: Int) {
        switch omath.boundaryPlacement {
        case .start?: return (0, omath.boundaryOrder ?? 0)
        case .end?: return (3, omath.boundaryOrder ?? 0)
        case nil:
            if let position = omath.sourcePosition, position > 0 {
                return (1, position)
            }
            return (2, 0)
        }
    }

    /// If the given OMath rawXML's opening tag lacks an `xmlns:<prefix>="<URI>"`
    /// declaration for the OMath namespace prefix, inject one with the standard URI.
    /// This is required when the source-side parser inherits xmlns from the parent
    /// `<w:p>` but `XMLElement.xmlString` doesn't carry inherited declarations
    /// — the extracted rawXML must be self-contained for round-trip correctness.
    static func ensureXmlnsDeclared(in xml: String) -> String {
        let prefix = OMathNamespace.extractPrefix(from: xml)
        // A local declaration using either legal quote style is already
        // self-contained. Do not rewrite its URI or lexical form.
        if OMathNamespace.extractURI(from: xml) != nil {
            return xml
        }

        // Inject after the tokenizer-bounded root QName. Prefix-less fragments
        // receive a default OMML namespace; prefixed fragments receive
        // xmlns:prefix. Attribute names/values cannot influence this boundary.
        guard let nameEnd = OMathNamespace.rootNameEnd(in: xml) else { return xml }
        let standardURI = "http://schemas.openxmlformats.org/officeDocument/2006/math"
        let injection = prefix.map { " xmlns:\($0)=\"\(standardURI)\"" }
            ?? " xmlns=\"\(standardURI)\""
        return String(xml[..<nameEnd]) + injection + String(xml[nameEnd...])
    }
}

// MARK: - Internal: namespace inspection helpers

internal enum OMathNamespace {
    private struct RootTag {
        let qualifiedName: String
        let nameEnd: String.Index
        let attributes: [String: String]
    }

    /// Extracts the `xmlns:` URI for the OMath prefix in the given XML fragment.
    /// Returns the URI string, or nil if no `xmlns:` declaration found.
    ///
    /// Parses the complete fragment with XmlTreeReader so namespace attribute
    /// character references are resolved and malformed XML fails closed.
    static func extractURI(from xml: String) -> String? {
        guard let tree = try? XmlTreeReader.parse(Data(xml.utf8)) else { return nil }
        return tree.root.namespaceURI
    }

    /// Extracts the OMath prefix (e.g. "m" or "mml") from the opening element.
    /// Returns nil if no prefix used (default namespace).
    static func extractPrefix(from xml: String) -> String? {
        guard let root = parseRootTag(in: xml) else { return nil }
        return prefix(fromQualifiedName: root.qualifiedName)
    }

    internal static func rootNameEnd(in xml: String) -> String.Index? {
        parseRootTag(in: xml)?.nameEnd
    }

    internal static func isWellFormed(_ xml: String) -> Bool {
        validatedFragmentRoot(in: xml) != nil
    }

    /// Parses an embeddable OMath fragment rather than a complete XML document.
    /// Wrapping the candidate makes document-level declarations and DTDs invalid
    /// at their actual syntactic positions while leaving the same text inside an
    /// OMath comment or CDATA section legal. The wrapper-child check also rejects
    /// leading/trailing comments, processing instructions, non-whitespace text,
    /// and multiple roots.
    internal static func validatedFragmentRoot(in xml: String) -> XMLElement? {
        let fragment = trimmingXMLS(in: xml)
        guard fragment.first == "<",
              fragment.last == ">",
              parseRootTag(in: String(fragment)) != nil,
              (try? XmlTreeReader.parse(Data(fragment.utf8))) != nil else {
            return nil
        }

        let wrapped = "<iddOMathFragment>\(fragment)</iddOMathFragment>"
        guard let document = try? XMLDocument(data: Data(wrapped.utf8)),
              let wrapper = document.rootElement(),
              wrapper.localName == "iddOMathFragment" else {
            return nil
        }

        var roots: [XMLElement] = []
        for child in wrapper.children ?? [] {
            switch child.kind {
            case .text:
                guard (child.stringValue ?? "")
                    .trimmingCharacters(in: .whitespacesAndNewlines)
                    .isEmpty else {
                    return nil
                }
            case .element:
                guard let element = child as? XMLElement else { return nil }
                roots.append(element)
            default:
                return nil
            }
        }

        guard roots.count == 1,
              let localName = roots[0].localName,
              localName == "oMath" || localName == "oMathPara" else {
            return nil
        }
        return roots[0]
    }

    /// XML 1.0 `S` is exactly space, tab, carriage return, and line feed.
    /// Foundation treats a leading U+FEFF as an ignorable BOM even when it is
    /// interpolated inside the synthetic wrapper; trimming only XML `S` and
    /// requiring the first remaining character to be `<` keeps admission,
    /// retained rawXML, and serialization consistent.
    private static func trimmingXMLS(in xml: String) -> Substring {
        let scalars = xml.unicodeScalars
        var lower = scalars.startIndex
        var upper = scalars.endIndex
        while lower < upper, isXMLS(scalars[lower]) {
            lower = scalars.index(after: lower)
        }
        while upper > lower {
            let previous = scalars.index(before: upper)
            guard isXMLS(scalars[previous]) else { break }
            upper = previous
        }
        return xml[lower..<upper]
    }

    private static func isXMLS(_ scalar: Unicode.Scalar) -> Bool {
        scalar.value == 0x20 || scalar.value == 0x09
            || scalar.value == 0x0D || scalar.value == 0x0A
    }

    private static func prefix(fromQualifiedName qualifiedName: String) -> String? {
        let parts = qualifiedName.split(separator: ":", omittingEmptySubsequences: false)
        guard parts.count == 2,
              !parts[0].isEmpty,
              !parts[1].isEmpty,
              parts.allSatisfy({ part in
                  part.allSatisfy { character in
                      character.isLetter || character.isNumber
                          || character == "_" || character == "-" || character == "."
                  }
              }) else {
            return nil
        }
        return String(parts[0])
    }

    /// Quote-aware tokenizer for only the root opening tag. It reads the QName
    /// before any attribute, then parses actual name="value" pairs. This avoids
    /// treating URI colons or fake `xmlns` text inside ordinary values as syntax.
    private static func parseRootTag(in xml: String) -> RootTag? {
        var cursor = xml.startIndex

        func advancePastWhitespace() {
            while cursor < xml.endIndex && xml[cursor].isWhitespace {
                cursor = xml.index(after: cursor)
            }
        }

        advancePastWhitespace()
        guard cursor < xml.endIndex, xml[cursor] == "<" else { return nil }
        cursor = xml.index(after: cursor)
        guard cursor < xml.endIndex, xml[cursor] != "?", xml[cursor] != "!", xml[cursor] != "/" else {
            return nil
        }

        let nameStart = cursor
        while cursor < xml.endIndex,
              !xml[cursor].isWhitespace,
              xml[cursor] != "/",
              xml[cursor] != ">" {
            cursor = xml.index(after: cursor)
        }
        guard cursor > nameStart else { return nil }
        let qualifiedName = String(xml[nameStart..<cursor])
        let nameEnd = cursor
        var attributes: [String: String] = [:]

        while cursor < xml.endIndex {
            advancePastWhitespace()
            guard cursor < xml.endIndex, xml[cursor] != ">", xml[cursor] != "/" else { break }

            let attributeStart = cursor
            while cursor < xml.endIndex,
                  !xml[cursor].isWhitespace,
                  xml[cursor] != "=",
                  xml[cursor] != ">",
                  xml[cursor] != "/" {
                cursor = xml.index(after: cursor)
            }
            let attributeName = String(xml[attributeStart..<cursor])
            advancePastWhitespace()
            guard !attributeName.isEmpty,
                  cursor < xml.endIndex,
                  xml[cursor] == "=" else {
                return nil
            }
            cursor = xml.index(after: cursor)
            advancePastWhitespace()
            guard cursor < xml.endIndex,
                  xml[cursor] == "\"" || xml[cursor] == "'" else {
                return nil
            }
            let quote = xml[cursor]
            cursor = xml.index(after: cursor)
            let valueStart = cursor
            while cursor < xml.endIndex, xml[cursor] != quote {
                cursor = xml.index(after: cursor)
            }
            guard cursor < xml.endIndex else { return nil }
            attributes[attributeName] = String(xml[valueStart..<cursor])
            cursor = xml.index(after: cursor)
        }

        return RootTag(
            qualifiedName: qualifiedName,
            nameEnd: nameEnd,
            attributes: attributes
        )
    }
}

// MARK: - Internal: semantic XML equivalence (#123)

/// Canonical semantic representation for the public save/reload fidelity
/// contract. Unlike lexical XML, this representation ignores prefix spelling,
/// declaration placement, attribute order/quote style, and entity spelling,
/// while preserving namespace-expanded names, values, text, and child order.
internal enum OMathSemanticXML {
    static func canonicalRepresentation(of xml: String) throws -> String {
        guard let root = OMathNamespace.validatedFragmentRoot(in: xml) else {
            throw OMathSpliceMalformedXMLError()
        }
        return canonicalElement(root)
    }

    static func isEquivalent(_ lhs: String, _ rhs: String) throws -> Bool {
        do {
            return try canonicalRepresentation(of: lhs) == canonicalRepresentation(of: rhs)
        } catch {
            return false
        }
    }

    /// Foundation/libxml2 supplies the XML-processor semantics that matter for
    /// this contract: namespace scope (including `xmlns=""` undeclaration),
    /// entity expansion, and XML attribute whitespace normalization. Prefixes,
    /// namespace declaration placement, quote style, and attribute order are
    /// intentionally absent from the representation.
    private static func canonicalElement(_ element: XMLElement) -> String {
        let elementName = namespaceKey(uri: element.uri) + framed(element.localName ?? element.name ?? "")
        let attributes = (element.attributes ?? [])
            .filter { attribute in
                attribute.name != "xmlns" && attribute.prefix != "xmlns"
            }
            .map { attribute -> String in
                let name = namespaceKey(uri: attribute.uri)
                    + framed(attribute.localName ?? attribute.name ?? "")
                return framed(name) + framed(attribute.stringValue ?? "")
            }
            .sorted()
            .joined()

        var children = ""
        var pendingText = ""
        func flushText() -> String {
            guard !pendingText.isEmpty else { return "" }
            defer { pendingText = "" }
            return "T" + framed(pendingText)
        }
        for child in element.children ?? [] {
            switch child.kind {
            case .text:
                pendingText += child.stringValue ?? ""
            case .element:
                children += flushText()
                if let childElement = child as? XMLElement {
                    children += canonicalElement(childElement)
                }
            case .comment:
                children += flushText()
                children += "M" + framed(child.stringValue ?? "")
            case .processingInstruction:
                children += flushText()
                children += "P" + framed(child.name ?? "") + framed(child.stringValue ?? "")
            default:
                children += flushText()
                children += "X" + framed(child.xmlString)
            }
        }
        children += flushText()
        return "E" + framed(elementName) + "A" + framed(attributes) + "C" + framed(children)
    }

    private static func namespaceKey(uri: String?) -> String {
        guard let uri, !uri.isEmpty else { return "N" }
        return "B" + framed(uri)
    }

    private static func framed(_ value: String) -> String {
        "\(value.utf8.count):\(value)"
    }
}

// MARK: - Internal: rPr propagation modes

internal extension RunProperties {
    /// Returns a copy of this `RunProperties` filtered by the given splice mode.
    ///
    /// - `.full`: returns self verbatim (deep copy via Equatable struct semantics).
    /// - `.omathOnly`: returns a new `RunProperties` with only OMath-rendering-relevant fields:
    ///   `rFonts`, `fontName`, `fontSize`, `bold`, `italic`, `lang`. Other fields (rStyle / color /
    ///   highlight / verticalAlign / etc.) are dropped.
    /// - `.discard`: returns `RunProperties()` (default-initialized).
    ///
    /// Spec: rPr propagation modes (Decision Q4)
    func filteredForOMathSplice(mode: OMathSpliceRpRMode) -> RunProperties {
        switch mode {
        case .full:
            return self
        case .omathOnly:
            var out = RunProperties()
            out.rFonts = self.rFonts
            out.fontName = self.fontName
            out.fontSize = self.fontSize
            out.bold = self.bold
            out.italic = self.italic
            out.lang = self.lang
            return out
        case .discard:
            return RunProperties()
        }
    }
}
