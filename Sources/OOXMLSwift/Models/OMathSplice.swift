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
    /// Original local name for a paragraph direct-child carrier.
    /// Nil for inline Run carriers.
    let directChildName: String?
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
            collected.append(ExtractedOMath(
                xml: ensureXmlnsDeclared(in: raw),
                kind: .inRun,
                directChildName: nil,
                sourcePosition: run.position,
                sourceRunProperties: run.properties
            ))
        }

        // Carrier 2: Paragraph.unrecognizedChildren (direct-child OMath)
        for child in paragraph.unrecognizedChildren where child.name == "oMath" || child.name == "oMathPara" {
            collected.append(ExtractedOMath(
                xml: ensureXmlnsDeclared(in: child.rawXML),
                kind: .directChild,
                directChildName: child.name,
                sourcePosition: child.position,
                sourceRunProperties: nil
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
    /// Heuristic: scans for `xmlns:<prefix>="<URI>"` where the prefix is the same
    /// one used in the opening element name (e.g. `<mml:oMath xmlns:mml="...">` → `mml`).
    /// Falls back to scanning for any `xmlns:` declaration if prefix detection fails.
    static func extractURI(from xml: String) -> String? {
        guard let root = parseRootTag(in: xml) else { return nil }
        if let prefix = prefix(fromQualifiedName: root.qualifiedName) {
            return root.attributes["xmlns:\(prefix)"]
        }
        return root.attributes["xmlns"]
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
    private static let xmlNamespaceURI = "http://www.w3.org/XML/1998/namespace"

    static func canonicalRepresentation(of xml: String) throws -> String {
        let data = Data(xml.utf8)
        let tree = try XmlTreeReader.parse(data)
        return canonicalNode(tree.root, inheritedNamespaces: [:])
    }

    static func isEquivalent(_ lhs: String, _ rhs: String) throws -> Bool {
        try canonicalRepresentation(of: lhs) == canonicalRepresentation(of: rhs)
    }

    private static func canonicalNode(
        _ node: XmlNode,
        inheritedNamespaces: [String: String]
    ) -> String {
        switch node.kind {
        case .element:
            var namespaces = inheritedNamespaces
            for attribute in node.attributes where attribute.isNamespaceDeclaration {
                namespaces[attribute.declaredNamespacePrefix ?? ""] = attribute.value
            }

            let elementName = expandedName(
                namespaceURI: node.namespaceURI ?? "",
                localName: node.localName
            )
            let attributes = node.attributes
                .filter { !$0.isNamespaceDeclaration }
                .map { attribute -> String in
                    let namespaceURI: String
                    if let prefix = attribute.prefix {
                        namespaceURI = prefix == "xml"
                            ? xmlNamespaceURI
                            : namespaces[prefix, default: "unbound:\(prefix)"]
                    } else {
                        // XML default namespaces never apply to attributes.
                        namespaceURI = ""
                    }
                    return framed(expandedName(
                        namespaceURI: namespaceURI,
                        localName: attribute.localName
                    )) + framed(attribute.value)
                }
                .sorted()
                .joined()
            let children = node.children
                .map { canonicalNode($0, inheritedNamespaces: namespaces) }
                .joined()
            return "E" + framed(elementName) + "A" + framed(attributes) + "C" + framed(children)

        case .text:
            return "T" + framed(node.textContent)

        case .comment:
            return "M" + framed(node.textContent)

        case .processingInstruction:
            return "P" + framed(node.processingInstructionTarget) + framed(node.textContent)
        }
    }

    private static func expandedName(namespaceURI: String, localName: String) -> String {
        "{\(namespaceURI)}\(localName)"
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
    ///   `rFonts`, `fontName`, `fontSize`, `bold`, `italic`. Other fields (rStyle / color /
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
            return out
        case .discard:
            return RunProperties()
        }
    }
}
