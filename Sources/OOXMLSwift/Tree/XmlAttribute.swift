// Lossless representation of an XML attribute on an `XmlNode`.
//
// Attribute order in the source XML is preserved by storing attributes as a
// `[XmlAttribute]` on `XmlNode`. Namespace declarations (`xmlns="..."` and
// `xmlns:*="..."`) are NOT split into a separate field — they appear as
// regular attributes whose `prefix == "xmlns"` (declaration of a prefixed NS)
// or whose `prefix == nil` and `localName == "xmlns"` (default NS declaration).
// This keeps round-trip serialization simple and order-preserving.

import Foundation

/// A single attribute on an XML element. Order in the parent's
/// `attributes` array reflects source XML order and SHALL be preserved
/// across read-then-write to satisfy the byte-equal round-trip contract.
public struct XmlAttribute: Equatable, Hashable, Sendable {

    /// Namespace prefix declared on the attribute, e.g. `"w"` for
    /// `w:val="..."`. `nil` for unprefixed attributes (`val="..."`).
    public var prefix: String?

    /// Local name without prefix.
    public var localName: String

    /// Attribute value as it appears after entity decoding. The reader
    /// resolves `&amp;` / `&lt;` / `&gt;` / `&quot;` / `&apos;` and
    /// numeric character references; the writer re-encodes on emit.
    public var value: String

    public init(prefix: String? = nil, localName: String, value: String) {
        self.prefix = prefix
        self.localName = localName
        self.value = value
    }

    // MARK: - Convenience predicates

    /// True when this attribute is a namespace declaration (`xmlns` or
    /// `xmlns:foo`). Useful for separating namespace decls from data
    /// attributes when comparing structural fingerprints.
    public var isNamespaceDeclaration: Bool {
        if prefix == "xmlns" {
            return true
        }
        if prefix == nil && localName == "xmlns" {
            return true
        }
        return false
    }

    /// The `xmlns` prefix this attribute declares (e.g. `"w"` for
    /// `xmlns:w="..."`), or empty string for the default-NS declaration
    /// (`xmlns="..."`). `nil` for non-namespace attributes.
    public var declaredNamespacePrefix: String? {
        if prefix == "xmlns" { return localName }
        if prefix == nil && localName == "xmlns" { return "" }
        return nil
    }

    /// True when this attribute is in the rsids identity-noise set.
    /// The `XmlNode.normalizedFingerprint()` API skips these when computing
    /// comparison-stable fingerprints (spec `ooxml-tree-io` requirement
    /// "Identity-noise normalization for diff comparison").
    public var isRsidNoise: Bool {
        guard prefix == "w" else { return false }
        switch localName {
        case "rsidR", "rsidRPr", "rsidP", "rsidRDefault", "rsidSect", "rsidTr":
            return true
        default:
            return false
        }
    }

    // MARK: - Qualified-name helpers

    /// Qualified name combining prefix and local name (e.g. `"w:val"`).
    /// Returns just `localName` when no prefix is set.
    public var qualifiedName: String {
        if let p = prefix, !p.isEmpty {
            return "\(p):\(localName)"
        }
        return localName
    }
}
