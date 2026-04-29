import Foundation

/// XML input hardening errors thrown by `DocxReader` and `DocxWriter` at the
/// raw-bytes / root-attribute boundary.
///
/// **Origin**: PsychQuant/ooxml-swift#7 — bundled hardening pass for findings
/// F10 (DTD reject), F12 (attr-name whitelist), and F14 (attr-value byte cap)
/// surfaced by the verification of PsychQuant/che-word-mcp#56.
///
/// **Caller guidance**: these throws fire on already-malformed or potentially
/// malicious input. Callers SHOULD propagate the error so the offending part
/// can be surfaced to the user. Silencing via `try?` swallows corruption
/// signal — only do so when degrade-not-crash is an explicit product
/// requirement (e.g., partial-recovery batch readers) and you log the error
/// elsewhere.
public enum XMLHardeningError: Error, LocalizedError, Equatable {
    /// Input bytes contained a `<!DOCTYPE` declaration. OOXML disallows DTDs
    /// (ECMA-376 part 1 §17.18.42 + part 2 packaging), so the pre-scan
    /// rejects the input before `XMLDocument(data:)` can begin entity
    /// expansion. The `part` argument identifies which OOXML part triggered
    /// the reject (e.g., `"word/document.xml"`, `"word/header1.xml"`).
    case dtdNotAllowed(part: String)

    /// A root-level attribute name failed the XML 1.0 NameChar whitelist
    /// regex `^[A-Za-z_:][A-Za-z0-9._:-]*$`. The `context` argument is a
    /// namespaced string identifying the throw site (e.g.,
    /// `"split-attributes"` for reader-side ingest, `"document root"` for
    /// writer-side emit) so the caller can disambiguate.
    case invalidAttributeName(name: String, context: String)

    /// A root-level attribute value exceeded the per-attribute byte cap
    /// (currently 64 KiB, far above any legitimate `mc:Ignorable` /
    /// `xmlns:*` value observed in the wild). The `byteSize` is the
    /// measured UTF-8 byte length; `cap` is the configured maximum (also
    /// included so callers can report the limit without re-deriving it).
    case attributeValueTooLarge(name: String, byteSize: Int, cap: Int)

    public var errorDescription: String? {
        switch self {
        case let .dtdNotAllowed(part):
            return "DTD declarations are not permitted in OOXML input (part: \(part))."
        case let .invalidAttributeName(name, context):
            return "Invalid XML attribute name '\(name)' in \(context); names must match XML 1.0 NameChar production."
        case let .attributeValueTooLarge(name, byteSize, cap):
            return "XML attribute '\(name)' value exceeds size cap: \(byteSize) bytes > \(cap) bytes."
        }
    }
}
