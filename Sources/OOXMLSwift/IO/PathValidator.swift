import Foundation

/// Validates that an OOXML relationship target is a safe relative path.
///
/// Background (che-word-mcp#55): `_rels/document.xml.rels` `Target` attributes
/// flow into `URL.appendingPathComponent`, which does NOT normalize `..`
/// traversal. Without validation a malicious .docx can read OR write outside
/// the intended `word/` directory at the user's UID.
///
/// Defense-in-depth: this validator runs at the parse boundary (DocxReader)
/// AND at property setters on `Header.originalFileName` / `Footer.originalFileName`.
///
/// Returns `true` when the path is acceptable for use as a relative OOXML
/// part filename. Rejects:
/// - Empty / oversized paths (DoS guard at 256 chars)
/// - Absolute paths (leading `/`)
/// - Parent-directory traversal (any `..` segment, including URL-encoded)
/// - Control characters (NUL, newline, etc. — would terminate path parsing in C-string sinks)
///
/// Accepts non-ASCII Unicode in the printable range (CJK-named header parts
/// are valid in Word installs with CJK locales).
public func isSafeRelativeOOXMLPath(_ path: String) -> Bool {
    // DoS guard
    guard !path.isEmpty, path.count <= 256 else {
        return false
    }

    // Absolute paths
    if path.hasPrefix("/") {
        return false
    }

    // URL-encoded traversal (decode-then-check: %2e%2e, %2E%2E, mixed case)
    let lowercased = path.lowercased()
    if lowercased.contains("%2e%2e") || lowercased.contains("%2f") || lowercased.contains("%5c") {
        return false
    }

    // Direct `..` segments (split by both `/` and `\`)
    let segments = path.split { $0 == "/" || $0 == "\\" }
    for segment in segments where segment == ".." {
        return false
    }

    // Control characters (NUL, newlines, tabs — anything < 0x20 or 0x7F)
    for scalar in path.unicodeScalars {
        if scalar.value < 0x20 || scalar.value == 0x7F {
            return false
        }
    }

    return true
}
