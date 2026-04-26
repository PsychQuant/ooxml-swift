import Foundation

/// Escapes a caller-provided String for safe interpolation into an OOXML attribute value.
///
/// Maps the five XML attribute special characters:
/// - `&` → `&amp;`
/// - `<` → `&lt;`
/// - `>` → `&gt;`
/// - `"` → `&quot;`
/// - `'` → `&apos;`
///
/// `&apos;` (not `&#39;`) matches Microsoft Word's emit output for byte-equivalent
/// round-trip with documents authored in Word.
internal func escapeXMLAttribute(_ s: String) -> String {
    var result = ""
    result.reserveCapacity(s.count)
    for c in s {
        switch c {
        case "&": result += "&amp;"
        case "<": result += "&lt;"
        case ">": result += "&gt;"
        case "\"": result += "&quot;"
        case "'": result += "&apos;"
        default: result.append(c)
        }
    }
    return result
}
