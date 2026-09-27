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
///
/// PsychQuant/ooxml-swift#165: also maps the three whitespace characters XML
/// 1.0 §3.3.3 requires a conformant reader to normalize to a plain space when
/// they appear LITERAL inside an attribute value — `\n`/`\r`/`\t` survive
/// intact only when written as character references:
/// - `\n` → `&#10;`
/// - `\r` → `&#13;`
/// - `\t` → `&#9;`
///
/// Without this, an attribute value containing a newline (e.g. base64
/// `o:gfxdata`, or user-entered `wp:docPr/@descr` alt text) written literally
/// survives one round-trip byte-identical, but any parser that re-reads it
/// normalizes the literal newline to a space — so the SECOND parse silently
/// changes the content. Same escaping `XmlTreeWriter.escapeAttributeValue`
/// already applies on the tree/raw channel (using decimal, not hex, refs —
/// both forms are valid XML and round-trip identically).
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
        case "\n": result += "&#10;"
        case "\r": result += "&#13;"
        case "\t": result += "&#9;"
        default: result.append(c)
        }
    }
    return result
}
