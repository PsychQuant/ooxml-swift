import Foundation

// MARK: - Container Root Tag rendering (PsychQuant/che-word-mcp#56 follow-up F4)
//
// Shared helper for rebuilding the root open tag of any OOXML container part
// (`<w:hdr>`, `<w:ftr>`, `<w:footnotes>`, `<w:endnotes>`) from the captured
// `rootAttributes: [String: String]` map populated by the Reader.
//
// Mirrors `DocxWriter.renderDocumentRootOpenTag` but generalized over the
// container element name so every part type — not just `word/document.xml` —
// preserves source `xmlns:*` declarations + `mc:Ignorable` through a no-op
// round-trip. Without this, headers / footers with VML watermarks (which
// commonly declare `mc`/`wp`/`w14`/`w15` beyond the hardcoded 5-namespace
// template) silently regress to the unbound-prefix bug from #56.

enum ContainerRootTag {
    /// Default minimal namespace template per container element. Used as the
    /// fallback when `attributes` is empty (API-built parts that never went
    /// through the Reader). Headers / footers need the VML namespace family
    /// (`v`/`o`/`w10`) for watermarks; footnotes / endnotes only need the
    /// core wordprocessingml + relationships pair.
    static func defaultAttributes(for elementName: String) -> [(String, String)] {
        switch elementName {
        case "w:hdr", "w:ftr":
            return [
                ("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
                ("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
                ("xmlns:v", "urn:schemas-microsoft-com:vml"),
                ("xmlns:o", "urn:schemas-microsoft-com:office:office"),
                ("xmlns:w10", "urn:schemas-microsoft-com:office:word"),
            ]
        case "w:footnotes", "w:endnotes":
            return [
                ("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
                ("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
            ]
        default:
            return [
                ("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
            ]
        }
    }

    /// Render the opening tag including every attribute, ending with `>`.
    ///
    /// **When `attributes` is non-empty** (Reader path): emit `xmlns:w` first
    /// (so the default prefix is parser-stable), then `xmlns:r`, then every
    /// remaining `xmlns:*` alphabetically, then every non-namespace attribute
    /// alphabetically. Mirrors `DocxWriter.renderDocumentRootOpenTag` order
    /// for consistency across the codebase.
    ///
    /// **When `attributes` is empty** (API-built path): emit the
    /// element-specific default template from `defaultAttributes(for:)`.
    static func render(elementName: String, attributes: [String: String]) -> String {
        if attributes.isEmpty {
            var pieces: [String] = ["<\(elementName)"]
            for (name, value) in defaultAttributes(for: elementName) {
                pieces.append("\(name)=\"\(escapeAttr(value))\"")
            }
            return pieces.joined(separator: " ") + ">"
        }

        var xmlnsW: String? = nil
        var xmlnsR: String? = nil
        var otherXmlns: [(String, String)] = []
        var nonNamespace: [(String, String)] = []

        for (name, value) in attributes {
            if name == "xmlns:w" {
                xmlnsW = value
            } else if name == "xmlns:r" {
                xmlnsR = value
            } else if name.hasPrefix("xmlns:") {
                otherXmlns.append((name, value))
            } else {
                nonNamespace.append((name, value))
            }
        }

        otherXmlns.sort { $0.0 < $1.0 }
        nonNamespace.sort { $0.0 < $1.0 }

        var pieces: [String] = ["<\(elementName)"]

        let xmlnsWValue = xmlnsW ?? "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        let xmlnsRValue = xmlnsR ?? "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        pieces.append("xmlns:w=\"\(escapeAttr(xmlnsWValue))\"")
        pieces.append("xmlns:r=\"\(escapeAttr(xmlnsRValue))\"")

        for (name, value) in otherXmlns {
            pieces.append("\(name)=\"\(escapeAttr(value))\"")
        }
        for (name, value) in nonNamespace {
            pieces.append("\(name)=\"\(escapeAttr(value))\"")
        }

        return pieces.joined(separator: " ") + ">"
    }

    private static func escapeAttr(_ s: String) -> String {
        return s
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
    }
}
