// Hand-rolled, byte-offset-tracking XML parser for OOXML parts.
//
// Foundation `XMLParser` is a SAX-style parser that does not expose byte
// offsets per token. The byte-equal round-trip contract requires every node
// to know its source byte range, so this module implements a small
// recursive-descent parser specifically for the well-formed OOXML subset
// that Word produces. It is NOT a general-purpose XML parser:
//   - Assumes input is well-formed UTF-8 (Word always writes UTF-8).
//   - Does not resolve external entities or DOCTYPE.
//   - Supports the XML 1.0 productions actually present in OOXML files:
//     processing instructions, comments, elements, attributes, character
//     references, mixed content.
//
// Pure-Swift implementation; no `Foundation.XMLDocument` or `libxml2`.

import Foundation

public enum XmlTreeReaderError: Error, Equatable {
    case emptyInput
    case unexpectedEOF(at: Int)
    case unexpectedCharacter(Character, at: Int)
    case malformedAttribute(at: Int)
    case malformedTag(at: Int)
    case mismatchedClosingTag(expected: String, found: String, at: Int)
    case invalidEntityReference(reference: String, at: Int)
    case invalidNumericReference(reference: String, at: Int)
    case unexpectedTrailingContent(at: Int)
}

/// Parses a UTF-8 XML document into a lossless `XmlTree`.
public enum XmlTreeReader {

    /// Parse the given UTF-8 bytes into an `XmlTree`. The returned tree's
    /// `root` corresponds to the document element of the input. The original
    /// `data` is retained on the tree so the writer can copy clean
    /// sub-trees byte-equal.
    public static func parse(_ data: Data) throws -> XmlTree {
        guard !data.isEmpty else { throw XmlTreeReaderError.emptyInput }

        var parser = Parser(bytes: Array(data))
        try parser.skipProlog()
        let root = try parser.parseElement()
        try parser.skipTrailingWhitespaceAndComments()
        guard parser.position == parser.bytes.count else {
            throw XmlTreeReaderError.unexpectedTrailingContent(at: parser.position)
        }
        // Resolve namespace URIs throughout the tree so callers can compare
        // by URI instead of prefix.
        resolveNamespaces(root, scope: NamespaceScope())
        return XmlTree(root: root, sourceBytes: data)
    }

    // MARK: - Namespace resolution

    private static func resolveNamespaces(_ node: XmlNode, scope: NamespaceScope) {
        guard node.kind == .element else { return }
        var nextScope = scope
        for attribute in node.attributes where attribute.isNamespaceDeclaration {
            // declaredNamespacePrefix returns "" for default-NS decls.
            let prefix = attribute.declaredNamespacePrefix ?? ""
            nextScope.bind(prefix: prefix, uri: attribute.value)
        }
        node.namespaceURI = nextScope.lookup(prefix: node.prefix ?? "")
        for child in node.children {
            resolveNamespaces(child, scope: nextScope)
        }
    }
}

// MARK: - Namespace scope (immutable up-stack lookups)

private struct NamespaceScope {
    private var bindings: [String: String] = [:]

    mutating func bind(prefix: String, uri: String) {
        bindings[prefix] = uri
    }

    func lookup(prefix: String) -> String? {
        bindings[prefix]
    }
}

// MARK: - Parser state machine

private struct Parser {
    let bytes: [UInt8]
    var position: Int = 0

    init(bytes: [UInt8]) {
        self.bytes = bytes
    }

    // MARK: - Prolog (XML decl + leading PI / comment / whitespace)

    mutating func skipProlog() throws {
        skipWhitespace()
        // Optional XML declaration: <?xml ... ?>
        if peekString("<?xml") {
            _ = try parseProcessingInstruction(asNode: false)
            skipWhitespace()
        }
        // Optional sequence of comments / PIs / whitespace before root.
        while position < bytes.count {
            if peekString("<!--") {
                _ = try parseComment()
                skipWhitespace()
            } else if peekString("<?") {
                _ = try parseProcessingInstruction(asNode: false)
                skipWhitespace()
            } else if peekString("<!DOCTYPE") {
                try skipDoctype()
                skipWhitespace()
            } else {
                break
            }
        }
    }

    mutating func skipTrailingWhitespaceAndComments() throws {
        while position < bytes.count {
            skipWhitespace()
            if peekString("<!--") {
                _ = try parseComment()
            } else if peekString("<?") {
                _ = try parseProcessingInstruction(asNode: false)
            } else {
                break
            }
        }
    }

    // MARK: - Element

    mutating func parseElement() throws -> XmlNode {
        let elementStart = position
        guard consumeByte(0x3C /* '<' */) else { // expect '<'
            throw XmlTreeReaderError.unexpectedCharacter(currentCharacter() ?? "?", at: position)
        }
        // Tag name
        let qualifiedName = try parseName()
        let (prefix, localName) = splitQualifiedName(qualifiedName)
        // Attributes
        var attributes: [XmlAttribute] = []
        while true {
            skipWhitespace()
            if position >= bytes.count {
                throw XmlTreeReaderError.unexpectedEOF(at: position)
            }
            let b = bytes[position]
            if b == 0x2F /* '/' */ {
                // Self-closing element: />
                position += 1
                guard consumeByte(0x3E /* '>' */) else {
                    throw XmlTreeReaderError.malformedTag(at: position)
                }
                let elementEnd = position
                return XmlNode.element(
                    prefix: prefix,
                    localName: localName,
                    attributes: attributes,
                    children: [],
                    sourceRange: elementStart..<elementEnd
                )
            }
            if b == 0x3E /* '>' */ {
                position += 1
                break
            }
            // Otherwise must be an attribute.
            attributes.append(try parseAttribute())
        }
        // Parse children until matching close tag.
        var children: [XmlNode] = []
        while true {
            if position >= bytes.count {
                throw XmlTreeReaderError.unexpectedEOF(at: position)
            }
            if peekString("</") {
                break
            }
            if peekString("<!--") {
                children.append(try parseComment())
                continue
            }
            if peekString("<?") {
                children.append(try parseProcessingInstruction(asNode: true))
                continue
            }
            if peekString("<![CDATA[") {
                children.append(try parseCDATA())
                continue
            }
            if bytes[position] == 0x3C /* '<' */ {
                children.append(try parseElement())
                continue
            }
            children.append(try parseText())
        }
        // Closing tag: </name>
        let closeStart = position
        guard consumeString("</") else {
            throw XmlTreeReaderError.malformedTag(at: position)
        }
        let closingName = try parseName()
        if closingName != qualifiedName {
            throw XmlTreeReaderError.mismatchedClosingTag(
                expected: qualifiedName, found: closingName, at: closeStart
            )
        }
        skipWhitespace()
        guard consumeByte(0x3E /* '>' */) else {
            throw XmlTreeReaderError.malformedTag(at: position)
        }
        let elementEnd = position
        return XmlNode.element(
            prefix: prefix,
            localName: localName,
            attributes: attributes,
            children: children,
            sourceRange: elementStart..<elementEnd
        )
    }

    // MARK: - Attribute

    mutating func parseAttribute() throws -> XmlAttribute {
        let qualifiedName = try parseName()
        let (prefix, localName) = splitQualifiedName(qualifiedName)
        skipWhitespace()
        guard consumeByte(0x3D /* '=' */) else {
            throw XmlTreeReaderError.malformedAttribute(at: position)
        }
        skipWhitespace()
        let quote: UInt8
        if position < bytes.count, bytes[position] == 0x22 /* '"' */ {
            quote = 0x22
        } else if position < bytes.count, bytes[position] == 0x27 /* '\'' */ {
            quote = 0x27
        } else {
            throw XmlTreeReaderError.malformedAttribute(at: position)
        }
        position += 1
        let valueStart = position
        while position < bytes.count, bytes[position] != quote {
            position += 1
        }
        guard position < bytes.count else {
            throw XmlTreeReaderError.unexpectedEOF(at: position)
        }
        let rawValue = stringFromBytes(start: valueStart, end: position)
        position += 1 // closing quote
        let decoded = try decodeXmlEntities(rawValue, position: valueStart)
        return XmlAttribute(prefix: prefix, localName: localName, value: decoded)
    }

    // MARK: - Comment

    mutating func parseComment() throws -> XmlNode {
        let start = position
        guard consumeString("<!--") else {
            throw XmlTreeReaderError.malformedTag(at: position)
        }
        let bodyStart = position
        while position + 2 < bytes.count {
            if bytes[position] == 0x2D /* '-' */ &&
               bytes[position + 1] == 0x2D /* '-' */ &&
               bytes[position + 2] == 0x3E /* '>' */ {
                let body = stringFromBytes(start: bodyStart, end: position)
                position += 3
                return XmlNode.comment(body, sourceRange: start..<position)
            }
            position += 1
        }
        throw XmlTreeReaderError.unexpectedEOF(at: position)
    }

    // MARK: - Processing instruction

    mutating func parseProcessingInstruction(asNode: Bool) throws -> XmlNode {
        let start = position
        guard consumeString("<?") else {
            throw XmlTreeReaderError.malformedTag(at: position)
        }
        let target = try parseName()
        skipWhitespace()
        let dataStart = position
        while position + 1 < bytes.count {
            if bytes[position] == 0x3F /* '?' */ &&
               bytes[position + 1] == 0x3E /* '>' */ {
                let data = stringFromBytes(start: dataStart, end: position)
                position += 2
                let node = XmlNode.processingInstruction(
                    target: target,
                    data: data,
                    sourceRange: start..<position
                )
                _ = asNode // consumed; XML decl is parsed identically but discarded by caller
                return node
            }
            position += 1
        }
        throw XmlTreeReaderError.unexpectedEOF(at: position)
    }

    // MARK: - DOCTYPE (skip; OOXML doesn't use it but be permissive)

    mutating func skipDoctype() throws {
        guard consumeString("<!DOCTYPE") else { return }
        var depth = 0
        while position < bytes.count {
            let b = bytes[position]
            if b == 0x5B /* '[' */ { depth += 1 }
            else if b == 0x5D /* ']' */ { depth -= 1 }
            else if b == 0x3E /* '>' */, depth <= 0 {
                position += 1
                return
            }
            position += 1
        }
        throw XmlTreeReaderError.unexpectedEOF(at: position)
    }

    // MARK: - CDATA

    mutating func parseCDATA() throws -> XmlNode {
        let start = position
        guard consumeString("<![CDATA[") else {
            throw XmlTreeReaderError.malformedTag(at: position)
        }
        let bodyStart = position
        while position + 2 < bytes.count {
            if bytes[position] == 0x5D /* ']' */ &&
               bytes[position + 1] == 0x5D /* ']' */ &&
               bytes[position + 2] == 0x3E /* '>' */ {
                let body = stringFromBytes(start: bodyStart, end: position)
                position += 3
                // CDATA content is character data that bypasses entity decoding.
                return XmlNode.text(body, sourceRange: start..<position)
            }
            position += 1
        }
        throw XmlTreeReaderError.unexpectedEOF(at: position)
    }

    // MARK: - Text content

    mutating func parseText() throws -> XmlNode {
        let start = position
        while position < bytes.count, bytes[position] != 0x3C /* '<' */ {
            position += 1
        }
        let raw = stringFromBytes(start: start, end: position)
        let decoded = try decodeXmlEntities(raw, position: start)
        return XmlNode.text(decoded, sourceRange: start..<position)
    }

    // MARK: - Name (Name production from XML 1.0)

    mutating func parseName() throws -> String {
        let start = position
        while position < bytes.count, isNameCharacter(bytes[position], isFirst: position == start) {
            position += 1
        }
        guard position > start else {
            throw XmlTreeReaderError.unexpectedCharacter(currentCharacter() ?? "?", at: position)
        }
        return stringFromBytes(start: start, end: position)
    }

    // MARK: - Low-level helpers

    func currentCharacter() -> Character? {
        guard position < bytes.count else { return nil }
        return Character(UnicodeScalar(bytes[position]))
    }

    func peekString(_ s: String) -> Bool {
        let s8 = Array(s.utf8)
        if position + s8.count > bytes.count { return false }
        for i in 0..<s8.count where bytes[position + i] != s8[i] { return false }
        return true
    }

    mutating func consumeByte(_ b: UInt8) -> Bool {
        guard position < bytes.count, bytes[position] == b else { return false }
        position += 1
        return true
    }

    mutating func consumeString(_ s: String) -> Bool {
        if peekString(s) {
            position += s.utf8.count
            return true
        }
        return false
    }

    mutating func skipWhitespace() {
        while position < bytes.count, isWhitespace(bytes[position]) {
            position += 1
        }
    }

    func stringFromBytes(start: Int, end: Int) -> String {
        guard start < end else { return "" }
        let slice = Array(bytes[start..<end])
        return String(decoding: slice, as: UTF8.self)
    }
}

// MARK: - Free functions

private func isWhitespace(_ b: UInt8) -> Bool {
    b == 0x20 || b == 0x09 || b == 0x0A || b == 0x0D
}

private func isNameStartCharacter(_ b: UInt8) -> Bool {
    // ASCII letters, underscore, colon — covers OOXML names. Word produces
    // ASCII tag names; Unicode names are theoretically allowed but unused.
    if b >= 0x41 && b <= 0x5A { return true } // A-Z
    if b >= 0x61 && b <= 0x7A { return true } // a-z
    if b == 0x5F { return true } // _
    if b == 0x3A { return true } // :
    return false
}

private func isNameCharacter(_ b: UInt8, isFirst: Bool) -> Bool {
    if isNameStartCharacter(b) { return true }
    if isFirst { return false }
    if b >= 0x30 && b <= 0x39 { return true } // 0-9
    if b == 0x2D { return true } // -
    if b == 0x2E { return true } // .
    return false
}

private func splitQualifiedName(_ qname: String) -> (prefix: String?, localName: String) {
    if let colonIndex = qname.firstIndex(of: ":") {
        let prefix = String(qname[..<colonIndex])
        let local = String(qname[qname.index(after: colonIndex)...])
        return (prefix, local)
    }
    return (nil, qname)
}

private func decodeXmlEntities(_ s: String, position: Int) throws -> String {
    if !s.contains("&") { return s }
    var result = ""
    result.reserveCapacity(s.count)
    var i = s.startIndex
    while i < s.endIndex {
        let c = s[i]
        if c != "&" {
            result.append(c)
            i = s.index(after: i)
            continue
        }
        guard let semi = s[i...].firstIndex(of: ";") else {
            throw XmlTreeReaderError.invalidEntityReference(
                reference: String(s[i...]), at: position
            )
        }
        let ref = String(s[s.index(after: i)..<semi])
        switch ref {
        case "amp": result.append("&")
        case "lt": result.append("<")
        case "gt": result.append(">")
        case "quot": result.append("\"")
        case "apos": result.append("'")
        default:
            if ref.hasPrefix("#x") || ref.hasPrefix("#X") {
                let hex = String(ref.dropFirst(2))
                guard let scalarValue = UInt32(hex, radix: 16),
                      let scalar = Unicode.Scalar(scalarValue) else {
                    throw XmlTreeReaderError.invalidNumericReference(
                        reference: ref, at: position
                    )
                }
                result.append(Character(scalar))
            } else if ref.hasPrefix("#") {
                let dec = String(ref.dropFirst())
                guard let scalarValue = UInt32(dec),
                      let scalar = Unicode.Scalar(scalarValue) else {
                    throw XmlTreeReaderError.invalidNumericReference(
                        reference: ref, at: position
                    )
                }
                result.append(Character(scalar))
            } else {
                throw XmlTreeReaderError.invalidEntityReference(
                    reference: ref, at: position
                )
            }
        }
        i = s.index(after: semi)
    }
    return result
}
