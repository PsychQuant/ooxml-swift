// Lossless XML writer for `XmlTree`.
//
// The writer satisfies the byte-equal round-trip contract for untouched
// sub-trees by copying `sourceBytes[node.sourceRange]` verbatim when a node
// is clean. Dirty nodes are re-serialized from their typed fields. Mixed
// dirty/clean trees produce output that differs from input only inside
// dirty sub-trees.
//
// Pure-Swift, no external dependencies.

import Foundation

public enum XmlTreeWriterError: Error, Equatable {
    /// A clean node referenced a `sourceRange` that falls outside the
    /// `sourceBytes` buffer. Indicates corruption between read and write.
    case sourceRangeOutOfBounds(start: Int, end: Int, bufferLength: Int)
    /// Synthesized node carried `sourceRange != nil`. Should never happen
    /// when callers use the public factory methods.
    case synthesizedNodeHasSourceRange
}

/// Serializes an `XmlTree` to UTF-8 bytes. The output matches the
/// original `tree.sourceBytes` byte-for-byte when no mutations occurred.
public enum XmlTreeWriter {

    /// Serialize the tree to a `Data` buffer.
    public static func serialize(_ tree: XmlTree) throws -> Data {
        // Pre-compute "subtree contains any dirty descendant" once for every
        // node so the emit pass can decide blob-copy vs re-emit in O(1).
        // Without this, a mutation deep in a subtree leaves all ancestors
        // structurally clean (their own fields didn't change), so the naive
        // ancestor-clean → blob-copy path would emit pre-mutation bytes for
        // the whole subtree, silently losing the mutation.
        var dirtyMap: [ObjectIdentifier: Bool] = [:]
        _ = computeSubtreeDirty(tree.root, into: &dirtyMap)

        var output = Data()
        if let rootRange = tree.root.sourceRange, !tree.sourceBytes.isEmpty {
            if rootRange.lowerBound > 0 {
                let prolog = tree.sourceBytes.subdata(in: 0..<rootRange.lowerBound)
                output.append(prolog)
            }
            try emit(tree.root, sourceBytes: tree.sourceBytes, dirtyMap: dirtyMap, into: &output)
            if rootRange.upperBound < tree.sourceBytes.count {
                let epilog = tree.sourceBytes.subdata(
                    in: rootRange.upperBound..<tree.sourceBytes.count
                )
                output.append(epilog)
            }
        } else {
            // Synthesized tree: emit a minimal declaration then the root.
            output.append(contentsOf: Array(#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#.utf8))
            output.append(0x0A) // newline
            try emit(tree.root, sourceBytes: tree.sourceBytes, dirtyMap: dirtyMap, into: &output)
        }
        return output
    }

    // MARK: - Subtree-dirty pre-pass

    @discardableResult
    private static func computeSubtreeDirty(
        _ node: XmlNode,
        into map: inout [ObjectIdentifier: Bool]
    ) -> Bool {
        var hasDirty = node.isDirty
        // Always recurse on every child so each node ends up in the map.
        for child in node.children {
            let childDirty = computeSubtreeDirty(child, into: &map)
            if childDirty { hasDirty = true }
        }
        map[ObjectIdentifier(node)] = hasDirty
        return hasDirty
    }

    // MARK: - Internal recursion

    private static func emit(
        _ node: XmlNode,
        sourceBytes: Data,
        dirtyMap: [ObjectIdentifier: Bool],
        into output: inout Data
    ) throws {
        // Fast path: clean node + clean subtree + valid source range = blob copy.
        let subtreeDirty = dirtyMap[ObjectIdentifier(node)] ?? node.isDirty
        if !subtreeDirty, let range = node.sourceRange {
            guard range.lowerBound >= 0 && range.upperBound <= sourceBytes.count else {
                throw XmlTreeWriterError.sourceRangeOutOfBounds(
                    start: range.lowerBound,
                    end: range.upperBound,
                    bufferLength: sourceBytes.count
                )
            }
            output.append(sourceBytes.subdata(in: range))
            return
        }
        // Re-serialize from typed fields. Children are recursed individually
        // so clean grand-descendants of a dirty parent can still blob-copy.
        switch node.kind {
        case .element:
            try emitElement(node, sourceBytes: sourceBytes, dirtyMap: dirtyMap, into: &output)
        case .text:
            output.append(contentsOf: Array(escapeText(node.textContent).utf8))
        case .comment:
            output.append(contentsOf: Array("<!--".utf8))
            output.append(contentsOf: Array(node.textContent.utf8))
            output.append(contentsOf: Array("-->".utf8))
        case .processingInstruction:
            output.append(contentsOf: Array("<?".utf8))
            output.append(contentsOf: Array(node.processingInstructionTarget.utf8))
            if !node.textContent.isEmpty {
                output.append(0x20) // space
                output.append(contentsOf: Array(node.textContent.utf8))
            }
            output.append(contentsOf: Array("?>".utf8))
        }
    }

    private static func emitElement(
        _ node: XmlNode,
        sourceBytes: Data,
        dirtyMap: [ObjectIdentifier: Bool],
        into output: inout Data
    ) throws {
        // Open tag: <prefix:localName attr1="..." attr2="..."
        output.append(0x3C) // '<'
        if let prefix = node.prefix, !prefix.isEmpty {
            output.append(contentsOf: Array(prefix.utf8))
            output.append(0x3A) // ':'
        }
        output.append(contentsOf: Array(node.localName.utf8))
        for attribute in node.attributes {
            output.append(0x20) // space
            if let prefix = attribute.prefix, !prefix.isEmpty {
                output.append(contentsOf: Array(prefix.utf8))
                output.append(0x3A) // ':'
            }
            output.append(contentsOf: Array(attribute.localName.utf8))
            output.append(0x3D) // '='
            output.append(0x22) // '"'
            output.append(contentsOf: Array(escapeAttributeValue(attribute.value).utf8))
            output.append(0x22) // '"'
        }
        if node.children.isEmpty {
            // Self-closing form: <tag .../>
            output.append(0x2F) // '/'
            output.append(0x3E) // '>'
            return
        }
        output.append(0x3E) // '>'
        for child in node.children {
            try emit(child, sourceBytes: sourceBytes, dirtyMap: dirtyMap, into: &output)
        }
        // Closing tag: </prefix:localName>
        output.append(0x3C) // '<'
        output.append(0x2F) // '/'
        if let prefix = node.prefix, !prefix.isEmpty {
            output.append(contentsOf: Array(prefix.utf8))
            output.append(0x3A) // ':'
        }
        output.append(contentsOf: Array(node.localName.utf8))
        output.append(0x3E) // '>'
    }

    // MARK: - Entity escaping

    /// Escape XML reserved characters in element character data. Quote and
    /// apostrophe do not need escaping inside text nodes.
    private static func escapeText(_ s: String) -> String {
        var result = ""
        result.reserveCapacity(s.count)
        for c in s {
            switch c {
            case "&": result.append("&amp;")
            case "<": result.append("&lt;")
            case ">": result.append("&gt;")
            default: result.append(c)
            }
        }
        return result
    }

    /// Escape XML reserved characters in attribute values. Quote escapes
    /// because the writer always uses double-quoted attribute syntax.
    private static func escapeAttributeValue(_ s: String) -> String {
        var result = ""
        result.reserveCapacity(s.count)
        for c in s {
            switch c {
            case "&": result.append("&amp;")
            case "<": result.append("&lt;")
            case ">": result.append("&gt;")
            case "\"": result.append("&quot;")
            default: result.append(c)
            }
        }
        return result
    }
}
