// Container pairing the root `XmlNode` of a parsed XML part with the
// original source bytes. Required by `XmlTreeWriter` to emit untouched
// sub-trees byte-equal: each clean `XmlNode` carries a `sourceRange` that
// indexes into `sourceBytes`.
//
// `XmlTree` is the unit of read/write exchange between the tree IO layer
// and the rest of the library. A docx contains many parts; each part is
// parsed into its own `XmlTree`.

import Foundation

/// A parsed XML part: root node + the original byte buffer the nodes
/// reference via `sourceRange`. Synthesized trees (built without reading
/// from a file) carry an empty `sourceBytes` buffer; their nodes must all
/// have `sourceRange == nil` and `isDirty == true`.
public struct XmlTree {

    /// Root node of the parsed XML. For an OOXML part this is the document
    /// element (e.g. `<w:document>`, `<w:settings>`, `<w:numbering>`).
    public var root: XmlNode

    /// Original part bytes. The writer reads `sourceBytes[node.sourceRange]`
    /// for clean nodes. Synthesized trees carry an empty `Data`.
    public let sourceBytes: Data

    public init(root: XmlNode, sourceBytes: Data) {
        self.root = root
        self.sourceBytes = sourceBytes
    }

    /// Constructs a synthesized tree for a node built in memory (no source).
    /// All nodes in the tree are by definition dirty.
    public static func synthesized(root: XmlNode) -> XmlTree {
        XmlTree(root: root, sourceBytes: Data())
    }
}
