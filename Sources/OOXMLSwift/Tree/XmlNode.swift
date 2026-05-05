// Generic, lossless XML DOM node used by the tree IO path of ooxml-swift.
//
// `XmlNode` is the foundation type for the word-aligned-state-sync architecture
// (Spectra change `word-aligned-state-sync`, Phase 0). It represents every
// well-formed OOXML element, attribute, namespace declaration, comment,
// processing instruction, and text node read from a docx, with no element
// class or attribute key dropped during reading.
//
// Round-trip preservation strategy:
//   - On read, each node records the byte range `sourceRange` it occupied in
//     the original part XML.
//   - On write, clean nodes are emitted by copying `sourceBytes[sourceRange]`
//     verbatim; dirty nodes are re-serialized from their typed fields.
//   - Mutation through the public API marks the touched node and all ancestors
//     dirty so the writer recomputes their bytes.
//
// This file implements the data type only (Phase 0 task 1.1). Reader, writer,
// and fingerprint live in sibling files in this directory.

import Foundation

/// A single node in the lossless OOXML DOM tree.
///
/// `XmlNode` is reference-typed (a class) because:
/// - sub-trees are shared as references when copying / repositioning,
/// - parent-traversal and identity comparisons are common,
/// - mutation marks identity-stable nodes dirty without re-allocating arrays
///   of value types up the spine.
public final class XmlNode {

    /// Kind discriminator for what the node represents.
    public enum Kind: Equatable, Sendable {
        /// An XML element with name + attributes + ordered children.
        case element
        /// A text node (character data between elements).
        case text
        /// An XML comment (`<!-- ... -->`).
        case comment
        /// An XML processing instruction (`<?target data?>`).
        case processingInstruction
    }

    // MARK: - Discriminator and shape

    public let kind: Kind

    // MARK: - Element-only fields (valid when kind == .element)

    /// Namespace prefix as it appeared in the source XML, e.g. `"w"` in `<w:p>`.
    /// `nil` for unprefixed elements. Distinct from `namespaceURI` so the writer
    /// can re-emit the same prefix decision the source used (see scenario
    /// "Namespace prefix decisions preserved" in spec `ooxml-tree-io`).
    public var prefix: String? {
        didSet { markDirty() }
    }

    /// Local name without namespace prefix, e.g. `"p"` for `<w:p>`.
    public var localName: String {
        didSet { markDirty() }
    }

    /// Resolved namespace URI looked up at parse time from the prefix
    /// declaration in scope. Stored on the node so consumers can compare
    /// nodes across parts that use different prefixes for the same namespace.
    /// This is *derived* metadata: changing it does NOT mark the node dirty,
    /// because the prefix (which IS bytes-affecting) is the source of truth
    /// for serialization. The reader populates this field after parsing so
    /// post-init namespace resolution does not falsely dirty clean nodes.
    public var namespaceURI: String?

    /// Attributes in source order, including namespace declarations
    /// (`xmlns` and `xmlns:*`). Source order preservation matters for
    /// byte-equal round-trip on untouched sub-trees.
    public var attributes: [XmlAttribute] {
        didSet { markDirty() }
    }

    /// Children in source order. Mixed content (text interleaved with elements)
    /// is preserved by including `XmlNode` instances of `.text` kind in this
    /// array at their original positions.
    public var children: [XmlNode] {
        didSet { markDirty() }
    }

    // MARK: - Text / comment / PI fields

    /// Character data for `.text` nodes, comment body for `.comment` nodes,
    /// PI data for `.processingInstruction` nodes. Empty for `.element`.
    public var textContent: String {
        didSet { markDirty() }
    }

    /// PI target for `.processingInstruction` nodes (the part before the
    /// space, e.g. `xml-stylesheet`). Empty for other kinds.
    public var processingInstructionTarget: String {
        didSet { markDirty() }
    }

    // MARK: - Round-trip preservation

    /// Byte range in the source part XML that this node occupied. `nil` when
    /// the node was synthesized in memory (created via factory methods rather
    /// than parsed) or when mutation has invalidated the range.
    public internal(set) var sourceRange: Range<Int>?

    /// True when this node or one of its mutated descendants requires
    /// re-serialization. Writer copies clean sub-trees from `sourceBytes`
    /// verbatim and re-emits dirty sub-trees from typed fields.
    public private(set) var isDirty: Bool

    // MARK: - Init (private — use factory methods)

    private init(
        kind: Kind,
        prefix: String?,
        localName: String,
        namespaceURI: String?,
        attributes: [XmlAttribute],
        children: [XmlNode],
        textContent: String,
        processingInstructionTarget: String,
        sourceRange: Range<Int>?
    ) {
        self.kind = kind
        self.prefix = prefix
        self.localName = localName
        self.namespaceURI = namespaceURI
        self.attributes = attributes
        self.children = children
        self.textContent = textContent
        self.processingInstructionTarget = processingInstructionTarget
        self.sourceRange = sourceRange
        // Synthesized nodes (no source range) are dirty by definition.
        self.isDirty = sourceRange == nil
    }

    // MARK: - Public factories

    /// Construct an `.element` node. Used both by the parser (with
    /// `sourceRange` set) and by callers building new content (with
    /// `sourceRange == nil`).
    public static func element(
        prefix: String? = nil,
        localName: String,
        namespaceURI: String? = nil,
        attributes: [XmlAttribute] = [],
        children: [XmlNode] = [],
        sourceRange: Range<Int>? = nil
    ) -> XmlNode {
        XmlNode(
            kind: .element,
            prefix: prefix,
            localName: localName,
            namespaceURI: namespaceURI,
            attributes: attributes,
            children: children,
            textContent: "",
            processingInstructionTarget: "",
            sourceRange: sourceRange
        )
    }

    /// Construct a `.text` node. Character data appears in the parent's
    /// `children` array at the position it occupied in the source XML.
    public static func text(
        _ value: String,
        sourceRange: Range<Int>? = nil
    ) -> XmlNode {
        XmlNode(
            kind: .text,
            prefix: nil,
            localName: "",
            namespaceURI: nil,
            attributes: [],
            children: [],
            textContent: value,
            processingInstructionTarget: "",
            sourceRange: sourceRange
        )
    }

    /// Construct a `.comment` node.
    public static func comment(
        _ value: String,
        sourceRange: Range<Int>? = nil
    ) -> XmlNode {
        XmlNode(
            kind: .comment,
            prefix: nil,
            localName: "",
            namespaceURI: nil,
            attributes: [],
            children: [],
            textContent: value,
            processingInstructionTarget: "",
            sourceRange: sourceRange
        )
    }

    /// Construct a `.processingInstruction` node.
    public static func processingInstruction(
        target: String,
        data: String,
        sourceRange: Range<Int>? = nil
    ) -> XmlNode {
        XmlNode(
            kind: .processingInstruction,
            prefix: nil,
            localName: "",
            namespaceURI: nil,
            attributes: [],
            children: [],
            textContent: data,
            processingInstructionTarget: target,
            sourceRange: sourceRange
        )
    }

    // MARK: - Stable identity

    /// Stable element ID derived from the OOXML attributes that act as
    /// document-level identifiers. Priority follows the `ooxml-operation-log`
    /// `ElementID derivation rules` requirement:
    ///   1. `w14:paraId` (paragraphs)
    ///   2. `w:bookmarkId` and `w:id` on bookmarks / comments
    ///   3. `r:id` on relationship-bearing elements
    ///   4. `w14:textId` as a paragraph tiebreaker
    /// Returns `nil` for elements with none of these; the caller may attach
    /// a library-generated UUID via `setLibraryUUID(_:)` for in-memory matching.
    public var stableID: String? {
        guard kind == .element else { return nil }
        // Order matters: paraId first, then comment/bookmark id, then r:id.
        if let v = attributeValue(prefix: "w14", localName: "paraId") { return "w14:paraId=\(v)" }
        if let v = attributeValue(prefix: "w", localName: "bookmarkId") { return "w:bookmarkId=\(v)" }
        if let v = attributeValue(prefix: "w", localName: "id") { return "w:id=\(v)" }
        if let v = attributeValue(prefix: "r", localName: "id") { return "r:id=\(v)" }
        if let v = attributeValue(prefix: "w14", localName: "textId") { return "w14:textId=\(v)" }
        return nil
    }

    /// Library-generated UUID assigned to elements lacking an OOXML stable ID.
    /// Lives only in memory; never written to the docx. Used by the operation
    /// log to address elements when no native ID exists.
    public var libraryUUID: UUID?

    // MARK: - Mutation

    /// Marks this node dirty and clears its source range so the writer
    /// re-emits it from typed fields on next serialization. Idempotent.
    /// Note: this does NOT propagate dirtiness to ancestors; the writer
    /// walks the tree and recurses anyway. Ancestor invalidation happens
    /// implicitly because the writer always inspects every dirty descendant.
    public func markDirty() {
        guard !isDirty else { return }
        isDirty = true
        sourceRange = nil
    }

    /// Looks up an attribute value by namespace prefix + local name.
    public func attributeValue(prefix: String?, localName: String) -> String? {
        attributes.first { $0.prefix == prefix && $0.localName == localName }?.value
    }

    /// Sets or replaces an attribute by namespace prefix + local name.
    /// Marks the node dirty.
    public func setAttribute(prefix: String?, localName: String, value: String) {
        if let index = attributes.firstIndex(where: { $0.prefix == prefix && $0.localName == localName }) {
            attributes[index].value = value
        } else {
            attributes.append(XmlAttribute(prefix: prefix, localName: localName, value: value))
        }
        // didSet on attributes already calls markDirty(); explicit call is a no-op
        // when already dirty but documents intent.
        markDirty()
    }
}

// MARK: - Equatable / Hashable by identity

extension XmlNode: Equatable {
    /// Identity equality. Two `XmlNode` references are equal iff they are the
    /// same instance. Structural equality is provided separately via
    /// `XmlNode.normalizedFingerprint()` (see `XmlTreeFingerprint.swift`).
    public static func == (lhs: XmlNode, rhs: XmlNode) -> Bool {
        lhs === rhs
    }
}

extension XmlNode: Hashable {
    public func hash(into hasher: inout Hasher) {
        hasher.combine(ObjectIdentifier(self))
    }
}
