import Foundation

// PsychQuant/ooxml-swift#133 (Decision (B), 2026-11-09): tree-level
// surgery for typed-dirty saves of `word/document.xml`.
//
// A typed mutator marks the whole part typed-dirty, and the writer used to
// re-serialize every body block from the typed model. The typed model loses
// shapes the reader never modelled (#133: math inside text-box runs,
// whitespace-only text, nested-table order, paragraph-level `sectPr`, …), so
// one edit damaged blocks nobody touched. Reading and writing the live tree
// is byte-exact, so the writer now keeps the source bytes of every block that
// is still the block it was read as, and re-serializes only the others.
//
// "Still the block it was read as" takes two checks:
// - identity: the reader stamps each block value with where it was read
//   (`BodySourceIdentity`). Copies of the value carry the stamp; values built
//   anew do not. Content alone cannot do this: typed-equal blocks with
//   different source bytes would lend each other their bytes, and a block
//   edited into the same text as a deleted one would take the deleted
//   block's bytes (verify R1).
// - no change: the block's typed serialization equals the one it had when
//   read. Comparing serializations rather than `==` keeps a field missing
//   from a hand-written `==` from turning an edit into a silent no-op.
// Whenever either check fails or is in doubt, the block is rewritten from the
// typed model, which is what the writer did for every block before.

/// Where a body block value was read: the reading it came from and its index
/// in the body as read. Copies of the value carry it. It is not part of `==`
/// (two values compare equal whatever their identities), so it never changes
/// what the typed model considers equal.
struct BodySourceIdentity: Equatable {
    var generation: UInt64?
    var index = 0

    static func == (lhs: BodySourceIdentity, rhs: BodySourceIdentity) -> Bool { true }
}

extension BodyChild {
    /// The identity `DocxReader` stamped on this block's value.
    var bodySourceIdentity: BodySourceIdentity {
        get {
            switch self {
            case .paragraph(let p): return p.bodySourceIdentity
            case .table(let t): return t.bodySourceIdentity
            case .contentControl(let c, _): return c.bodySourceIdentity
            case .bookmarkMarker(let m): return m.bodySourceIdentity
            case .rawBlockElement(let r): return r.bodySourceIdentity
            }
        }
        set {
            switch self {
            case .paragraph(var p): p.bodySourceIdentity = newValue; self = .paragraph(p)
            case .table(var t): t.bodySourceIdentity = newValue; self = .table(t)
            case .contentControl(var c, let children):
                c.bodySourceIdentity = newValue; self = .contentControl(c, children: children)
            case .bookmarkMarker(var m): m.bodySourceIdentity = newValue; self = .bookmarkMarker(m)
            case .rawBlockElement(var r): r.bodySourceIdentity = newValue; self = .rawBlockElement(r)
            }
        }
    }
}

/// The body of `word/document.xml` as `DocxReader` read it.
struct BodySourceBaseline {
    /// Identifies this reading; a block stamped by another reading (a value
    /// taken from another document) never counts as one of these blocks.
    let generation: UInt64
    /// Root and body of the live tree at read time (identity checked at write).
    let root: XmlNode
    let body: XmlNode
    /// `body.children` at read time, in order (identity checked at write).
    let bodyChildren: [XmlNode]
    /// For each block in `snapshot`, its index in `bodyChildren`.
    let blockNodeIndex: [Int]
    /// `document.body.children` at read time.
    let snapshot: [BodyChild]
    /// `document.sectionProperties` at read time.
    let sectionProperties: SectionProperties
    /// `document.documentRootAttributes` at read time. Surgery keeps the
    /// source root start tag, so a change to these attributes falls back to
    /// the full writer, which rebuilds the tag from them.
    let rootAttributes: [String: String]

    /// Stamp every body block with its identity and record the baseline.
    /// The document gets no baseline (so surgery never applies) when the
    /// typed body and the tree's body do not line up one to one.
    static func install(on document: inout WordDocument) {
        guard let (root, body, index) = pairing(document) else {
            document.bodySourceBaseline = nil
            return
        }
        let generation = UInt64.random(in: 1...UInt64.max)
        for k in document.body.children.indices {
            document.body.children[k].bodySourceIdentity = BodySourceIdentity(generation: generation, index: k)
        }
        document.bodySourceBaseline = BodySourceBaseline(
            generation: generation, root: root, body: body, bodyChildren: body.children,
            blockNodeIndex: index, snapshot: document.body.children,
            sectionProperties: document.sectionProperties,
            rootAttributes: document.documentRootAttributes)
    }

    /// The tree's root and body and, for each typed block, the index of its
    /// node among the body's children, paired by the rules of
    /// `DocxReader.parseBodyChildren`. Nil when they do not line up.
    private static func pairing(_ document: WordDocument) -> (XmlNode, XmlNode, [Int])? {
        guard let tree = document.xmlTrees["word/document.xml"],
              let bodyAt = tree.root.children.firstIndex(where: { $0.kind == .element && $0.localName == "body" })
        else { return nil }
        // The reader takes the first `body` element in document order.
        guard !tree.root.children[..<bodyAt].contains(where: containsBody) else { return nil }
        let body = tree.root.children[bodyAt]
        let children = body.children
        let elements = children.indices.filter { children[$0].kind == .element }
        // The reader takes the last body-level `sectPr`, and the splice puts
        // new trailing blocks before it: require at most one, at the end.
        let sections = elements.filter { children[$0].localName == "sectPr" }
        guard sections.count <= 1, sections.allSatisfy({ $0 == elements.last }) else { return nil }
        var index: [Int] = []
        for i in elements {
            let node = children[i]
            switch node.localName {
            case "sectPr":
                continue
            case "bookmarkStart", "bookmarkEnd":
                // The reader keeps a body-level bookmark marker only when its
                // `w:id` is an integer.
                guard let id = node.attributes.first(where: { $0.prefix == "w" && $0.localName == "id" })?.value,
                      Int(id) != nil else { continue }
                index.append(i)
            default:
                index.append(i)
            }
        }
        let blocks = document.body.children
        guard index.count == blocks.count else { return nil }
        for (k, i) in index.enumerated() where !kindMatches(children[i].localName, blocks[k]) {
            return nil
        }
        return (tree.root, body, index)
    }

    private static func containsBody(_ node: XmlNode) -> Bool {
        node.kind == .element && (node.localName == "body" || node.children.contains(where: containsBody))
    }

    private static func kindMatches(_ localName: String, _ block: BodyChild) -> Bool {
        switch (localName, block) {
        case ("p", .paragraph), ("tbl", .table), ("sdt", .contentControl),
             ("bookmarkStart", .bookmarkMarker), ("bookmarkEnd", .bookmarkMarker):
            return true
        case (_, .rawBlockElement):
            return !["p", "tbl", "sdt", "bookmarkStart", "bookmarkEnd", "sectPr"].contains(localName)
        default:
            return false
        }
    }
}

enum BodySurgery {

    /// `word/document.xml` with only the changed blocks rewritten, or nil when
    /// surgery does not apply (no baseline, the tree was touched since read,
    /// the root attributes changed, a block cannot be serialized, or a
    /// rewritten block uses a namespace prefix with no known URI). Nil means:
    /// use the full typed serializer.
    static func surgicalDocumentXML(_ document: WordDocument) throws -> Data? {
        guard let base = document.bodySourceBaseline,
              let tree = document.xmlTrees["word/document.xml"],
              tree.root === base.root,
              tree.root.sourceRange != nil,
              !tree.sourceBytes.isEmpty,
              document.documentRootAttributes == base.rootAttributes,
              base.body.children.count == base.bodyChildren.count,
              zip(base.body.children, base.bodyChildren).allSatisfy({ $0 === $1 }),
              !base.bodyChildren.isEmpty,
              base.bodyChildren.allSatisfy({ $0.sourceRange != nil }),
              !hasDirtyNode(tree.root)
        else { return nil }

        // Typed serializations of the current blocks and of the section
        // properties. A block that cannot be serialized leaves the decision
        // to the full writer (which reports it).
        let current = document.body.children
        let after: [String]
        let sectionBefore: String
        let sectionAfter: String
        do {
            after = try current.map { try DocxWriter.xmlForBodyChild($0) }
            sectionBefore = base.sectionProperties.toXML()
            sectionAfter = document.sectionProperties.toXML()
        } catch {
            return nil
        }
        guard let matched = keptBlocks(current, after, base) else { return nil }

        let rootBindings = Dictionary(
            tree.root.attributes.filter { $0.prefix == "xmlns" }.map { ($0.localName, $0.value) },
            uniquingKeysWith: { first, _ in first })
        func fragment(_ xml: String) -> Data? {
            repairNamespaces(xml, rootBindings: rootBindings).map { Data($0.utf8) }
        }

        let bytes = tree.sourceBytes
        let nodes = base.bodyChildren
        var blockAt: [Int: Int] = [:]
        for (k, i) in base.blockNodeIndex.enumerated() { blockAt[i] = k }
        // For each block as read, the current index of the first kept block
        // at or after it (`after.count` when none). The blocks that replace a
        // removed or rewritten block are written where it stood, so siblings
        // the typed model does not see (comments, unpaired bookmarks) keep
        // their place relative to them.
        var nextKept = [Int](repeating: after.count, count: base.snapshot.count + 1)
        for k in base.snapshot.indices.reversed() {
            nextKept[k] = matched[k] ?? nextKept[k + 1]
        }
        let lastMatchedNode = base.blockNodeIndex.indices.last(where: { matched[$0] != nil })
            .map { base.blockNodeIndex[$0] } ?? -1
        let sectionChanged = sectionBefore != sectionAfter
        var sectionWritten = false

        // span(i): node i's bytes plus whatever the source holds between it
        // and the node before, so untouched stretches reproduce the source.
        func span(_ i: Int) -> Data {
            let start = i == 0 ? nodes[0].sourceRange!.lowerBound : nodes[i - 1].sourceRange!.upperBound
            return bytes.subdata(in: start..<nodes[i].sourceRange!.upperBound)
        }

        var out = bytes.subdata(in: 0..<nodes[0].sourceRange!.lowerBound)
        var nextCurrent = 0
        func emitCurrent(upTo end: Int) -> Bool {
            while nextCurrent < end {
                guard let data = fragment(after[nextCurrent]) else { return false }
                out.append(data)
                nextCurrent += 1
            }
            return true
        }

        for i in nodes.indices {
            if let k = blockAt[i] {
                guard let c = matched[k] else {
                    // Removed or rewritten: what replaces it goes here.
                    guard emitCurrent(upTo: nextKept[k]) else { return nil }
                    continue
                }
                guard emitCurrent(upTo: c) else { return nil }
                out.append(span(i))
                nextCurrent = c + 1
                continue
            }
            let node = nodes[i]
            let isSection = node.kind == .element && node.localName == "sectPr"
            if isSection, i > lastMatchedNode, !sectionWritten {
                // Blocks added after the last kept block go before the
                // body-level section properties.
                guard emitCurrent(upTo: after.count) else { return nil }
            }
            if isSection, sectionChanged, !sectionWritten {
                let start = i == 0 ? node.sourceRange!.lowerBound : nodes[i - 1].sourceRange!.upperBound
                out.append(bytes.subdata(in: start..<node.sourceRange!.lowerBound))
                guard let data = fragment(sectionAfter) else { return nil }
                out.append(data)
                sectionWritten = true
                continue
            }
            if isSection { sectionWritten = true }
            out.append(span(i))
        }
        guard emitCurrent(upTo: after.count) else { return nil }
        if sectionChanged, !sectionWritten {
            guard let data = fragment(sectionAfter) else { return nil }
            out.append(data)
        }
        out.append(bytes.subdata(in: nodes[nodes.count - 1].sourceRange!.upperBound..<bytes.count))
        return out
    }

    /// For each block as read, the index of the current block that is still
    /// that block, or nil when it was removed or must be rewritten. A current
    /// block is still block k when it carries this reading's identity k and
    /// its typed serialization is unchanged. When two current blocks qualify
    /// for the same k (an unchanged copy next to the original), they cannot
    /// be told apart, so both are rewritten. Kept blocks must stay in source
    /// order: the longest run in order is kept and a moved block is rewritten.
    /// Nil when a block as read cannot be serialized.
    static func keptBlocks(_ current: [BodyChild], _ after: [String], _ base: BodySourceBaseline) -> [Int?]? {
        var before: [Int: String] = [:]
        var candidates: [Int: [Int]] = [:]
        for (c, block) in current.enumerated() {
            let id = block.bodySourceIdentity
            guard id.generation == base.generation, base.snapshot.indices.contains(id.index) else { continue }
            if before[id.index] == nil {
                guard let xml = try? DocxWriter.xmlForBodyChild(base.snapshot[id.index]) else { return nil }
                before[id.index] = xml
            }
            if after[c] == before[id.index] { candidates[id.index, default: []].append(c) }
        }
        let pairs = candidates.compactMap { k, cs in cs.count == 1 ? (k, cs[0]) : nil }
            .sorted { $0.1 < $1.1 }
        var matched = [Int?](repeating: nil, count: base.snapshot.count)
        for (k, c) in longestIncreasing(pairs) { matched[k] = c }
        return matched
    }

    /// The longest subsequence of `pairs` (read index, current index), sorted
    /// by current index, whose read indices also increase.
    static func longestIncreasing(_ pairs: [(Int, Int)]) -> [(Int, Int)] {
        var tails: [Int] = []
        var previous = [Int](repeating: -1, count: pairs.count)
        for (i, pair) in pairs.enumerated() {
            var lo = 0, hi = tails.count
            while lo < hi {
                let mid = (lo + hi) / 2
                if pairs[tails[mid]].0 < pair.0 { lo = mid + 1 } else { hi = mid }
            }
            if lo > 0 { previous[i] = tails[lo - 1] }
            if lo == tails.count { tails.append(i) } else { tails[lo] = i }
        }
        var out: [(Int, Int)] = []
        var i = tails.last ?? -1
        while i >= 0 {
            out.append(pairs[i])
            i = previous[i]
        }
        return out.reversed()
    }

    /// The typed serialization of one block, with an `xmlns:*` declaration
    /// added to its first start tag for every prefix it uses that neither the
    /// block itself nor the document root binds. A prefix the root binds keeps
    /// the root's URI (in a Strict document `w` stays Strict), as when the
    /// full typed writer wrote the block. Nil when a prefix has no known URI.
    static func repairNamespaces(_ xml: String, rootBindings: [String: String]) -> String? {
        var decls = WordDocument.standardWordNamespaces
        for (k, v) in rootBindings { decls[k] = v }
        decls["xml"] = nil
        let header = decls.keys.sorted().map { "xmlns:\($0)=\"\(decls[$0]!)\"" }.joined(separator: " ")
        guard let parsed = try? XmlTreeReader.parse(Data(("<w:graft \(header)>" + xml + "</w:graft>").utf8)) else {
            return nil
        }
        let elements = parsed.root.children.filter { $0.kind == .element }
        var unbound = Set<String>()
        for element in elements { unbound.formUnion(unboundPrefixes(element, scope: [])) }
        var needed: [(String, String)] = []
        for prefix in unbound.sorted() where rootBindings[prefix] == nil {
            guard let uri = WordDocument.standardWordNamespaces[prefix] else { return nil }
            needed.append((prefix, uri))
        }
        guard !needed.isEmpty else { return xml }
        // One block serializes to one element; anything else cannot be
        // repaired by editing a single start tag.
        guard elements.count == 1, xml.first == "<" else { return nil }
        guard let nameEnd = xml.dropFirst().firstIndex(where: { $0 == " " || $0 == "/" || $0 == ">" || $0 == "\n" || $0 == "\t" || $0 == "\r" })
        else { return nil }
        let addition = needed.map { " xmlns:\($0.0)=\"\($0.1)\"" }.joined()
        return String(xml[..<nameEnd]) + addition + String(xml[nameEnd...])
    }

    /// Prefixes an element subtree uses on element and attribute names that no
    /// `xmlns:*` declaration within the subtree binds where they are used
    /// (`xmlns` itself and the reserved `xml` prefix excluded).
    static func unboundPrefixes(_ node: XmlNode, scope: Set<String>) -> Set<String> {
        guard node.kind == .element else { return [] }
        var scope = scope
        for a in node.attributes where a.prefix == "xmlns" { scope.insert(a.localName) }
        var out = Set<String>()
        if let p = node.prefix, !p.isEmpty, p != "xml", !scope.contains(p) { out.insert(p) }
        for a in node.attributes {
            if let p = a.prefix, !p.isEmpty, p != "xmlns", p != "xml", !scope.contains(p) { out.insert(p) }
        }
        for child in node.children { out.formUnion(unboundPrefixes(child, scope: scope)) }
        return out
    }

    private static func hasDirtyNode(_ node: XmlNode) -> Bool {
        if node.isDirty { return true }
        return node.children.contains(where: hasDirtyNode)
    }
}
