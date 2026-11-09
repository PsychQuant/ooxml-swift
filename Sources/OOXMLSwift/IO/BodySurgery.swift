import Foundation

// PsychQuant/ooxml-swift#133 (Decision (B), 2026-11-09): tree-level
// surgery for typed-dirty saves of `word/document.xml`.
//
// A typed mutator marks the whole part typed-dirty, and the writer used to
// re-serialize every body block from the typed model. The typed model loses
// shapes the reader never modelled (#133: math inside text-box runs,
// whitespace-only text, nested-table order, paragraph-level `sectPr`, …), so
// one edit damaged blocks nobody touched. Reading and writing the live tree
// is byte-exact, so the writer now keeps the tree for every block whose typed
// serialization is the same as when the document was read, and re-serializes
// only the blocks that changed.
//
// "Unchanged" is decided by comparing typed serializations (as read vs now),
// not `==`: a wrong "changed" only falls back to what the writer did before,
// while `==` has hand-maintained field lists where a missing field would turn
// an edit into a silent no-op.

/// The body of `word/document.xml` as `DocxReader` read it.
struct BodySourceBaseline {
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

    /// Pair the typed body with the tree's body element children, using the
    /// same rules as `DocxReader.parseBodyChildren`. Returns nil when the two
    /// do not line up one to one (then surgery never applies).
    static func record(from document: WordDocument) -> BodySourceBaseline? {
        guard let tree = document.xmlTrees["word/document.xml"],
              let body = tree.root.children.first(where: { $0.kind == .element && $0.localName == "body" })
        else { return nil }
        let children = body.children
        var index: [Int] = []
        for (i, node) in children.enumerated() where node.kind == .element {
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
        return BodySourceBaseline(root: tree.root, body: body, bodyChildren: children,
                                  blockNodeIndex: index, snapshot: blocks,
                                  sectionProperties: document.sectionProperties)
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

    /// Above this many blocks in the changed middle of the body, alignment
    /// gives up and re-serializes the middle (what the writer did before):
    /// `difference(from:)` is Myers, O((N+M)·D).
    static let alignmentLimit = 4000

    /// `word/document.xml` with only the changed blocks rewritten, or nil when
    /// surgery does not apply (no baseline, the tree was touched since read,
    /// a block cannot be serialized, or a rewritten block uses a namespace
    /// prefix with no known URI). Nil means: use the full typed serializer.
    static func surgicalDocumentXML(_ document: WordDocument) throws -> Data? {
        guard let base = document.bodySourceBaseline,
              let tree = document.xmlTrees["word/document.xml"],
              tree.root === base.root,
              tree.root.sourceRange != nil,
              !tree.sourceBytes.isEmpty,
              base.body.children.count == base.bodyChildren.count,
              zip(base.body.children, base.bodyChildren).allSatisfy({ $0 === $1 }),
              !base.bodyChildren.isEmpty,
              base.bodyChildren.allSatisfy({ $0.sourceRange != nil }),
              !hasDirtyNode(tree.root)
        else { return nil }

        // Typed serializations, as read and now. A block that cannot be
        // serialized leaves the decision to the full writer (which reports it).
        let before: [String]
        let after: [String]
        let sectionBefore: String
        let sectionAfter: String
        do {
            before = try base.snapshot.map { try DocxWriter.xmlForBodyChild($0) }
            after = try document.body.children.map { try DocxWriter.xmlForBodyChild($0) }
            sectionBefore = base.sectionProperties.toXML()
            sectionAfter = document.sectionProperties.toXML()
        } catch {
            return nil
        }
        let bytes = tree.sourceBytes
        let matched = align(unambiguousKeys(before, after, nodeBytes: base.blockNodeIndex.map {
            bytes.subdata(in: base.bodyChildren[$0].sourceRange!)
        }), after)

        let rootBindings = Dictionary(
            tree.root.attributes.filter { $0.prefix == "xmlns" }.map { ($0.localName, $0.value) },
            uniquingKeysWith: { first, _ in first })
        func fragment(_ xml: String) -> Data? {
            repairNamespaces(xml, rootBindings: rootBindings).map { Data($0.utf8) }
        }

        let nodes = base.bodyChildren
        var blockAt: [Int: Int] = [:]
        for (k, i) in base.blockNodeIndex.enumerated() { blockAt[i] = k }
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
                guard let c = matched[k] else { continue } // block removed or rewritten
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

    /// `before` with every serialization that cannot be attributed to one
    /// source block replaced by a key nothing in `after` equals, so those
    /// blocks are rewritten from the typed model instead of borrowing bytes.
    /// Content equality cannot tell typed-equal blocks apart: if the edit
    /// deleted one of two such blocks whose source bytes differ, matching the
    /// survivor to the wrong one would keep the deleted block's unmodelled
    /// content. Ambiguous means: the serialization occurs more than once as
    /// read with differing source bytes and its count changed (one of them
    /// was deleted or rewritten), or once as read but more than once now.
    /// When the count of a repeated serialization is unchanged, order decides
    /// the pairing: a new paragraph is stamped with a fresh paraId, so it
    /// cannot take the place of an existing one. A body that did not change
    /// at all keeps the identity mapping.
    static func unambiguousKeys(_ before: [String], _ after: [String], nodeBytes: [Data]) -> [String] {
        guard before != after else { return before }
        var bytesByKey: [String: Set<Data>] = [:]
        var countBefore: [String: Int] = [:]
        for (k, key) in before.enumerated() {
            bytesByKey[key, default: []].insert(nodeBytes[k])
            countBefore[key, default: 0] += 1
        }
        var countAfter: [String: Int] = [:]
        for key in after { countAfter[key, default: 0] += 1 }
        return before.enumerated().map { k, key in
            let ambiguous = ((bytesByKey[key]?.count ?? 0) > 1 && countBefore[key] != countAfter[key])
                || (countBefore[key] == 1 && (countAfter[key] ?? 0) > 1)
            // U+0000 cannot occur in serialized XML, so this never matches.
            return ambiguous ? "\u{0}ambiguous-\(k)" : key
        }
    }

    /// For each block as read, the index of the current block it equals
    /// (same typed serialization), keeping order. Unmatched entries were
    /// removed or rewritten.
    static func align(_ before: [String], _ after: [String]) -> [Int?] {
        var matched = [Int?](repeating: nil, count: before.count)
        let n = before.count, m = after.count
        var head = 0
        while head < n, head < m, before[head] == after[head] {
            matched[head] = head
            head += 1
        }
        var tail = 0
        while tail < n - head, tail < m - head, before[n - 1 - tail] == after[m - 1 - tail] {
            matched[n - 1 - tail] = m - 1 - tail
            tail += 1
        }
        let oldMiddle = Array(before[head..<(n - tail)])
        let newMiddle = Array(after[head..<(m - tail)])
        guard !oldMiddle.isEmpty, !newMiddle.isEmpty else { return matched }
        if oldMiddle.count == newMiddle.count {
            for d in oldMiddle.indices where oldMiddle[d] == newMiddle[d] {
                matched[head + d] = head + d
            }
            return matched
        }
        guard oldMiddle.count + newMiddle.count <= alignmentLimit else { return matched }
        let diff = newMiddle.difference(from: oldMiddle)
        var removed = Set<Int>(), inserted = Set<Int>()
        for change in diff {
            switch change {
            case .remove(let offset, _, _): removed.insert(offset)
            case .insert(let offset, _, _): inserted.insert(offset)
            }
        }
        let keptOld = oldMiddle.indices.filter { !removed.contains($0) }
        let keptNew = newMiddle.indices.filter { !inserted.contains($0) }
        for (o, nw) in zip(keptOld, keptNew) { matched[head + o] = head + nw }
        return matched
    }

    /// The typed serialization of one block, with an `xmlns:*` declaration
    /// added to its first start tag for every prefix it uses that the
    /// document root does not bind to the expected URI (same rule as the
    /// #132 graft). Nil when a prefix has no known URI.
    static func repairNamespaces(_ xml: String, rootBindings: [String: String]) -> String? {
        var decls = rootBindings
        for (k, v) in WordDocument.standardWordNamespaces { decls[k] = v }
        let header = decls.keys.sorted().map { "xmlns:\($0)=\"\(decls[$0]!)\"" }.joined(separator: " ")
        guard let parsed = try? XmlTreeReader.parse(Data(("<w:graft \(header)>" + xml + "</w:graft>").utf8)) else {
            return nil
        }
        let elements = parsed.root.children.filter { $0.kind == .element }
        var needed: [(String, String)] = []
        var used = Set<String>()
        for element in elements { used.formUnion(WordDocument.prefixesUsed(in: element)) }
        for prefix in used.sorted() {
            guard let wanted = WordDocument.standardWordNamespaces[prefix] ?? rootBindings[prefix] else { return nil }
            if rootBindings[prefix] != wanted { needed.append((prefix, wanted)) }
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

    private static func hasDirtyNode(_ node: XmlNode) -> Bool {
        if node.isDirty { return true }
        return node.children.contains(where: hasDirtyNode)
    }
}
