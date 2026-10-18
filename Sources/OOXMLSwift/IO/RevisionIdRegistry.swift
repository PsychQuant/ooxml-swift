import Foundation

/// PsychQuant/che-word-mcp#275 (verify round 3): a run carries a revision id
/// but no type, so two revisions of different types sharing one id cannot be
/// told apart once read. Every id-only lookup — the writer's wrapper lookup,
/// accept / reject inside a paragraph, the document-level list — resolves to
/// whichever comes first: a deletion sharing a paragraph mark's id was saved
/// as `<w:ins>` holding `<w:delText>`, and one accept or reject acted on both.
/// Word never writes such files; other producers and crafted files do.
///
/// While a package is read, each typed revision claims its id for its kind.
/// A later revision of another kind gets a fresh id; a revision of the same
/// kind shares it (che-word-mcp writes a paragraph mark and its runs under
/// one id, and a deletion split by a marker is several `<w:del>` with one id).
/// Fresh ids avoid every revision id in the file, typed or not.
final class RevisionIdRegistry {
    private struct Kind: Hashable {
        let type: RevisionType
        /// A wrapper kept as raw XML (around a hyperlink, field, …) is
        /// resolved by a different path than one whose runs are typed.
        let mixed: Bool
    }

    private var claimed: [Int: Kind] = [:]

    /// Every revision id in the file plus every id handed out.
    private(set) var used: Set<Int>

    init(fileIds: Set<Int>) {
        used = fileIds
    }

    /// The id the revision read as `id` keeps, or a fresh one when `id` is
    /// already taken by a revision of another kind.
    func claim(_ id: Int, _ type: RevisionType, mixed: Bool = false) -> Int {
        let kind = Kind(type: type, mixed: mixed)
        guard let owner = claimed[id] else {
            claimed[id] = kind
            used.insert(id)
            return id
        }
        if owner == kind { return id }
        let fresh = nextFreeId()
        claimed[fresh] = kind
        used.insert(fresh)
        return fresh
    }

    private func nextFreeId() -> Int {
        let maxId = used.max() ?? 0
        if maxId < Int.max / 2 { return maxId + 1 }
        var candidate = 1
        while used.contains(candidate) { candidate += 1 }
        return candidate
    }

    // MARK: - File scan

    /// `w:id` of every track-change element: insertions, deletions, moves
    /// and their ranges, every `*Change`, cell revisions.
    private static let revisionIdPattern = try! NSRegularExpression(
        pattern: #"<w:(?:ins|del|moveFrom|moveTo|moveFromRangeStart|moveToRangeStart|cellIns|cellDel|cellMerge|[A-Za-z]+Change)\b[^>]*?\bw:id=["'](\d+)["']"#)

    static func revisionIds(in xml: String) -> Set<Int> {
        let ns = xml as NSString
        var ids: Set<Int> = []
        for match in revisionIdPattern.matches(in: xml, range: NSRange(location: 0, length: ns.length)) {
            if let id = Int(ns.substring(with: match.range(at: 1))) { ids.insert(id) }
        }
        return ids
    }

    /// Revision ids in every XML part under `word/` of an unpacked package.
    static func revisionIds(inPackageAt root: URL) -> Set<Int> {
        let word = root.appendingPathComponent("word")
        guard let walker = FileManager.default.enumerator(at: word, includingPropertiesForKeys: nil) else { return [] }
        var ids: Set<Int> = []
        for case let url as URL in walker where url.pathExtension == "xml" {
            // Best effort: a part that cannot be read here fails the reader
            // proper when it is parsed, or is not one the model types.
            guard let data = try? Data(contentsOf: url) else { continue }
            ids.formUnion(revisionIds(in: String(decoding: data, as: UTF8.self)))
        }
        return ids
    }
}
