// FormDeletion.swift — `--remove`, `--remove-table` and `--truncate-from` of
// `macdoc word form apply` (macdoc#250, macdoc#253): the deletion half of
// FormOperations, split out of FormOperations.swift without changes (#238).

import Foundation

extension FormOperations {
    // MARK: Remove and truncate

    /// Checks everything the command deletes as one set (macdoc#253 verify R1, R2):
    /// fields and comment references against the union of the `--remove`
    /// paragraphs and the `--truncate-from` range, and each cell or text box
    /// against what is left after all of its named paragraphs go.
    static func checkDeletions(_ ids: [String], tables: [String] = [], truncateFrom: String?,
                               in index: XMLTagIndex) throws -> [FormRanges.DeletionSpan] {
        try deletionPlan(ids, tables: tables, truncateFrom: truncateFrom, in: index).spans
    }

    /// What `checkDeletions` judges, plus whether the body needs an empty
    /// final paragraph afterwards (rule (b) of `checkBodyStructure`).
    static func deletionPlan(_ ids: [String], tables: [String], truncateFrom: String?,
                             in index: XMLTagIndex) throws -> (spans: [FormRanges.DeletionSpan], addFinalParagraph: Bool) {
        // Nothing to delete: no pairing or container question arises, and the
        // whole-part scans below are skipped (macdoc#253 verify R3).
        guard !ids.isEmpty || !tables.isEmpty || truncateFrom != nil else { return ([], false) }
        var distinct: Set<String> = []
        let ids = ids.filter { distinct.insert($0).inserted }
        for id in ids { try checkRemove(id, in: index) }
        let removed = try ids.map { try locate($0, in: index) }
        var spans = zip(ids, removed).map { FormRanges.DeletionSpan(id: $0, range: index.elementRange($1)) }
        // A part with more than one body is not a document the structure
        // rules below can judge (macdoc#250 verify R1).
        if index.elements(named: "w:body").count > 1 {
            throw FormError.malformedPart("有多個 w:body，不是一般的 Word 文件，拒絕刪除")
        }
        // Each named table once, under the first id naming it (macdoc#250).
        let removedTables = try uniqueTables(tables, in: index)
        // A paragraph inside the table carrying its own section break would go
        // with it, as `--remove` refuses for a paragraph (macdoc#250 verify R1).
        // Only the table's own tags are walked, so the cost is the size of the
        // removed tables, not tables times paragraphs (macdoc#250 verify R2).
        for entry in removedTables {
            for t in entry.table...index.tags[entry.table].match! where index.tags[t].kind != .close
                && index.tags[t].name == "w:p" && ownSectPr(of: t, in: index) != nil {
                throw FormError.tableCarriesSectionBreak(entry.id)
            }
        }
        let tableSpans = removedTables.map { FormRanges.DeletionSpan(id: $0.id, range: index.elementRange($0.table)) }
        for (id, p) in zip(ids, removed) {
            if let table = tableSpans.first(where: { $0.range.contains(index.tags[p].start) }) {
                throw FormError.conflictingOperations(id, detail: "位在 --remove-table \(table.id) 會刪掉的表格裡")
            }
        }
        spans += tableSpans
        // The cut is located and checked once; every later use reads it here.
        let cut = try truncateFrom.map { try checkTruncate($0, in: index) }
        if let id = truncateFrom, let cut {
            let cutStart = index.tags[cut.paragraph].start
            if let table = tableSpans.first(where: { $0.range.lowerBound >= cutStart }) {
                throw FormError.conflictingOperations(table.id, detail: "位在 --truncate-from \(id) 會截掉的範圍內")
            }
            spans.append(.init(id: id, range: cutStart..<bodyEnd(of: cut.body, in: index)))
        }
        try FormRanges.checkPairedMarkup(spans, index: index)
        let removedSet = Set(removed)
        var seen: Set<Int> = []
        for p in removed.sorted() {
            guard let container = blockContainer(of: p, in: index),
                  ["w:tc", "w:txbxContent"].contains(index.tags[container].name), seen.insert(container).inserted else { continue }
            // checkRemove has refused every paragraph that breaks the
            // container alone, so what is left is a combination. Name the
            // paragraphs that cause it, in document order (tag order).
            let named = removed.filter { blockContainer(of: $0, in: index) == container }.sorted()
            let remaining = flattenBlocks(index.children(of: container), in: index).filter { !removedSet.contains($0) }
            let ids = { (ps: [Int]) in ps.map { index.attribute("w14:paraId", of: $0) ?? "" } }
            if !remaining.contains(where: { index.tags[$0].name == "w:p" }) {
                throw FormError.removalsLeaveNoParagraph(ids(named))
            }
            if let last = remaining.last, index.tags[last].name != "w:p" {
                throw FormError.removalsEndWithoutParagraph(ids(named.filter { $0 > last }))
            }
        }
        let deletedBlocks = removedSet.union(removedTables.map(\.table))
        if let id = truncateFrom, let cut, !deletedBlocks.isEmpty {
            let before = index.children(of: cut.body).filter { index.tags[$0].start < index.tags[cut.paragraph].start }
            if contentBlocks(before, in: index).allSatisfy({ deletedBlocks.contains($0) }) {
                throw FormError.removalsLeaveNothingBeforeCut(id)
            }
        }
        var owners: [Int: String] = [:]
        for (id, p) in zip(ids, removed) where owners[p] == nil { owners[p] = id }
        for entry in removedTables { owners[entry.table] = entry.id }
        let addFinalParagraph = try checkBodyStructure(owners, truncateFrom: truncateFrom.flatMap { id in
            cut.map { (id, $0.paragraph, $0.body) }
        }, in: index)
        // A kept marker goes where deleteAll puts it; refuse when that would
        // change the Markup Compatibility rules applying to it (R4).
        let deleted = spans.map(\.range)
        for p in removedSet.sorted() {
            // A named paragraph with no parent is the root of a crafted part;
            // the result check reports it, there is no context to compare.
            guard let destination = survivingParagraphBefore(p, removed: removedSet, in: index) ?? index.tags[p].parent else { continue }
            for marker in FormRanges.keptMarkers(in: index.elementRange(p), deleted: deleted, index: index)
            where FormRanges.compatibilityContextChanges(marker, destination: destination, index: index) {
                throw FormError.markerUnderCompatibilityRules(index.attribute("w14:paraId", of: p) ?? "")
            }
        }
        for entry in removedTables {
            guard let parent = index.tags[entry.table].parent else { continue }
            for marker in FormRanges.keptMarkers(in: index.elementRange(entry.table), deleted: deleted, index: index)
            where FormRanges.compatibilityContextChanges(marker, destination: parent, index: index) {
                throw FormError.markerUnderCompatibilityRules(entry.id)
            }
        }
        if let id = truncateFrom, let body = cut?.body, let span = spans.last(where: { $0.id == id }) {
            for marker in FormRanges.keptMarkers(in: span.range, deleted: deleted, index: index)
            where FormRanges.compatibilityContextChanges(marker, destination: body, index: index) {
                throw FormError.markerUnderCompatibilityRules(id)
            }
        }
        return (spans, addFinalParagraph)
    }

    /// The table `--remove-table` deletes for a paragraph: the nearest table
    /// around it, which must be a body-level block (looking through
    /// body-level content controls and custom XML) (macdoc#250).
    static func tableToRemove(_ id: String, in index: XMLTagIndex) throws -> Int {
        let table: Int
        switch enclosingTable(of: try locate(id, in: index), in: index) {
        case .table(let t): table = t
        case .notInTable: throw FormError.notInTable(id)
        case .inTextBox: throw FormError.inTextBox(id)
        }
        guard let container = blockContainer(of: table, in: index), index.tags[container].name == "w:body" else {
            throw FormError.tableNotBodyLevel(id)
        }
        return table
    }

    enum TableLookup: Equatable {
        case table(Int)
        case notInTable
        /// The paragraph is in a text box inside a table.
        case inTextBox
    }

    /// The nearest table around a paragraph. A text box is its own story:
    /// the table around it is not the paragraph's table, so the walk stops
    /// there (macdoc#250 verify R1).
    static func enclosingTable(of p: Int, in index: XMLTagIndex) -> TableLookup {
        var current = index.tags[p].parent
        while let c = current, index.tags[c].name != "w:tbl" {
            if index.tags[c].name == "w:txbxContent" {
                return index.ancestor(of: c, named: "w:tbl") == nil ? .notInTable : .inTextBox
            }
            current = index.tags[c].parent
        }
        return current.map { .table($0) } ?? .notInTable
    }

    /// Each table named by `ids` once, under the first id naming it, in
    /// document order.
    static func uniqueTables(_ ids: [String], in index: XMLTagIndex) throws -> [(id: String, table: Int)] {
        var seen: Set<Int> = []
        var out: [(id: String, table: Int)] = []
        for id in ids {
            let table = try tableToRemove(id, in: index)
            if seen.insert(table).inserted { out.append((id, table)) }
        }
        return out.sorted { $0.table < $1.table }
    }

    /// The body-level block structure every deletion keeps (macdoc#250), judged on
    /// the body's content blocks (looking through content controls and
    /// custom XML) before and after everything named goes. Only what this
    /// command causes is refused; a part that already had the shape passes.
    /// (a) Two tables that something other than a table separated must not
    ///     become adjacent: Word would join them.
    /// (b) The body must keep content; deleting every block is refused. A body
    ///     that ended with a paragraph and would end with a table (or another
    ///     block that is not a paragraph) gets an empty final paragraph; the
    ///     return value says so.
    /// (c) Without a truncation, the last section must keep content: the last
    ///     kept block must not be a paragraph with its own section break when
    ///     the body has its own final section properties. (Truncation folds
    ///     such a break into the final section, macdoc#249.)
    @discardableResult
    static func checkBodyStructure(_ owners: [Int: String], truncateFrom cut: (id: String, paragraph: Int, body: Int)?,
                                   in index: XMLTagIndex) throws -> Bool {
        guard let body = cut?.body ?? index.elements(named: "w:body").first else { return false }
        let blocks = contentBlocks(index.children(of: body), in: index)
        guard let lastBlock = blocks.indices.last else { return false }
        let cutStart = cut.map { index.tags[$0.paragraph].start }
        func owner(_ i: Int) -> String? {
            if let id = owners[blocks[i]] { return id }
            if let cut, let cutStart, index.tags[blocks[i]].start >= cutStart { return cut.id }
            return nil
        }
        func names(_ positions: ClosedRange<Int>) -> [String] {
            var seen: Set<String> = []
            return positions.compactMap(owner).filter { seen.insert($0).inserted }
        }
        let isTable = { (i: Int) in index.tags[blocks[i]].name == "w:tbl" }
        let isParagraph = { (i: Int) in index.tags[blocks[i]].name == "w:p" }
        let kept = blocks.indices.filter { owner($0) == nil }
        for (a, b) in zip(kept, kept.dropFirst()) where b > a + 1 && isTable(a) && isTable(b) {
            if ((a + 1)..<b).contains(where: { !isTable($0) }) { throw FormError.tablesWouldJoin(names((a + 1)...(b - 1))) }
        }
        guard let last = kept.last else { throw FormError.removalsLeaveEmptyBody(names(0...lastBlock)) }
        if cut == nil, last < lastBlock, isParagraph(last),
           ownSectPr(of: blocks[last], in: index) != nil,
           index.children(of: body).contains(where: { index.tags[$0].name == "w:sectPr" }) {
            throw FormError.lastSectionWouldBeEmpty(names((last + 1)...lastBlock))
        }
        return isParagraph(lastBlock) && !isParagraph(last)
    }

    /// The only entry point that deletes (macdoc#253 verify R2, R3): checks the whole
    /// deletion, removes every named paragraph and table in one edit, then
    /// truncates.
    /// A range marker whose partner stays is kept where its paragraph was;
    /// when no surviving block follows the paragraph in its cell or text box,
    /// it goes to the end of the nearest surviving paragraph before it, so the
    /// container still ends with a paragraph whatever order the ids came in.
    /// The truncation's kept markers are taken on the part before the removal
    /// edit, where the partners of removed markers can still be seen.
    @discardableResult
    static func deleteAll(_ ids: [String], tables: [String] = [], truncateFrom: String?, editor: inout XMLEditor) throws -> Bool {
        let index = editor.index
        let (spans, addFinalParagraph) = try deletionPlan(ids, tables: tables, truncateFrom: truncateFrom, in: index)
        guard !spans.isEmpty else { return false }
        let deleted = spans.map(\.range)
        let removed = Set(try ids.map { try locate($0, in: index) })
        var edits: [Edit] = []
        var appended: [Int: String] = [:]
        for p in removed.sorted() {
            let span = index.elementRange(p)
            if let destination = survivingParagraphBefore(p, removed: removed, in: index) {
                appended[destination, default: ""] += FormRanges.keptMarkerXML(in: span, deleted: deleted,
                                                                               destination: destination, index: index)
                edits.append(.remove(span))
            } else {
                edits.append(.replace(span, with: FormRanges.keptMarkerXML(in: span, deleted: deleted,
                                                                           destination: index.tags[p].parent, index: index)))
            }
        }
        // A removed table leaves its kept range markers where it was (macdoc#250).
        for entry in try uniqueTables(tables, in: index) {
            let span = index.elementRange(entry.table)
            edits.append(.replace(span, with: FormRanges.keptMarkerXML(in: span, deleted: deleted,
                                                                       destination: index.tags[entry.table].parent, index: index)))
        }
        for (destination, markers) in appended where !markers.isEmpty {
            edits.append(index.tags[destination].kind == .selfClosing
                ? .replace(index.elementRange(destination), with: openTag(of: destination, in: index) + markers + "</w:p>")
                : .insert(markers, at: index.tags[index.tags[destination].match!].start))
        }
        let truncation = try truncateFrom.map { id -> (id: String, markers: String) in
            guard let span = spans.last(where: { $0.id == id }) else { return (id, "") }
            return (id, FormRanges.keptMarkerXML(in: span.range, deleted: deleted,
                                                 destination: try checkTruncate(id, in: index).body, index: index))
        }
        // Rule (b): the body ends with a paragraph again, as its last block
        // right before the body's own section properties; the cut writes it
        // itself after the kept range markers.
        if addFinalParagraph, truncation == nil, let body = index.elements(named: "w:body").first {
            let at = index.children(of: body).last(where: { index.tags[$0].name == "w:sectPr" })
                .map { index.tags[$0].start } ?? bodyEnd(of: body, in: index)
            edits.append(.insert("<w:p/>", at: at))
        }
        if !edits.isEmpty { try editor.apply(edits) }
        if let truncation {
            // The paragraph goes last, right before the final section
            // properties, as it does without a cut.
            try truncate(from: truncation.id, keptMarkers: truncation.markers + (addFinalParagraph ? "<w:p/>" : ""),
                         editor: &editor)
        }
        return addFinalParagraph
    }

    /// For a removed paragraph with no surviving block after it in its cell or
    /// text box: the nearest surviving paragraph before it.
    static func survivingParagraphBefore(_ p: Int, removed: Set<Int>, in index: XMLTagIndex) -> Int? {
        guard let container = blockContainer(of: p, in: index),
              ["w:tc", "w:txbxContent"].contains(index.tags[container].name) else { return nil }
        let blocks = flattenBlocks(index.children(of: container), in: index)
        guard let at = blocks.firstIndex(of: p),
              !blocks[(at + 1)...].contains(where: { !removed.contains($0) }),
              let before = blocks[..<at].last(where: { !removed.contains($0) }),
              index.tags[before].name == "w:p" else { return nil }
        return before
    }

    /// Reached only through `deleteAll`, which has judged the whole deletion;
    /// here only the cut's own structure is checked again, not the pairing.
    private static func truncate(from id: String, keptMarkers: String, editor: inout XMLEditor) throws {
        let (p, body) = try checkTruncate(id, in: editor.index)
        let index = editor.index
        let cut = index.tags[p].start
        let bodyEnd = bodyEnd(of: body, in: index)
        // Only live section properties count: a body-level paragraph's own
        // w:pPr/w:sectPr, or the body's own w:sectPr. A w:sectPr kept inside
        // w:pPrChange is revision history and must not become the result.
        var section = ""
        search: for child in index.children(of: body) where index.tags[child].start >= cut {
            if index.tags[child].name == "w:sectPr" {
                section = index.elementString(child)
                break
            }
            // Body-level paragraphs, including those inside body-level content controls.
            for paragraph in flattenBlocks([child], in: index) where index.tags[paragraph].name == "w:p" {
                if let own = ownSectPr(of: paragraph, in: index) {
                    section = index.elementString(own)
                    break search
                }
            }
        }
        // When the kept content ends with its own section break, the cut
        // removed a whole section. Keeping both would leave a final section
        // with no content (an extra blank page), so the kept break becomes the
        // body's final w:sectPr and the removed range's one is dropped.
        var edits: [Edit] = []
        let kept = contentBlocks(index.children(of: body).filter { index.tags[$0].start < cut }, in: index)
        if let last = kept.last, index.tags[last].name == "w:p", let own = ownSectPr(of: last, in: index),
           let pPr = ownPPr(of: last, in: index) {
            section = index.elementString(own)
            edits.append(.replace(index.elementRange(pPr),
                                  with: try Fragment.removeChildren(index.elementString(pPr), named: ["w:sectPr"])))
        }
        // Range markers whose partner is kept end the kept content, before the
        // final section properties (macdoc#253); the plan serializes them.
        edits.append(.replace(cut..<bodyEnd, with: keptMarkers + section))
        try editor.apply(edits)
    }
}
