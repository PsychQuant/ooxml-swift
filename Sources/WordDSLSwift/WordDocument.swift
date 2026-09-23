// WordDocument.swift
// word-aligned-state-sync Phase 4 task 5.3 — root container of an `.mdocx`
// script: op-log generation ("the compiler emits operations") and the
// atomic three-file save.

import Foundation
import OOXMLSwift

/// Loud emission failure — a DSL element whose op-log channel doesn't exist
/// yet must fail the build, never silently drop content ("apply errors are
/// reported, not swallowed" discipline).
public enum DSLEmissionError: Error, Equatable {
    case unsupportedElement(element: String, reason: String)
}

/// Root container of an `.mdocx` script.
public struct WordDocument {

    public let sections: [Section]

    /// Result-builder entry point: `WordDocument { Section(id: "main") { … } }`.
    public init(@WordBuilder content: () -> [Section]) {
        self.sections = content()
    }

    /// Empty document (`WordDocument { }` also resolves here via the
    /// zero-statement builder block).
    public init() {
        self.sections = []
    }

    // MARK: - Op-log generation

    /// Emits the document's operations in declaration order — the execution
    /// semantics of the DSL (`ooxml-script-transcode`: a script's execution
    /// against an empty log reproduces the log).
    ///
    /// Emission rules (v0.34 slice):
    /// - style define-on-first-use: the first paragraph referencing a
    ///   `WordStyle` emits one `defineStyle`; later references don't
    /// - plain paragraph (String-only body) → `appendParagraph` carrying the
    ///   joined text
    /// - body with formatted `Run`s → `appendParagraph(text: "")` +
    ///   `setRuns` with the full ordered run list
    /// - inline atoms → canonical run-targeted `insertTab`/`insertBreak`/
    ///   `insertNoBreakHyphen` ops AFTER the paragraph's structural content
    ///   op; stable run IDs preserve declaration order on replay.
    public func buildLog() throws -> OperationLog {
        var log = OperationLog()
        var definedStyles = Set<String>()
        var nextBookmarkID = 0
        var nextRelationshipID = 1
        var blockBookmarkIDs: [String: Int] = [:]

        for (sectionIndex, section) in sections.enumerated() {
            var lastParagraphID: String?
            for child in section.children {
                switch child {
                case .paragraph(let paragraph):
                    try emit(paragraph: paragraph, into: &log,
                             definedStyles: &definedStyles,
                             nextBookmarkID: &nextBookmarkID,
                             nextRelationshipID: &nextRelationshipID)
                    lastParagraphID = paragraph.id
                case .component(let type, let id, let body):
                    // mdocx-grammar "Component-aware op log": paired envelope
                    // bracketing the body's operations. Body paragraphs append
                    // to the document body (in: nil) — the component id is an
                    // envelope identity, not a tree node (reducer treats the
                    // markers as no-ops).
                    log.append(.beginComponent(type: type,
                                               id: ElementID(rawString: id)), source: .swift)
                    for paragraph in body {
                        try emit(paragraph: paragraph, into: &log,
                                 definedStyles: &definedStyles,
                                 nextBookmarkID: &nextBookmarkID,
                                 nextRelationshipID: &nextRelationshipID)
                        lastParagraphID = paragraph.id
                    }
                    log.append(.endComponent(id: ElementID(rawString: id)), source: .swift)
                case .table(let table):
                    let columns = table.rows.map(\.cells.count).max() ?? 0
                    guard !table.rows.isEmpty, columns > 0,
                          table.rows.allSatisfy({ $0.cells.count == columns }) else {
                        throw DSLEmissionError.unsupportedElement(
                            element: "Table(id: \(table.id))",
                            reason: "table must be a non-empty rectangular grid")
                    }
                    let cells = table.rows.map { row in
                        row.cells.map { cell in
                            cell.paragraphs.map(paragraphText).joined(separator: "\n")
                        }
                    }
                    for row in table.rows {
                        for cell in row.cells {
                            for paragraph in cell.paragraphs {
                                if let style = paragraph.style,
                                   !definedStyles.contains(style.styleId) {
                                    definedStyles.insert(style.styleId)
                                    log.append(.defineStyle(payload: style.payload), source: .swift)
                                }
                            }
                        }
                    }
                    var tableAtomOperations: [OOXMLSwift.Operation] = []
                    var tableRelationships: [OOXMLSwift.Operation] = []
                    let richCells = table.rows.map { row in
                        row.cells.map { cell in
                            cell.paragraphs.map { paragraph in
                                let runs = paragraph.children.compactMap { child -> RunPayload? in
                                    switch child {
                                    case .text(let text): return RunPayload(text: text)
                                    case .run(let run):
                                        return RunPayload(text: run.text, bold: run.bold,
                                                          italic: run.italic, color: run.color)
                                    default: return nil
                                    }
                                }
                                let hasStructuredInline = paragraph.children.contains {
                                    switch $0 {
                                    case .bookmark, .hyperlink, .tab, .lineBreak, .noBreakHyphen:
                                        return true
                                    default:
                                        return false
                                    }
                                }
                                let items = hasStructuredInline ? inlineItems(
                                    paragraph.children,
                                    nextBookmarkID: &nextBookmarkID,
                                    nextRelationshipID: &nextRelationshipID,
                                    relationships: &tableRelationships,
                                    atomOperations: &tableAtomOperations) : nil
                                return TableParagraphPayload(
                                    paragraph: paragraphPayload(paragraph, text: paragraphText(paragraph)),
                                    runs: hasStructuredInline ? nil : runs,
                                    items: items)
                            }
                        }
                    }
                    log.append(.appendTable(
                        in: nil,
                        table: TablePayload(rows: table.rows.count,
                                            columns: columns, cells: cells,
                                            richCells: richCells,
                                            tableId: table.id,
                                            rowIds: table.rows.map(\.id),
                                            cellIds: table.rows.map {
                                                $0.cells.map(\.id)
                                            })), source: .swift)
                    for atom in tableAtomOperations {
                        log.append(atom, source: .swift)
                    }
                    for relationship in tableRelationships {
                        log.append(relationship, source: .swift)
                    }
                case .bookmarkStart(let marker):
                    let bookmarkID = nextBookmarkID
                    nextBookmarkID += 1
                    blockBookmarkIDs[marker.id] = bookmarkID
                    log.append(.appendBlockMarker(marker: InlineMarker(
                        localName: "bookmarkStart",
                        attributes: [
                            RootAttribute(prefix: "w", localName: "id",
                                          value: String(bookmarkID)),
                            RootAttribute(prefix: "w", localName: "name",
                                          value: marker.id),
                        ])), source: .swift)
                case .bookmarkEnd(let marker):
                    guard let bookmarkID = blockBookmarkIDs.removeValue(forKey: marker.id) else {
                        throw DSLEmissionError.unsupportedElement(
                            element: "BookmarkEnd(id: \(marker.id))",
                            reason: "bookmark end has no matching start")
                    }
                    log.append(.appendBlockMarker(marker: InlineMarker(
                        localName: "bookmarkEnd",
                        attributes: [RootAttribute(
                            prefix: "w", localName: "id",
                            value: String(bookmarkID))])), source: .swift)
                }
            }
            let isFinalSection = sectionIndex == sections.count - 1
            let anchor = isFinalSection
                ? nil
                : lastParagraphID.map { ElementID(rawString: "w14:paraId=\($0)") }
            log.append(.setSectionProperties(
                at: anchor,
                section: SectionPayload(sectionType: section.type?.rawValue)), source: .swift)
        }
        if let unmatched = blockBookmarkIDs.keys.sorted().first {
            throw DSLEmissionError.unsupportedElement(
                element: "BookmarkStart(id: \(unmatched))",
                reason: "bookmark start has no matching end")
        }
        return log
    }

    /// Emits one paragraph's ops (define-on-first-use + append + runs + atoms).
    private func emit(paragraph: Paragraph, into log: inout OperationLog,
                      definedStyles: inout Set<String>,
                      nextBookmarkID: inout Int,
                      nextRelationshipID: inout Int) throws {
                if let style = paragraph.style, !definedStyles.contains(style.styleId) {
                    definedStyles.insert(style.styleId)
                    log.append(.defineStyle(payload: style.payload), source: .swift)
                }

                let texts: [String] = paragraph.children.compactMap {
                    if case .text(let s) = $0 { return s } else { return nil }
                }
                let hasFormattedRuns = paragraph.children.contains {
                    if case .run = $0 { return true } else { return false }
                }
                let hasStructuredInline = paragraph.children.contains {
                    switch $0 {
                    case .bookmark, .hyperlink, .tab, .lineBreak, .noBreakHyphen:
                        return true
                    default: return false
                    }
                }
                let target = ElementID(rawString: "w14:paraId=\(paragraph.id)")

                log.append(.appendParagraph(
                    in: nil,
                    paragraph: paragraphPayload(
                        paragraph,
                        text: (hasFormattedRuns || hasStructuredInline)
                            ? "" : texts.joined())), source: .swift)

                if hasStructuredInline {
                    var relationships: [OOXMLSwift.Operation] = []
                    var atomOperations: [OOXMLSwift.Operation] = []
                    let items = inlineItems(
                        paragraph.children,
                        nextBookmarkID: &nextBookmarkID,
                        nextRelationshipID: &nextRelationshipID,
                        relationships: &relationships,
                        atomOperations: &atomOperations)
                    log.append(.setParagraphContent(target: target, items: items), source: .swift)
                    for atom in atomOperations {
                        log.append(atom, source: .swift)
                    }
                    for relationship in relationships {
                        log.append(relationship, source: .swift)
                    }
                } else if hasFormattedRuns {
                    let runs: [RunPayload] = paragraph.children.compactMap {
                        switch $0 {
                        case .text(let s): return RunPayload(text: s)
                        case .run(let r):
                            return RunPayload(text: r.text, bold: r.bold,
                                              italic: r.italic, color: r.color)
                        default: return nil
                        }
                    }
                    log.append(.setRuns(target: target, runs: runs), source: .swift)
                }

    }

    private func paragraphText(_ paragraph: Paragraph) -> String {
        inlineText(paragraph.children)
    }

    private func paragraphPayload(_ paragraph: Paragraph, text: String) -> ParagraphPayload {
        ParagraphPayload(
            text: text,
            styleId: paragraph.style?.styleId,
            paraId: paragraph.id,
            numId: paragraph.numbering?.rawValue,
            numLevel: paragraph.level)
    }

    private func inlineText(_ children: [InlineChild]) -> String {
        children.map { child in
            switch child {
            case .text(let text): return text
            case .run(let run): return run.text
            case .tab: return "\t"
            case .lineBreak: return "\n"
            case .noBreakHyphen: return "‑"
            case .bookmark(let bookmark): return inlineText(bookmark.children)
            case .hyperlink(let hyperlink): return inlineText(hyperlink.children)
            }
        }.joined()
    }

    private func inlineItems(
        _ children: [InlineChild],
        nextBookmarkID: inout Int,
        nextRelationshipID: inout Int,
        relationships: inout [OOXMLSwift.Operation],
        atomOperations: inout [OOXMLSwift.Operation]
    ) -> [InlineItem] {
        var items: [InlineItem] = []
        var lastRunID: UUID?

        func appendRun(_ payload: RunPayload) {
            let id = UUID()
            items.append(.run(payload, id: id))
            lastRunID = id
        }

        func appendAtom(_ operation: (ElementID) -> OOXMLSwift.Operation) {
            if lastRunID == nil { appendRun(RunPayload(text: "")) }
            atomOperations.append(operation(ElementID(libraryUUID: lastRunID!)))
        }

        for child in children {
            switch child {
            case .text(let text):
                appendRun(RunPayload(text: text))
            case .run(let run):
                appendRun(RunPayload(
                    text: run.text, bold: run.bold,
                    italic: run.italic, color: run.color))
            case .bookmark(let bookmark):
                let id = nextBookmarkID
                nextBookmarkID += 1
                items.append(.marker(InlineMarker(
                    localName: "bookmarkStart",
                    attributes: [
                        RootAttribute(prefix: "w", localName: "id", value: String(id)),
                        RootAttribute(prefix: "w", localName: "name", value: bookmark.id),
                    ])))
                items.append(contentsOf: inlineItems(
                    bookmark.children,
                    nextBookmarkID: &nextBookmarkID,
                    nextRelationshipID: &nextRelationshipID,
                    relationships: &relationships,
                    atomOperations: &atomOperations))
                items.append(.marker(InlineMarker(
                    localName: "bookmarkEnd",
                    attributes: [RootAttribute(
                        prefix: "w", localName: "id", value: String(id))])))
                // A following atom belongs to the paragraph again, not to
                // the run that preceded the bookmark container. Clear the
                // outer carrier so Tab/Break synthesizes a new wrapping run
                // after bookmarkEnd in declaration order.
                lastRunID = nil
            case .hyperlink(let hyperlink):
                let hyperlinkItems = inlineItems(
                    hyperlink.children,
                    nextBookmarkID: &nextBookmarkID,
                    nextRelationshipID: &nextRelationshipID,
                    relationships: &relationships,
                    atomOperations: &atomOperations)
                let runs = hyperlink.children.compactMap { item -> RunPayload? in
                    switch item {
                    case .text(let text): return RunPayload(text: text)
                    case .run(let run):
                        return RunPayload(text: run.text, bold: run.bold,
                                          italic: run.italic, color: run.color)
                    default: return nil
                    }
                }
                switch hyperlink.target {
                case .anchor(let anchor):
                    items.append(.hyperlink(HyperlinkPayload(
                        targetKind: .anchor, target: anchor, runs: runs,
                        items: hyperlinkItems)))
                case .url(let url):
                    let rId = "rIdMdocx\(nextRelationshipID)"
                    nextRelationshipID += 1
                    items.append(.hyperlink(HyperlinkPayload(
                        targetKind: .external, target: url,
                        relationshipId: rId, runs: runs,
                        items: hyperlinkItems)))
                    relationships.append(.addRelationship(
                        part: "word/_rels/document.xml.rels",
                        id: rId,
                        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                        target: url,
                        targetMode: "External"))
                case .mailto(let address):
                    let target = address.hasPrefix("mailto:")
                        ? address : "mailto:\(address)"
                    let rId = "rIdMdocx\(nextRelationshipID)"
                    nextRelationshipID += 1
                    items.append(.hyperlink(HyperlinkPayload(
                        targetKind: .external, target: target,
                        relationshipId: rId, runs: runs,
                        items: hyperlinkItems)))
                    relationships.append(.addRelationship(
                        part: "word/_rels/document.xml.rels",
                        id: rId,
                        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                        target: target,
                        targetMode: "External"))
                }
                // A paragraph-level atom following a hyperlink has no
                // preceding run in the same container, so it must synthesize
                // its own wrapper rather than target inside the hyperlink.
                lastRunID = nil
            case .tab:
                appendAtom { .insertTab(in: $0) }
            case .lineBreak:
                appendAtom { .insertBreak(in: $0) }
            case .noBreakHyphen:
                appendAtom { .insertNoBreakHyphen(in: $0) }
            }
        }
        return items
    }

    // MARK: - save(to:) atomic three-file write (mdocx-grammar requirement)

    /// Writes `<name>.docx` + the op-log and snapshot sidecars as one logical
    /// state. On failure of ANY of the three writes the file system is
    /// restored to its pre-save state. Refuses while Word holds the docx
    /// open (`~$` owner file present).
    ///
    /// Sidecars follow the full-name grammar:
    /// `<name>.docx.oplog.jsonl` and `<name>.docx.snapshot.json`.
    public func save(to url: URL) throws {
        if WordLock.isLockedByWord(url) {
            throw SyncError.fileLockedByWord(lockURL: WordLock.lockFileURL(for: url))
        }

        let log = try buildLog()
        var doc = OOXMLSwift.WordDocument.emptyAuthoringDocument()
        try doc.apply(log: log)

        let fm = FileManager.default
        let targets = [url, SidecarStore.oplogURL(for: url), SidecarStore.snapshotURL(for: url)]
        // Backup capture distinguishes "absent" (nil — rollback removes the
        // freshly created file) from "present but unreadable" (throw — abort
        // the save BEFORE any write; `try?` here conflated the two, and a
        // transient read failure at backup time would have let rollback
        // DELETE the user's pre-existing file. 7.5 verify panel P1.)
        let backups: [Data?] = try targets.map { target in
            guard fm.fileExists(atPath: target.path) else { return nil }
            return try Data(contentsOf: target)
        }

        do {
            try doc.writeAuthoringPackage(to: url)
            try SidecarStore.saveLog(doc.operationLog, alongside: url)

            let docxData = try Data(contentsOf: url)
            var fingerprints: [String: String] = [:]
            for (partPath, tree) in doc.xmlTrees {
                fingerprints[partPath] = tree.root.normalizedFingerprint()
            }
            var documentXML: String?
            if let docTree = doc.xmlTrees["word/document.xml"],
               let serialized = try? XmlTreeWriter.serialize(docTree) {
                documentXML = String(decoding: serialized, as: UTF8.self)
            }
            try SidecarStore.saveSnapshot(SyncSnapshot(
                docxSHA256: SidecarStore.sha256Hex(of: docxData),
                savedAt: Date(),
                opCount: doc.operationLog.entries.count,
                partFingerprints: fingerprints,
                documentXML: documentXML,
                partSHA256: SidecarStore.partSHA256(
                    try RawPartChannel.readAllParts(from: url))), alongside: url)
        } catch {
            // Roll back: restore pre-existing bytes, remove freshly created files.
            for (i, target) in targets.enumerated() {
                if let backup = backups[i] {
                    try? backup.write(to: target)
                } else {
                    try? fm.removeItem(at: target)
                }
            }
            throw error
        }
    }
}
