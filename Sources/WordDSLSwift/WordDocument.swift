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
    /// - inline atoms → paragraph-targeted `insertTab`/`insertBreak`/
    ///   `insertNoBreakHyphen` ops AFTER the paragraph's content ops.
    ///   KNOWN LIMITATION (5.5 scope): atoms replay at the end of the
    ///   paragraph, so an atom interleaved BETWEEN runs serializes after
    ///   them; declaration order inside the log is preserved for the
    ///   reverse transcoder either way.
    public func buildLog() throws -> OperationLog {
        var log = OperationLog()
        var definedStyles = Set<String>()

        for section in sections {
            for child in section.children {
                switch child {
                case .paragraph(let paragraph):
                    try emit(paragraph: paragraph, into: &log, definedStyles: &definedStyles)
                case .component(let type, let id, let body):
                    // mdocx-grammar "Component-aware op log": paired envelope
                    // bracketing the body's operations. Body paragraphs append
                    // to the document body (in: nil) — the component id is an
                    // envelope identity, not a tree node (reducer treats the
                    // markers as no-ops).
                    log.append(.beginComponent(type: type,
                                               id: ElementID(rawString: id)), source: .swift)
                    for paragraph in body {
                        try emit(paragraph: paragraph, into: &log, definedStyles: &definedStyles)
                    }
                    log.append(.endComponent(id: ElementID(rawString: id)), source: .swift)
                case .table(let table):
                    throw DSLEmissionError.unsupportedElement(
                        element: "Table(id: \(table.id))",
                        reason: "table emission awaits an appendTable authoring op + reducer support (post-v0.34 taxonomy increment); the DSL type compiles so scripts stay source-stable")
                case .bookmarkStart(let marker):
                    throw DSLEmissionError.unsupportedElement(
                        element: "BookmarkStart(id: \(marker.id))",
                        reason: "cross-paragraph bookmark emission awaits the reducer's insertBookmark implementation (Phase 2c residue)")
                case .bookmarkEnd(let marker):
                    throw DSLEmissionError.unsupportedElement(
                        element: "BookmarkEnd(id: \(marker.id))",
                        reason: "cross-paragraph bookmark emission awaits the reducer's insertBookmark implementation (Phase 2c residue)")
                }
            }
        }
        return log
    }

    /// Emits one paragraph's ops (define-on-first-use + append + runs + atoms).
    private func emit(paragraph: Paragraph, into log: inout OperationLog,
                      definedStyles: inout Set<String>) throws {
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
                let target = ElementID(rawString: "w14:paraId=\(paragraph.id)")

                log.append(.appendParagraph(in: nil, paragraph: ParagraphPayload(
                    text: hasFormattedRuns ? "" : texts.joined(),
                    styleId: paragraph.style?.styleId,
                    paraId: paragraph.id)), source: .swift)

                if hasFormattedRuns {
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

                for child in paragraph.children {
                    switch child {
                    case .tab: log.append(.insertTab(in: target), source: .swift)
                    case .lineBreak: log.append(.insertBreak(in: target), source: .swift)
                    case .noBreakHyphen: log.append(.insertNoBreakHyphen(in: target), source: .swift)
                    case .bookmark(let b):
                        throw DSLEmissionError.unsupportedElement(
                            element: "Bookmark(id: \(b.id))",
                            reason: "bookmark emission awaits the reducer's insertBookmark implementation (Phase 2c residue); the DSL type compiles so scripts stay source-stable")
                    case .hyperlink:
                        throw DSLEmissionError.unsupportedElement(
                            element: "Hyperlink",
                            reason: "hyperlink emission awaits an authoring-op channel (EditAlgebra lowering not yet reachable from buildLog)")
                    default: break
                    }
                }
    }

    // MARK: - save(to:) atomic three-file write (mdocx-grammar requirement)

    /// Writes `<name>.docx` + the op-log and snapshot sidecars as one logical
    /// state. On failure of ANY of the three writes the file system is
    /// restored to its pre-save state. Refuses while Word holds the docx
    /// open (`~$` owner file present).
    ///
    /// Sidecar naming note: `mdocx-grammar`'s scenario text spells the
    /// sidecars `<name>.docx.oplog.jsonl`; the implemented convention is the
    /// `ooxml-word-sync` stem form (`<name>.oplog.jsonl`, via `SidecarStore`)
    /// — the two specs disagree and the shipped SidecarStore wins for
    /// consistency with `SyncOrchestrator`. Flagged for spec errata.
    public func save(to url: URL) throws {
        if WordLock.isLockedByWord(url) {
            throw SyncError.fileLockedByWord(lockURL: WordLock.lockFileURL(for: url))
        }

        let log = try buildLog()
        var doc = OOXMLSwift.WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: log.entries.map(\.op), source: .swift)

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
                documentXML: documentXML), alongside: url)
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
