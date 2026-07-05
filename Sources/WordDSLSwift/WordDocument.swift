// WordDocument.swift
// word-aligned-state-sync Phase 4 task 5.3 — root container of an `.mdocx`
// script: op-log generation ("the compiler emits operations") and the
// atomic three-file save.

import Foundation
import OOXMLSwift

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
    public func buildLog() -> OperationLog {
        var log = OperationLog()
        var definedStyles = Set<String>()

        for section in sections {
            for paragraph in section.paragraphs {
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
                    default: break
                    }
                }
            }
        }
        return log
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

        let log = buildLog()
        var doc = OOXMLSwift.WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: log.entries.map(\.op), source: .swift)

        let fm = FileManager.default
        let targets = [url, SidecarStore.oplogURL(for: url), SidecarStore.snapshotURL(for: url)]
        let backups: [Data?] = targets.map { try? Data(contentsOf: $0) }

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
