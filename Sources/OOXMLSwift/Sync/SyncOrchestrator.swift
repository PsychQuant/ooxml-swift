// SyncOrchestrator.swift
// word-aligned-state-sync Phase 3 tasks 4.1 + 4.7 — bidirectional state
// alignment between the operation log and the on-disk docx
// (`ooxml-word-sync` Requirements "SyncOrchestrator coordinates Word and
// Swift writers", "Sidecar persistence of snapshot and log", "Bootstrap
// from existing docx").
//
// Turn-based, not concurrent (design Non-Goal: no CRDT/OT): at any instant
// either Word or Swift owns the file. The orchestrator's job is detecting
// whose turn ended (watcher + lock-file lifecycle), importing Word's edits
// as `source: "word"` operations, and flushing Swift's pending operations
// back to disk.
//
// Known deviation (documented): `flush()` persists through
// `WordDocument.saveWithSidecars` → `DocxWriter.write` (typed writer +
// overlay byte-copy) rather than serializing every part via `XmlTreeWriter`
// directly. Behavior is equivalent for the mutation surface Phase 2/3
// support (mutations resync the typed view and mark their part dirty);
// the full tree-writer migration is Phase 5 (v1.0, one IO path).

import Foundation

public final class SyncOrchestrator {

    /// In-memory document: xmlTrees + operationLog are canonical.
    public private(set) var document: WordDocument
    public let docxURL: URL

    /// Baseline `word/document.xml` tree = the last state both writers
    /// agreed on (bootstrap read, last import, or last flush). Word-import
    /// diffs run against this.
    private var baselineDocumentTree: XmlTree

    /// Log length at the last flush/bootstrap: entries beyond this index
    /// with `source == .swift` are "pending" (not yet on disk) for
    /// conflict-detection purposes.
    private var flushedOpCount: Int

    private var changeDetector: DocxChangeDetector

    private init(document: WordDocument, docxURL: URL,
                 baseline: XmlTree, flushedOpCount: Int,
                 changeDetector: DocxChangeDetector) {
        self.document = document
        self.docxURL = docxURL
        self.baselineDocumentTree = baseline
        self.flushedOpCount = flushedOpCount
        self.changeDetector = changeDetector
    }

    // MARK: - Bootstrap (task 4.7)

    /// Initializes a sync session from any docx (spec Requirement
    /// "Bootstrap from existing docx"):
    /// - no sidecars → current docx becomes the initial snapshot, empty
    ///   log, sidecars created immediately;
    /// - existing sidecars → log + snapshot loaded; when the docx content
    ///   hash differs from the snapshot's (Word edited between sessions),
    ///   an import diff runs to capture the intervening changes.
    public static func bootstrapFromDocx(
        url: URL, policy: SyncPolicy = .abortOnConflict
    ) throws -> SyncOrchestrator {
        var document = try DocxReader.read(from: url, wireTreeBackedViews: true)
        guard let currentTree = document.xmlTrees["word/document.xml"] else {
            throw SyncError.missingDocumentTree(partPath: "word/document.xml")
        }

        let existingLog = try SidecarStore.loadLog(alongside: url)
        let existingSnapshot = try SidecarStore.loadSnapshot(alongside: url)

        if let log = existingLog { document.operationLog = log }

        let detector = try DocxChangeDetector(url: url)

        guard let snapshot = existingSnapshot else {
            // Fresh start: sidecars created with the docx's current state.
            let orchestrator = SyncOrchestrator(
                document: document, docxURL: url,
                baseline: currentTree.deepCopy(),
                flushedOpCount: document.operationLog.entries.count,
                changeDetector: detector)
            try document.saveWithSidecars(to: url)
            // Re-baseline the detector: saveWithSidecars rewrote the docx.
            orchestrator.changeDetector = try DocxChangeDetector(url: url)
            return orchestrator
        }

        // Existing sidecars. Baseline = snapshot's stored document.xml when
        // present (true last-synced state), else the current tree.
        var baseline = currentTree.deepCopy()
        if let storedXML = snapshot.documentXML,
           let parsed = try? XmlTreeReader.parse(Data(storedXML.utf8)) {
            baseline = parsed
        }

        let orchestrator = SyncOrchestrator(
            document: document, docxURL: url,
            baseline: baseline,
            flushedOpCount: document.operationLog.entries.count,
            changeDetector: detector)

        // Stale snapshot: Word (or anything) changed the docx since the
        // snapshot was taken — import the intervening changes now.
        let currentHash = SidecarStore.sha256Hex(of: try Data(contentsOf: url))
        if currentHash != snapshot.docxSHA256 {
            try orchestrator.importFromDisk(policy: policy)
        }
        return orchestrator
    }

    // MARK: - Watcher (task 4.5 integration)

    /// Polls the docx for a real content change (mtime fast-path + SHA-256).
    public func checkForExternalChange() throws -> Bool {
        try changeDetector.poll()
    }

    /// True while Word's `~$` owner file is present next to the docx.
    public var isLockedByWord: Bool {
        WordLock.isLockedByWord(docxURL)
    }

    // MARK: - Import (Word → log)

    /// Reads the docx from disk, diffs it against the baseline, resolves
    /// conflicts per `policy`, appends the surviving operations to the log
    /// with `source: "word"`, materializes them onto the in-memory tree,
    /// and advances the baseline. Returns the appended operations.
    @discardableResult
    public func importFromDisk(policy: SyncPolicy = .abortOnConflict) throws -> [Operation] {
        let onDisk = try DocxReader.read(from: docxURL, wireTreeBackedViews: false)
        guard let diskTree = onDisk.xmlTrees["word/document.xml"] else {
            throw SyncError.missingDocumentTree(partPath: "word/document.xml")
        }

        let diff = WordImport.diff(snapshot: baselineDocumentTree, current: diskTree)
        let pending = Array(document.operationLog.entries.dropFirst(flushedOpCount))
        let resolved = try SyncConflict.resolve(
            wordOps: diff.operations, pendingSwiftOps: pending, policy: policy)

        if !resolved.isEmpty {
            try document.appendAndMaterialize(resolved, source: .word)
            document.resyncBodyFromDocumentTree()
            try SidecarStore.saveLog(document.operationLog, alongside: docxURL)
        }

        // Disk state becomes the new diff baseline; the detector's baseline
        // is already advanced by the poll (or re-anchored here for direct
        // importFromDisk calls without a preceding poll).
        baselineDocumentTree = diskTree.deepCopy()
        changeDetector = try DocxChangeDetector(url: docxURL)
        return resolved
    }

    // MARK: - Flush (log → docx)

    /// Serializes the in-memory state to the docx and refreshes both
    /// sidecars. Refuses while Word holds the file open (spec scenario
    /// "Swift write while Word holds lock").
    public func flush() throws {
        if isLockedByWord {
            throw SyncError.fileLockedByWord(lockURL: WordLock.lockFileURL(for: docxURL))
        }
        try document.saveWithSidecars(to: docxURL)
        flushedOpCount = document.operationLog.entries.count
        if let tree = document.xmlTrees["word/document.xml"] {
            baselineDocumentTree = tree.deepCopy()
        }
        // Our own write must not read back as an external change.
        changeDetector = try DocxChangeDetector(url: docxURL)
    }

    // MARK: - Swift mutations through the orchestrator

    /// Convenience: the task-3.15 typed setter, applied to the
    /// orchestrator-owned document (pending until `flush()`).
    public func setParagraphText(id: ElementID, _ text: String) throws {
        try document.setParagraphText(id: id, text)
    }
}
