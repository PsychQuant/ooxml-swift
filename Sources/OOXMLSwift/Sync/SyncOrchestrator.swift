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

    /// Byte-level baseline for every package part. Optional only while
    /// upgrading a pre-part-hash snapshot; the first successful import or
    /// flush refreshes it.
    private var baselinePartSHA256: [String: String]?
    /// SHA-256 of the exact docx generation this session has accepted.
    /// `flush()` compares it immediately before replacement so an external
    /// save cannot be silently overwritten.
    private var baselineDocxSHA256: String

    /// Swift entries present in the canonical log/in-memory tree but not yet
    /// materialized in the docx. IDs are required because imported Word ops
    /// can follow a pending Swift op while already being present on disk.
    private var pendingSwiftOpIDs: Set<UUID>

    private var changeDetector: DocxChangeDetector
    private var isClosed = false

    /// Releases the reader-owned extracted package. Idempotent; callers may
    /// close a session explicitly, while `deinit` remains the safety net.
    public func close() {
        guard !isClosed else { return }
        document.close()
        isClosed = true
    }

    deinit {
        document.close()
    }

    private init(document: WordDocument, docxURL: URL,
                 baseline: XmlTree, baselinePartSHA256: [String: String]?,
                 baselineDocxSHA256: String,
                 pendingSwiftOpIDs: Set<UUID>,
                 changeDetector: DocxChangeDetector) {
        self.document = document
        self.docxURL = docxURL
        self.baselineDocumentTree = baseline
        self.baselinePartSHA256 = baselinePartSHA256
        self.baselineDocxSHA256 = baselineDocxSHA256
        self.pendingSwiftOpIDs = pendingSwiftOpIDs
        self.changeDetector = changeDetector
    }

    private func requireOpen() throws {
        if isClosed { throw SyncError.sessionClosed }
    }

    private func requireDiskGeneration(_ expectedSHA256: String) throws {
        if isLockedByWord {
            throw SyncError.fileLockedByWord(
                lockURL: WordLock.lockFileURL(for: docxURL))
        }
        let actual = SidecarStore.sha256Hex(of: try Data(contentsOf: docxURL))
        guard actual == expectedSHA256 else {
            throw SyncError.externalGenerationChanged(
                expectedSHA256: expectedSHA256, actualSHA256: actual)
        }
    }

    private static func requireDiskGeneration(
        at url: URL, expectedSHA256: String
    ) throws {
        if WordLock.isLockedByWord(url) {
            throw SyncError.fileLockedByWord(
                lockURL: WordLock.lockFileURL(for: url))
        }
        let actual = SidecarStore.sha256Hex(of: try Data(contentsOf: url))
        guard actual == expectedSHA256 else {
            throw SyncError.externalGenerationChanged(
                expectedSHA256: expectedSHA256, actualSHA256: actual)
        }
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
        if WordLock.isLockedByWord(url) {
            throw SyncError.fileLockedByWord(
                lockURL: WordLock.lockFileURL(for: url))
        }
        let bootstrapData = try Data(contentsOf: url)
        let currentDocxSHA256 = SidecarStore.sha256Hex(of: bootstrapData)
        let generationDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent("ooxml-sync-bootstrap-\(UUID().uuidString)")
        try FileManager.default.createDirectory(
            at: generationDirectory, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: generationDirectory) }
        let generationURL = generationDirectory.appendingPathComponent("generation.docx")
        try bootstrapData.write(to: generationURL)

        var document = try DocxReader.read(
            from: generationURL, wireTreeBackedViews: true)
        var transferredDocument = false
        defer {
            if !transferredDocument { document.close() }
        }
        guard let currentTree = document.xmlTrees["word/document.xml"] else {
            throw SyncError.missingDocumentTree(partPath: "word/document.xml")
        }

        let existingState = try SidecarStore.loadSyncState(alongside: url)
        let existingLog = existingState.log
        let existingSnapshot = existingState.snapshot

        if let log = existingLog { document.operationLog = log }

        let detector = try DocxChangeDetector(url: url)
        try requireDiskGeneration(at: url, expectedSHA256: currentDocxSHA256)

        guard let snapshot = existingSnapshot else {
            // Fresh start: sidecars created with the docx's current state.
            let orchestrator = SyncOrchestrator(
                document: document, docxURL: url,
                baseline: currentTree.deepCopy(),
                baselinePartSHA256: SidecarStore.partSHA256(
                    try RawPartChannel.readAllParts(from: generationURL)),
                baselineDocxSHA256: currentDocxSHA256,
                pendingSwiftOpIDs: [],
                changeDetector: detector)
            try document.saveWithSidecars(
                to: url, expectedDocxSHA256: currentDocxSHA256)
            // Re-baseline the detector: saveWithSidecars rewrote the docx.
            orchestrator.changeDetector = try DocxChangeDetector(url: url)
            orchestrator.baselinePartSHA256 = SidecarStore.partSHA256(
                try RawPartChannel.readAllParts(from: url))
            orchestrator.baselineDocxSHA256 = SidecarStore.sha256Hex(
                of: try Data(contentsOf: url))
            transferredDocument = true
            return orchestrator
        }

        // Existing sidecars. Baseline = snapshot's stored document.xml when
        // present (true last-synced state), else the current tree.
        var baseline = currentTree.deepCopy()
        if let storedXML = snapshot.documentXML,
           let parsed = try? XmlTreeReader.parse(Data(storedXML.utf8)) {
            baseline = parsed
        }

        let pendingIDs = Set(snapshot.pendingSwiftOpIDs ?? [])
        let pendingEntries = document.operationLog.entries.filter {
            pendingIDs.contains($0.opID) && $0.source == .swift
        }
        let restoredPendingIDs = Set(pendingEntries.map(\.opID))
        let missingPendingIDs = pendingIDs.subtracting(restoredPendingIDs)
        guard missingPendingIDs.isEmpty else {
            throw SyncError.missingPendingOperations(opIDs: missingPendingIDs.sorted {
                $0.uuidString < $1.uuidString
            })
        }
        if !pendingEntries.isEmpty, currentDocxSHA256 != snapshot.docxSHA256 {
            // The live docx contains Word edits newer than the snapshot.
            // Reconstruct canonical memory from the snapshot baseline first;
            // otherwise importing the stale delta onto the already-new disk
            // tree duplicates non-idempotent inserts.
            document.xmlTrees["word/document.xml"] = baseline.deepCopy()
            document.resyncBodyFromDocumentTree()
        }
        if !pendingEntries.isEmpty {
            try document.appendAndMaterialize(
                pendingEntries.map(\.op),
                source: .swift,
                replayOpIDs: pendingEntries.map(\.opID),
                replayTimestamps: pendingEntries.map(\.timestamp),
                appendToLog: false)
            document.resyncBodyFromDocumentTree()
        }

        let orchestrator = SyncOrchestrator(
            document: document, docxURL: url,
            baseline: baseline,
            baselinePartSHA256: snapshot.partSHA256,
            baselineDocxSHA256: currentDocxSHA256,
            pendingSwiftOpIDs: pendingIDs,
            changeDetector: detector)

        // Stale snapshot: Word (or anything) changed the docx since the
        // snapshot was taken — import the intervening changes now.
        if currentDocxSHA256 != snapshot.docxSHA256 {
            try orchestrator.importFromDisk(policy: policy)
        }
        transferredDocument = true
        return orchestrator
    }

    // MARK: - Watcher (task 4.5 integration)

    /// Polls the docx for a real content change (mtime fast-path + SHA-256).
    public func checkForExternalChange() throws -> Bool {
        try requireOpen()
        return try changeDetector.poll()
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
        try importFromDisk(
            policy: policy, immediatelyBeforeSnapshotWrite: nil)
    }

    /// Test injection is closure-valued and passed per call; production sync
    /// has no global mutable hook. It lets transaction tests force the second
    /// sidecar write to fail deterministically.
    @discardableResult
    internal func importFromDisk(
        policy: SyncPolicy,
        immediatelyBeforeSnapshotWrite: (() throws -> Void)?
    ) throws -> [Operation] {
        try requireOpen()
        if isLockedByWord {
            throw SyncError.fileLockedByWord(lockURL: WordLock.lockFileURL(for: docxURL))
        }
        // Parse and hash one immutable byte generation. Reading the live path
        // separately for the tree and raw parts can otherwise mix two Word
        // saves into one baseline.
        let generationData = try Data(contentsOf: docxURL)
        let generationSHA256 = SidecarStore.sha256Hex(of: generationData)
        let generationDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent("ooxml-sync-generation-\(UUID().uuidString)")
        try FileManager.default.createDirectory(
            at: generationDirectory, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: generationDirectory) }
        let generationURL = generationDirectory.appendingPathComponent("generation.docx")
        try generationData.write(to: generationURL)

        var onDisk = try DocxReader.read(from: generationURL, wireTreeBackedViews: false)
        var adoptedOnDisk = false
        defer { if !adoptedOnDisk { onDisk.close() } }
        guard let diskTree = onDisk.xmlTrees["word/document.xml"] else {
            throw SyncError.missingDocumentTree(partPath: "word/document.xml")
        }

        let diff = WordImport.diff(snapshot: baselineDocumentTree, current: diskTree)
        let pending = document.operationLog.entries.filter {
            pendingSwiftOpIDs.contains($0.opID) && $0.source == .swift
        }
        let resolved = try SyncConflict.resolve(
            wordOps: diff.operations, pendingSwiftOps: pending, policy: policy)

        // Prove whether the semantic paragraph ops reproduce the complete
        // normalized document tree. Anything outside that surface (formatting,
        // tables, sections, ID-less edits, vendor nodes) rides a raw part op.
        var semanticCoversDocument = baselineDocumentTree.root.normalizedFingerprint()
            == diskTree.root.normalizedFingerprint()
        if !diff.operations.isEmpty {
            var trialLog = OperationLog()
            for op in diff.operations { trialLog.append(op, source: .word) }
            if let trial = try? OperationReducer.materialize(
                log: trialLog, base: baselineDocumentTree.deepCopy()) {
                semanticCoversDocument = trial.root.normalizedFingerprint()
                    == diskTree.root.normalizedFingerprint()
            }
        }

        let diskParts = try RawPartChannel.readAllParts(from: generationURL)
        let diskPartSHA256 = SidecarStore.partSHA256(diskParts)
        let changedSiblingPaths: [String]
        if let baselinePartSHA256 {
            let allPaths = Set(baselinePartSHA256.keys).union(diskPartSHA256.keys)
            changedSiblingPaths = allPaths.filter {
                $0 != "word/document.xml"
                    && baselinePartSHA256[$0] != diskPartSHA256[$0]
            }.sorted()
        } else {
            // An old snapshot has no raw-part hashes. Preserve every current
            // sibling rather than silently treating current disk bytes as the
            // previous baseline. The next snapshot upgrades the sidecar.
            changedSiblingPaths = diskParts.keys.filter {
                $0 != "word/document.xml"
            }.sorted()
        }

        var carryOps: [Operation] = []
        if !semanticCoversDocument || diff.requiresDocumentCarry
            || !diff.unrepresentedChanges.isEmpty {
            if let bytes = diskParts["word/document.xml"],
               let xml = String(data: bytes, encoding: .utf8) {
                carryOps.append(.carryPart(partPath: "word/document.xml", xml: xml))
            }
        }
        for path in changedSiblingPaths {
            // Deletions cannot yet be represented by the raw channel. Fail
            // loudly below instead of pretending the baseline advanced.
            guard let bytes = diskParts[path] else { continue }
            if let xml = String(data: bytes, encoding: .utf8) {
                carryOps.append(.carryPart(partPath: path, xml: xml))
            } else {
                carryOps.append(.carryBinaryPart(
                    partPath: path, base64: bytes.base64EncodedString()))
            }
        }

        let deletedParts = changedSiblingPaths.filter { diskParts[$0] == nil }
        let unsafeExternalParts = carryOps.map { op -> String in
            switch op {
            case .carryPart(let path, _), .carryBinaryPart(let path, _): return path
            default: return ""
            }
        } + deletedParts
        if !deletedParts.isEmpty {
            throw SyncError.unrepresentedExternalChanges(partPaths: deletedParts)
        }
        if !pending.isEmpty && !unsafeExternalParts.isEmpty {
            throw SyncError.unrepresentedExternalChanges(
                partPaths: Array(Set(unsafeExternalParts)).sorted())
        }

        // Recheck immediately before advancing canonical state/sidecars. A
        // concurrent replace invalidates this import instead of committing a
        // tree, raw parts, and hash taken from different generations.
        try requireDiskGeneration(generationSHA256)

        var candidateDocument = document
        if pending.isEmpty {
            // No Swift state needs merging: the freshly read Word package is
            // the exact canonical state. Adopt it wholesale, append semantic
            // history plus raw operations for unrepresentable/sibling parts,
            // and retire the old preserved archive.
            let appended = resolved + carryOps
            var log = document.operationLog
            for op in appended { log.append(op, source: .word) }
            onDisk.operationLog = log
            onDisk.carriedParts = document.carriedParts
            for op in carryOps {
                switch op {
                case .carryPart(let path, let xml):
                    onDisk.carriedParts[path] = Data(xml.utf8)
                case .carryBinaryPart(let path, let base64):
                    onDisk.carriedParts[path] = Data(base64Encoded: base64)
                default:
                    break
                }
            }
            candidateDocument = onDisk
        } else {
            if !resolved.isEmpty {
                try candidateDocument.appendAndMaterialize(resolved, source: .word)
                candidateDocument.resyncBodyFromDocumentTree()
            }
            // normalizedFingerprint deliberately ignores rsid churn, but the
            // raw attributes still belong to Word's accepted generation.
            // Merge that identity noise onto the pending tree before flush;
            // otherwise a same-session import→flush silently erases it.
            if let candidateTree = candidateDocument.xmlTrees["word/document.xml"] {
                let mergedTree = candidateTree.deepCopy()
                if Self.mergeRsidNoise(from: diskTree.root, into: mergedTree.root) {
                    candidateDocument.xmlTrees["word/document.xml"] = mergedTree
                    candidateDocument.modifiedParts.insert("word/document.xml")
                    candidateDocument.treeFreshParts.insert("word/document.xml")
                    candidateDocument.resyncBodyFromDocumentTree()
                }
            }
        }

        // Build and persist the new log/snapshot generation before advancing
        // any in-memory baseline. The pair writer rolls both paths back on a
        // second-write failure, so restart never sees a new log with an old
        // snapshot (or vice versa).
        var fingerprints: [String: String] = [:]
        for (partPath, tree) in onDisk.xmlTrees {
            fingerprints[partPath] = tree.root.normalizedFingerprint()
        }
        let serializedBaseline = try? XmlTreeWriter.serialize(diskTree)
        let nextSnapshot = SyncSnapshot(
            docxSHA256: generationSHA256,
            savedAt: Date(),
            opCount: candidateDocument.operationLog.entries.count,
            partFingerprints: fingerprints,
            documentXML: serializedBaseline.map { String(decoding: $0, as: UTF8.self) },
            partSHA256: diskPartSHA256,
            pendingSwiftOpIDs: pendingSwiftOpIDs.sorted {
                $0.uuidString < $1.uuidString
            })
        try requireDiskGeneration(generationSHA256)
        let nextDetector = try DocxChangeDetector(url: docxURL)
        try SidecarStore.saveSyncState(
            log: candidateDocument.operationLog,
            snapshot: nextSnapshot,
            alongside: docxURL,
            immediatelyBeforeSnapshotWrite: immediatelyBeforeSnapshotWrite)

        if pending.isEmpty {
            document.close()
            document = candidateDocument
            adoptedOnDisk = true
        } else {
            document = candidateDocument
        }
        baselineDocumentTree = diskTree.deepCopy()
        baselinePartSHA256 = diskPartSHA256
        baselineDocxSHA256 = generationSHA256
        changeDetector = nextDetector
        return resolved + carryOps
    }

    /// Copies Word's rsid identity-noise attributes across structurally
    /// aligned nodes while leaving semantic attributes/content untouched.
    @discardableResult
    private static func mergeRsidNoise(from disk: XmlNode, into memory: XmlNode) -> Bool {
        guard disk.kind == memory.kind,
              disk.localName == memory.localName else { return false }
        let diskRsid = disk.attributes.filter(\.isRsidNoise)
        let memoryRsid = memory.attributes.filter(\.isRsidNoise)
        var changed = diskRsid != memoryRsid
        if changed {
            memory.attributes = memory.attributes.filter { !$0.isRsidNoise } + diskRsid
        }

        guard disk.kind == .element else { return changed }
        for (index, memoryChild) in memory.children.enumerated() {
            let diskChild: XmlNode?
            if index < disk.children.count,
               disk.children[index].kind == memoryChild.kind,
               disk.children[index].localName == memoryChild.localName {
                diskChild = disk.children[index]
            } else if let stableID = memoryChild.stableID {
                diskChild = disk.children.first { $0.stableID == stableID }
            } else {
                diskChild = nil
            }
            if let diskChild,
               mergeRsidNoise(from: diskChild, into: memoryChild) {
                changed = true
            }
        }
        return changed
    }

    // MARK: - Flush (log → docx)

    /// Serializes the in-memory state to the docx and refreshes both
    /// sidecars. Refuses while Word holds the file open (spec scenario
    /// "Swift write while Word holds lock").
    public func flush() throws {
        try requireOpen()
        if isLockedByWord {
            throw SyncError.fileLockedByWord(lockURL: WordLock.lockFileURL(for: docxURL))
        }
        try requireDiskGeneration(baselineDocxSHA256)
        try document.saveWithSidecars(
            to: docxURL,
            pendingSwiftOpIDs: [],
            expectedDocxSHA256: baselineDocxSHA256)
        pendingSwiftOpIDs.removeAll()
        if let tree = document.xmlTrees["word/document.xml"] {
            baselineDocumentTree = tree.deepCopy()
        }
        baselinePartSHA256 = SidecarStore.partSHA256(
            try RawPartChannel.readAllParts(from: docxURL))
        baselineDocxSHA256 = SidecarStore.sha256Hex(
            of: try Data(contentsOf: docxURL))
        // Our own write must not read back as an external change.
        changeDetector = try DocxChangeDetector(url: docxURL)
    }

    // MARK: - Swift mutations through the orchestrator

    /// Convenience: the task-3.15 typed setter, applied to the
    /// orchestrator-owned document (pending until `flush()`).
    public func setParagraphText(id: ElementID, _ text: String) throws {
        try requireOpen()
        let previousCount = document.operationLog.entries.count
        try document.setParagraphText(id: id, text)
        for entry in document.operationLog.entries.dropFirst(previousCount)
        where entry.source == .swift {
            pendingSwiftOpIDs.insert(entry.opID)
        }
    }
}
