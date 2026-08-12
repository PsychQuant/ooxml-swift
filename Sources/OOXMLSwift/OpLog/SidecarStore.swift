// SidecarStore.swift
// word-aligned-state-sync Phase 2 task 3.16 — sidecar file management.
//
// "Decision 5: Sidecar persistence, not in-document metadata": Word strips
// unknown attributes and unknown parts on save, so sync state cannot live
// inside the docx. The operation log and the last-synced snapshot are
// persisted as plain files next to the docx:
//
//   report.docx            ← canonical OOXML, zero sync metadata
//   report.docx.oplog.jsonl   ← canonical edit history (append-friendly JSONL)
//   report.docx.snapshot.json ← last-synced state marker (Phase 3 diff baseline)
//
// Naming follows the `mdocx-grammar` full-name convention.
//
// Sidecars are strictly opt-in (design Open Question Q1 working answer):
// `DocxWriter.write` / `DocxReader.read` never touch them; callers use
// `WordDocument.saveWithSidecars(to:)` / `WordDocument.openWithSidecars(from:)`.

import Foundation
import CryptoKit

public enum SidecarStoreError: Error, Equatable {
    /// Neither naming generation contains a complete log/snapshot pair.
    /// Mixing one file from each generation would corrupt history/baseline
    /// alignment, so sync bootstrap stops instead.
    case incompleteSidecarPair
    /// Both sidecar files exist, but they describe different saved history
    /// generations (for example, a process stopped after replacing the log
    /// and before replacing the snapshot).
    case sidecarGenerationMismatch(logEntryCount: Int, snapshotOpCount: Int)
}

/// Last-synced state marker persisted as `<name>.docx.snapshot.json`.
///
/// Phase 2 records enough for Phase 3's `SyncOrchestrator` to decide
/// whether the docx changed since the last Swift-side sync (content hash)
/// and which parts changed (normalized fingerprints). The full baseline
/// tree for deep diffing is the docx itself at snapshot time — Phase 3
/// task 4.7 extends this shape if the diff needs more.
public struct SyncSnapshot: Codable, Equatable {
    /// SHA-256 (hex) of the docx bytes as written at snapshot time.
    public let docxSHA256: String
    /// Wall-clock time of the snapshot.
    public let savedAt: Date
    /// Number of entries in the op log at snapshot time.
    public let opCount: Int
    /// `normalizedFingerprint()` per part path, for cheap changed-part
    /// detection on import (identity noise like rsids already excluded).
    public let partFingerprints: [String: String]
    /// Serialized `word/document.xml` at snapshot time — the baseline tree
    /// `SyncOrchestrator` diffs Word saves against across sessions
    /// (Phase 3 task 4.7). Optional for backward compatibility with
    /// pre-Phase-3 snapshot files.
    public let documentXML: String?
    /// SHA-256 per raw package part. Unlike `partFingerprints`, this covers
    /// binary parts and byte-level XML changes, allowing a new sync session
    /// to identify Word edits that happened after the previous snapshot.
    /// Optional for backward compatibility with older snapshot files.
    public let partSHA256: [String: String]?
    /// Swift-originated log entries that are present in the canonical log but
    /// not yet materialized in the docx bytes. This is an ID set rather than
    /// a prefix count because an imported Word op can follow a pending Swift
    /// op while already being present on disk.
    public let pendingSwiftOpIDs: [UUID]?

    public init(docxSHA256: String, savedAt: Date, opCount: Int,
                partFingerprints: [String: String], documentXML: String? = nil,
                partSHA256: [String: String]? = nil,
                pendingSwiftOpIDs: [UUID]? = nil) {
        self.docxSHA256 = docxSHA256
        self.savedAt = savedAt
        self.opCount = opCount
        self.partFingerprints = partFingerprints
        self.documentXML = documentXML
        self.partSHA256 = partSHA256
        self.pendingSwiftOpIDs = pendingSwiftOpIDs
    }
}

/// Path derivation + load/save for the two sidecar files.
public enum SidecarStore {

    /// `/dir/report.docx` → `/dir/report.docx.oplog.jsonl`
    public static func oplogURL(for docxURL: URL) -> URL {
        docxURL.appendingPathExtension("oplog.jsonl")
    }

    /// `/dir/report.docx` → `/dir/report.docx.snapshot.json`
    public static func snapshotURL(for docxURL: URL) -> URL {
        docxURL.appendingPathExtension("snapshot.json")
    }

    /// Pre-full-name convention used by earlier releases. New saves never
    /// write these paths, but loads accept them so an upgrade does not make
    /// existing edit history disappear.
    private static func legacyOplogURL(for docxURL: URL) -> URL {
        docxURL.deletingPathExtension().appendingPathExtension("oplog.jsonl")
    }

    private static func legacySnapshotURL(for docxURL: URL) -> URL {
        docxURL.deletingPathExtension().appendingPathExtension("snapshot.json")
    }

    // MARK: - Operation log sidecar

    /// Writes the full log as JSONL next to the docx (atomic replace).
    ///
    /// Whole-file rewrite rather than `O_APPEND` incremental append: Phase 2
    /// saves happen at document-save granularity where the in-memory log is
    /// the source of truth. Phase 3's live `SyncOrchestrator` adds the
    /// incremental append path when it owns a long-running session.
    public static func saveLog(_ log: OperationLog, alongside docxURL: URL) throws {
        try log.encodeJSONL().write(to: oplogURL(for: docxURL), options: .atomic)
    }

    /// Loads the log sidecar. `nil` when the file does not exist
    /// (fresh-start semantics — absence is not an error). Malformed content
    /// throws (loud, per apply-errors-are-reported discipline).
    public static func loadLog(alongside docxURL: URL) throws -> OperationLog? {
        let canonical = oplogURL(for: docxURL)
        let legacy = legacyOplogURL(for: docxURL)
        let url = FileManager.default.fileExists(atPath: canonical.path)
            ? canonical : legacy
        guard FileManager.default.fileExists(atPath: url.path) else { return nil }
        return try OperationLog.decodeJSONL(try Data(contentsOf: url))
    }

    // MARK: - Snapshot sidecar

    public static func saveSnapshot(_ snapshot: SyncSnapshot, alongside docxURL: URL) throws {
        try encodedSnapshot(snapshot).write(
            to: snapshotURL(for: docxURL), options: .atomic)
    }

    private static func encodedSnapshot(_ snapshot: SyncSnapshot) throws -> Data {
        let encoder = JSONEncoder()
        encoder.outputFormatting = [.prettyPrinted, .sortedKeys]
        encoder.dateEncodingStrategy = .iso8601
        return try encoder.encode(snapshot)
    }

    /// `nil` when the snapshot sidecar does not exist; malformed JSON throws.
    public static func loadSnapshot(alongside docxURL: URL) throws -> SyncSnapshot? {
        let canonical = snapshotURL(for: docxURL)
        let legacy = legacySnapshotURL(for: docxURL)
        let url = FileManager.default.fileExists(atPath: canonical.path)
            ? canonical : legacy
        guard FileManager.default.fileExists(atPath: url.path) else { return nil }
        let decoder = JSONDecoder()
        decoder.dateDecodingStrategy = .iso8601
        return try decoder.decode(SyncSnapshot.self, from: try Data(contentsOf: url))
    }

    /// Loads log + snapshot from one naming generation. Canonical wins only
    /// when both files exist; otherwise a complete legacy pair is selected.
    /// Half-written/mixed generations fail loudly instead of pairing history
    /// from one save with a baseline from another.
    internal static func loadSyncState(
        alongside docxURL: URL
    ) throws -> (log: OperationLog?, snapshot: SyncSnapshot?) {
        let fm = FileManager.default
        let canonical = (oplogURL(for: docxURL), snapshotURL(for: docxURL))
        let legacy = (legacyOplogURL(for: docxURL), legacySnapshotURL(for: docxURL))
        let canonicalExists = (
            fm.fileExists(atPath: canonical.0.path),
            fm.fileExists(atPath: canonical.1.path))
        let legacyExists = (
            fm.fileExists(atPath: legacy.0.path),
            fm.fileExists(atPath: legacy.1.path))

        let selected: (URL, URL)?
        if canonicalExists.0 && canonicalExists.1 {
            selected = canonical
        } else if legacyExists.0 && legacyExists.1 {
            selected = legacy
        } else if !canonicalExists.0 && !canonicalExists.1
                    && !legacyExists.0 && !legacyExists.1 {
            selected = nil
        } else {
            throw SidecarStoreError.incompleteSidecarPair
        }
        guard let selected else { return (nil, nil) }
        let log = try OperationLog.decodeJSONL(try Data(contentsOf: selected.0))
        let decoder = JSONDecoder()
        decoder.dateDecodingStrategy = .iso8601
        let snapshot = try decoder.decode(
            SyncSnapshot.self, from: try Data(contentsOf: selected.1))
        guard log.entries.count == snapshot.opCount else {
            throw SidecarStoreError.sidecarGenerationMismatch(
                logEntryCount: log.entries.count,
                snapshotOpCount: snapshot.opCount)
        }
        return (log, snapshot)
    }

    /// Atomically-at-API-level replaces the two sync sidecars. If either
    /// write fails, both paths are restored to their exact pre-call bytes.
    internal static func saveSyncState(
        log: OperationLog,
        snapshot: SyncSnapshot,
        alongside docxURL: URL,
        immediatelyBeforeSnapshotWrite: (() throws -> Void)? = nil
    ) throws {
        let fm = FileManager.default
        let targets = [oplogURL(for: docxURL), snapshotURL(for: docxURL)]
        let backups: [Data?] = try targets.map { target in
            guard fm.fileExists(atPath: target.path) else { return nil }
            return try Data(contentsOf: target)
        }
        do {
            try log.encodeJSONL().write(to: targets[0], options: .atomic)
            try immediatelyBeforeSnapshotWrite?()
            try encodedSnapshot(snapshot).write(to: targets[1], options: .atomic)
        } catch {
            for (index, target) in targets.enumerated() {
                if let backup = backups[index] {
                    try? backup.write(to: target, options: .atomic)
                } else {
                    try? fm.removeItem(at: target)
                }
            }
            throw error
        }
    }

    // MARK: - Hashing

    /// SHA-256 hex digest (CryptoKit — native framework per
    /// native-macos-compat).
    public static func sha256Hex(of data: Data) -> String {
        SHA256.hash(data: data).map { String(format: "%02x", $0) }.joined()
    }

    public static func partSHA256(_ parts: [String: Data]) -> [String: String] {
        parts.mapValues(sha256Hex(of:))
    }
}

extension WordDocument {

    /// Opt-in sidecar save: writes the docx (identical bytes to a plain
    /// `DocxWriter.write`), then persists the op log and a fresh snapshot
    /// alongside it. The docx itself carries zero sync metadata.
    public func saveWithSidecars(
        to url: URL, pendingSwiftOpIDs: [UUID] = [],
        expectedDocxSHA256: String? = nil
    ) throws {
        try saveWithSidecars(
            to: url,
            pendingSwiftOpIDs: pendingSwiftOpIDs,
            expectedDocxSHA256: expectedDocxSHA256,
            afterBackupsCaptured: nil,
            immediatelyBeforeGenerationCheck: nil,
            immediatelyAfterDocxWrite: nil)
    }

    /// Injection points are internal and closure-valued to keep production
    /// state concurrency-safe while allowing deterministic TOCTOU tests.
    internal func saveWithSidecars(
        to url: URL,
        pendingSwiftOpIDs: [UUID],
        expectedDocxSHA256: String?,
        afterBackupsCaptured: (() throws -> Void)?,
        immediatelyBeforeGenerationCheck: (() throws -> Void)?,
        immediatelyAfterDocxWrite: (() throws -> Void)? = nil
    ) throws {
        // Reject a generation already known stale before capturing rollback
        // bytes. A precondition failure performs no write and must never
        // restore an older target over an external save.
        if let expectedDocxSHA256 {
            let actual = SidecarStore.sha256Hex(of: try Data(contentsOf: url))
            guard actual == expectedDocxSHA256 else {
                throw SyncError.externalGenerationChanged(
                    expectedSHA256: expectedDocxSHA256,
                    actualSHA256: actual)
            }
        }
        // 7.3 verify P2 (torn-write window): capture pre-state and roll all
        // three files back on any failure. Absent -> nil (rollback removes
        // the fresh file); present-but-unreadable -> throw BEFORE any write.
        let fm = FileManager.default
        let targets = [url, SidecarStore.oplogURL(for: url), SidecarStore.snapshotURL(for: url)]
        let backups: [Data?] = try targets.map { target in
            guard fm.fileExists(atPath: target.path) else { return nil }
            return try Data(contentsOf: target)
        }
        try afterBackupsCaptured?()
        var docxWriteCompleted = false
        var writtenDocxSHA256: String?
        do {
            try saveWithSidecarsBody(
                to: url,
                pendingSwiftOpIDs: pendingSwiftOpIDs,
                expectedDocxSHA256: expectedDocxSHA256,
                immediatelyBeforeGenerationCheck: immediatelyBeforeGenerationCheck,
                immediatelyAfterDocxWrite: immediatelyAfterDocxWrite,
                docxWriteCompleted: &docxWriteCompleted,
                writtenDocxSHA256: &writtenDocxSHA256)
        } catch {
            // DocxWriter guarantees the target is untouched until its atomic
            // rename. A generation/lock rejection before that point must not
            // roll old backups over a newly-written external generation.
            if docxWriteCompleted {
                for (i, target) in targets.enumerated() {
                    if i == 0, let writtenDocxSHA256,
                       let current = try? Data(contentsOf: target),
                       SidecarStore.sha256Hex(of: current) != writtenDocxSHA256 {
                        // Another writer replaced the docx after our rename.
                        // Never roll the older backup over that newer
                        // generation; sidecars can still be restored below.
                        continue
                    }
                    if let backup = backups[i] {
                        try? backup.write(to: target)
                    } else {
                        try? fm.removeItem(at: target)
                    }
                }
            }
            throw error
        }
    }

    private func saveWithSidecarsBody(
        to url: URL, pendingSwiftOpIDs: [UUID],
        expectedDocxSHA256: String?,
        immediatelyBeforeGenerationCheck: (() throws -> Void)?,
        immediatelyAfterDocxWrite: (() throws -> Void)?,
        docxWriteCompleted: inout Bool,
        writtenDocxSHA256: inout String?
    ) throws {
        let ownGenerationSHA256 = try DocxWriter.write(
            self,
            to: url,
            expectedDocxSHA256: expectedDocxSHA256,
            immediatelyBeforeGenerationCheck: immediatelyBeforeGenerationCheck)
        docxWriteCompleted = true
        writtenDocxSHA256 = ownGenerationSHA256
        try immediatelyAfterDocxWrite?()
        let currentSHA256 = SidecarStore.sha256Hex(of: try Data(contentsOf: url))
        guard currentSHA256 == ownGenerationSHA256 else {
            throw SyncError.externalGenerationChanged(
                expectedSHA256: ownGenerationSHA256,
                actualSHA256: currentSHA256)
        }
        try SidecarStore.saveLog(operationLog, alongside: url)

        let docxData = try Data(contentsOf: url)
        let observedSHA256 = SidecarStore.sha256Hex(of: docxData)
        guard observedSHA256 == ownGenerationSHA256 else {
            throw SyncError.externalGenerationChanged(
                expectedSHA256: ownGenerationSHA256,
                actualSHA256: observedSHA256)
        }
        var fingerprints: [String: String] = [:]
        for (partPath, tree) in xmlTrees {
            fingerprints[partPath] = tree.root.normalizedFingerprint()
        }
        var documentXML: String?
        if let docTree = xmlTrees["word/document.xml"],
           let serialized = try? XmlTreeWriter.serialize(docTree) {
            documentXML = String(decoding: serialized, as: UTF8.self)
        }
        let partSHA256 = SidecarStore.partSHA256(
            try RawPartChannel.readAllParts(from: url))
        let finalSHA256 = SidecarStore.sha256Hex(
            of: try Data(contentsOf: url))
        guard finalSHA256 == ownGenerationSHA256 else {
            throw SyncError.externalGenerationChanged(
                expectedSHA256: ownGenerationSHA256,
                actualSHA256: finalSHA256)
        }
        let snapshot = SyncSnapshot(
            docxSHA256: ownGenerationSHA256,
            savedAt: Date(),
            opCount: operationLog.entries.count,
            partFingerprints: fingerprints,
            documentXML: documentXML,
            partSHA256: partSHA256,
            pendingSwiftOpIDs: pendingSwiftOpIDs)
        try SidecarStore.saveSnapshot(snapshot, alongside: url)
    }

    /// Opt-in sidecar open: reads the docx and, when a log sidecar exists,
    /// restores it onto `operationLog`. Absent sidecars mean fresh start
    /// (empty log) — never an error (`bootstrapFromDocx` semantics).
    public static func openWithSidecars(
        from url: URL, wireTreeBackedViews: Bool = false
    ) throws -> WordDocument {
        var document = try DocxReader.read(from: url, wireTreeBackedViews: wireTreeBackedViews)
        var transferredDocument = false
        defer {
            if !transferredDocument { document.close() }
        }
        let state = try SidecarStore.loadSyncState(alongside: url)
        if let log = state.log {
            document.operationLog = log
        }
        transferredDocument = true
        return document
    }
}
