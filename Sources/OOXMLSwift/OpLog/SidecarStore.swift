// SidecarStore.swift
// word-aligned-state-sync Phase 2 task 3.16 — sidecar file management.
//
// "Decision 5: Sidecar persistence, not in-document metadata": Word strips
// unknown attributes and unknown parts on save, so sync state cannot live
// inside the docx. The operation log and the last-synced snapshot are
// persisted as plain files next to the docx:
//
//   report.docx            ← canonical OOXML, zero sync metadata
//   report.oplog.jsonl     ← canonical edit history (append-friendly JSONL)
//   report.snapshot.json   ← last-synced state marker (Phase 3 diff baseline)
//
// Naming follows the `ooxml-word-sync` spec's stem convention
// (`<docx-stem>.oplog.jsonl`), NOT `<full-name>.oplog.jsonl`.
//
// Sidecars are strictly opt-in (design Open Question Q1 working answer):
// `DocxWriter.write` / `DocxReader.read` never touch them; callers use
// `WordDocument.saveWithSidecars(to:)` / `WordDocument.openWithSidecars(from:)`.

import Foundation
import CryptoKit

/// Last-synced state marker persisted as `<stem>.snapshot.json`.
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

    public init(docxSHA256: String, savedAt: Date, opCount: Int,
                partFingerprints: [String: String], documentXML: String? = nil) {
        self.docxSHA256 = docxSHA256
        self.savedAt = savedAt
        self.opCount = opCount
        self.partFingerprints = partFingerprints
        self.documentXML = documentXML
    }
}

/// Path derivation + load/save for the two sidecar files.
public enum SidecarStore {

    /// `/dir/report.docx` → `/dir/report.oplog.jsonl`
    public static func oplogURL(for docxURL: URL) -> URL {
        docxURL.deletingPathExtension().appendingPathExtension("oplog.jsonl")
    }

    /// `/dir/report.docx` → `/dir/report.snapshot.json`
    public static func snapshotURL(for docxURL: URL) -> URL {
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
        let url = oplogURL(for: docxURL)
        guard FileManager.default.fileExists(atPath: url.path) else { return nil }
        return try OperationLog.decodeJSONL(try Data(contentsOf: url))
    }

    // MARK: - Snapshot sidecar

    public static func saveSnapshot(_ snapshot: SyncSnapshot, alongside docxURL: URL) throws {
        let encoder = JSONEncoder()
        encoder.outputFormatting = [.prettyPrinted, .sortedKeys]
        encoder.dateEncodingStrategy = .iso8601
        try encoder.encode(snapshot).write(to: snapshotURL(for: docxURL), options: .atomic)
    }

    /// `nil` when the snapshot sidecar does not exist; malformed JSON throws.
    public static func loadSnapshot(alongside docxURL: URL) throws -> SyncSnapshot? {
        let url = snapshotURL(for: docxURL)
        guard FileManager.default.fileExists(atPath: url.path) else { return nil }
        let decoder = JSONDecoder()
        decoder.dateDecodingStrategy = .iso8601
        return try decoder.decode(SyncSnapshot.self, from: try Data(contentsOf: url))
    }

    // MARK: - Hashing

    /// SHA-256 hex digest (CryptoKit — native framework per
    /// native-macos-compat).
    public static func sha256Hex(of data: Data) -> String {
        SHA256.hash(data: data).map { String(format: "%02x", $0) }.joined()
    }
}

extension WordDocument {

    /// Opt-in sidecar save: writes the docx (identical bytes to a plain
    /// `DocxWriter.write`), then persists the op log and a fresh snapshot
    /// alongside it. The docx itself carries zero sync metadata.
    public func saveWithSidecars(to url: URL) throws {
        // 7.3 verify P2 (torn-write window): capture pre-state and roll all
        // three files back on any failure. Absent -> nil (rollback removes
        // the fresh file); present-but-unreadable -> throw BEFORE any write.
        let fm = FileManager.default
        let targets = [url, SidecarStore.oplogURL(for: url), SidecarStore.snapshotURL(for: url)]
        let backups: [Data?] = try targets.map { target in
            guard fm.fileExists(atPath: target.path) else { return nil }
            return try Data(contentsOf: target)
        }
        do {
            try saveWithSidecarsBody(to: url)
        } catch {
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

    private func saveWithSidecarsBody(to url: URL) throws {
        try DocxWriter.write(self, to: url)
        try SidecarStore.saveLog(operationLog, alongside: url)

        let docxData = try Data(contentsOf: url)
        var fingerprints: [String: String] = [:]
        for (partPath, tree) in xmlTrees {
            fingerprints[partPath] = tree.root.normalizedFingerprint()
        }
        var documentXML: String?
        if let docTree = xmlTrees["word/document.xml"],
           let serialized = try? XmlTreeWriter.serialize(docTree) {
            documentXML = String(decoding: serialized, as: UTF8.self)
        }
        let snapshot = SyncSnapshot(
            docxSHA256: SidecarStore.sha256Hex(of: docxData),
            savedAt: Date(),
            opCount: operationLog.entries.count,
            partFingerprints: fingerprints,
            documentXML: documentXML)
        try SidecarStore.saveSnapshot(snapshot, alongside: url)
    }

    /// Opt-in sidecar open: reads the docx and, when a log sidecar exists,
    /// restores it onto `operationLog`. Absent sidecars mean fresh start
    /// (empty log) — never an error (`bootstrapFromDocx` semantics).
    public static func openWithSidecars(
        from url: URL, wireTreeBackedViews: Bool = false
    ) throws -> WordDocument {
        var document = try DocxReader.read(from: url, wireTreeBackedViews: wireTreeBackedViews)
        if let log = try SidecarStore.loadLog(alongside: url) {
            document.operationLog = log
        }
        return document
    }
}
