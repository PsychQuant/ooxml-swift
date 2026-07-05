// SyncPolicy.swift
// word-aligned-state-sync Phase 3 tasks 4.3 + 4.4 — conflict detection and
// typed conflict policy ("Decision 7: Conflict policy is opt-in and
// explicit"; `ooxml-word-sync` Requirements "Conflict detection on
// overlapping mutations" + "Typed conflict policy").

import Foundation

/// Per-element resolution returned by an `.askUser` handler.
public enum ConflictResolution: Equatable, Sendable {
    /// Keep Swift's pending operation; drop the Word-inferred op for this element.
    case takeSwift
    /// Keep the Word-inferred op (appended after Swift's — last write wins on replay).
    case takeWord
}

/// Conflict policy for `SyncOrchestrator` imports. The default everywhere is
/// `.abortOnConflict` — no silent data loss; merging is opt-in.
public enum SyncPolicy {
    /// Throw `SyncError.conflict(report:)`; the log is not modified.
    case abortOnConflict
    /// Keep non-conflicting Word ops; drop Word ops that touch elements with
    /// pending Swift mutations.
    case swiftWins
    /// Keep every Word op — appended after Swift's pending ops, so replay
    /// order makes Word's value win for touched elements.
    case wordWins
    /// Invoke the handler synchronously with the full report; the handler
    /// returns a per-element `ConflictResolution`. Elements missing from the
    /// returned map default to `.takeSwift` (the conservative side).
    case askUser(handler: (ConflictReport) -> [ElementID: ConflictResolution])
}

/// Structured error surface for sync operations.
public enum SyncError: Error {
    /// Import found overlapping mutations and the policy was `.abortOnConflict`.
    case conflict(report: ConflictReport)
    /// A Swift write was refused because Word holds the docx open
    /// (`~$` owner file present). Recovery: close Word, retry.
    case fileLockedByWord(lockURL: URL)
    /// `bootstrapFromDocx` / import could not load a required part tree.
    case missingDocumentTree(partPath: String)
}

/// One import's worth of overlapping-mutation findings.
public struct ConflictReport: Equatable {
    public struct Entry: Equatable {
        /// The element both writers touched.
        public let elementID: ElementID
        /// `opID` of the pending (unflushed) Swift-originated log entry.
        public let swiftOpID: UUID
        /// Swift's pending operation for the element.
        public let swiftOp: Operation
        /// The Word-inferred operation for the same element.
        public let wordOp: Operation

        public init(elementID: ElementID, swiftOpID: UUID,
                    swiftOp: Operation, wordOp: Operation) {
            self.elementID = elementID
            self.swiftOpID = swiftOpID
            self.swiftOp = swiftOp
            self.wordOp = wordOp
        }
    }

    public var entries: [Entry]
    public init(entries: [Entry] = []) { self.entries = entries }
}

public enum SyncConflict {

    /// Detects overlaps: Word-inferred ops whose target element is also
    /// targeted by a pending (not-yet-flushed) Swift-originated log entry.
    public static func detect(
        pendingSwiftOps: [LogEntry], wordOps: [Operation]
    ) -> ConflictReport {
        var swiftByTarget: [ElementID: LogEntry] = [:]
        for entry in pendingSwiftOps where entry.source == .swift {
            for id in OperationReducer.referencedElementIDs(in: entry.op)
            where swiftByTarget[id] == nil {
                swiftByTarget[id] = entry
            }
        }

        var entries: [ConflictReport.Entry] = []
        for wordOp in wordOps {
            for id in OperationReducer.referencedElementIDs(in: wordOp) {
                if let swiftEntry = swiftByTarget[id] {
                    entries.append(ConflictReport.Entry(
                        elementID: id, swiftOpID: swiftEntry.opID,
                        swiftOp: swiftEntry.op, wordOp: wordOp))
                }
            }
        }
        return ConflictReport(entries: entries)
    }

    /// Applies the policy to the Word-inferred op set and returns the ops
    /// that should actually be appended to the log. Throws for
    /// `.abortOnConflict` with a non-empty report.
    public static func resolve(
        wordOps: [Operation],
        pendingSwiftOps: [LogEntry],
        policy: SyncPolicy
    ) throws -> [Operation] {
        let report = detect(pendingSwiftOps: pendingSwiftOps, wordOps: wordOps)
        guard !report.entries.isEmpty else { return wordOps }

        switch policy {
        case .abortOnConflict:
            throw SyncError.conflict(report: report)

        case .wordWins:
            return wordOps

        case .swiftWins:
            let conflictingWordOps = report.entries.map(\.wordOp)
            return wordOps.filter { op in !conflictingWordOps.contains(op) }

        case .askUser(let handler):
            let resolutions = handler(report)
            // Word ops for elements resolved .takeSwift (or unresolved —
            // conservative default) are dropped; .takeWord keeps them.
            var dropOps: [Operation] = []
            for entry in report.entries {
                if resolutions[entry.elementID, default: .takeSwift] == .takeSwift {
                    dropOps.append(entry.wordOp)
                }
            }
            return wordOps.filter { op in !dropOps.contains(op) }
        }
    }
}
