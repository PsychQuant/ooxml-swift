// OperationLog — append-only collection of LogEntry records.
//
// Spectra change: operation-log-scaffold-impl, target ooxml-swift v0.31.3.
// Capability: ooxml-operation-log
//
// Value-type with `private(set) var entries` to enforce the append-only
// invariant. The `batch(_:label:_:)` helper emits BatchBegin/BatchEnd markers
// around the closure body — Phase 2b reducer uses these markers to group
// related ops for undo/redo. JSONL serialization lives in the +JSONL
// extension file.

import Foundation

/// Append-only collection of log entries representing the mutation history.
///
/// Construct with `init()`, append via `append(_:source:)` or `batch(_:_:_:)`.
/// `entries` is exposed read-only externally so external code cannot mutate
/// the history out-of-band — the only sanctioned ways to grow the log are
/// the two mutating methods on this type.
public struct OperationLog: Equatable, Sendable {

    /// Append-only history of log entries in source order.
    ///
    /// Externally read-only — `private(set)` enforces the append-only invariant
    /// so callers cannot remove, replace, or reorder entries directly. The
    /// only sanctioned ways to grow the log are the `append(_:source:)` and
    /// `batch(_:label:_:)` mutating methods below.
    public private(set) var entries: [LogEntry] = []

    public init() {}

    /// Appends a single operation to the log with explicit source attribution.
    ///
    /// `opID` defaults to a fresh UUID v4 — production callers omit it; tests
    /// supply deterministic UUIDs for byte-equal JSONL round-trip assertions.
    /// `timestamp` defaults to `Date()` for the same reason.
    public mutating func append(
        _ op: Operation,
        source: OpSource,
        opID: UUID = UUID(),
        at timestamp: Date = Date()
    ) {
        entries.append(LogEntry(opID: opID, op: op, source: source, timestamp: timestamp))
    }

    /// Wraps a closure body in `batchBegin(label:)` / `batchEnd` markers.
    ///
    /// The body's appends to `self` are sandwiched between the markers. If the
    /// body throws, `batchEnd` is NOT appended (the spec scenario "batch closes
    /// its end marker on throw" pins this best-effort behavior — rollback is a
    /// Phase 2b reducer concern, not a data-structure concern).
    public mutating func batch(
        _ source: OpSource,
        label: String? = nil,
        _ body: (inout OperationLog) throws -> Void
    ) rethrows {
        append(.batchBegin(label: label), source: source)
        try body(&self)
        append(.batchEnd, source: source)
    }
}

// MARK: - LogEntry

/// One row in the operation log.
///
/// Carries the op itself plus three pieces of metadata: a unique `opID`, the
/// `source` attribution (Swift code or Word app), and a wall-clock timestamp.
/// Phase 2b reducer uses `opID` to address ops for undo/redo; Phase 3
/// SyncOrchestrator filters by `source` when computing import diffs.
public struct LogEntry: Equatable, Sendable {
    public let opID: UUID
    public let op: Operation
    public let source: OpSource
    public let timestamp: Date

    public init(opID: UUID, op: Operation, source: OpSource, timestamp: Date) {
        self.opID = opID
        self.op = op
        self.source = source
        self.timestamp = timestamp
    }
}

// `LogEntry: Codable` and `Operation: Codable` conformances live in
// `OperationLog+JSONL.swift` (task 1.4) where they're co-located with the
// JSONL field-order and forward-compat logic that drives the encode/decode
// shape. Auto-synthesis cannot apply because `Operation`'s associated-value
// cases need custom `init(from:)` / `encode(to:)`.

// MARK: - OpSource

/// Source attribution for an operation. `swift` = the local Swift code emitted
/// the op (typed-view setter, batch helper). `word` = the op was reconstructed
/// from a Word-app edit by Phase 3's SyncOrchestrator import path.
public enum OpSource: String, Equatable, Sendable, Codable {
    case swift
    case word
}
