// OperationReducerCache — actor-isolated tail-replay cache for the reducer.
//
// Spectra change: operation-reducer-impl, target ooxml-swift v0.31.4.
// Capability: ooxml-operation-reducer (Snapshot caching avoids full replay).
//
// Phase 2b read-path optimization. The pure reducer's `materialize` is O(N)
// in log length; for read-heavy workloads (typed-view getters that re-derive
// state on every access) this is wasteful. The cache memoizes the last
// materialized tree per `base` and, on the next call with an extended log,
// replays only the tail (`log.entries[cached.logLength..<log.entries.count]`)
// against a deep-clone of the cached tree.
//
// Concurrency: backed by a Swift `actor` so concurrent readers serialize
// without explicit locks. The cache is keyed by `ObjectIdentifier(base.root)`
// so distinct base trees get distinct cache entries automatically.
//
// Cache miss / stale conditions:
//   - First call for a given `base.root` identity (miss).
//   - `cached.logLength > log.entries.count` (log shrank — stale; full replay).
//
// On hit, the tail-replay path uses `OperationReducer.applyOrInterpret` so it
// honors the same `.undo` / `.batchBegin` / `.batchEnd` / `.unknown` semantics
// as the full-replay path.

import Foundation

/// Actor-isolated cache for `OperationReducer.materialize`. Stores the most
/// recently materialized tree per `base.root` identity and tail-replays new
/// log entries on subsequent calls.
public actor OperationReducerCache {

    /// One cache entry per distinct `base` tree (keyed by `ObjectIdentifier`
    /// of `base.root`). `logLength` records how many `log.entries` were
    /// materialized into `materializedTree`.
    private struct CacheEntry {
        let logLength: Int
        let materializedTree: XmlTree
    }

    private var cached: [ObjectIdentifier: CacheEntry] = [:]

    public init() {}

    /// Returns a freshly-cloned `XmlTree` representing `log` materialized on
    /// `base`. Same semantics as `OperationReducer.materialize(log:base:)`,
    /// just memoized.
    public func materialize(log: OperationLog, base: XmlTree) async throws -> XmlTree {
        let key = ObjectIdentifier(base.root)
        if let entry = cached[key], entry.logLength <= log.entries.count {
            // Tail-replay: clone the cached tree, replay only the new entries.
            var working = entry.materializedTree.deepCopy()
            for i in entry.logLength..<log.entries.count {
                try OperationReducer.applyOrInterpret(
                    entry: log.entries[i],
                    entryIndex: i,
                    log: log,
                    to: &working
                )
            }
            cached[key] = CacheEntry(
                logLength: log.entries.count,
                materializedTree: working
            )
            return working.deepCopy()
        }

        // Miss or stale (log shrank): full replay.
        let working = try OperationReducer.materialize(log: log, base: base)
        cached[key] = CacheEntry(
            logLength: log.entries.count,
            materializedTree: working
        )
        return working.deepCopy()
    }
}
