// ParaIdGenerator.swift
// authoring-canonical-conformance (design D3): generates `w14:paraId` values
// for paragraphs created through the authoring chokepoints. Word's paraId is
// an 8-uppercase-hex token whose numeric value must lie strictly between
// 0x00000000 and 0x80000000 (exclusive) and be unique within the document —
// it is the stable `setRuns` addressing key across script round-trips, so
// the transcoder hard-requires it (`paragraph-no-paraId`).

import Foundation

/// Generates Word-conforming `w14:paraId` tokens. The RNG is injectable so
/// tests can pin deterministic sequences; production uses the system RNG.
public struct ParaIdGenerator {

    private var rng: any RandomNumberGenerator

    public init() {
        self.rng = SystemRandomNumberGenerator()
    }

    public init(rng: some RandomNumberGenerator) {
        self.rng = rng
    }

    /// Next paraId not present in `existing`. Draws map into the open
    /// interval (0x00000000, 0x80000000); out-of-range and colliding draws
    /// are skipped. Termination: the space holds 2^31 - 1 candidates while a
    /// document holds orders of magnitude fewer paragraphs, so a free value
    /// is always reachable.
    public mutating func next(excluding existing: Set<String>) -> String {
        while true {
            let raw = UInt32(truncatingIfNeeded: rng.next()) & 0x7FFF_FFFF
            guard raw != 0 else { continue }
            let id = String(format: "%08X", raw)
            if !existing.contains(id) { return id }
        }
    }
}
