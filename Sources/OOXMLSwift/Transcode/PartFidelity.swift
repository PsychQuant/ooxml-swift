// PartFidelity.swift
// format-alignment-engine Phase A task 1.1 — dual-track byte-diff + coverage
// accounting for the reference→rebuild pipeline
// (`format-alignment-pipeline` capability, Decision 2: raw channel is the
// byte-equal floor, DSL coverage is the imitation-ability score).
//
// Two independent measurement axes:
//   - Verdict: per-part byte comparison (Stage A) and full part-set equality
//     (Stage B — the final acceptance stage; Stage C zip-container equality is
//     out of contract).
//   - Coverage: for each part, how many bytes were rebuilt through the typed
//     DSL channel versus the raw channel, aggregated to DSL-form coverage %.
//
// The type is pure value logic over `[String: Data]` part maps (partPath →
// bytes) and performs NO docx I/O. Callers read a package into a part map and
// hand it here, keeping the metric independently unit-testable.

import Foundation

public enum PartFidelity {

    // MARK: - Stage A/B byte comparison

    /// Verdict for a single XML part path across the reference and rebuilt
    /// packages.
    public struct PartVerdict: Equatable {
        public let partPath: String
        public let status: Status

        public enum Status: Equatable {
            /// Byte-identical in both packages.
            case equal
            /// Present in both but differing; `firstDivergenceOffset` is the
            /// 0-based index of the first differing byte, or the length of the
            /// shorter blob when one is a strict prefix of the other.
            case differ(firstDivergenceOffset: Int)
            /// Present in reference, absent in rebuilt.
            case missingInRebuilt
            /// Present in rebuilt, absent in reference.
            case unexpectedInRebuilt
        }

        public init(partPath: String, status: Status) {
            self.partPath = partPath
            self.status = status
        }

        /// True only for a byte-identical part.
        public var isEqual: Bool { status == .equal }
    }

    /// Stage A: per-part verdicts over the union of both part sets, sorted by
    /// part path for stable reporting.
    public static func compareParts(reference: [String: Data],
                                    rebuilt: [String: Data]) -> [PartVerdict] {
        let allPaths = Set(reference.keys).union(rebuilt.keys).sorted()
        return allPaths.map { path in
            switch (reference[path], rebuilt[path]) {
            case let (ref?, reb?):
                if ref == reb {
                    return PartVerdict(partPath: path, status: .equal)
                }
                return PartVerdict(partPath: path,
                                   status: .differ(firstDivergenceOffset: firstDivergence(ref, reb)))
            case (_?, nil):
                return PartVerdict(partPath: path, status: .missingInRebuilt)
            case (nil, _?):
                return PartVerdict(partPath: path, status: .unexpectedInRebuilt)
            case (nil, nil):
                // Unreachable: `path` came from the union of both key sets.
                return PartVerdict(partPath: path, status: .equal)
            }
        }
    }

    /// Stage B: the final acceptance stage — every reference part is present in
    /// rebuilt and byte-equal, with no unexpected parts.
    public static func stageB(reference: [String: Data],
                              rebuilt: [String: Data]) -> Bool {
        compareParts(reference: reference, rebuilt: rebuilt).allSatisfy { $0.isEqual }
    }

    /// 0-based index of the first differing byte between two blobs. When the
    /// shorter blob is a prefix of the longer, the divergence is reported at the
    /// shorter blob's length (the first position where lengths force a mismatch).
    ///
    /// Normalizing through `[UInt8]` guarantees logical (0-based) offsets even
    /// when a `Data` argument is a slice with a non-zero `startIndex`.
    private static func firstDivergence(_ a: Data, _ b: Data) -> Int {
        let aBytes = [UInt8](a)
        let bBytes = [UInt8](b)
        let n = min(aBytes.count, bBytes.count)
        var i = 0
        while i < n {
            if aBytes[i] != bBytes[i] { return i }
            i += 1
        }
        return n
    }

    // MARK: - Coverage accounting

    /// DSL-vs-raw byte split for a single XML part.
    public struct PartCoverage: Equatable {
        public let partPath: String
        /// Bytes rebuilt through the typed DSL channel.
        public let dslBytes: Int
        /// Bytes carried verbatim on the raw channel.
        public let rawBytes: Int

        public init(partPath: String, dslBytes: Int, rawBytes: Int) {
            self.partPath = partPath
            self.dslBytes = dslBytes
            self.rawBytes = rawBytes
        }

        public var totalBytes: Int { dslBytes + rawBytes }

        /// DSL-form coverage for this part, in `[0, 1]`. A zero-byte part is 0
        /// (never NaN).
        public var coverageRatio: Double {
            totalBytes == 0 ? 0 : Double(dslBytes) / Double(totalBytes)
        }
    }

    /// Aggregate DSL-form coverage across all XML parts.
    public struct CoverageReport: Equatable {
        /// Per-part coverage, sorted by part path.
        public let parts: [PartCoverage]

        public init(parts: [PartCoverage]) {
            self.parts = parts
        }

        public var aggregateDSLBytes: Int { parts.reduce(0) { $0 + $1.dslBytes } }
        public var aggregateTotalBytes: Int { parts.reduce(0) { $0 + $1.totalBytes } }

        /// Aggregate DSL-form coverage = DSL bytes ÷ total XML bytes across all
        /// parts, in `[0, 1]` (0 when there are no bytes; never NaN).
        public var aggregateRatio: Double {
            aggregateTotalBytes == 0 ? 0 : Double(aggregateDSLBytes) / Double(aggregateTotalBytes)
        }
    }

    /// Build a coverage report from per-part splits, sorting parts by path.
    public static func coverage(_ parts: [PartCoverage]) -> CoverageReport {
        CoverageReport(parts: parts.sorted { $0.partPath < $1.partPath })
    }
}
