// ParaIdGeneratorTests.swift
// authoring-canonical-conformance task 1.1 (design D3) — the injectable
// per-document w14:paraId generator. Spec: "Authoring chokepoints stamp
// w14:paraId" — 8 uppercase hex, numeric value strictly inside
// (0x00000000, 0x80000000), unique against a provided existing-ID set,
// deterministic under an injected RNG.

import XCTest
@testable import OOXMLSwift

/// Deterministic RNG for pinning generator sequences (SplitMix64).
struct SplitMix64: RandomNumberGenerator {
    private var state: UInt64
    init(seed: UInt64) { self.state = seed }
    mutating func next() -> UInt64 {
        state &+= 0x9E37_79B9_7F4A_7C15
        var z = state
        z = (z ^ (z >> 30)) &* 0xBF58_476D_1CE4_E5B9
        z = (z ^ (z >> 27)) &* 0x94D0_49BB_1331_11EB
        return z ^ (z >> 31)
    }
}

/// RNG that replays a fixed word queue — forces exact draw sequences so
/// boundary skips and collision skips are testable without probability.
struct FixedSequenceRNG: RandomNumberGenerator {
    private var queue: [UInt64]
    init(_ values: [UInt64]) { self.queue = values }
    mutating func next() -> UInt64 {
        precondition(!queue.isEmpty, "FixedSequenceRNG exhausted — test drew more values than queued")
        return queue.removeFirst()
    }
}

final class ParaIdGeneratorTests: XCTestCase {

    func testFormatIsEightUppercaseHex() {
        var gen = ParaIdGenerator(rng: SplitMix64(seed: 0xC0FFEE))
        for _ in 0..<100 {
            let id = gen.next(excluding: [])
            XCTAssertNotNil(id.range(of: "^[0-9A-F]{8}$", options: .regularExpression),
                            "generated paraId '\(id)' must be 8 uppercase hex characters")
        }
    }

    func testValuesStayStrictlyInsideWordRange() {
        var gen = ParaIdGenerator(rng: SplitMix64(seed: 42))
        for _ in 0..<1000 {
            let id = gen.next(excluding: [])
            let value = UInt32(id, radix: 16)!
            XCTAssertGreaterThan(value, 0x0000_0000, "0x00000000 must never be emitted")
            XCTAssertLessThan(value, 0x8000_0000, "values >= 0x80000000 must never be emitted")
        }
    }

    func testBoundaryDrawsAreSkipped() {
        // Raw draws of 0 and 0x80000000 both map outside the open interval
        // and must be skipped in favor of the next draw.
        var gen = ParaIdGenerator(rng: FixedSequenceRNG([0x0000_0000, 0x8000_0000, 0x2AB4_C9F0]))
        XCTAssertEqual(gen.next(excluding: []), "2AB4C9F0")
    }

    func testCollisionSkipAgainstExistingSet() {
        // Spec example: document already uses 11111111 and 2AB4C9F0 — draws
        // colliding with either are skipped; the first free draw is returned.
        var gen = ParaIdGenerator(rng: FixedSequenceRNG([0x2AB4_C9F0, 0x1111_1111, 0x3F2A_0001]))
        let id = gen.next(excluding: ["11111111", "2AB4C9F0"])
        XCTAssertEqual(id, "3F2A0001")
        XCTAssertNotNil(id.range(of: "^[0-9A-F]{8}$", options: .regularExpression))
    }

    func testSeededSequenceIsReproducible() {
        var a = ParaIdGenerator(rng: SplitMix64(seed: 7))
        var b = ParaIdGenerator(rng: SplitMix64(seed: 7))
        let seqA = (0..<10).map { _ in a.next(excluding: []) }
        let seqB = (0..<10).map { _ in b.next(excluding: []) }
        XCTAssertEqual(seqA, seqB, "same seed must yield the same paraId sequence")
    }

    func testDefaultInitializerProducesConformingValues() {
        var gen = ParaIdGenerator()
        let id = gen.next(excluding: [])
        XCTAssertNotNil(id.range(of: "^[0-9A-F]{8}$", options: .regularExpression))
        let value = UInt32(id, radix: 16)!
        XCTAssertGreaterThan(value, 0)
        XCTAssertLessThan(value, 0x8000_0000)
    }
}
