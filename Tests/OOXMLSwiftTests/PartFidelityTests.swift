// PartFidelityTests.swift
// format-alignment-engine Phase A task 1.1 — dual-track byte-diff + coverage
// accounting for the rebuild pipeline (`format-alignment-pipeline` capability:
// «Dual-track acceptance — byte-equal floor and DSL coverage score»,
// «DSL-form coverage measurement»; Decision 2).

import XCTest
@testable import OOXMLSwift

final class PartFidelityTests: XCTestCase {

    // MARK: - Stage A/B byte comparison

    func testEqualPartSetsProduceEqualVerdictsAndStageBPasses() {
        let ref: [String: Data] = [
            "word/document.xml": Data("<doc/>".utf8),
            "word/styles.xml": Data("<styles/>".utf8),
        ]
        let verdicts = PartFidelity.compareParts(reference: ref, rebuilt: ref)
        XCTAssertEqual(verdicts.map(\.status), [.equal, .equal])
        XCTAssertTrue(PartFidelity.stageB(reference: ref, rebuilt: ref))
    }

    func testDifferingPartReportsFirstDivergenceOffset() {
        // "hello" vs "help!" — first mismatch at 0-based index 3 (l ≠ p).
        let ref = ["word/document.xml": Data("hello".utf8)]
        let reb = ["word/document.xml": Data("help!".utf8)]
        let verdicts = PartFidelity.compareParts(reference: ref, rebuilt: reb)
        XCTAssertEqual(verdicts.count, 1)
        XCTAssertEqual(verdicts[0].partPath, "word/document.xml")
        XCTAssertEqual(verdicts[0].status, .differ(firstDivergenceOffset: 3))
        XCTAssertFalse(PartFidelity.stageB(reference: ref, rebuilt: reb))
    }

    func testMissingPartInRebuiltIsReportedAndFailsStageB() {
        let ref = [
            "word/document.xml": Data("a".utf8),
            "word/styles.xml": Data("b".utf8),
        ]
        let reb = ["word/document.xml": Data("a".utf8)]
        let verdicts = PartFidelity.compareParts(reference: ref, rebuilt: reb)
        // Verdicts are sorted by part path: document.xml (equal), styles.xml (missing).
        XCTAssertEqual(verdicts.map(\.partPath), ["word/document.xml", "word/styles.xml"])
        XCTAssertEqual(verdicts[1].status, .missingInRebuilt)
        XCTAssertFalse(PartFidelity.stageB(reference: ref, rebuilt: reb))
    }

    func testUnexpectedPartInRebuiltIsReportedAndFailsStageB() {
        let ref = ["word/document.xml": Data("a".utf8)]
        let reb = [
            "word/document.xml": Data("a".utf8),
            "word/glossary/document.xml": Data("x".utf8),
        ]
        let verdicts = PartFidelity.compareParts(reference: ref, rebuilt: reb)
        let unexpected = verdicts.first { $0.partPath == "word/glossary/document.xml" }
        XCTAssertEqual(unexpected?.status, .unexpectedInRebuilt)
        XCTAssertFalse(PartFidelity.stageB(reference: ref, rebuilt: reb))
    }

    func testPrefixLengthMismatchDivergesAtBoundary() {
        // "abc" vs "ab" — common prefix, rebuilt shorter → divergence at index 2.
        let ref = ["word/document.xml": Data("abc".utf8)]
        let reb = ["word/document.xml": Data("ab".utf8)]
        let verdicts = PartFidelity.compareParts(reference: ref, rebuilt: reb)
        XCTAssertEqual(verdicts[0].status, .differ(firstDivergenceOffset: 2))
    }

    func testDivergenceOffsetIsZeroBasedForSlicedData() {
        // A Data slice carries a non-zero startIndex; the reported offset must be
        // relative to logical content (0-based), never the slice's startIndex.
        let refSlice = Data("PADhelp!".utf8).dropFirst(3)   // "help!" startIndex=3
        let rebSlice = Data("XXXXhello".utf8).dropFirst(4)  // "hello" startIndex=4
        let verdicts = PartFidelity.compareParts(
            reference: ["word/document.xml": refSlice],
            rebuilt: ["word/document.xml": rebSlice]
        )
        // "help!" vs "hello": diverge at logical index 3 (p ≠ l).
        XCTAssertEqual(verdicts[0].status, .differ(firstDivergenceOffset: 3))
    }

    // MARK: - Coverage accounting

    func testCoverageArithmeticMatchesSpecExample() {
        // format-alignment-pipeline spec §Example: coverage arithmetic.
        // document.xml = 70,000 DSL + 10,000 raw; styles.xml = 20,000 all raw.
        let report = PartFidelity.coverage([
            .init(partPath: "word/document.xml", dslBytes: 70_000, rawBytes: 10_000),
            .init(partPath: "word/styles.xml", dslBytes: 0, rawBytes: 20_000),
        ])
        let doc = report.parts.first { $0.partPath == "word/document.xml" }!
        let styles = report.parts.first { $0.partPath == "word/styles.xml" }!
        XCTAssertEqual(doc.coverageRatio, 0.875, accuracy: 1e-9)
        XCTAssertEqual(styles.coverageRatio, 0.0, accuracy: 1e-9)
        XCTAssertEqual(report.aggregateDSLBytes, 70_000)
        XCTAssertEqual(report.aggregateTotalBytes, 100_000)
        XCTAssertEqual(report.aggregateRatio, 0.70, accuracy: 1e-9)
    }

    func testEmptyPartHasZeroCoverageNotNaN() {
        // Divide-by-zero guard: a zero-byte part must not produce NaN.
        let report = PartFidelity.coverage([
            .init(partPath: "word/empty.xml", dslBytes: 0, rawBytes: 0),
        ])
        XCTAssertEqual(report.parts[0].coverageRatio, 0.0)
        XCTAssertEqual(report.aggregateRatio, 0.0)
        XCTAssertFalse(report.aggregateRatio.isNaN)
    }

    func testCoverageReportPartsAreSortedByPath() {
        let report = PartFidelity.coverage([
            .init(partPath: "word/styles.xml", dslBytes: 1, rawBytes: 0),
            .init(partPath: "word/document.xml", dslBytes: 1, rawBytes: 0),
        ])
        XCTAssertEqual(report.parts.map(\.partPath), ["word/document.xml", "word/styles.xml"])
    }
}
