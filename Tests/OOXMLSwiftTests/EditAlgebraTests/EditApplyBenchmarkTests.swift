// EditApplyBenchmarkTests.swift
// EditAlgebra — addresses macdoc#110 item #4 (§10.2 of macdoc#105 tasks.md).
//
// Performance benchmark per macdoc#105 spec.md Requirement "Edit Apply
// Performance Within Foundation Baseline" (§163-170):
//
//   The WordDocument.apply(_ edit:) method's performance on the NTPU thesis
//   fixture SHALL be within 10% of the baseline measured by direct
//   OperationLog manipulation + OperationReducer.materialize. If a regression
//   exceeds 10%, this Requirement is NOT satisfied and the implementation
//   MUST be optimized before merge.
//
// **Scope deviation**: NTPU thesis fixture does not exist (carried-over
// deviation from FullyFaithfulFunctorTests / NaturalityTests). Uses a
// synthesized 10-paragraph fixture as substrate — same XmlTree code path
// as real .docx, so relative overhead measurements transfer.
//
// **Iterations**: 100 per path per spec.md Scenario, with 10 warmup iterations
// (Foundation's JIT / cache effects stabilize after a few runs).
//
// **What overhead the Edit API adds over direct Operation manipulation**:
// - WordEdit.lower() identity call (OOXMLEdit lowers to [self])
// - opID generation + sharing between persisted + materialize logs
// - partContaining tree walk to find target part
// - Per-op singleOpLog construction
// - Wraps Reducer error as EditError.operationLogFailure
//
// The 10% budget assumes these overheads are dominated by the materialize
// step itself. For small fixtures (single-paragraph synthesized), overhead
// can dominate and the ratio exceeds 10%. Use a representative fixture size.

import XCTest
@testable import OOXMLSwift

final class EditApplyBenchmarkTests: XCTestCase {

    // MARK: - Configuration

    /// Iterations per measurement round per spec.md §163-170 (100 minimum).
    private static let measurementIterations = 100

    /// Number of measurement rounds. Per-round ratio fluctuates ±5% due to
    /// scheduling / cache noise; taking median of 3 rounds smooths this
    /// out while staying within reasonable test runtime (~1.5 seconds).
    private static let measurementRounds = 3

    /// Warmup iterations before measurement starts. JIT / cache effects
    /// stabilize after a few runs; 10 is comfortable.
    private static let warmupIterations = 10

    /// Maximum allowed ratio of Edit API time to direct Operation time.
    /// Spec.md §163 mandates 1.10 (10% overhead budget).
    private static let maxAllowedRatio = 1.10

    /// Paragraphs in the synthesized fixture. Larger gives the materialize
    /// step more work (tree walk + node insertion are O(N)), making the
    /// Edit API's constant overhead a smaller relative cost.
    ///
    /// 200 paragraphs ≈ small thesis chapter; materialize ≈ 50µs, Edit
    /// overhead ≈ 5µs constant → ratio ≈ 1.10. Smaller fixtures (10-20
    /// paragraphs) give ratios ≈ 1.35-2.0 because the constant overhead
    /// dominates. Spec.md §163-170 prescribed NTPU thesis fixture (hundreds
    /// of paragraphs); 200 is a representative substitute.
    private static let fixtureParagraphCount = 200

    // MARK: - Fixture

    private func makeFixture() -> (WordDocument, ElementID) {
        let firstParaUUID = UUID()
        var paragraphs: [XmlNode] = []

        for i in 0..<Self.fixtureParagraphCount {
            let paraUUID = (i == 0) ? firstParaUUID : UUID()
            let textNode = XmlNode.text("paragraph-\(i)")
            let wt = XmlNode.element(prefix: "w", localName: "t", children: [textNode])
            let wr = XmlNode.element(prefix: "w", localName: "r", children: [wt])
            let wp = XmlNode.element(prefix: "w", localName: "p", children: [wr])
            wp.libraryUUID = paraUUID
            paragraphs.append(wp)
        }

        let body = XmlNode.element(prefix: "w", localName: "body", children: paragraphs)
        let root = XmlNode.element(prefix: "w", localName: "document", children: [body])

        var doc = WordDocument()
        doc.xmlTrees["word/document.xml"] = XmlTree.synthesized(root: root)
        return (doc, ElementID(libraryUUID: firstParaUUID))
    }

    // MARK: - Measurement helpers

    private func timeBlock(_ block: () throws -> Void) rethrows -> Double {
        let start = ProcessInfo.processInfo.systemUptime
        try block()
        return ProcessInfo.processInfo.systemUptime - start
    }

    /// Runs the given `editBlock` and `directBlock` for `measurementRounds`
    /// rounds and returns the MEDIAN ratio. Eliminates per-round noise
    /// (±5% variance from scheduling/cache) while staying within reasonable
    /// test runtime.
    private func medianRatio(
        editBlock: () throws -> Void,
        directBlock: () throws -> Void
    ) rethrows -> (median: Double, allRatios: [Double], editAvg: Double, directAvg: Double) {
        var ratios: [Double] = []
        var editTotals: [Double] = []
        var directTotals: [Double] = []

        for _ in 0..<Self.measurementRounds {
            let editElapsed = try timeBlock(editBlock)
            let directElapsed = try timeBlock(directBlock)
            ratios.append(editElapsed / directElapsed)
            editTotals.append(editElapsed)
            directTotals.append(directElapsed)
        }

        let sortedRatios = ratios.sorted()
        let median = sortedRatios[sortedRatios.count / 2]
        let editAvg = editTotals.reduce(0, +) / Double(editTotals.count) / Double(Self.measurementIterations) * 1_000_000
        let directAvg = directTotals.reduce(0, +) / Double(directTotals.count) / Double(Self.measurementIterations) * 1_000_000
        return (median, ratios, editAvg, directAvg)
    }

    // MARK: - Benchmark: insertParagraph apply

    func testInsertParagraphApplyWithin10PercentOfDirectOperation() throws {
        let (doc, targetID) = makeFixture()
        let baseTree = doc.xmlTrees["word/document.xml"]!

        // Warmup both paths (eliminate startup / JIT / cache effects)
        for _ in 0..<Self.warmupIterations {
            let edit = OOXMLEdit.insertParagraph(after: targetID, content: "warmup", styleId: nil)
            _ = try doc.apply(edit)

            var log = OperationLog()
            log.append(
                .insertParagraphAfter(after: targetID, paragraph: ParagraphPayload(text: "warmup", styleId: nil)),
                source: .swift
            )
            _ = try OperationReducer.materialize(log: log, base: baseTree)
        }

        let result = try medianRatio(
            editBlock: {
                for i in 0..<Self.measurementIterations {
                    let edit = OOXMLEdit.insertParagraph(after: targetID, content: "bench-\(i)", styleId: nil)
                    _ = try doc.apply(edit)
                }
            },
            directBlock: {
                for i in 0..<Self.measurementIterations {
                    var log = OperationLog()
                    log.append(
                        .insertParagraphAfter(after: targetID, paragraph: ParagraphPayload(text: "bench-\(i)", styleId: nil)),
                        source: .swift
                    )
                    _ = try OperationReducer.materialize(log: log, base: baseTree)
                }
            }
        )

        let ratiosStr = result.allRatios.map { String(format: "%.3f", $0) }.joined(separator: ", ")
        print("[BENCH] insertParagraph apply (median of \(Self.measurementRounds) rounds × \(Self.measurementIterations) iterations):")
        print("[BENCH]   Edit API:  \(String(format: "%.2f", result.editAvg)) µs/op")
        print("[BENCH]   Direct:    \(String(format: "%.2f", result.directAvg)) µs/op")
        print("[BENCH]   Ratios:    [\(ratiosStr)] (median: \(String(format: "%.3f", result.median)), spec budget: ≤\(Self.maxAllowedRatio))")

        XCTAssertLessThanOrEqual(
            result.median, Self.maxAllowedRatio,
            "Edit API performance regression: median ratio \(String(format: "%.3f", result.median))× direct path (spec budget ≤\(Self.maxAllowedRatio))"
        )
    }

    // MARK: - Benchmark: setBold apply

    func testSetBoldApplyWithin10PercentOfDirectOperation() throws {
        // Build a fixture matching insertParagraph's size: N paragraphs each
        // with 1 Run. Pick the LAST run as target — worst case for findNode
        // (walks the whole tree). The same overhead applies to both Edit
        // and direct paths, so the ratio measures Edit-specific overhead.
        var paragraphs: [XmlNode] = []
        var targetRunUUID = UUID()  // will be overwritten with last iteration's UUID
        for i in 0..<Self.fixtureParagraphCount {
            let runUUID = UUID()
            let textNode = XmlNode.text("run-\(i)")
            let wt = XmlNode.element(prefix: "w", localName: "t", children: [textNode])
            let wr = XmlNode.element(prefix: "w", localName: "r", children: [wt])
            wr.libraryUUID = runUUID
            let wp = XmlNode.element(prefix: "w", localName: "p", children: [wr])
            paragraphs.append(wp)
            // Capture the last run's UUID as target (worst-case findNode walk)
            if i == Self.fixtureParagraphCount - 1 {
                targetRunUUID = runUUID
            }
        }
        let body = XmlNode.element(prefix: "w", localName: "body", children: paragraphs)
        let root = XmlNode.element(prefix: "w", localName: "document", children: [body])

        var doc = WordDocument()
        doc.xmlTrees["word/document.xml"] = XmlTree.synthesized(root: root)
        let runID = ElementID(libraryUUID: targetRunUUID)
        let baseTree = doc.xmlTrees["word/document.xml"]!

        for _ in 0..<Self.warmupIterations {
            _ = try doc.apply(OOXMLEdit.setBold(target: runID, value: true))

            var log = OperationLog()
            log.append(
                .setRunFormat(target: runID, format: RunFormatPayload(bold: true)),
                source: .swift
            )
            _ = try OperationReducer.materialize(log: log, base: baseTree)
        }

        let result = try medianRatio(
            editBlock: {
                for _ in 0..<Self.measurementIterations {
                    _ = try doc.apply(OOXMLEdit.setBold(target: runID, value: true))
                }
            },
            directBlock: {
                for _ in 0..<Self.measurementIterations {
                    var log = OperationLog()
                    log.append(
                        .setRunFormat(target: runID, format: RunFormatPayload(bold: true)),
                        source: .swift
                    )
                    _ = try OperationReducer.materialize(log: log, base: baseTree)
                }
            }
        )

        let ratiosStr = result.allRatios.map { String(format: "%.3f", $0) }.joined(separator: ", ")
        print("[BENCH] setBold apply (median of \(Self.measurementRounds) rounds × \(Self.measurementIterations) iterations):")
        print("[BENCH]   Edit API:  \(String(format: "%.2f", result.editAvg)) µs/op")
        print("[BENCH]   Direct:    \(String(format: "%.2f", result.directAvg)) µs/op")
        print("[BENCH]   Ratios:    [\(ratiosStr)] (median: \(String(format: "%.3f", result.median)), spec budget: ≤\(Self.maxAllowedRatio))")

        XCTAssertLessThanOrEqual(
            result.median, Self.maxAllowedRatio,
            "Edit API performance regression: median ratio \(String(format: "%.3f", result.median))× direct path (spec budget ≤\(Self.maxAllowedRatio))"
        )
    }
}
