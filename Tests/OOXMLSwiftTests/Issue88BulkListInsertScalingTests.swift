import Foundation
import XCTest
@testable import OOXMLSwift

/// Regression test for [PsychQuant/ooxml-swift#88](https://github.com/PsychQuant/ooxml-swift/issues/88).
///
/// `withStampedParaId(_:)` (used by every path that inserts a paragraph
/// into the document — `appendParagraph`, `insertParagraph`) is O(current
/// document size): it calls `getAllParagraphs()`, a full walk of body +
/// every table cell + every content control, EVERY time, to build the
/// exclusion set for a fresh random paraId. That is correct and cheap for
/// one paragraph. `insertBulletList`/`insertNumberedList` call
/// `appendParagraph`/`insertParagraph` once PER ITEM in a loop — for M
/// items into an N-paragraph document, that is O(M × (N + M)): each of the
/// M items re-walks a document that just grew by the previous items,
/// effectively O(M²) when M dominates (bulk authoring into an otherwise
/// small/empty document — exactly `insertBulletList`'s own common case).
///
/// Fix: `withStampedParaIds(_:)` (plural) scans the existing-id exclusion
/// set ONCE, then stamps the whole batch from a single running generator
/// + growing exclusion set. `insertBulletList`/`insertNumberedList` call it
/// once on their full batch before handing paragraphs one at a time to
/// `insertParagraph`/`appendParagraph` — those calls' own internal
/// `withStampedParaId` then hits its "already has an id" fast path
/// instead of re-scanning. O(M²) → O(N + M).
///
/// Uses this repo's existing `XCTAssertScalesLinearly` (`ScalingAssertion.swift`,
/// #174) rather than an absolute wall-clock threshold — see that file's own
/// header for why (load-sensitive absolute-time assertions flake; this
/// compares CPU time at two input scales instead, which a load spike
/// affects roughly equally at both scales).
final class Issue88BulkListInsertScalingTests: XCTestCase {

    /// `baseSize` chosen so this scale's run clears `ScalingAssertion`'s
    /// 50ms noise floor with margin on the FIXED implementation. Verified
    /// empirically against the pre-fix (reverted, per-item-loop) code at a
    /// smaller scale, where it already fails decisively — baseSize=250,
    /// 4n=1000: measured ratio 15.86 (n=250 → ~0.073s, 4n=1000 → ~1.17s;
    /// linear would be ~4×, this is within 1% of the 16× a genuine O(M²)
    /// predicts). At this test's actual baseSize=4000 (4n=16000) the
    /// pre-fix code does not merely fail the ratio check — it did not
    /// finish within a 550-second wall-clock budget in manual testing,
    /// which is why the smaller scale is what pre-fix evidence cites.
    func testInsertBulletListScalesLinearlyNotQuadratically() throws {
        try XCTAssertScalesLinearly(
            "insertBulletList",
            baseSize: 4000,
            prepare: { n in
                (document: WordDocument(), items: (0..<n).map { "item \($0)" })
            },
            teardown: { _ in }
        ) { input in
            var doc = input.document
            _ = doc.insertBulletList(items: input.items)
        }
    }

    func testInsertNumberedListScalesLinearlyNotQuadratically() throws {
        try XCTAssertScalesLinearly(
            "insertNumberedList",
            baseSize: 4000,
            prepare: { n in
                (document: WordDocument(), items: (0..<n).map { "item \($0)" })
            },
            teardown: { _ in }
        ) { input in
            var doc = input.document
            _ = doc.insertNumberedList(items: input.items)
        }
    }

    // MARK: - Correctness: batch-stamped ids are unique and stable

    func testBulkStampedParagraphsAllGetDistinctParaIds() throws {
        var doc = WordDocument()
        _ = doc.insertBulletList(items: (0..<200).map { "item \($0)" })

        let paraIds = doc.getAllParagraphs().compactMap(\.w14ParaId)
        XCTAssertEqual(paraIds.count, 200)
        XCTAssertEqual(Set(paraIds).count, 200, "every bulk-inserted paragraph must get a DISTINCT w14:paraId")
    }

    /// A paragraph created with a caller-preset `w14:paraId` must keep it
    /// verbatim (matches the single-paragraph `withStampedParaId`'s own
    /// "caller-preset value kept verbatim" contract) even when it appears
    /// inside a batch alongside auto-stamped paragraphs.
    func testWithStampedParaIdsPreservesCallerPresetIds() {
        var doc = WordDocument()
        var preset = Paragraph(text: "has id")
        preset.w14ParaId = "DEADBEEF"
        let unstamped = Paragraph(text: "no id")

        let stamped = doc.withStampedParaIds([preset, unstamped])
        XCTAssertEqual(stamped[0].w14ParaId, "DEADBEEF")
        XCTAssertNotNil(stamped[1].w14ParaId)
        XCTAssertNotEqual(stamped[1].w14ParaId, "DEADBEEF")
    }

    /// New ids allocated by a batch must not collide with ids ALREADY
    /// present in the document before the batch started.
    func testWithStampedParaIdsAvoidsCollisionWithExistingDocumentIds() {
        var doc = WordDocument()
        var existing = Paragraph(text: "existing")
        existing.w14ParaId = "00000001"
        doc.body.children = [.paragraph(existing)]

        let batch = (0..<50).map { Paragraph(text: "new \($0)") }
        let stamped = doc.withStampedParaIds(batch)

        XCTAssertTrue(stamped.allSatisfy { $0.w14ParaId != "00000001" })
        XCTAssertEqual(Set(stamped.compactMap(\.w14ParaId)).count, 50)
    }
}
