import Foundation
import XCTest
@testable import OOXMLSwift

/// Regression tests for [PsychQuant/ooxml-swift#180](https://github.com/PsychQuant/ooxml-swift/issues/180).
///
/// An independent review (che-word-mcp#234, real release binary) measured
/// `Table(rowCount:columnCount:)`'s blank cells at ~2.2 KB resident each
/// (65536×63 → 9.0 GB resident, 14.2 GB save-time peak, for a `.docx` that
/// serializes to 0.8 MB) — memory scaling linearly with cell count for
/// content that is, by construction, empty.
///
/// Root cause, found by measuring `MemoryLayout<T>.stride` for the types
/// involved (see the fix report for the full numbers): `Paragraph`'s
/// stride was 4128 bytes, of which `ParagraphProperties` (1944 bytes)
/// accounted for TWICE — once via the always-embedded `properties` field,
/// once via the always-embedded (even though almost always `nil`)
/// `previousProperties: ParagraphProperties?` field. Every `TableCell()`
/// (an empty cell) holds exactly one such `Paragraph` (OOXML requires at
/// least one `<w:p>` per `<w:tc>`) in a FRESHLY allocated one-element
/// array — `TableCell.init()` built a brand new `[Paragraph()]` array
/// literal every single call, and `Table.init(rowCount:columnCount:)`
/// built a brand new `[TableCell]` array for every row.
///
/// Fix (all three, together — "shared defaults" per the issue's own
/// suggestion):
///
/// 1. `Paragraph.previousProperties` is now boxed (a private reference
///    type) instead of directly embedded — collapses ~1944 bytes to 8
///    bytes for the (overwhelmingly common) `nil` case, dropping
///    `Paragraph`'s stride to 2192 bytes. This helps EVERY paragraph
///    everywhere, not just blank table cells.
/// 2. `TableCell.init()` now assigns a `static let` shared `[Paragraph()]`
///    array instead of a fresh array literal — every blank cell's
///    paragraph array is the SAME retained buffer (Array is
///    copy-on-write) until that specific cell's paragraphs are actually
///    replaced.
/// 3. `Table.init(rowCount:columnCount:)` now builds ONE blank row's
///    `[TableCell]` array and shares it (again via COW) across every row,
///    instead of calling `TableCell()` fresh `rowCount * columnCount`
///    times into `rowCount` separate array allocations.
///
/// Measured effect (debug build, `MemoryLayout` + RSS delta via
/// `task_info`, 200,000 cells / 20,000 rows × 10 columns): resident memory
/// for constructing the table dropped from ~5895 bytes/cell to ~5
/// bytes/cell — three orders of magnitude, consistent with "most cells now
/// share one underlying allocation instead of each owning its own."
final class Issue180BlankTableMemoryTests: XCTestCase {

    // MARK: - Size regression (structural — catches a future field re-inlining previousProperties)

    func testParagraphStrideStaysSmallAfterBoxingPreviousProperties() {
        // Before the fix this was 4128 (2x ParagraphProperties.stride,
        // embedded once directly, once via the un-boxed optional). A
        // regression that re-inlines `previousProperties` would silently
        // double this again.
        XCTAssertLessThan(
            MemoryLayout<Paragraph>.stride,
            MemoryLayout<ParagraphProperties>.stride + 512,
            "Paragraph's stride should be roughly ParagraphProperties's own size plus its other (small) fields, not 2x+ that — a regression likely re-inlined previousProperties")
    }

    // MARK: - Correctness: COW sharing must not leak mutations across cells

    /// The core safety property of the fix: sharing the same underlying
    /// array across many `TableCell`/`TableRow` values via COW is only
    /// correct if mutating ONE cell's content never affects any other
    /// cell — Swift's copy-on-write guarantees this for plain array
    /// reassignment, but this test pins it down explicitly rather than
    /// trusting the guarantee blindly.
    func testMutatingOneCellDoesNotAffectSiblingCellsOrOtherRows() {
        var table = Table(rowCount: 3, columnCount: 3)

        // Mutate exactly one cell (row 1, col 1) — replace its paragraphs
        // entirely, matching how `TableCell.paragraphs`'s setter works.
        var middleCell = table.rows[1].cells[1]
        middleCell.paragraphs = [Paragraph(text: "mutated")]
        table.rows[1].cells[1] = middleCell

        // Every OTHER cell (same row, other rows) must remain the
        // original blank placeholder.
        for r in 0..<3 {
            for c in 0..<3 {
                let cell = table.rows[r].cells[c]
                if r == 1 && c == 1 {
                    XCTAssertEqual(cell.getText(), "mutated")
                } else {
                    XCTAssertEqual(cell.getText(), "", "cell (\(r),\(c)) must be unaffected by mutating (1,1)")
                }
            }
        }
    }

    /// Mutating an entire ROW's cells (replacing the row's `[TableCell]`
    /// wholesale) must not affect sibling rows that still share the
    /// original blank-row-cells buffer.
    func testReplacingOneRowsCellsDoesNotAffectOtherRows() {
        var table = Table(rowCount: 4, columnCount: 5)

        var row2 = table.rows[2]
        row2.cells = (0..<5).map { TableCell(text: "row2-\($0)") }
        table.rows[2] = row2

        for r in 0..<4 {
            for c in 0..<5 {
                let text = table.rows[r].cells[c].getText()
                if r == 2 {
                    XCTAssertEqual(text, "row2-\(c)")
                } else {
                    XCTAssertEqual(text, "", "row \(r) must be unaffected by replacing row 2's cells")
                }
            }
        }
    }

    // MARK: - Correctness: previousProperties get/set round-trips through the box

    func testPreviousPropertiesRoundTripsThroughBox() {
        var para = Paragraph(text: "hi")
        XCTAssertNil(para.previousProperties, "default must still be nil")

        var props = ParagraphProperties()
        props.alignment = .center
        para.previousProperties = props
        XCTAssertEqual(para.previousProperties?.alignment, .center)

        para.previousProperties = nil
        XCTAssertNil(para.previousProperties)
    }

    /// Two `Paragraph` VALUES that both set the SAME `previousProperties`
    /// content, independently, must remain EQUAL and independently
    /// mutable (no accidental box-sharing across genuinely distinct
    /// paragraphs — each `set` creates its own fresh box).
    func testTwoParagraphsWithEqualPreviousPropertiesAreIndependent() {
        var propsA = ParagraphProperties()
        propsA.alignment = .center
        var a = Paragraph(text: "a")
        a.previousProperties = propsA

        var b = a  // struct copy — shares the box's REFERENCE until mutated
        XCTAssertEqual(a.previousProperties, b.previousProperties)

        var propsB = ParagraphProperties()
        propsB.alignment = .right
        b.previousProperties = propsB

        XCTAssertEqual(a.previousProperties?.alignment, .center, "original must be unaffected by copy's mutation")
        XCTAssertEqual(b.previousProperties?.alignment, .right)
    }

    // MARK: - Correctness: serialized output unchanged (this is a storage-only optimization)

    func testBlankTableSerializesIdenticallyToBeforeTheOptimization() {
        let table = Table(rowCount: 2, columnCount: 3)
        let xml = table.toXML()

        // 2 rows x 3 cols = 6 cells, each with exactly one empty <w:p>.
        XCTAssertEqual(xml.components(separatedBy: "<w:tr>").count - 1, 2)
        XCTAssertEqual(xml.components(separatedBy: "<w:tc>").count - 1, 6)
        XCTAssertEqual(xml.components(separatedBy: "<w:p").count - 1, 6, "each of the 6 cells must still emit its own <w:p>")
    }

    func testMutatedCellStillRoundTripsCorrectlyAmongSharedBlankSiblings() {
        var table = Table(rowCount: 2, columnCount: 2)
        var cell = table.rows[0].cells[0]
        cell.paragraphs = [Paragraph(text: "only this one")]
        table.rows[0].cells[0] = cell

        let xml = table.toXML()
        XCTAssertTrue(xml.contains("only this one"))
        // The other 3 cells' <w:p> must still be present and empty (no
        // stray content leaked in from the shared default, and no
        // corruption of the shared buffer for the cells that never
        // touched it).
        XCTAssertEqual(xml.components(separatedBy: "<w:p").count - 1, 4)
    }

    // MARK: - Memory: qualitative, generous-threshold regression guard

    /// Not a precise byte-count assertion (RSS measurement varies with
    /// allocator behavior across platforms/OS versions) — a generous
    /// regression guard: before the fix, 50,000 cells cost north of
    /// 5,800 bytes RESIDENT each (~290 MB total); after, well under 100
    /// bytes/cell (~5 MB total). This asserts the total stays under 20 MB
    /// — comfortably above the fixed measured value, comfortably below
    /// the pre-fix behavior, so a regression that re-inflates per-cell
    /// cost trips it long before ever approaching the old ~290 MB.
    func testBlankTableConstructionMemoryStaysSmall() {
        func residentBytes() -> Int {
            var info = mach_task_basic_info()
            var count = mach_msg_type_number_t(MemoryLayout<mach_task_basic_info>.size) / 4
            let result = withUnsafeMutablePointer(to: &info) {
                $0.withMemoryRebound(to: integer_t.self, capacity: Int(count)) {
                    task_info(mach_task_self_, task_flavor_t(MACH_TASK_BASIC_INFO), $0, &count)
                }
            }
            return result == KERN_SUCCESS ? Int(info.resident_size) : -1
        }

        let before = residentBytes()
        var table = Table(rowCount: 5000, columnCount: 10)
        let after = residentBytes()
        XCTAssertGreaterThanOrEqual(before, 0)
        XCTAssertGreaterThanOrEqual(after, 0)

        let delta = after - before
        XCTAssertLessThan(delta, 20_000_000,
                           "50,000-cell blank table grew resident memory by \(delta) bytes — expected well under 20MB after the shared-defaults fix (pre-fix this was ~290MB)")
        _ = table.rows.count  // keep alive until after measurement
    }
}
