// RenderGeometryTests.swift
// render-effect-semantics task 1.1 — geometry extraction over rendered PDFs
// (`docx-visual-diff-testing`, «Geometry measurement extends the harness
// beyond pixel ratios»). Committed fixtures with known ground truth
// (generated via CoreText at exact baselines): NO Word dependency, NO
// RUN_WORD_INTEGRATION gate — plain `swift test` must exercise these.
//
// Fixture ground truth (Fixtures/render/, US Letter 612x792):
//   render-five-lines.pdf  — 5 lines, 10pt Helvetica, baselines 14pt apart
//   render-single-line.pdf — 1 line (median line pitch undefined → nil)

import XCTest
import PDFKit
@testable import OOXMLSwift

final class RenderGeometryTests: XCTestCase {

    private func fixtureURL(_ name: String) -> URL {
        URL(fileURLWithPath: #filePath)
            .deletingLastPathComponent()
            .appendingPathComponent("Fixtures/render/\(name)")
    }

    private func document(_ name: String) throws -> PDFDocument {
        let url = fixtureURL(name)
        return try XCTUnwrap(PDFDocument(url: url),
                             "fixture \(name) must load — regenerate via task 1.1 generator")
    }

    /// Spec scenario "Page box and page count without Word".
    func testPageCountAndPageBox() throws {
        let pdf = try document("render-five-lines.pdf")
        XCTAssertEqual(RenderGeometry.pageCount(pdf: pdf), 1)
        let box = try XCTUnwrap(RenderGeometry.pageBox(pdf: pdf, page: 0))
        XCTAssertEqual(box.width, 612, accuracy: 0.01, "US Letter width")
        XCTAssertEqual(box.height, 792, accuracy: 0.01, "US Letter height")
    }

    /// Spec scenario "Line geometry from a committed fixture": line-box count
    /// matches the known 5 lines, and median line pitch matches the known
    /// 14pt baseline distance within measurement tolerance.
    func testLineBoxesAndMedianPitchOnFiveLineFixture() throws {
        let pdf = try document("render-five-lines.pdf")
        let boxes = RenderGeometry.lineBoxes(pdf: pdf, page: 0)
        XCTAssertEqual(boxes.count, 5, "fixture draws exactly 5 text lines")

        let pitch = try XCTUnwrap(RenderGeometry.medianLinePitch(pdf: pdf, page: 0),
                                  "5-line page must yield a pitch")
        XCTAssertEqual(pitch, 14.0, accuracy: 0.5,
                       "baselines are 14pt apart in the fixture")
    }

    /// Spec scenario "Sparse page yields nil line pitch": fewer than two text
    /// lines → nil (consumers that expected text treat nil as failure).
    func testSingleLinePageYieldsNilPitch() throws {
        let pdf = try document("render-single-line.pdf")
        XCTAssertEqual(RenderGeometry.lineBoxes(pdf: pdf, page: 0).count, 1)
        XCTAssertNil(RenderGeometry.medianLinePitch(pdf: pdf, page: 0),
                     "median pitch is undefined with fewer than 2 lines")
    }

    /// Out-of-range page index degrades to empty/nil, never traps.
    func testOutOfRangePageIsSafe() throws {
        let pdf = try document("render-single-line.pdf")
        XCTAssertNil(RenderGeometry.pageBox(pdf: pdf, page: 7))
        XCTAssertTrue(RenderGeometry.lineBoxes(pdf: pdf, page: 7).isEmpty)
        XCTAssertNil(RenderGeometry.medianLinePitch(pdf: pdf, page: 7))
    }
}
