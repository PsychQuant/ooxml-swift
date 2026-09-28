import Foundation
import XCTest
@testable import OOXMLSwift

/// Regression tests for [PsychQuant/ooxml-swift#178](https://github.com/PsychQuant/ooxml-swift/issues/178).
///
/// `Drawing.from(widthPx:heightPx:imageId:name:)`, `Drawing.anchorFromPixels`,
/// and `WordDocument.updateImage` all multiplied a caller-supplied pixel
/// value by 9525 (px → EMU @ 96 DPI) with no range check. For a pixel value
/// beyond roughly `Int.max / 9525` (~9.7×10¹⁴), that multiplication TRAPS —
/// an integer-overflow crash that takes down the entire host process, not
/// just the one call (che-word-mcp#234 found this first, on the MCP side,
/// where a single malformed tool-call parameter could crash the whole
/// server — including every other open document's unsaved edits). Negative
/// or zero pixel values were not rejected either, though `wp:extent/@cx`/`@cy`
/// (`ST_PositiveCoordinate`) only allows `0` through `27273042316900`.
///
/// Fix, per the issue's own two options (chosen: keep the original
/// signatures, since changing them to `throws` would be source-breaking for
/// every existing call site and force a major version bump):
///
/// - `Drawing.validated(widthPx:heightPx:imageId:name:)` and
///   `Drawing.anchorValidated(widthPx:heightPx:imageId:position:name:)` are
///   NEW throwing constructors that validate before converting, throwing
///   `ImageError.dimensionOutOfRange` for anything non-positive or whose EMU
///   equivalent would overflow `Int` or exceed `ST_PositiveCoordinate`'s max.
/// - `Drawing.from`/`Drawing.anchorFromPixels` are kept (deprecated) for
///   source compatibility, but no longer TRAP: they clamp to
///   `Drawing.maxPositiveCoordinateEMU` (or `0`) instead of crashing.
/// - `WordDocument.updateImage` already throws, so it gained full
///   validation directly — no clamped fallback needed.
final class Issue178PixelToEMURangeCheckTests: XCTestCase {

    /// The smallest pixel value whose `px * 9525` genuinely overflows `Int`
    /// — one past `Int.max / 9525`.
    private var firstOverflowingPixelValue: Int {
        Int.max / Drawing.emuPerPixel + 2
    }

    // MARK: - Drawing.validated (new throwing API)

    func testValidatedNormalPixelValueSucceeds() throws {
        let drawing = try Drawing.validated(widthPx: 100, heightPx: 200, imageId: "rId1")
        XCTAssertEqual(drawing.width, 100 * Drawing.emuPerPixel)
        XCTAssertEqual(drawing.height, 200 * Drawing.emuPerPixel)
    }

    func testValidatedZeroPixelThrows() {
        XCTAssertThrowsError(try Drawing.validated(widthPx: 0, heightPx: 100, imageId: "rId1")) { error in
            guard case ImageError.dimensionOutOfRange = error else {
                XCTFail("expected ImageError.dimensionOutOfRange, got \(error)")
                return
            }
        }
    }

    func testValidatedNegativePixelThrows() {
        XCTAssertThrowsError(try Drawing.validated(widthPx: -5, heightPx: 100, imageId: "rId1")) { error in
            guard case ImageError.dimensionOutOfRange = error else {
                XCTFail("expected ImageError.dimensionOutOfRange, got \(error)")
                return
            }
        }
    }

    /// The exact crash from the issue: a pixel value whose conversion
    /// overflows `Int`. Before the fix, `Drawing.from` (the only
    /// constructor that existed) TRAPPED here — this is the RED case.
    /// `validated` must throw instead, cleanly.
    func testValidatedOverflowingPixelThrowsInsteadOfTrapping() {
        XCTAssertThrowsError(try Drawing.validated(widthPx: firstOverflowingPixelValue, heightPx: 100, imageId: "rId1")) { error in
            guard case ImageError.dimensionOutOfRange = error else {
                XCTFail("expected ImageError.dimensionOutOfRange, got \(error)")
                return
            }
        }
    }

    func testValidatedJustOverSTPositiveCoordinateMaxThrows() {
        // One pixel past the largest value whose EMU conversion still fits
        // ST_PositiveCoordinate's range.
        let px = Drawing.maxPositiveCoordinateEMU / Drawing.emuPerPixel + 1
        XCTAssertThrowsError(try Drawing.validated(widthPx: px, heightPx: 100, imageId: "rId1")) { error in
            guard case ImageError.dimensionOutOfRange = error else {
                XCTFail("expected ImageError.dimensionOutOfRange, got \(error)")
                return
            }
        }
    }

    func testValidatedAtSTPositiveCoordinateMaxSucceeds() throws {
        let px = Drawing.maxPositiveCoordinateEMU / Drawing.emuPerPixel
        let drawing = try Drawing.validated(widthPx: px, heightPx: 100, imageId: "rId1")
        XCTAssertLessThanOrEqual(drawing.width, Drawing.maxPositiveCoordinateEMU)
    }

    func testAnchorValidatedOverflowingPixelThrowsInsteadOfTrapping() {
        XCTAssertThrowsError(
            try Drawing.anchorValidated(widthPx: firstOverflowingPixelValue, heightPx: 100, imageId: "rId1")
        ) { error in
            guard case ImageError.dimensionOutOfRange = error else {
                XCTFail("expected ImageError.dimensionOutOfRange, got \(error)")
                return
            }
        }
    }

    func testAnchorValidatedNormalPixelValueSucceeds() throws {
        let drawing = try Drawing.anchorValidated(widthPx: 50, heightPx: 60, imageId: "rId1")
        XCTAssertEqual(drawing.type, .anchor)
        XCTAssertEqual(drawing.width, 50 * Drawing.emuPerPixel)
    }

    // MARK: - Deprecated `from`/`anchorFromPixels`: must not trap, must clamp

    /// The deprecated, still-source-compatible constructor: must not crash
    /// the process for an overflowing pixel value (unlike before the fix),
    /// even though — being non-throwing — it cannot report the problem to
    /// the caller either. It clamps instead.
    func testDeprecatedFromOverflowingPixelDoesNotTrapAndClamps() {
        let drawing = Drawing.from(widthPx: firstOverflowingPixelValue, heightPx: 100, imageId: "rId1")
        XCTAssertEqual(drawing.width, Drawing.maxPositiveCoordinateEMU)
    }

    func testDeprecatedFromNormalPixelValueUnaffected() {
        let drawing = Drawing.from(widthPx: 100, heightPx: 200, imageId: "rId1")
        XCTAssertEqual(drawing.width, 100 * Drawing.emuPerPixel)
        XCTAssertEqual(drawing.height, 200 * Drawing.emuPerPixel)
    }

    func testDeprecatedFromNegativePixelDoesNotTrapAndClampsToZero() {
        let drawing = Drawing.from(widthPx: -5, heightPx: 100, imageId: "rId1")
        XCTAssertEqual(drawing.width, 0)
    }

    func testDeprecatedAnchorFromPixelsOverflowingPixelDoesNotTrapAndClamps() {
        let drawing = Drawing.anchorFromPixels(widthPx: firstOverflowingPixelValue, heightPx: 100, imageId: "rId1")
        XCTAssertEqual(drawing.width, Drawing.maxPositiveCoordinateEMU)
    }

    // MARK: - WordDocument.updateImage (already throws — full validation, no clamping)

    func testUpdateImageOverflowingPixelThrowsInsteadOfTrapping() throws {
        var doc = WordDocument()
        let drawing = Drawing(width: 100, height: 100, imageId: "rId1")
        let run = Run.withDrawing(drawing)
        doc.body.children = [.paragraph(Paragraph(runs: [run]))]

        XCTAssertThrowsError(try doc.updateImage(imageId: "rId1", widthPx: firstOverflowingPixelValue)) { error in
            guard case ImageError.dimensionOutOfRange = error else {
                XCTFail("expected ImageError.dimensionOutOfRange, got \(error)")
                return
            }
        }
    }

    func testUpdateImageZeroPixelThrows() throws {
        var doc = WordDocument()
        let drawing = Drawing(width: 100, height: 100, imageId: "rId1")
        let run = Run.withDrawing(drawing)
        doc.body.children = [.paragraph(Paragraph(runs: [run]))]

        XCTAssertThrowsError(try doc.updateImage(imageId: "rId1", widthPx: 0)) { error in
            guard case ImageError.dimensionOutOfRange = error else {
                XCTFail("expected ImageError.dimensionOutOfRange, got \(error)")
                return
            }
        }
    }

    func testUpdateImageNormalPixelValueSucceeds() throws {
        var doc = WordDocument()
        let drawing = Drawing(width: 100, height: 100, imageId: "rId1")
        let run = Run.withDrawing(drawing)
        doc.body.children = [.paragraph(Paragraph(runs: [run]))]

        try doc.updateImage(imageId: "rId1", widthPx: 300, heightPx: 400)
        guard case .paragraph(let para) = doc.body.children[0],
              let updated = para.runs[0].drawing else {
            XCTFail("expected updated drawing")
            return
        }
        XCTAssertEqual(updated.width, 300 * Drawing.emuPerPixel)
        XCTAssertEqual(updated.height, 400 * Drawing.emuPerPixel)
    }

    /// Sanity: `updateImage` failing on the WIDTH validation must not have
    /// partially applied the height, and must not have marked the document
    /// dirty (mirrors the "zero side effects on refusal" discipline used
    /// elsewhere in this codebase, e.g. #196's pre-write validation).
    func testUpdateImageWidthValidationFailureLeavesDrawingUnchanged() throws {
        var doc = WordDocument()
        let drawing = Drawing(width: 111, height: 222, imageId: "rId1")
        let run = Run.withDrawing(drawing)
        doc.body.children = [.paragraph(Paragraph(runs: [run]))]

        XCTAssertThrowsError(try doc.updateImage(imageId: "rId1", widthPx: firstOverflowingPixelValue, heightPx: 999))
        guard case .paragraph(let para) = doc.body.children[0],
              let unchanged = para.runs[0].drawing else {
            XCTFail("expected drawing to still be present")
            return
        }
        XCTAssertEqual(unchanged.width, 111)
        XCTAssertEqual(unchanged.height, 222)
    }
}
