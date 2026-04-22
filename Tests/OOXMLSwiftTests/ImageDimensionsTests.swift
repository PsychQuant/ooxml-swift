import XCTest
@testable import OOXMLSwift

final class ImageDimensionsTests: XCTestCase {

    private func writeTempFile(bytes: [UInt8], ext: String) throws -> String {
        let url = FileManager.default.temporaryDirectory.appendingPathComponent("img-\(UUID().uuidString).\(ext)")
        try Data(bytes).write(to: url)
        return url.path
    }

    // MARK: PNG

    func testDetectPNG800x600() throws {
        // Signature + IHDR length + "IHDR" + width(0x0320=800) + height(0x0258=600) + rest of IHDR
        let bytes: [UInt8] = [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0x00, 0x00, 0x00, 0x0D,
            0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x03, 0x20,
            0x00, 0x00, 0x02, 0x58,
            0x08, 0x02, 0x00, 0x00, 0x00
        ]
        let path = try writeTempFile(bytes: bytes, ext: "png")
        defer { try? FileManager.default.removeItem(atPath: path) }

        let dims = try ImageDimensions.detect(path: path)
        XCTAssertEqual(dims.widthPx, 800)
        XCTAssertEqual(dims.heightPx, 600)
        XCTAssertEqual(dims.aspectRatio, 800.0 / 600.0, accuracy: 0.0001)
    }

    func testInvalidPNGThrows() throws {
        let bytes: [UInt8] = [0x00, 0x01, 0x02]  // not PNG signature
        let path = try writeTempFile(bytes: bytes, ext: "png")
        defer { try? FileManager.default.removeItem(atPath: path) }

        XCTAssertThrowsError(try ImageDimensions.detect(path: path)) { error in
            XCTAssertEqual(error as? ImageDimensionsError, .invalidPNG)
        }
    }

    // MARK: JPEG

    func testDetectJPEG800x600() throws {
        // SOI + APP0(JFIF) + SOF0(height=600 width=800) + EOI
        let bytes: [UInt8] = [
            0xFF, 0xD8,                                     // SOI
            0xFF, 0xE0, 0x00, 0x10,                         // APP0 length=16
            0x4A, 0x46, 0x49, 0x46, 0x00,                   // "JFIF\0"
            0x01, 0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00,
            0xFF, 0xC0,                                     // SOF0
            0x00, 0x11,                                     // length 17
            0x08,                                           // precision 8
            0x02, 0x58,                                     // height 600
            0x03, 0x20,                                     // width 800
            0x03,                                           // 3 components
            0x01, 0x22, 0x00,
            0x02, 0x11, 0x01,
            0x03, 0x11, 0x01,
            0xFF, 0xD9                                      // EOI
        ]
        let path = try writeTempFile(bytes: bytes, ext: "jpg")
        defer { try? FileManager.default.removeItem(atPath: path) }

        let dims = try ImageDimensions.detect(path: path)
        XCTAssertEqual(dims.widthPx, 800)
        XCTAssertEqual(dims.heightPx, 600)
    }

    // MARK: Error paths

    func testUnsupportedFormatThrows() throws {
        let bytes: [UInt8] = [0x00]
        let path = try writeTempFile(bytes: bytes, ext: "tiff")
        defer { try? FileManager.default.removeItem(atPath: path) }

        XCTAssertThrowsError(try ImageDimensions.detect(path: path)) { error in
            guard case ImageDimensionsError.unsupportedFormat(let ext) = error else {
                XCTFail("Expected unsupportedFormat, got \(error)"); return
            }
            XCTAssertEqual(ext, "tiff")
        }
    }

    func testFileNotFoundThrows() {
        XCTAssertThrowsError(try ImageDimensions.detect(path: "/tmp/definitely-not-here-\(UUID().uuidString).png")) { error in
            guard case ImageDimensionsError.fileNotFound = error else {
                XCTFail("Expected fileNotFound, got \(error)"); return
            }
        }
    }

    func testAspectRatioZeroHeightReturnsZero() {
        let d = ImageDimensions(widthPx: 100, heightPx: 0)
        XCTAssertEqual(d.aspectRatio, 0)
    }
}
