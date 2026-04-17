import XCTest
@testable import OOXMLSwift

final class DocxWriterWriteDataTests: XCTestCase {

    private var tempDir: URL!

    override func setUp() {
        super.setUp()
        tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("DocxWriterWriteDataTests-\(UUID().uuidString)")
        try? FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
    }

    override func tearDown() {
        try? FileManager.default.removeItem(at: tempDir)
        super.tearDown()
    }

    // MARK: - Task 1.1: writeData returns a valid ZIP archive

    func testWriteDataReturnsZipMagicBytes() throws {
        let document = WordDocument()
        let data = try DocxWriter.writeData(document)

        XCTAssertGreaterThanOrEqual(data.count, 4, "writeData SHALL return non-trivial data")
        XCTAssertEqual(data[0], 0x50, "byte 0 SHALL be 'P'")
        XCTAssertEqual(data[1], 0x4B, "byte 1 SHALL be 'K'")
        XCTAssertEqual(data[2], 0x03, "byte 2 SHALL be 0x03 (ZIP local file header)")
        XCTAssertEqual(data[3], 0x04, "byte 3 SHALL be 0x04 (ZIP local file header)")
    }

    // MARK: - Task 1.4: writeData and write produce identical bytes

    func testWriteDataAndWriteProduceByteEqualOutput() throws {
        let document = WordDocument()

        let writeURL = tempDir.appendingPathComponent("via-write.docx")
        try DocxWriter.write(document, to: writeURL)
        let fileBytes = try Data(contentsOf: writeURL)

        let dataBytes = try DocxWriter.writeData(document)

        XCTAssertEqual(dataBytes, fileBytes,
                       "writeData SHALL produce byte-equal output to write(_:to:)")
    }

    // MARK: - Task 1.5: writeData performs no disk I/O that persists

    func testWriteDataLeavesNoFilesInTempDir() throws {
        let macdocTempRoot = FileManager.default.temporaryDirectory
            .appendingPathComponent("che-word-mcp")

        let beforeFiles = (try? FileManager.default.contentsOfDirectory(
            at: macdocTempRoot, includingPropertiesForKeys: nil)) ?? []

        let document = WordDocument()
        _ = try DocxWriter.writeData(document)

        let afterFiles = (try? FileManager.default.contentsOfDirectory(
            at: macdocTempRoot, includingPropertiesForKeys: nil)) ?? []

        let created = Set(afterFiles.map(\.lastPathComponent))
            .subtracting(beforeFiles.map(\.lastPathComponent))

        XCTAssertTrue(created.isEmpty,
                      "writeData SHALL NOT leave files in che-word-mcp tempdir; leaked: \(created)")
    }
}
