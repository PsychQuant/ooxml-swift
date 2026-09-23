import XCTest
import ZIPFoundation
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
        // The reader namespace is shared by every process on the machine (a
        // second `swift test`, a running MCP server) and `TMPDIR` cannot move it
        // on macOS (`NSTemporaryDirectory()` ignores it — verify R6 regression
        // N-R6-2), so a bare before/after listing flaps under concurrent
        // readers (N-R6-1). The document carries a per-test marker; only an
        // item that contains it counts as OUR leak.
        let macdocTempRoot = FileManager.default.temporaryDirectory
            .appendingPathComponent(ZipHelper.readerNamespace)
        let marker = "writeData-leak-marker-\(UUID().uuidString)"

        let beforeFiles = Set(((try? FileManager.default.contentsOfDirectory(
            at: macdocTempRoot, includingPropertiesForKeys: nil)) ?? []).map(\.lastPathComponent))

        var document = WordDocument()
        document.appendParagraph(Paragraph(text: marker))
        _ = try DocxWriter.writeData(document)

        let afterFiles = ((try? FileManager.default.contentsOfDirectory(
            at: macdocTempRoot, includingPropertiesForKeys: nil)) ?? [])
        let markerBytes = Data(marker.utf8)
        func containsMarker(_ item: URL) -> Bool {
            let enumerator = FileManager.default.enumerator(at: item, includingPropertiesForKeys: [.isRegularFileKey])
            var files = [item]
            while let next = enumerator?.nextObject() as? URL { files.append(next) }
            return files.contains { url in
                guard (try? url.resourceValues(forKeys: [.isRegularFileKey]))?.isRegularFile == true,
                      let bytes = try? Data(contentsOf: url) else { return false }
                if bytes.range(of: markerBytes) != nil { return true }
                // A leaked private copy is a zip: look inside it too.
                guard let archive = try? Archive(data: bytes, accessMode: .read), let entry = archive["word/document.xml"] else { return false }
                var inner = Data(); _ = try? archive.extract(entry) { inner.append($0) }
                return inner.range(of: markerBytes) != nil
            }
        }
        let leaked = afterFiles.filter { !beforeFiles.contains($0.lastPathComponent) && containsMarker($0) }.map(\.lastPathComponent)

        XCTAssertTrue(leaked.isEmpty,
                      "writeData SHALL NOT leave files in the reader's extraction namespace; leaked: \(leaked)")
    }
}
