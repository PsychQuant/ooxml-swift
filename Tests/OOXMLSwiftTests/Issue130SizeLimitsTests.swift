import XCTest
import Foundation
import ZIPFoundation
@testable import OOXMLSwift

/// PsychQuant/ooxml-swift#130 — bounds on what a package may expand to.
///
/// Two axes, and the issue's own numbers belong to different ones: the 3000:1
/// amplification is disk, the 1.7 GB RSS is memory (one expanded part read into
/// `Data` before parsing). Both are bounded here, and each is refused on its
/// own evidence — declared sizes before extraction, actual sizes after it.
final class Issue130SizeLimitsTests: XCTestCase {

    private var saved = ZipHelper.Limits()
    override func setUp() { super.setUp(); saved = ZipHelper.limits }
    override func tearDown() { ZipHelper.limits = saved; super.tearDown() }

    /// A package whose entries are honestly declared and huge.
    private func bomb(entryBytes: Int, entries: Int = 1) throws -> Data {
        let archive = try Archive(accessMode: .create)
        let payload = Data(repeating: 0x41, count: entryBytes)          // compresses to almost nothing
        try archive.addEntry(with: "word/document.xml", type: .file, uncompressedSize: Int64(payload.count),
                             compressionMethod: .deflate,
                             provider: { pos, size in payload.subdata(in: Int(pos)..<Int(pos) + size) })
        for i in 1..<entries {
            try archive.addEntry(with: "word/media/blob\(i).bin", type: .file, uncompressedSize: Int64(payload.count),
                                 compressionMethod: .deflate,
                                 provider: { pos, size in payload.subdata(in: Int(pos)..<Int(pos) + size) })
        }
        return archive.data ?? Data()
    }

    private func namespaceEntries(_ ns: String) -> Int {
        let dir = FileManager.default.temporaryDirectory.appendingPathComponent(ns)
        return (try? FileManager.default.contentsOfDirectory(atPath: dir.path))?.count ?? 0
    }

    func testAnEntryOverTheLimitIsRefusedBeforeAnythingIsWritten() throws {
        ZipHelper.limits.maximumEntryBytes = 1 << 20                    // 1 MB
        ZipHelper.limits.maximumCompressionRatio = .infinity            // the per-entry size is what is under test here
        let data = try bomb(entryBytes: 4 << 20)                        // 4 MB, honestly declared
        let ns = "i130-\(UUID().uuidString)"
        XCTAssertThrowsError(try ZipHelper.unzip(data: data, namespace: ns)) { error in
            let message = String(describing: error)
            XCTAssertTrue(message.contains("over the"), message)
            XCTAssertTrue(message.contains("limit"), message)
            XCTAssertTrue(message.contains("document.xml"), "the entry is named: \(message)")
        }
        // Refused from the declaration alone: nothing was extracted, so the
        // namespace holds nothing to clean up.
        XCTAssertEqual(namespaceEntries(ns), 0, "no extraction directory is created for a refused package")
        try? FileManager.default.removeItem(at: FileManager.default.temporaryDirectory.appendingPathComponent(ns))
    }

    func testTheTotalIsBoundedEvenWhenEveryEntryFits() throws {
        ZipHelper.limits.maximumEntryBytes = 4 << 20
        ZipHelper.limits.maximumTotalBytes = 6 << 20
        ZipHelper.limits.maximumCompressionRatio = .infinity            // the total is what is under test here
        let data = try bomb(entryBytes: 2 << 20, entries: 6)            // 6 x 2 MB = 12 MB total, each entry legal
        let ns = "i130-\(UUID().uuidString)"
        XCTAssertThrowsError(try ZipHelper.unzip(data: data, namespace: ns)) { error in
            XCTAssertTrue(String(describing: error).contains("in total"), String(describing: error))
        }
        try? FileManager.default.removeItem(at: FileManager.default.temporaryDirectory.appendingPathComponent(ns))
    }

    func testAnExtremeCompressionRatioIsRefused() throws {
        ZipHelper.limits.maximumEntryBytes = 64 << 20                   // not the binding limit here
        ZipHelper.limits.maximumTotalBytes = 64 << 20
        ZipHelper.limits.maximumCompressionRatio = 50                   // real corpus p99.9 is 32x
        let data = try bomb(entryBytes: 8 << 20)                        // a run of one byte: ratio far over 50x
        let ns = "i130-\(UUID().uuidString)"
        XCTAssertThrowsError(try ZipHelper.unzip(data: data, namespace: ns)) { error in
            XCTAssertTrue(String(describing: error).contains("expands"), String(describing: error))
        }
        try? FileManager.default.removeItem(at: FileManager.default.temporaryDirectory.appendingPathComponent(ns))
    }

    func testALegitimateDocumentIsNotRefusedByTheDefaults() throws {
        // The defaults come from the upper edge of a 738-document corpus
        // (largest entry 17.5 MB, largest package 26.4 MB, highest ratio 118x).
        // A package at that upper edge must still open.
        let d = ZipHelper.Limits()
        XCTAssertGreaterThan(d.maximumEntryBytes, Int64(18 << 20), "17.5 MB was a real entry")
        XCTAssertGreaterThan(d.maximumTotalBytes, Int64(27 << 20), "26.4 MB was a real package")
        XCTAssertGreaterThan(d.maximumCompressionRatio, 118, "117.8x was a real ratio")
        let archive = try Archive(accessMode: .create)
        let body = Data("<?xml version=\"1.0\"?><w:document xmlns:w=\"x\"><w:body/></w:document>".utf8)
        try archive.addEntry(with: "word/document.xml", type: .file, uncompressedSize: Int64(body.count),
                             provider: { pos, size in body.subdata(in: Int(pos)..<Int(pos) + size) })
        let ns = "i130-\(UUID().uuidString)"
        let out = try ZipHelper.unzip(data: archive.data ?? Data(), namespace: ns)
        ZipHelper.cleanup(out)
        try? FileManager.default.removeItem(at: FileManager.default.temporaryDirectory.appendingPathComponent(ns))
    }

    func testAPartTooLargeToHoldIsRefusedBeforeItIsRead() throws {
        // The memory axis. The file is written directly, so the extraction
        // limits are not what refuses it — `readPart` is.
        ZipHelper.limits.maximumPartBytes = 1 << 20
        let root = FileManager.default.temporaryDirectory.appendingPathComponent("i130-part-\(UUID().uuidString)")
        defer { ZipHelper.removeTreeForcibly(root) }
        try FileManager.default.createDirectory(at: root.appendingPathComponent("word"), withIntermediateDirectories: true)
        let part = root.appendingPathComponent("word/document.xml")
        try Data(repeating: 0x41, count: 4 << 20).write(to: part)
        XCTAssertThrowsError(try ZipHelper.readPart(at: part, describedAs: "word/document.xml")) { error in
            let message = String(describing: error)
            XCTAssertTrue(message.contains("limit for a single part"), message)
            // The part NAME legitimately contains "/" — what must not appear is
            // the extraction path it happens to live at.
            XCTAssertTrue(message.contains("word/document.xml"), message)
            XCTAssertFalse(message.contains(root.lastPathComponent), "no temporary path in the message: \(message)")
            XCTAssertFalse(message.contains("/var/"), message)
        }
        // Under the limit it reads normally.
        ZipHelper.limits.maximumPartBytes = 64 << 20
        XCTAssertEqual(try ZipHelper.readPart(at: part, describedAs: "word/document.xml").count, 4 << 20)
    }
}
