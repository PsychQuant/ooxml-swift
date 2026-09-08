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

    /// Limits are passed to each call, never installed process-wide: a shared
    /// mutable policy let one test's override leak into another's extraction.
    private func limits(entry: Int64 = 256 << 20, total: Int64 = 512 << 20,
                        ratio: Double = 500, part: Int64 = 256 << 20) -> ZipHelper.Limits {
        ZipHelper.Limits(maximumEntryBytes: entry, maximumTotalBytes: total,
                         maximumCompressionRatio: ratio, maximumPartBytes: part)
    }

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
        let l = limits(entry: 1 << 20, ratio: .infinity)                // the per-entry size is what is under test here
        let data = try bomb(entryBytes: 4 << 20)                        // 4 MB, honestly declared
        let ns = "i130-\(UUID().uuidString)"
        XCTAssertThrowsError(try ZipHelper.unzip(data: data, namespace: ns, limits: l)) { error in
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
        let l = limits(entry: 4 << 20, total: 6 << 20, ratio: .infinity) // the total is what is under test here
        let data = try bomb(entryBytes: 2 << 20, entries: 6)            // 6 x 2 MB = 12 MB total, each entry legal
        let ns = "i130-\(UUID().uuidString)"
        XCTAssertThrowsError(try ZipHelper.unzip(data: data, namespace: ns, limits: l)) { error in
            XCTAssertTrue(String(describing: error).contains("in total"), String(describing: error))
        }
        try? FileManager.default.removeItem(at: FileManager.default.temporaryDirectory.appendingPathComponent(ns))
    }

    func testAnExtremeCompressionRatioIsRefused() throws {
        let l = limits(entry: 64 << 20, total: 64 << 20, ratio: 50)     // real corpus p99.9 is 32x
        let data = try bomb(entryBytes: 8 << 20)                        // a run of one byte: ratio far over 50x
        let ns = "i130-\(UUID().uuidString)"
        XCTAssertThrowsError(try ZipHelper.unzip(data: data, namespace: ns, limits: l)) { error in
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
        let small = limits(part: 1 << 20)
        let root = FileManager.default.temporaryDirectory.appendingPathComponent("i130-part-\(UUID().uuidString)")
        defer { ZipHelper.removeTreeForcibly(root) }
        try FileManager.default.createDirectory(at: root.appendingPathComponent("word"), withIntermediateDirectories: true)
        let part = root.appendingPathComponent("word/document.xml")
        try Data(repeating: 0x41, count: 4 << 20).write(to: part)
        XCTAssertThrowsError(try ZipHelper.readPart(at: part, describedAs: "word/document.xml", limits: small)) { error in
            let message = String(describing: error)
            XCTAssertTrue(message.contains("limit for a single part"), message)
            // The part NAME legitimately contains "/" — what must not appear is
            // the extraction path it happens to live at.
            XCTAssertTrue(message.contains("word/document.xml"), message)
            XCTAssertFalse(message.contains(root.lastPathComponent), "no temporary path in the message: \(message)")
            XCTAssertFalse(message.contains("/var/"), message)
        }
        // Under the limit it reads normally.
        XCTAssertEqual(try ZipHelper.readPart(at: part, describedAs: "word/document.xml",
                                              limits: limits(part: 64 << 20)).count, 4 << 20)
    }

    // MARK: - A central directory that lies

    /// Build a package, then rewrite its central directory to declare
    /// `uncompressedSize` for the one entry. The archive's real bytes are
    /// untouched; only the declaration changes — which is the case the code's
    /// own comment calls out as possible.
    private func packageDeclaring(_ declared: UInt64, actualBytes: Int, stored: Bool = false) throws -> Data {
        let archive = try Archive(accessMode: .create)
        let payload = Data(repeating: 0x41, count: actualBytes)
        try archive.addEntry(with: "word/document.xml", type: .file, uncompressedSize: Int64(payload.count),
                             compressionMethod: stored ? .none : .deflate,
                             provider: { pos, size in payload.subdata(in: Int(pos)..<Int(pos) + size) })
        var raw = [UInt8](archive.data ?? Data())

        // Central directory header: usize is 4 bytes at +24, name length at +28,
        // extra length at +30, the extra field itself after the name.
        guard let i = raw.firstRange(of: Array("PK\u{01}\u{02}".utf8)) else {
            XCTFail("no central directory"); return Data()
        }
        let base = i.lowerBound
        func u16(_ at: Int) -> Int { Int(raw[at]) | Int(raw[at + 1]) << 8 }
        let nameLen = u16(base + 28), extraLen = u16(base + 30)

        if declared <= UInt64(UInt32.max) {
            for k in 0..<4 { raw[base + 24 + k] = UInt8((declared >> (8 * UInt64(k))) & 0xFF) }
            return Data(raw)
        }
        // Over 32 bits: signal ZIP64 in the 32-bit field and carry the real
        // value in a ZIP64 extended-information extra field (header 0x0001).
        for k in 0..<4 { raw[base + 24 + k] = 0xFF }
        var extra: [UInt8] = [0x01, 0x00, 0x08, 0x00]
        for k in 0..<8 { extra.append(UInt8((declared >> (8 * UInt64(k))) & 0xFF)) }
        raw.insert(contentsOf: extra, at: base + 46 + nameLen + extraLen)
        let newExtra = extraLen + extra.count
        raw[base + 30] = UInt8(newExtra & 0xFF); raw[base + 31] = UInt8((newExtra >> 8) & 0xFF)
        if let e = raw.firstRange(of: Array("PK\u{05}\u{06}".utf8), in: base..<raw.count) {
            let at = e.lowerBound + 12
            let size = Int(raw[at]) | Int(raw[at+1]) << 8 | Int(raw[at+2]) << 16 | Int(raw[at+3]) << 24
            let grown = size + extra.count
            for k in 0..<4 { raw[at + k] = UInt8((grown >> (8 * k)) & 0xFF) }
        }
        return Data(raw)
    }

    /// A forged **compressedSize** must be refused too — and this one reaches
    /// the crash at the SHIPPING DEFAULTS, with no `.max` limits anywhere.
    ///
    /// Measured: a 161-byte package whose ZIP64 record declares
    /// `compressedSize = UInt64.max` and an honest `uncompressedSize` of 64
    /// kills `main` (v3.7.0 as released) with "Not enough bits to represent the
    /// passed value". Nothing in the pre-scan could catch it before this
    /// change: the only place `compressed` was touched is the ratio check, and
    /// `declared / compressed` tends to zero as `compressed` grows, so it
    /// always passed. `Limits` has no field bounding compressed size at all.
    ///
    /// The trap is inside ZIPFoundation (`Archive+Helpers.swift`), whose
    /// `guard size <= .max` on a `UInt64` is vacuously true — the same shape as
    /// `readUncompressed`'s. `Archive+Reading.swift`'s `guard entry.dataOffset
    /// <= .max` is a third instance of that spelling; it is NOT covered here,
    /// because this pre-scan guards the two size fields only.
    func testAForgedCompressedSizeIsRefusedAtTheDefaults() throws {
        let data = try packageDeclaringCompressedSize(UInt64.max, uncompressed: 64)
        let ns = "i130-\(UUID().uuidString)"
        XCTAssertThrowsError(try ZipHelper.unzip(data: data, namespace: ns)) { error in
            XCTAssertTrue(String(describing: error).contains("larger than this library can represent"),
                          String(describing: error))
        }
        try? FileManager.default.removeItem(at: FileManager.default.temporaryDirectory.appendingPathComponent(ns))
    }

    /// A package whose ZIP64 record declares an honest uncompressed size and a
    /// forged compressed one.
    private func packageDeclaringCompressedSize(_ compressed: UInt64, uncompressed: UInt64) throws -> Data {
        let archive = try Archive(accessMode: .create)
        let payload = Data(repeating: 0x41, count: 64)
        try archive.addEntry(with: "word/document.xml", type: .file, uncompressedSize: Int64(payload.count),
                             compressionMethod: .deflate,
                             provider: { pos, size in payload.subdata(in: Int(pos)..<Int(pos) + size) })
        var raw = [UInt8](archive.data ?? Data())
        guard let r = raw.firstRange(of: Array("PK\u{01}\u{02}".utf8)) else { XCTFail("no central directory"); return Data() }
        let base = r.lowerBound
        func u16(_ at: Int) -> Int { Int(raw[at]) | Int(raw[at + 1]) << 8 }
        let nameLen = u16(base + 28), extraLen = u16(base + 30)
        for k in 0..<4 { raw[base + 20 + k] = 0xFF }        // csize → ZIP64
        for k in 0..<4 { raw[base + 24 + k] = 0xFF }        // usize → ZIP64
        var extra: [UInt8] = [0x01, 0x00, 0x10, 0x00]       // header 0x0001, 16 bytes
        for k in 0..<8 { extra.append(UInt8((uncompressed >> (8 * UInt64(k))) & 0xFF)) }
        for k in 0..<8 { extra.append(UInt8((compressed   >> (8 * UInt64(k))) & 0xFF)) }
        raw.insert(contentsOf: extra, at: base + 46 + nameLen + extraLen)
        let ne = extraLen + extra.count
        raw[base + 30] = UInt8(ne & 0xFF); raw[base + 31] = UInt8((ne >> 8) & 0xFF)
        if let e = raw.firstRange(of: Array("PK\u{05}\u{06}".utf8), in: base..<raw.count) {
            let at = e.lowerBound + 12
            let size = Int(raw[at]) | Int(raw[at+1]) << 8 | Int(raw[at+2]) << 16 | Int(raw[at+3]) << 24
            let grown = size + extra.count
            for k in 0..<4 { raw[at + k] = UInt8((grown >> (8 * k)) & 0xFF) }
        }
        return Data(raw)
    }

    /// A declaration above `Int64.max` must be REFUSED, not trap — under the
    /// DEFAULT limits, where the size check would also have caught it.
    ///
    /// `Int64(entry.uncompressedSize)` on the UInt64 the central directory
    /// declares crashed the process — "Not enough bits to represent the passed
    /// value" — on a 152-byte package, before reaching the check meant to
    /// refuse it, and on a package v3.7.0 extracted without incident. The
    /// unlimited-limits case, where nothing else would catch it, is
    /// `testAnUnrepresentableDeclarationIsRefusedForBothCompressionMethods`.
    func testADeclarationAboveInt64MaxIsRefusedRatherThanTrapping() throws {
        let data = try packageDeclaring(UInt64.max, actualBytes: 32)
        let ns = "i130-\(UUID().uuidString)"
        XCTAssertThrowsError(try ZipHelper.unzip(data: data, namespace: ns, limits: limits())) { error in
            let m = String(describing: error)
            XCTAssertTrue(m.contains("larger than this library can represent") || m.contains("over the"), m)
        }
        try? FileManager.default.removeItem(at: FileManager.default.temporaryDirectory.appendingPathComponent(ns))
    }

    /// An archive that UNDERSTATES its entry is caught by measuring the tree,
    /// not by trusting the declaration. Deleting the post-extraction
    /// accounting leaves every other test in this file green.
    func testAnArchiveThatUnderstatesItsEntryIsStillRefused() throws {
        let data = try packageDeclaring(1024, actualBytes: 4 << 20)      // declares 1 KB, writes 4 MB
        let ns = "i130-\(UUID().uuidString)"
        let l = limits(entry: 2 << 20, total: 2 << 20, ratio: .infinity) // the declaration passes these; the reality does not
        XCTAssertThrowsError(try ZipHelper.unzip(data: data, namespace: ns, limits: l)) { error in
            let m = String(describing: error)
            // Refused on what it actually WROTE, not on its declaration — the
            // declaration passed. Either actual check may fire first (per-entry
            // or total), which is why both spellings are accepted; what must
            // not happen is the package being accepted.
            XCTAssertTrue(m.contains("on disk") || m.contains("over the"),
                          "refused on what it actually wrote, not on its declaration: \(m)")
        }
        try? FileManager.default.removeItem(at: FileManager.default.temporaryDirectory.appendingPathComponent(ns))
    }

    // MARK: - The limits reach the callers, not just ZipHelper

    /// Driving a limit through `DocxReader.read`. Reverting its call site to a
    /// plain `Data(contentsOf:)` / unlimited `unzip` leaves the direct
    /// `ZipHelper` tests green; this one fails.
    func testDocxReaderRefusesAPackageOverItsLimits() throws {
        let data = try bomb(entryBytes: 4 << 20)
        let file = FileManager.default.temporaryDirectory.appendingPathComponent("i130-\(UUID().uuidString).docx")
        try data.write(to: file)
        defer { try? FileManager.default.removeItem(at: file) }
        XCTAssertThrowsError(try DocxReader.read(from: file, limits: limits(entry: 1 << 20, ratio: .infinity))) { error in
            XCTAssertTrue(String(describing: error).contains("over the"), String(describing: error))
        }
    }

    /// The same through `PackageInspector`, whose extraction is a second,
    /// independent call site.
    func testPackageInspectorRefusesAPackageOverItsLimits() throws {
        let data = try bomb(entryBytes: 4 << 20)
        XCTAssertThrowsError(try PackageInspector.imageConsistencyReport(of: data,
                                                                        limits: limits(entry: 1 << 20, ratio: .infinity))) { error in
            XCTAssertTrue(String(describing: error).contains("could not be extracted") || String(describing: error).contains("over the"),
                          String(describing: error))
        }
    }

    /// The DEFAULT values bind. Raising every default to effectively unlimited
    /// left every other test here green, because each one names its own limits.
    func testTheDefaultsThemselvesRefuseAnAmplifiedPackage() throws {
        let d = ZipHelper.defaultLimits
        XCTAssertLessThan(d.maximumEntryBytes, Int64(1) << 40, "a default must actually bind")
        XCTAssertLessThan(d.maximumCompressionRatio, 100_000)
        // 600 MB of one byte: over the 512 MB default total, declared honestly.
        let data = try bomb(entryBytes: 100 << 20, entries: 6)
        let ns = "i130-\(UUID().uuidString)"
        XCTAssertThrowsError(try ZipHelper.unzip(data: data, namespace: ns)) { error in
            XCTAssertTrue(String(describing: error).contains("in total") || String(describing: error).contains("expands"),
                          String(describing: error))
        }
        try? FileManager.default.removeItem(at: FileManager.default.temporaryDirectory.appendingPathComponent(ns))
    }

    /// Two declarations of `Int64.max - 1` under limits expressed as `.max`.
    ///
    /// `declaredTotal += declared` TRAPPED here — Swift traps on overflow, and
    /// `.max` is the only way `Limits` offers to say "no limit", a spelling this
    /// repo's own pathological-fixture tests already use. Measured SIGTRAP on
    /// the second entry. An overflow is "over the limit", so it is refused.
    func testATotalThatOverflowsIsRefusedRatherThanTrapping() throws {
        // two entries each declaring Int64.max - 1, with limits expressed as .max
        let arch = try Archive(accessMode: .create)
        let payload = Data(repeating: 0x41, count: 64)
        for n in ["word/document.xml", "word/media/a.bin"] {
            try arch.addEntry(with: n, type: .file, uncompressedSize: Int64(payload.count),
                              compressionMethod: .deflate,
                              provider: { pos, size in payload.subdata(in: Int(pos)..<Int(pos)+size) })
        }
        var raw = [UInt8](arch.data ?? Data())
        // patch both central-directory records to declare Int64.max - 1
        var idx = 0, patched = 0
        while let r = raw[idx...].firstRange(of: Array("PK\u{01}\u{02}".utf8)) {
            let base = r.lowerBound
            let v = UInt64(bitPattern: Int64.max - 1)
            let nameLen = Int(raw[base+28]) | Int(raw[base+29]) << 8
            let extraLen = Int(raw[base+30]) | Int(raw[base+31]) << 8
            for k in 0..<4 { raw[base + 24 + k] = 0xFF }
            var extra: [UInt8] = [0x01, 0x00, 0x08, 0x00]
            for k in 0..<8 { extra.append(UInt8((v >> (8 * UInt64(k))) & 0xFF)) }
            raw.insert(contentsOf: extra, at: base + 46 + nameLen + extraLen)
            let ne = extraLen + extra.count
            raw[base+30] = UInt8(ne & 0xFF); raw[base+31] = UInt8((ne >> 8) & 0xFF)
            idx = base + 46 + nameLen + ne
            patched += 1
            if patched >= 2 { break }
        }
        let unlimited = ZipHelper.Limits(maximumEntryBytes: .max, maximumTotalBytes: .max,
                                         maximumCompressionRatio: .infinity, maximumPartBytes: .max)
        let ns = "f5-\(UUID().uuidString)"
        XCTAssertEqual(patched, 2, "both entries must carry the forged declaration")
        XCTAssertThrowsError(try ZipHelper.unzip(data: Data(raw), namespace: ns, limits: unlimited)) { error in
            XCTAssertTrue(String(describing: error).contains("in total"), String(describing: error))
        }
        try? FileManager.default.removeItem(at: FileManager.default.temporaryDirectory.appendingPathComponent(ns))
    }

    /// A declaration that does not fit in `Int64` is refused, for BOTH
    /// compression methods, before anything converts it.
    ///
    /// Clamping it instead let it pass this pre-scan whenever the caller's
    /// limit was `.max`, and then ZIPFoundation's own unclamped conversions
    /// trapped: `totalUnitCountForReading` (live as soon as a progress object
    /// is passed) and `readUncompressed` (whose `guard size <= .max` on a
    /// `UInt64` is vacuously true).
    func testAnUnrepresentableDeclarationIsRefusedForBothCompressionMethods() throws {
        for stored in [false, true] {
            let data = try packageDeclaring(UInt64.max, actualBytes: 32, stored: stored)
            let ns = "i130-\(UUID().uuidString)"
            let unlimited = ZipHelper.Limits(maximumEntryBytes: .max, maximumTotalBytes: .max,
                                             maximumCompressionRatio: .infinity, maximumPartBytes: .max)
            XCTAssertThrowsError(try ZipHelper.unzip(data: data, namespace: ns, limits: unlimited),
                                 "stored=\(stored)") { error in
                XCTAssertTrue(String(describing: error).contains("larger than this library can represent"),
                              String(describing: error))
            }
            try? FileManager.default.removeItem(at: FileManager.default.temporaryDirectory.appendingPathComponent(ns))
        }
    }
}
