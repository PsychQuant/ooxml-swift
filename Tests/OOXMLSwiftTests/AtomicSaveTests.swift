import XCTest
@testable import OOXMLSwift

/// Phase 1 of `che-word-mcp-save-durability-stack` (closes che-word-mcp#36).
///
/// Spec: `openspec/changes/che-word-mcp-save-durability-stack/specs/ooxml-atomic-save/spec.md`
/// Requirement: "DocxWriter.write performs atomic-rename save".
///
/// 4 scenarios from the spec:
/// 1. Successful save replaces target atomically (size+SHA256 transition cleanly)
/// 2. Throw mid-write preserves original (no orphan tmp file remains)
/// 3. Process killed mid-write preserves original (orphan tmp may remain, original intact)
/// 4. Fresh write to non-existent path (no extra files in parent)
/// + temp-file naming pattern `<url>.tmp.<UUID>` invariant.
final class AtomicSaveTests: XCTestCase {

    private var tempDir: URL!

    override func setUp() {
        super.setUp()
        tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("AtomicSaveTests-\(UUID().uuidString)")
        try? FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
    }

    override func tearDown() {
        if let tempDir = tempDir, FileManager.default.fileExists(atPath: tempDir.path) {
            // Restore writability before removal in case a test chmod'd the dir.
            _ = try? FileManager.default.setAttributes(
                [.posixPermissions: 0o755], ofItemAtPath: tempDir.path
            )
            try? FileManager.default.removeItem(at: tempDir)
        }
        super.tearDown()
    }

    // MARK: - Scenario 1: Successful save replaces target atomically

    func testSuccessfulSaveReplacesTargetAtomically() throws {
        let url = tempDir.appendingPathComponent("test.docx")

        // Seed the target with "original" docx bytes (using a fresh document A).
        var originalDoc = WordDocument()
        originalDoc.body.children.append(.paragraph(Paragraph(text: "ORIGINAL")))
        let originalBytes = try DocxWriter.writeData(originalDoc)
        try originalBytes.write(to: url)
        let originalSha = sha256(originalBytes)

        // Now write a different document via the API under test.
        var modifiedDoc = WordDocument()
        modifiedDoc.body.children.append(.paragraph(Paragraph(text: "MODIFIED — longer content for distinct bytes")))

        try DocxWriter.write(modifiedDoc, to: url)

        // Post-condition: target has new bytes, distinct from original.
        let postBytes = try Data(contentsOf: url)
        XCTAssertNotEqual(sha256(postBytes), originalSha,
                          "Target SHA256 SHALL change after successful write")
        XCTAssertGreaterThan(postBytes.count, 0,
                             "Target SHALL have non-zero size after successful write")
        XCTAssertEqual(postBytes[0], 0x50, "Target SHALL be valid ZIP (PK)")
        XCTAssertEqual(postBytes[1], 0x4B, "Target SHALL be valid ZIP (PK)")
    }

    // MARK: - Scenario 2: Throw mid-write preserves original

    func testThrowMidWritePreservesOriginalAndNoOrphanTempRemains() throws {
        let url = tempDir.appendingPathComponent("test.docx")

        // Seed original.
        var originalDoc = WordDocument()
        originalDoc.body.children.append(.paragraph(Paragraph(text: "ORIGINAL")))
        let originalBytes = try DocxWriter.writeData(originalDoc)
        try originalBytes.write(to: url)
        let originalSha = sha256(originalBytes)

        // Make parent dir read-only so temp file write throws.
        try FileManager.default.setAttributes(
            [.posixPermissions: 0o555], ofItemAtPath: tempDir.path
        )

        // Attempt to write — must throw.
        var modifiedDoc = WordDocument()
        modifiedDoc.body.children.append(.paragraph(Paragraph(text: "MODIFIED")))
        XCTAssertThrowsError(try DocxWriter.write(modifiedDoc, to: url),
                             "Read-only parent SHALL cause write to throw")

        // Restore writability for inspection.
        try FileManager.default.setAttributes(
            [.posixPermissions: 0o755], ofItemAtPath: tempDir.path
        )

        // Post-condition 1: original file's bytes unchanged.
        let postBytes = try Data(contentsOf: url)
        XCTAssertEqual(sha256(postBytes), originalSha,
                       "Original target SHALL be byte-preserved when write throws")

        // Post-condition 2: no orphan `.tmp.*` files in parent dir.
        let entries = try FileManager.default.contentsOfDirectory(atPath: tempDir.path)
        let orphans = entries.filter { $0.contains(".tmp.") }
        XCTAssertTrue(orphans.isEmpty,
                      "No orphan tmp files SHALL remain in parent dir; found \(orphans)")
    }

    // MARK: - Scenario 3: Process killed mid-write preserves original

    /// Simulates the post-SIGKILL state by manually planting an orphan
    /// `<url>.tmp.<UUID>` file (as if a prior crashed write left one) and
    /// verifying both that the original at `url` is intact and that a
    /// subsequent successful write completes cleanly.
    func testProcessKilledMidWritePreservesOriginal() throws {
        let url = tempDir.appendingPathComponent("test.docx")

        // Seed original.
        var originalDoc = WordDocument()
        originalDoc.body.children.append(.paragraph(Paragraph(text: "ORIGINAL")))
        let originalBytes = try DocxWriter.writeData(originalDoc)
        try originalBytes.write(to: url)
        let originalSha = sha256(originalBytes)

        // Plant an orphan temp file matching the spec's naming pattern.
        let orphanURL = url.appendingPathExtension("tmp.\(UUID().uuidString)")
        try Data("partial-write-from-killed-process".utf8).write(to: orphanURL)
        XCTAssertTrue(FileManager.default.fileExists(atPath: orphanURL.path),
                      "Orphan tmp SHALL be plantable for simulation")

        // Invariant: original file untouched by the simulated crash.
        let preBytes = try Data(contentsOf: url)
        XCTAssertEqual(sha256(preBytes), originalSha,
                       "Original SHALL be byte-preserved post-simulated-SIGKILL")

        // Subsequent successful write succeeds + cleans up its OWN tmp.
        var modifiedDoc = WordDocument()
        modifiedDoc.body.children.append(.paragraph(Paragraph(text: "MODIFIED")))
        try DocxWriter.write(modifiedDoc, to: url)

        let postBytes = try Data(contentsOf: url)
        XCTAssertNotEqual(sha256(postBytes), originalSha,
                          "Target SHA256 SHALL change after recovery write")

        // Spec allows at most the prior orphan to remain (cleanup-on-throw
        // cannot retroactively remove someone else's temp).
        let entries = try FileManager.default.contentsOfDirectory(atPath: tempDir.path)
        let orphans = entries.filter { $0.contains(".tmp.") }
        XCTAssertEqual(orphans.count, 1,
                       "Only the planted orphan SHALL remain — DocxWriter.write SHALL clean up its own temp; got \(orphans)")
        XCTAssertEqual(orphans.first, orphanURL.lastPathComponent,
                       "Remaining orphan SHALL be the planted one, not a new one from this write")
    }

    // MARK: - Scenario 4: Fresh write to non-existent path

    func testFreshWriteToNonExistentPathProducesOnlyTarget() throws {
        let subdir = tempDir.appendingPathComponent("fresh")
        try FileManager.default.createDirectory(at: subdir, withIntermediateDirectories: true)
        let url = subdir.appendingPathComponent("new.docx")

        XCTAssertFalse(FileManager.default.fileExists(atPath: url.path),
                       "Pre-condition: target SHALL NOT exist")

        var doc = WordDocument()
        doc.body.children.append(.paragraph(Paragraph(text: "FRESH")))
        try DocxWriter.write(doc, to: url)

        XCTAssertTrue(FileManager.default.fileExists(atPath: url.path),
                      "Target SHALL exist after fresh write")
        let bytes = try Data(contentsOf: url)
        XCTAssertGreaterThan(bytes.count, 0, "Target SHALL have non-zero size")

        // Parent dir SHALL contain exactly one file (the target).
        let entries = try FileManager.default.contentsOfDirectory(atPath: subdir.path)
        XCTAssertEqual(entries, [url.lastPathComponent],
                       "Parent dir SHALL contain only the new target; got \(entries)")
    }

    // MARK: - RED-revealing atomicity test (concurrent observer)

    /// Pre-v0.13.2 `DocxWriter.write` deletes the target file (line 19-20)
    /// BEFORE computing new bytes. A concurrent observer polling
    /// `fileExists(atPath:)` will detect the gap when the file is absent.
    ///
    /// The atomic-rename refactor writes new bytes to `<url>.tmp.<UUID>` first,
    /// then `replaceItemAt`. The target at `url` is observable as either the
    /// full original or the full new bytes — never absent or zero-byte.
    ///
    /// This test is the primary RED indicator for the bug behind che-word-mcp#36.
    func testTargetIsAlwaysObservableDuringSuccessfulWrite() throws {
        let url = tempDir.appendingPathComponent("observed.docx")

        // Seed original (use a sizable doc so writeData takes measurable time).
        var originalDoc = WordDocument()
        for i in 0..<200 {
            originalDoc.body.children.append(.paragraph(Paragraph(text: "Original paragraph \(i)")))
        }
        try DocxWriter.write(originalDoc, to: url)

        // Modified doc: also sizable so writeData takes a few ms.
        var modifiedDoc = WordDocument()
        for i in 0..<200 {
            modifiedDoc.body.children.append(.paragraph(Paragraph(text: "Modified paragraph \(i)")))
        }

        // Concurrent observer: poll fileExists at high frequency during write.
        let stopFlag = AtomicFlag()
        var sawAbsent = false
        let lock = NSLock()
        let observer = Thread {
            while !stopFlag.isSet {
                if !FileManager.default.fileExists(atPath: url.path) {
                    lock.lock()
                    sawAbsent = true
                    lock.unlock()
                }
                // No sleep — poll as fast as possible to maximize gap-detection probability.
            }
        }
        observer.start()

        // Run the write under observation.
        try DocxWriter.write(modifiedDoc, to: url)

        stopFlag.set()
        // Give observer a tick to exit cleanly.
        Thread.sleep(forTimeInterval: 0.01)

        lock.lock()
        let absentObserved = sawAbsent
        lock.unlock()

        XCTAssertFalse(absentObserved,
                       "External observer SHALL NEVER see the target absent during DocxWriter.write (RED on pre-v0.13.2)")
    }

    // MARK: - Temp file naming pattern + cleanup invariant

    func testNoTempOrphanRemainsAfterSuccessfulOverwrite() throws {
        let url = tempDir.appendingPathComponent("test.docx")

        var doc1 = WordDocument()
        doc1.body.children.append(.paragraph(Paragraph(text: "V1")))
        try DocxWriter.write(doc1, to: url)

        var doc2 = WordDocument()
        doc2.body.children.append(.paragraph(Paragraph(text: "V2 — distinct content")))
        try DocxWriter.write(doc2, to: url)

        let entries = try FileManager.default.contentsOfDirectory(atPath: tempDir.path)
        let orphans = entries.filter { $0.contains(".tmp.") }
        XCTAssertTrue(orphans.isEmpty,
                      "No orphan tmp files SHALL remain after successful overwrite; got \(orphans)")
        XCTAssertEqual(entries, [url.lastPathComponent],
                       "Parent SHALL contain exactly the target; got \(entries)")
    }

    // MARK: - Helpers

    private func sha256(_ data: Data) -> String {
        var hash = [UInt8](repeating: 0, count: 32)
        data.withUnsafeBytes { bytes in
            _ = SHA256_simple(bytes.baseAddress!, bytes.count, &hash)
        }
        return hash.map { String(format: "%02x", $0) }.joined()
    }
}

// Minimal SHA-256 wrapper using CommonCrypto. Local to this test file to avoid
// a public surface change in OOXMLSwift.
import CommonCrypto

private func SHA256_simple(_ data: UnsafeRawPointer, _ len: Int, _ out: UnsafeMutablePointer<UInt8>) -> Bool {
    CC_SHA256(data, CC_LONG(len), out)
    return true
}

/// Simple atomic boolean for cross-thread stop signaling in observer tests.
private final class AtomicFlag {
    private let lock = NSLock()
    private var flag = false
    var isSet: Bool {
        lock.lock(); defer { lock.unlock() }
        return flag
    }
    func set() {
        lock.lock(); defer { lock.unlock() }
        flag = true
    }
}
