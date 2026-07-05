import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// word-aligned-state-sync Phase 3 tasks 4.1 + 4.7 (+ 4.9 at orchestrator
/// level) — `ooxml-word-sync` Requirements "SyncOrchestrator coordinates
/// Word and Swift writers", "Sidecar persistence of snapshot and log",
/// "Bootstrap from existing docx".
///
/// "Word saves" are simulated by rewriting `word/document.xml` inside the
/// docx zip out-of-band (no orchestrator involvement) — the same observable
/// the real Word produces: new bytes at the same path. The live-Word
/// AppleScript variant is task 4.8's gated integration test.
final class SyncOrchestratorTests: XCTestCase {

    // MARK: - Fixture

    private static let initialDocumentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="0AB7C123"><w:r><w:t>original first</w:t></w:r></w:p><w:p w14:paraId="0DEF4567"><w:r><w:t>original second</w:t></w:r></w:p></w:body></w:document>
        """

    /// Builds the docx in its own temp directory (sidecars land next to it).
    private func buildFixture() throws -> URL {
        let dir = FileManager.default.temporaryDirectory
            .appendingPathComponent("sync-orch-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        let staging = dir.appendingPathComponent("staging")

        func write(_ content: String, to relativePath: String) throws {
            let url = staging.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(), withIntermediateDirectories: true)
            try content.write(to: url, atomically: true, encoding: .utf8)
        }
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="xml" ContentType="application/xml"/>
                <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>
            """, to: "[Content_Types].xml")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """, to: "_rels/.rels")
        try write(Self.initialDocumentXML, to: "word/document.xml")

        let docxURL = dir.appendingPathComponent("report.docx")
        let archive = try Archive(url: docxURL, accessMode: .create)
        let base = staging.resolvingSymlinksInPath().path
        let enumerator = FileManager.default.enumerator(
            at: staging, includingPropertiesForKeys: [.isDirectoryKey])!
        for case let fileURL as URL in enumerator {
            let isDir = (try? fileURL.resourceValues(forKeys: [.isDirectoryKey]))?.isDirectory ?? false
            if isDir { continue }
            let entry = String(fileURL.resolvingSymlinksInPath().path.dropFirst(base.count + 1))
            try archive.addEntry(with: entry, fileURL: fileURL, compressionMethod: .deflate)
        }
        try FileManager.default.removeItem(at: staging)
        return docxURL
    }

    private func cleanup(_ docxURL: URL) {
        try? FileManager.default.removeItem(at: docxURL.deletingLastPathComponent())
    }

    /// Simulates a Word save: rewrites `word/document.xml` inside the zip
    /// with `transform` applied, touching nothing else and creating no
    /// sidecars — exactly the observable a real Word save produces.
    private func simulateWordSave(at docxURL: URL, transform: (String) -> String) throws {
        let readArchive = try Archive(url: docxURL, accessMode: .read)
        var parts: [(String, Data)] = []
        for entry in readArchive {
            var data = Data()
            _ = try readArchive.extract(entry) { data.append($0) }
            parts.append((entry.path, data))
        }
        let tmpURL = docxURL.deletingLastPathComponent()
            .appendingPathComponent("word-save-\(UUID().uuidString).docx")
        let writeArchive = try Archive(url: tmpURL, accessMode: .create)
        for (path, data) in parts {
            var out = data
            if path == "word/document.xml" {
                out = Data(transform(String(decoding: data, as: UTF8.self)).utf8)
            }
            try writeArchive.addEntry(
                with: path, type: .file, uncompressedSize: Int64(out.count),
                provider: { position, size in
                    out.subdata(in: Int(position)..<(Int(position) + size))
                })
        }
        _ = try FileManager.default.replaceItemAt(docxURL, withItemAt: tmpURL)
    }

    // MARK: - 4.7 Bootstrap

    func testBootstrapFreshCreatesSidecars() throws {
        let docxURL = try buildFixture()
        defer { cleanup(docxURL) }

        let orch = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)

        XCTAssertTrue(FileManager.default.fileExists(
            atPath: SidecarStore.oplogURL(for: docxURL).path),
            "fresh bootstrap must create the oplog sidecar")
        XCTAssertTrue(FileManager.default.fileExists(
            atPath: SidecarStore.snapshotURL(for: docxURL).path),
            "fresh bootstrap must create the snapshot sidecar")
        XCTAssertTrue(orch.document.operationLog.entries.isEmpty,
                      "fresh bootstrap starts with an empty log")

        let snapshot = try SidecarStore.loadSnapshot(alongside: docxURL)
        XCTAssertNotNil(snapshot?.documentXML,
                        "snapshot must store the baseline document.xml for cross-session diffs")
    }

    func testBootstrapReusesExistingSidecars() throws {
        let docxURL = try buildFixture()
        defer { cleanup(docxURL) }

        // Session 1: bootstrap + a Swift mutation + flush.
        let first = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)
        try first.setParagraphText(id: ElementID(rawString: "w14:paraId=0AB7C123"), "swift v1")
        try first.flush()

        // Session 2: log history must be restored.
        let second = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)
        XCTAssertEqual(second.document.operationLog.entries.count, 1,
                       "existing oplog sidecar must be reused across sessions")
        XCTAssertEqual(second.document.operationLog.entries.first?.source, .swift)
    }

    func testBootstrapWithStaleSnapshotImportsInterveningChanges() throws {
        // Spec scenario "Existing sidecars are reused": docx changed after
        // the snapshot → bootstrap runs an import diff.
        let docxURL = try buildFixture()
        defer { cleanup(docxURL) }

        _ = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)   // creates sidecars
        try simulateWordSave(at: docxURL) {
            $0.replacingOccurrences(of: "original second", with: "word edited between sessions")
        }

        let orch = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)
        XCTAssertEqual(orch.document.operationLog.entries.count, 1,
                       "stale snapshot must trigger an intervening-change import")
        XCTAssertEqual(orch.document.operationLog.entries.first?.source, .word)
    }

    // MARK: - 4.1 Word save detected and imported

    func testWordSaveDetectedAndImportedAsWordSourcedOps() throws {
        let docxURL = try buildFixture()
        defer { cleanup(docxURL) }
        let orch = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)

        try simulateWordSave(at: docxURL) {
            $0.replacingOccurrences(of: "original first", with: "edited in Word")
        }

        XCTAssertTrue(try orch.checkForExternalChange(),
                      "watcher must detect the Word save (content hash changed)")
        let imported = try orch.importFromDisk()

        XCTAssertEqual(imported.count, 1)
        guard case .setText(let target, let text) = imported[0] else {
            return XCTFail("expected SetText from the Word edit, got \(imported)")
        }
        XCTAssertEqual(target.raw, "w14:paraId=0AB7C123")
        XCTAssertEqual(text, "edited in Word")
        XCTAssertEqual(orch.document.operationLog.entries.last?.source, .word,
                       "imported ops must carry source word")

        // In-memory typed view reflects Word's edit after import.
        if case .paragraph(let p) = orch.document.body.children.first {
            XCTAssertEqual(p.text, "edited in Word")
        } else {
            XCTFail("expected paragraph view after import resync")
        }
    }

    func testRsidOnlyWordSaveImportsNothing() throws {
        // 4.9 at orchestrator level: rsid renumbering only → empty import.
        let docxURL = try buildFixture()
        defer { cleanup(docxURL) }
        let orch = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)

        try simulateWordSave(at: docxURL) {
            $0.replacingOccurrences(
                of: #"<w:p w14:paraId="0AB7C123">"#,
                with: #"<w:p w14:paraId="0AB7C123" w:rsidR="00FF00AA" w:rsidRDefault="00FF00AA">"#)
        }

        XCTAssertTrue(try orch.checkForExternalChange(),
                      "bytes changed, watcher fires")
        let imported = try orch.importFromDisk()
        XCTAssertTrue(imported.isEmpty,
                      "rsid-only Word save must import an empty op set")
    }

    // MARK: - Conflict path

    func testConflictingEditsAbortByDefault() throws {
        let docxURL = try buildFixture()
        defer { cleanup(docxURL) }
        let orch = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)

        // Pending Swift edit (not flushed) on the same paragraph Word edits.
        try orch.setParagraphText(id: ElementID(rawString: "w14:paraId=0AB7C123"), "swift version")
        try simulateWordSave(at: docxURL) {
            $0.replacingOccurrences(of: "original first", with: "word version")
        }

        XCTAssertThrowsError(try orch.importFromDisk()) { error in
            guard case SyncError.conflict(let report) = error else {
                return XCTFail("expected SyncError.conflict, got \(error)")
            }
            XCTAssertEqual(report.entries.count, 1)
            XCTAssertEqual(report.entries[0].elementID.raw, "w14:paraId=0AB7C123")
        }
    }

    // MARK: - 4.6 flush refuses while Word holds the lock

    func testFlushThrowsWhileWordLockPresent() throws {
        let docxURL = try buildFixture()
        defer { cleanup(docxURL) }
        let orch = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)

        let lockURL = WordLock.lockFileURL(for: docxURL)
        try Data("locked".utf8).write(to: lockURL)
        defer { try? FileManager.default.removeItem(at: lockURL) }

        XCTAssertThrowsError(try orch.flush()) { error in
            guard case SyncError.fileLockedByWord = error else {
                return XCTFail("expected fileLockedByWord, got \(error)")
            }
        }
    }

    // MARK: - Flush round-trip + own-write suppression

    func testFlushPersistsSwiftEditAndDoesNotSelfTrigger() throws {
        let docxURL = try buildFixture()
        defer { cleanup(docxURL) }
        let orch = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)

        try orch.setParagraphText(id: ElementID(rawString: "w14:paraId=0AB7C123"), "flushed text")
        try orch.flush()

        XCTAssertFalse(try orch.checkForExternalChange(),
                       "the orchestrator's own flush must not read back as an external change")

        let reread = try DocxReader.read(from: docxURL)
        if case .paragraph(let p) = reread.body.children.first {
            XCTAssertEqual(p.text, "flushed text", "flush must persist the Swift edit to disk")
        } else {
            XCTFail("expected paragraph after re-read")
        }
    }
}
