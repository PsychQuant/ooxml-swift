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
        func write(_ content: Data, to relativePath: String) throws {
            let url = staging.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(), withIntermediateDirectories: true)
            try content.write(to: url)
        }
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="xml" ContentType="application/xml"/>
                <Default Extension="bin" ContentType="application/octet-stream"/>
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
        try write(Data([0x00, 0x01, 0xFE, 0xFF]), to: "word/media/preserved.bin")

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

    private func extractedPackageDirectories() -> Set<String> {
        let root = FileManager.default.temporaryDirectory
            .appendingPathComponent("che-word-mcp")
        let names = (try? FileManager.default.contentsOfDirectory(
            atPath: root.path)) ?? []
        return Set(names)
    }

    /// Simulates a Word save: rewrites `word/document.xml` inside the zip
    /// with `transform` applied, touching nothing else and creating no
    /// sidecars — exactly the observable a real Word save produces.
    private func simulateWordSave(at docxURL: URL, transform: (String) -> String) throws {
        try simulateWordPartSave(at: docxURL, partPath: "word/document.xml") { data in
            Data(transform(String(decoding: data, as: UTF8.self)).utf8)
        }
    }

    private func simulateWordPartSave(
        at docxURL: URL,
        partPath: String,
        transform: (Data) -> Data
    ) throws {
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
            if path == partPath { out = transform(data) }
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
        XCTAssertNotNil(snapshot?.partSHA256?["word/media/preserved.bin"],
                        "snapshot must store a byte-level baseline for binary parts")
    }

    func testCloseReleasesBootstrapArchiveDirectory() throws {
        let docxURL = try buildFixture()
        defer { cleanup(docxURL) }

        let orchestrator = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)
        let archiveURL = try XCTUnwrap(orchestrator.document.archiveTempDir)
        XCTAssertTrue(FileManager.default.fileExists(atPath: archiveURL.path))

        orchestrator.close()

        XCTAssertFalse(FileManager.default.fileExists(atPath: archiveURL.path))
        XCTAssertNil(orchestrator.document.archiveTempDir)
    }

    func testClosedSessionRejectsFlushWithoutChangingDocx() throws {
        let docxURL = try buildFixture()
        defer { cleanup(docxURL) }
        let orchestrator = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)
        let before = try Data(contentsOf: docxURL)

        orchestrator.close()

        XCTAssertThrowsError(try orchestrator.flush()) { error in
            guard case SyncError.sessionClosed = error else {
                return XCTFail("expected sessionClosed, got \(error)")
            }
        }
        XCTAssertEqual(try Data(contentsOf: docxURL), before)
    }

    func testMalformedSidecarBootstrapReleasesArchiveDirectory() throws {
        let docxURL = try buildFixture()
        defer { cleanup(docxURL) }
        try Data("{malformed-jsonl\n".utf8).write(
            to: SidecarStore.oplogURL(for: docxURL))
        let before = extractedPackageDirectories()

        XCTAssertThrowsError(try SyncOrchestrator.bootstrapFromDocx(url: docxURL))

        XCTAssertEqual(extractedPackageDirectories(), before,
                       "a failed bootstrap must release its extracted package")
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

    func testBootstrapWithStaleSnapshotImportsBinarySiblingChange() throws {
        let docxURL = try buildFixture()
        defer { cleanup(docxURL) }

        _ = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)
        let expected = Data([0xF0, 0x0D, 0xBA, 0xBE])
        try simulateWordPartSave(
            at: docxURL,
            partPath: "word/media/preserved.bin") { _ in expected }

        let orch = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)
        XCTAssertTrue(orch.document.operationLog.entries.contains {
            if case .carryBinaryPart(let path, _) = $0.op {
                return path == "word/media/preserved.bin"
            }
            return false
        }, "cross-session binary edits must be recorded as Word-sourced raw ops")

        try orch.flush()
        let parts = try RawPartChannel.readAllParts(from: docxURL)
        XCTAssertEqual(parts["word/media/preserved.bin"], expected)
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

    func testFormattingOnlyWordEditSurvivesImportThenFlush() throws {
        let docxURL = try buildFixture()
        defer { cleanup(docxURL) }
        let orch = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)

        try simulateWordSave(at: docxURL) {
            $0.replacingOccurrences(
                of: "<w:r><w:t>original first</w:t></w:r>",
                with: "<w:r><w:rPr><w:b/></w:rPr><w:t>original first</w:t></w:r>")
        }

        _ = try orch.importFromDisk()
        try orch.flush()

        let archive = try Archive(url: docxURL, accessMode: .read)
        var data = Data()
        _ = try archive.extract(archive["word/document.xml"]!) { data.append($0) }
        XCTAssertTrue(String(decoding: data, as: UTF8.self).contains("<w:b"),
                      "an imported formatting-only Word edit must not be overwritten by flush")
    }

    func testIDLessWordTextEditSurvivesImportThenFlush() throws {
        let docxURL = try buildFixture()
        defer { cleanup(docxURL) }
        let orch = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)

        try simulateWordSave(at: docxURL) {
            $0.replacingOccurrences(of: #" w14:paraId="0AB7C123""#, with: "")
                .replacingOccurrences(of: "original first", with: "ID-less Word edit")
        }

        _ = try orch.importFromDisk()
        try orch.flush()

        let archive = try Archive(url: docxURL, accessMode: .read)
        var data = Data()
        _ = try archive.extract(archive["word/document.xml"]!) { data.append($0) }
        XCTAssertTrue(String(decoding: data, as: UTF8.self).contains("ID-less Word edit"),
                      "an ID-less Word edit must be carried rather than silently dropped")
    }

    func testBinarySiblingWordEditSurvivesImportThenFlush() throws {
        let docxURL = try buildFixture()
        defer { cleanup(docxURL) }
        let orch = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)
        let expected = Data([0xFF, 0x00, 0xAA, 0x55, 0x10])

        try simulateWordPartSave(
            at: docxURL,
            partPath: "word/media/preserved.bin") { _ in expected }

        let imported = try orch.importFromDisk()
        XCTAssertTrue(imported.contains {
            if case .carryBinaryPart(let path, _) = $0 {
                return path == "word/media/preserved.bin"
            }
            return false
        })
        try orch.flush()

        let archive = try Archive(url: docxURL, accessMode: .read)
        var actual = Data()
        _ = try archive.extract(archive["word/media/preserved.bin"]!) { actual.append($0) }
        XCTAssertEqual(actual, expected)
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

    func testPendingSwiftEditSurvivesNonconflictingWordImportAndRestart() throws {
        let docxURL = try buildFixture()
        defer { cleanup(docxURL) }
        let first = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)

        try first.setParagraphText(
            id: ElementID(rawString: "w14:paraId=0AB7C123"),
            "pending Swift")
        try simulateWordSave(at: docxURL) {
            $0.replacingOccurrences(
                of: "original second", with: "imported Word")
        }
        _ = try first.importFromDisk()
        let snapshot = try XCTUnwrap(
            SidecarStore.loadSnapshot(alongside: docxURL))
        XCTAssertEqual(snapshot.pendingSwiftOpIDs?.count, 1)
        first.close()

        let second = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)
        defer { second.close() }
        let paragraphs = second.document.body.children.compactMap { child -> String? in
            guard case .paragraph(let paragraph) = child else { return nil }
            return paragraph.text
        }
        XCTAssertEqual(paragraphs, ["pending Swift", "imported Word"])

        try second.flush()
        var reread = try DocxReader.read(from: docxURL)
        defer { reread.close() }
        let persisted = reread.body.children.compactMap { child -> String? in
            guard case .paragraph(let paragraph) = child else { return nil }
            return paragraph.text
        }
        XCTAssertEqual(persisted, ["pending Swift", "imported Word"])
    }

    func testPendingSwiftEditSurvivesEmptyWordImportAndRestart() throws {
        let docxURL = try buildFixture()
        defer { cleanup(docxURL) }
        let first = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)
        try first.setParagraphText(
            id: ElementID(rawString: "w14:paraId=0AB7C123"),
            "pending Swift")
        try simulateWordSave(at: docxURL) {
            $0.replacingOccurrences(
                of: #"<w:p w14:paraId="0DEF4567">"#,
                with: #"<w:p w14:paraId="0DEF4567" w:rsidR="00FF00AA">"#)
        }
        XCTAssertTrue(try first.importFromDisk().isEmpty)
        first.close()

        let second = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)
        defer { second.close() }
        guard case .paragraph(let firstParagraph) = second.document.body.children.first else {
            return XCTFail("expected first paragraph")
        }
        XCTAssertEqual(firstParagraph.text, "pending Swift")
    }

    func testPendingSwiftEditMergesRsidOnlyWordSaveBeforeSameSessionFlush() throws {
        let docxURL = try buildFixture()
        defer { cleanup(docxURL) }
        let orchestrator = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)
        try orchestrator.setParagraphText(
            id: .init(rawString: "w14:paraId=0AB7C123"), "pending Swift")
        try simulateWordSave(at: docxURL) {
            $0.replacingOccurrences(
                of: #"<w:p w14:paraId="0DEF4567">"#,
                with: #"<w:p w14:paraId="0DEF4567" w:rsidR="AABBCCDD">"#)
        }

        XCTAssertTrue(try orchestrator.importFromDisk().isEmpty)
        try orchestrator.flush()

        let xml = String(decoding: try RawPartChannel.readAllParts(
            from: docxURL)["word/document.xml"]!, as: UTF8.self)
        XCTAssertTrue(xml.contains("pending Swift"))
        XCTAssertTrue(xml.contains(#"w:rsidR="AABBCCDD""#),
                      "identity-noise from Word must survive the pending merge")
    }

    func testRestartWithPendingStateDoesNotDuplicateNewWordInsertion() throws {
        let docxURL = try buildFixture()
        defer { cleanup(docxURL) }
        let first = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)
        try first.setParagraphText(
            id: .init(rawString: "w14:paraId=0AB7C123"), "P1-PENDING")
        try simulateWordSave(at: docxURL) {
            $0.replacingOccurrences(of: "original second", with: "P2-WORD1")
        }
        _ = try first.importFromDisk()
        first.close()

        try simulateWordSave(at: docxURL) {
            $0.replacingOccurrences(
                of: "</w:body>",
                with: #"<w:p w14:paraId="0EEE9999"><w:r><w:t>P3-WORD2</w:t></w:r></w:p></w:body>"#)
        }

        let second = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)
        defer { second.close() }
        let texts = second.document.body.children.compactMap { child -> String? in
            guard case .paragraph(let paragraph) = child else { return nil }
            return paragraph.text
        }
        XCTAssertEqual(texts, ["P1-PENDING", "P2-WORD1", "P3-WORD2"])
    }

    func testImportSidecarFailureRollsBackPairAndCanonicalMemory() throws {
        let docxURL = try buildFixture()
        defer { cleanup(docxURL) }
        let orchestrator = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)
        let oldLog = try Data(contentsOf: SidecarStore.oplogURL(for: docxURL))
        let oldSnapshot = try Data(contentsOf: SidecarStore.snapshotURL(for: docxURL))
        try simulateWordSave(at: docxURL) {
            $0.replacingOccurrences(of: "original second", with: "transactional Word")
        }

        XCTAssertThrowsError(try orchestrator.importFromDisk(
            policy: .abortOnConflict,
            immediatelyBeforeSnapshotWrite: {
                throw NSError(domain: "forced-sidecar-failure", code: 1)
            }))
        XCTAssertEqual(try Data(contentsOf: SidecarStore.oplogURL(for: docxURL)), oldLog)
        XCTAssertEqual(try Data(contentsOf: SidecarStore.snapshotURL(for: docxURL)), oldSnapshot)
        XCTAssertTrue(orchestrator.document.operationLog.entries.isEmpty,
                      "failed import must not commit candidate history in memory")

        let imported = try orchestrator.importFromDisk()
        XCTAssertEqual(imported.count, 1)
        XCTAssertEqual(orchestrator.document.operationLog.entries.count, 1)
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

    func testFlushRejectsUnimportedExternalGeneration() throws {
        let docxURL = try buildFixture()
        defer { cleanup(docxURL) }
        let orchestrator = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)
        try orchestrator.setParagraphText(
            id: ElementID(rawString: "w14:paraId=0AB7C123"),
            "pending Swift")
        try simulateWordSave(at: docxURL) {
            $0.replacingOccurrences(
                of: "original second", with: "unimported Word")
        }
        let externalBytes = try Data(contentsOf: docxURL)

        XCTAssertThrowsError(try orchestrator.flush()) { error in
            guard case SyncError.externalGenerationChanged = error else {
                return XCTFail("expected externalGenerationChanged, got \(error)")
            }
        }
        XCTAssertEqual(try Data(contentsOf: docxURL), externalBytes,
                       "rejected flush must not overwrite the external generation")
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
