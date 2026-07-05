import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// word-aligned-state-sync Phase 3 task 4.8 — live Microsoft Word
/// round-trip: open a fixture in the real Word (scripted via
/// `osascript`), edit one paragraph, save; assert the orchestrator
/// captures the edit as a non-empty op set with `source: "word"`.
///
/// Gated: requires `RUN_WORD_INTEGRATION=1` in the environment AND
/// Microsoft Word installed — skipped otherwise so CI / clean machines
/// stay green. Run locally with:
///
///     RUN_WORD_INTEGRATION=1 swift test --filter WordLiveRoundTripTests
final class WordLiveRoundTripTests: XCTestCase {

    private static let initialDocumentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="0AB7C123"><w:r><w:t>original from swift</w:t></w:r></w:p><w:p w14:paraId="0DEF4567"><w:r><w:t>second paragraph stays</w:t></w:r></w:p></w:body></w:document>
        """

    private func buildFixture(in dir: URL) throws -> URL {
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

        let docxURL = dir.appendingPathComponent("live-word-roundtrip.docx")
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

    @discardableResult
    private func runAppleScript(_ script: String) throws -> (status: Int32, output: String) {
        let process = Process()
        process.executableURL = URL(fileURLWithPath: "/usr/bin/osascript")
        process.arguments = ["-e", script]
        let pipe = Pipe()
        process.standardOutput = pipe
        process.standardError = pipe
        try process.run()
        process.waitUntilExit()
        let data = pipe.fileHandleForReading.readDataToEndOfFile()
        return (process.terminationStatus, String(decoding: data, as: UTF8.self))
    }

    func testLiveWordEditIsCapturedAsWordSourcedOps() throws {
        guard ProcessInfo.processInfo.environment["RUN_WORD_INTEGRATION"] == "1" else {
            throw XCTSkip("live Word integration gated behind RUN_WORD_INTEGRATION=1")
        }
        guard FileManager.default.fileExists(atPath: "/Applications/Microsoft Word.app") else {
            throw XCTSkip("Microsoft Word not installed")
        }

        // Word's sandbox is friendlier to user-domain paths than /tmp for
        // scripted open+save; use a scratch dir under the user's home.
        let dir = FileManager.default.homeDirectoryForCurrentUser
            .appendingPathComponent(".cache/ooxml-swift-word-integration-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: dir) }

        let docxURL = try buildFixture(in: dir)
        let orch = try SyncOrchestrator.bootstrapFromDocx(url: docxURL)

        // Drive the real Word: open, replace paragraph 1's text, save, close.
        let script = """
        tell application "Microsoft Word"
            open POSIX file "\(docxURL.path)"
            set theDoc to active document
            set content of text object of paragraph 1 of theDoc to "edited by live Word"
            save theDoc
            close theDoc saving no
        end tell
        """
        let result = try runAppleScript(script)
        guard result.status == 0 else {
            throw XCTSkip("osascript could not drive Word (automation permission?): \(result.output)")
        }

        // Word save is synchronous from AppleScript's perspective, but give
        // the filesystem a beat before polling.
        var changed = false
        for _ in 0..<20 {
            if try orch.checkForExternalChange() { changed = true; break }
            Thread.sleep(forTimeInterval: 0.5)
        }
        XCTAssertTrue(changed, "the watcher must detect the live Word save")

        let imported = try orch.importFromDisk()

        XCTAssertFalse(imported.isEmpty,
                       "the live Word edit must import as a non-empty op set")
        XCTAssertEqual(orch.document.operationLog.entries.last?.source, .word,
                       "imported ops must carry source word")
        let setTexts: [String] = imported.compactMap {
            if case .setText(_, let text) = $0 { return text }
            return nil
        }
        XCTAssertTrue(setTexts.contains { $0.contains("edited by live Word") },
                      "the paragraph edit must surface as a SetText carrying Word's text; imported: \(imported)")
    }
}
