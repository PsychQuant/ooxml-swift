import XCTest
@testable import OOXMLSwift

final class DocumentProfileStoreTests: XCTestCase {
    private func config(_ text: String? = nil) throws -> URL {
        let dir = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString)
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        addTeardownBlock { try? FileManager.default.removeItem(at: dir) }
        let url = dir.appendingPathComponent("config.json")
        if let text { try Data(text.utf8).write(to: url) }
        return url
    }

    func testSelectionAndNoUnnecessaryConfigurationReads() throws {
        let url = try config("broken")
        let store = DocumentProfileStore(configURL: url)
        XCTAssertNil(try store.resolve(explicit: nil, context: .existingDocument))
        XCTAssertEqual(try store.resolve(explicit: .inherit, context: .newDocument), .inherit)
        XCTAssertEqual(try store.resolve(explicit: .inherit, context: .existingDocument), .inherit)
        XCTAssertThrowsError(try store.resolve(explicit: nil, context: .newDocument))
        try FileManager.default.removeItem(at: url)
        XCTAssertEqual(try store.resolve(explicit: nil, context: .newDocument), .inherit)
        try Data(#"{"document":{"defaultProfile":"official"}}"#.utf8).write(to: url)
        XCTAssertThrowsError(try store.resolve(explicit: nil, context: .newDocument))
        XCTAssertNil(try store.resolve(explicit: nil, context: .existingDocument))
        XCTAssertEqual(try store.resolve(explicit: .inherit, context: .newDocument), .inherit)
    }

    func testInvalidSettingsDoNotOverwriteConfiguration() throws {
        for json in ["[]", "null", "invalid", #"{"document":null}"#, #"{"document":[]}"#,
                     #"{"document":{"defaultProfile":1}}"#, #"{"document":{"defaultProfile":"unknown"}}"#,
                     #"{"document":{"officialSnapshot":null}}"#, #"{"document":{"officialSnapshot":"../outside.json"}}"#,
                     #"{"document":{"officialSnapshot":"/tmp/outside.json"}}"#] {
            let url = try config(json)
            let store = DocumentProfileStore(configURL: url)
            XCTAssertThrowsError(try store.settings(), json)
            XCTAssertThrowsError(try store.setDefaultProfile(.inherit), json)
            XCTAssertEqual(try String(contentsOf: url), json)
        }
    }

    func testSetDefaultPreservesUnrelatedAndUnknownDocumentFields() throws {
        let url = try config(#"{"agent":"codex","ocrHosts":{"a":"b"},"extension":[1,true],"document":{"custom":{"v":3}}}"#)
        let store = DocumentProfileStore(configURL: url)
        try store.setDefaultProfile(.official)
        XCTAssertEqual(try store.settings().defaultProfile, .official)
        let root = try XCTUnwrap(JSONSerialization.jsonObject(with: Data(contentsOf: url)) as? [String: Any])
        XCTAssertEqual(root["agent"] as? String, "codex")
        XCTAssertEqual(root["ocrHosts"] as? [String: String], ["a": "b"])
        XCTAssertEqual((root["document"] as? [String: Any])?["custom"] as? [String: Int], ["v": 3])
        XCTAssertEqual((root["extension"] as? NSArray), [1, true] as NSArray)
    }

    func testImportImmutableSnapshotSurvivesSourceChangesAndDoesNotSelectDefault() throws {
        let url = try config(#"{"document":{"custom":"kept"},"secret":"private"}"#)
        let store = DocumentProfileStore(configURL: url)
        let fixture = DocumentFormattingProfileTests()
        let template = try fixture.template()
        defer { try? FileManager.default.removeItem(at: template.deletingLastPathComponent()) }
        let sourceBytes = try Data(contentsOf: template)
        try store.importOfficial(from: template)
        let first = try XCTUnwrap(store.settings().officialSnapshot)
        let firstURL = url.deletingLastPathComponent().appendingPathComponent(first)
        let firstBytes = try Data(contentsOf: firstURL)
        XCTAssertEqual(try store.settings().defaultProfile, .inherit)
        XCTAssertEqual(try Data(contentsOf: template), sourceBytes)
        XCTAssertFalse(String(decoding: firstBytes, as: UTF8.self).contains(template.path))
        try store.importOfficial(from: template)
        XCTAssertNotEqual(try store.settings().officialSnapshot, first)
        XCTAssertEqual(try Data(contentsOf: firstURL), firstBytes)
        try Data("broken source".utf8).write(to: template)
        let before = try Data(contentsOf: url)
        XCTAssertThrowsError(try store.importOfficial(from: template))
        XCTAssertEqual(try Data(contentsOf: url), before)
        XCTAssertEqual(try Data(contentsOf: firstURL), firstBytes)
        try FileManager.default.removeItem(at: template)
        XCTAssertThrowsError(try store.importOfficial(from: template))
        XCTAssertEqual(try Data(contentsOf: url), before)
        XCTAssertEqual(try store.resolve(explicit: .official, context: .newDocument)?.kind, .official)
        try store.setDefaultProfile(.official)
        XCTAssertEqual(try store.resolve(explicit: nil, context: .newDocument)?.kind, .official)
    }

    func testOfficialRejectsMissingCorruptWrongKindAndUnknownVersion() throws {
        let url = try config(#"{"document":{"defaultProfile":"official","officialSnapshot":"snapshot.json"}}"#)
        let store = DocumentProfileStore(configURL: url)
        let snapshot = url.deletingLastPathComponent().appendingPathComponent("snapshot.json")
        XCTAssertThrowsError(try store.resolve(explicit: .official, context: .existingDocument))
        for bytes in [Data("broken".utf8), try JSONEncoder().encode(DocumentFormattingProfile.inherit),
                      Data(#"{"schemaVersion":999,"kind":"official"}"#.utf8)] {
            try bytes.write(to: snapshot)
            XCTAssertThrowsError(try store.resolve(explicit: nil, context: .newDocument))
        }
    }

    func testSnapshotWriteFailureKeepsPreviousReferenceAndBytes() throws {
        let url = try config(#"{"document":{"officialSnapshot":"old.json"}}"#)
        let dir = url.deletingLastPathComponent()
        let old = dir.appendingPathComponent("old.json")
        try Data("old snapshot".utf8).write(to: old)
        try Data("blocked directory".utf8).write(to: dir.appendingPathComponent("profiles"))
        let before = try Data(contentsOf: url)
        let fixture = DocumentFormattingProfileTests()
        let template = try fixture.template()
        defer { try? FileManager.default.removeItem(at: template.deletingLastPathComponent()) }
        XCTAssertThrowsError(try DocumentProfileStore(configURL: url).importOfficial(from: template))
        XCTAssertEqual(try Data(contentsOf: url), before)
        XCTAssertEqual(try String(contentsOf: old), "old snapshot")
    }

    func testConfigPublicationFailureKeepsOldSnapshotAndCleansUnreferencedNewFile() throws {
        let url = try config(#"{"document":{"officialSnapshot":"old.json"}}"#)
        let dir = url.deletingLastPathComponent(), profiles = dir.appendingPathComponent("profiles")
        let old = dir.appendingPathComponent("old.json")
        try Data("old snapshot".utf8).write(to: old)
        try FileManager.default.createDirectory(at: profiles, withIntermediateDirectories: true)
        let before = try Data(contentsOf: url)
        let fixture = DocumentFormattingProfileTests()
        let template = try fixture.template()
        defer { try? FileManager.default.removeItem(at: template.deletingLastPathComponent()) }
        try FileManager.default.setAttributes([.posixPermissions: 0o500], ofItemAtPath: dir.path)
        defer { try? FileManager.default.setAttributes([.posixPermissions: 0o700], ofItemAtPath: dir.path) }
        XCTAssertThrowsError(try DocumentProfileStore(configURL: url).importOfficial(from: template))
        XCTAssertEqual(try Data(contentsOf: url), before)
        XCTAssertEqual(try String(contentsOf: old), "old snapshot")
        XCTAssertEqual(try FileManager.default.contentsOfDirectory(atPath: profiles.path), [])
    }

    // MARK: - PsychQuant/macdoc#194

    private func permissions(_ url: URL) throws -> Int {
        try XCTUnwrap(FileManager.default.attributesOfItem(atPath: url.path)[.posixPermissions] as? Int)
    }

    private func officialTemplate() throws -> URL {
        let template = try DocumentFormattingProfileTests().template()
        addTeardownBlock { try? FileManager.default.removeItem(at: template.deletingLastPathComponent()) }
        return template
    }

    func testResolveReportsMissingSnapshotFileAsDomainError() throws {
        let url = try config(#"{"document":{"defaultProfile":"official","officialSnapshot":"profiles/official-gone.json"}}"#)
        let store = DocumentProfileStore(configURL: url)
        for explicit in [DocumentFormattingProfile.Kind.official, nil] {
            XCTAssertThrowsError(try store.resolve(explicit: explicit, context: .newDocument)) { error in
                XCTAssertEqual(error as? DocumentProfileStoreError, .snapshotFileMissing("profiles/official-gone.json"))
            }
        }
        XCTAssertTrue(DocumentProfileStoreError.snapshotFileMissing("profiles/x.json").localizedDescription.contains("profiles/x.json"))
    }

    /// Under a permissive umask the store still creates its profiles
    /// directory 0700 and every snapshot and config file 0600; an existing
    /// world-readable config is tightened on its next write.
    func testStoreCreatesOwnerOnlyDirectoriesAndFilesRegardlessOfUmask() throws {
        let previous = umask(0o022)
        defer { umask(previous) }
        let root = try config().deletingLastPathComponent()
        let url = root.appendingPathComponent("nested/config.json")
        let store = DocumentProfileStore(configURL: url)
        try store.importOfficial(from: officialTemplate())
        let snapshot = url.deletingLastPathComponent().appendingPathComponent(try XCTUnwrap(store.settings().officialSnapshot))
        XCTAssertEqual(try permissions(url.deletingLastPathComponent()), 0o700)
        XCTAssertEqual(try permissions(snapshot.deletingLastPathComponent()), 0o700)
        XCTAssertEqual(try permissions(snapshot), 0o600)
        XCTAssertEqual(try permissions(url), 0o600)

        let existing = try config(#"{"agent":"codex"}"#)
        try FileManager.default.setAttributes([.posixPermissions: 0o644], ofItemAtPath: existing.path)
        try DocumentProfileStore(configURL: existing).setDefaultProfile(.official)
        XCTAssertEqual(try permissions(existing), 0o600)
        let oldProfiles = existing.deletingLastPathComponent().appendingPathComponent("profiles")
        try FileManager.default.createDirectory(at: oldProfiles, withIntermediateDirectories: true, attributes: [.posixPermissions: 0o755])
        try DocumentProfileStore(configURL: existing).importOfficial(from: officialTemplate())
        XCTAssertEqual(try permissions(oldProfiles), 0o700)
        let root2 = try XCTUnwrap(JSONSerialization.jsonObject(with: Data(contentsOf: existing)) as? [String: Any])
        XCTAssertEqual(root2["agent"] as? String, "codex")
    }

    /// Snapshots keep `.withoutOverwriting` semantics: an existing name —
    /// a file or a planted symlink — is never written through.
    func testSnapshotWriteNeverOverwritesAnExistingName() throws {
        let url = try config()
        let store = DocumentProfileStore(configURL: url)
        let relative = "profiles/official-fixed.json"
        let target = url.deletingLastPathComponent().appendingPathComponent(relative)
        try store.writeNewSnapshot(Data("first".utf8), relativePath: relative)
        XCTAssertThrowsError(try store.writeNewSnapshot(Data("second".utf8), relativePath: relative))
        XCTAssertEqual(try String(contentsOf: target), "first")
        let victim = url.deletingLastPathComponent().appendingPathComponent("victim.txt")
        try Data("victim".utf8).write(to: victim)
        let planted = url.deletingLastPathComponent().appendingPathComponent("profiles/official-link.json")
        try FileManager.default.createSymbolicLink(at: planted, withDestinationURL: victim)
        XCTAssertThrowsError(try store.writeNewSnapshot(Data("through link".utf8), relativePath: "profiles/official-link.json"))
        XCTAssertEqual(try String(contentsOf: victim), "victim")
    }

    func testGarbageCollectionListsThenDeletesOnlyUnreferencedOfficialSnapshots() throws {
        let url = try config(#"{"agent":"codex"}"#)
        let dir = url.deletingLastPathComponent(), profiles = dir.appendingPathComponent("profiles")
        let store = DocumentProfileStore(configURL: url)
        let template = try officialTemplate()
        var imported: [String] = []
        for _ in 0..<3 {
            try store.importOfficial(from: template)
            imported.append(try XCTUnwrap(store.settings().officialSnapshot))
        }
        let referenced = imported[2]
        let bystanders = ["profiles/notes.txt", "profiles/other.json", "profiles/official-keep.txt", "old-official-x.json"]
        for path in bystanders { try Data("keep".utf8).write(to: dir.appendingPathComponent(path)) }
        try FileManager.default.createDirectory(at: profiles.appendingPathComponent("official-directory.json"), withIntermediateDirectories: true)
        let link = profiles.appendingPathComponent("official-link.json")
        try FileManager.default.createSymbolicLink(at: link, withDestinationURL: dir.appendingPathComponent(imported[0]))
        let config = try Data(contentsOf: url)

        let listed = try store.garbageCollectOfficialSnapshots(dryRun: true)
        XCTAssertEqual(listed, imported.prefix(2).sorted())
        for path in imported { XCTAssertTrue(FileManager.default.fileExists(atPath: dir.appendingPathComponent(path).path), path) }

        XCTAssertEqual(try store.garbageCollectOfficialSnapshots(dryRun: false), listed)
        for path in listed { XCTAssertFalse(FileManager.default.fileExists(atPath: dir.appendingPathComponent(path).path), path) }
        XCTAssertTrue(FileManager.default.fileExists(atPath: dir.appendingPathComponent(referenced).path))
        for path in bystanders { XCTAssertEqual(try String(contentsOf: dir.appendingPathComponent(path)), "keep", path) }
        XCTAssertTrue(FileManager.default.fileExists(atPath: profiles.appendingPathComponent("official-directory.json").path))
        XCTAssertEqual(try FileManager.default.destinationOfSymbolicLink(atPath: link.path), dir.appendingPathComponent(imported[0]).path)
        XCTAssertEqual(try Data(contentsOf: url), config)
        XCTAssertEqual(try store.resolve(explicit: .official, context: .newDocument)?.kind, .official)
        XCTAssertEqual(try store.garbageCollectOfficialSnapshots(dryRun: false), [])
    }

    func testGarbageCollectionRefusesInvalidConfigurationAndToleratesMissingProfiles() throws {
        let empty = try config()
        XCTAssertEqual(try DocumentProfileStore(configURL: empty).garbageCollectOfficialSnapshots(dryRun: false), [])
        let url = try config(#"{"document":{"officialSnapshot":"../outside.json"}}"#)
        let profiles = url.deletingLastPathComponent().appendingPathComponent("profiles")
        try FileManager.default.createDirectory(at: profiles, withIntermediateDirectories: true)
        let orphan = profiles.appendingPathComponent("official-orphan.json")
        try Data("orphan".utf8).write(to: orphan)
        for dryRun in [true, false] {
            XCTAssertThrowsError(try DocumentProfileStore(configURL: url).garbageCollectOfficialSnapshots(dryRun: dryRun))
        }
        XCTAssertTrue(FileManager.default.fileExists(atPath: orphan.path))
    }
}
