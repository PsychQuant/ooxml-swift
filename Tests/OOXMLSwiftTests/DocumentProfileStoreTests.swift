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
        // The lock file already exists (as after any earlier write), so the
        // failure is the config publication itself, after the snapshot was
        // written — not the lock creation (PsychQuant/macdoc#204).
        try Data().write(to: URL(fileURLWithPath: url.path + ".lock"))
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

    // MARK: - PsychQuant/macdoc#204 cross-process config lock

    private func holdLock(_ configURL: URL) throws -> Int32 {
        let fd = open(configURL.path + ".lock", O_CREAT | O_RDWR | O_CLOEXEC, 0o600)
        XCTAssertGreaterThanOrEqual(fd, 0)
        XCTAssertEqual(flock(fd, LOCK_EX | LOCK_NB), 0, "the test must win the lock first")
        return fd
    }

    /// While another descriptor holds LOCK_EX on `<config>.lock`, every
    /// writer times out with a domain error and leaves the config, the
    /// profiles directory and the other writer's lock untouched.
    func testWritersTimeOutWhileAnotherDescriptorHoldsTheLock() throws {
        let url = try config(#"{"agent":"codex","document":{"custom":"kept"}}"#)
        let before = try Data(contentsOf: url)
        let fd = try holdLock(url)
        defer { flock(fd, LOCK_UN); close(fd) }
        let store = DocumentProfileStore(configURL: url, lockTimeout: 0.2, lockPollInterval: 0.02)
        let template = try officialTemplate()
        let started = Date()
        let attempts: [(String, () throws -> Void)] = [
            ("setDefaultProfile", { try store.setDefaultProfile(.official) }),
            ("importOfficial", { try store.importOfficial(from: template) }),
            ("garbageCollect", { _ = try store.garbageCollectOfficialSnapshots(dryRun: false) })
        ]
        for (name, attempt) in attempts {
            XCTAssertThrowsError(try attempt(), name) { error in
                XCTAssertEqual(error as? DocumentProfileStoreError, .configLockTimeout(url.path + ".lock"), name)
            }
        }
        XCTAssertLessThan(Date().timeIntervalSince(started), 3, "injected timeout must be honoured")
        XCTAssertEqual(try Data(contentsOf: url), before)
        XCTAssertFalse(FileManager.default.fileExists(atPath: url.deletingLastPathComponent().appendingPathComponent("profiles").path))
    }

    func testLockFileIsOwnerOnlyAndNeverDeleted() throws {
        let previous = umask(0o022)
        defer { umask(previous) }
        let url = try config()
        let store = DocumentProfileStore(configURL: url)
        try store.setDefaultProfile(.official)
        let lock = URL(fileURLWithPath: url.path + ".lock")
        XCTAssertEqual(try permissions(lock), 0o600)
        try store.setDefaultProfile(.inherit)
        XCTAssertTrue(FileManager.default.fileExists(atPath: lock.path))
        XCTAssertEqual(try store.settings().defaultProfile, .inherit)
    }

    /// The config directory is created (0700) before the lock is taken,
    /// because the lock file lives inside it.
    func testUpdateSucceedsWhenConfigDirectoryDoesNotExistYet() throws {
        let root = try config().deletingLastPathComponent()
        let url = root.appendingPathComponent("first-run/macdoc/config.json")
        try DocumentProfileStore(configURL: url).setDefaultProfile(.official)
        XCTAssertEqual(try DocumentProfileStore(configURL: url).settings().defaultProfile, .official)
        XCTAssertTrue(FileManager.default.fileExists(atPath: url.path + ".lock"))
        XCTAssertEqual(try permissions(url.deletingLastPathComponent()), 0o700)
    }

    /// A failing critical section still releases the lock.
    func testFailedUpdateReleasesTheLock() throws {
        let url = try config("invalid")
        XCTAssertThrowsError(try DocumentProfileStore(configURL: url).setDefaultProfile(.official))
        try Data("{}".utf8).write(to: url)
        XCTAssertNoThrow(try DocumentProfileStore(configURL: url, lockTimeout: 0.2, lockPollInterval: 0.02).setDefaultProfile(.official))
    }

    /// Concurrent writers — DocumentProfileStore instances and a writer that
    /// follows the same protocol for top-level keys, as pdf-to-latex-swift's
    /// AIConfig.save does — each set distinct keys; none is lost.
    func testConcurrentWritersEachSettingDistinctKeysAllSurvive() throws {
        let url = try config("{}")
        let writers = 8, rounds = 12
        @Sendable func otherWriter(_ key: String, _ value: Int) throws {
            try ConfigFileLock.withLock(forConfigAt: url.path) { () throws -> Void in
                var root = try JSONSerialization.jsonObject(with: Data(contentsOf: url)) as? [String: Any] ?? [:]
                root[key] = value
                try JSONSerialization.data(withJSONObject: root, options: [.sortedKeys]).write(to: url, options: .atomic)
            }
        }
        DispatchQueue.concurrentPerform(iterations: writers * 2) { index in
            let store = DocumentProfileStore(configURL: url)
            for round in 0..<rounds {
                do {
                    if index < writers { try store.updateDocument { $0["writer\(index)-\(round)"] = round } }
                    else { try otherWriter("other\(index)-\(round)", round) }
                } catch { XCTFail("writer \(index) round \(round): \(error)") }
            }
        }
        let root = try XCTUnwrap(JSONSerialization.jsonObject(with: Data(contentsOf: url)) as? [String: Any])
        let document = try XCTUnwrap(root["document"] as? [String: Any])
        for index in 0..<(writers * 2) {
            for round in 0..<rounds {
                if index < writers { XCTAssertEqual(document["writer\(index)-\(round)"] as? Int, round) }
                else { XCTAssertEqual(root["other\(index)-\(round)"] as? Int, round) }
            }
        }
    }

    // MARK: - PsychQuant/macdoc#194 resolve vs. garbage collection

    private final class Recorder: @unchecked Sendable {
        private let lock = NSLock()
        private var items: [String] = []
        func record(_ item: String) { lock.lock(); items.append(item); lock.unlock() }
        var recorded: [String] { lock.lock(); defer { lock.unlock() }; return items }
    }

    /// resolve holds the config lock from reading the snapshot reference
    /// until the snapshot bytes are read. An import + garbage collection
    /// attempted inside that window cannot get the lock, so the referenced
    /// snapshot cannot be switched away from and deleted mid-read.
    func testResolveHoldsTheLockFromReferenceReadToSnapshotRead() throws {
        let url = try config()
        let template = try officialTemplate()
        try DocumentProfileStore(configURL: url).importOfficial(from: template)
        let first = try XCTUnwrap(DocumentProfileStore(configURL: url).settings().officialSnapshot)
        let other = DocumentProfileStore(configURL: url, lockTimeout: 0.2, lockPollInterval: 0.02)
        let recorder = Recorder()
        let store = DocumentProfileStore(configURL: url, lockTimeout: 5, lockPollInterval: 0.05, afterSnapshotReferenceRead: {
            do {
                try other.importOfficial(from: template)
                _ = try other.garbageCollectOfficialSnapshots(dryRun: false)
                recorder.record("import and collection ran")
            } catch {
                recorder.record("\(error)")
            }
        })
        XCTAssertEqual(try store.resolve(explicit: .official, context: .newDocument)?.kind, .official)
        XCTAssertEqual(recorder.recorded, ["\(DocumentProfileStoreError.configLockTimeout(url.path + ".lock"))"])
        XCTAssertTrue(FileManager.default.fileExists(atPath: url.deletingLastPathComponent().appendingPathComponent(first).path))
        XCTAssertEqual(try DocumentProfileStore(configURL: url).settings().officialSnapshot, first)
    }

    /// Where no process can create the lock file (a read-only config
    /// directory) no writer can race either, so resolve reads without it.
    func testResolveReadsWithoutLockInAReadOnlyConfigDirectory() throws {
        let url = try config()
        try DocumentProfileStore(configURL: url).importOfficial(from: officialTemplate())
        let dir = url.deletingLastPathComponent()
        try FileManager.default.removeItem(atPath: url.path + ".lock")
        try FileManager.default.setAttributes([.posixPermissions: 0o500], ofItemAtPath: dir.path)
        defer { try? FileManager.default.setAttributes([.posixPermissions: 0o700], ofItemAtPath: dir.path) }
        XCTAssertEqual(try DocumentProfileStore(configURL: url).resolve(explicit: .official, context: .newDocument)?.kind, .official)
        XCTAssertFalse(FileManager.default.fileExists(atPath: url.path + ".lock"))
    }

    // MARK: - PsychQuant/macdoc#204 cross-process interoperability

    /// A separate process running an independent implementation of the
    /// protocol (perl's flock on `<config>.lock`) excludes this writer; the
    /// writer proceeds once that process exits and releases the lock.
    func testIndependentProcessHoldingTheLockExcludesTheWriter() throws {
        let perl = "/usr/bin/perl"
        guard FileManager.default.isExecutableFile(atPath: perl) else { throw XCTSkip("\(perl) is not available") }
        let url = try config(#"{"agent":"codex"}"#)
        let child = Process()
        child.executableURL = URL(fileURLWithPath: perl)
        child.arguments = ["-e", #"use Fcntl qw(:flock); open(my $f, ">>", $ARGV[0]) or die; flock($f, LOCK_EX) or die; print "locked\n"; $|=1; sleep 3"#,
                           url.path + ".lock"]
        let output = Pipe()
        child.standardOutput = output
        try child.run()
        addTeardownBlock { if child.isRunning { child.terminate() } }
        var announced = Data()
        while !String(decoding: announced, as: UTF8.self).contains("locked\n") {
            let chunk = output.fileHandleForReading.availableData
            if chunk.isEmpty { break }
            announced.append(chunk)
        }
        XCTAssertEqual(String(decoding: announced, as: UTF8.self), "locked\n")

        let store = DocumentProfileStore(configURL: url, lockTimeout: 0.3, lockPollInterval: 0.02)
        XCTAssertThrowsError(try store.updateDocument { $0["whileHeld"] = true }) { error in
            XCTAssertEqual(error as? DocumentProfileStoreError, .configLockTimeout(url.path + ".lock"))
        }
        XCTAssertTrue(child.isRunning, "the child must still hold the lock when the writer times out")
        child.waitUntilExit()
        XCTAssertEqual(child.terminationStatus, 0)
        try store.updateDocument { $0["afterRelease"] = true }
        let root = try XCTUnwrap(JSONSerialization.jsonObject(with: Data(contentsOf: url)) as? [String: Any])
        let document = try XCTUnwrap(root["document"] as? [String: Any])
        XCTAssertEqual(root["agent"] as? String, "codex")
        XCTAssertEqual(document["afterRelease"] as? Bool, true)
        XCTAssertNil(document["whileHeld"])
    }
}
