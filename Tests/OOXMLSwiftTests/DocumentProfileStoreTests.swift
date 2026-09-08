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
}
