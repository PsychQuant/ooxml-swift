import Foundation

public struct DocumentProfileSettings: Equatable, Sendable {
    public let defaultProfile: DocumentFormattingProfile.Kind
    public let officialSnapshot: String?
}

public enum DocumentProfileStoreError: Error, LocalizedError {
    case invalidConfiguration(String)
    case officialNotImported
    case invalidOfficialSnapshot

    public var errorDescription: String? {
        switch self {
        case .invalidConfiguration(let reason): return "文件格式設定無效：\(reason)"
        case .officialNotImported: return "尚未匯入 official 格式快照；請先執行 config document import-official。"
        case .invalidOfficialSnapshot: return "official 格式快照的 kind 必須是 official。"
        }
    }
}

/// Application adapters opt in explicitly; generic document writers never access settings.
public struct DocumentProfileStore: Sendable {
    public let configURL: URL

    public init(configURL: URL) { self.configURL = configURL }

    public static var defaultConfigURL: URL {
        FileManager.default.homeDirectoryForCurrentUser
            .appendingPathComponent(".config/macdoc/config.json")
    }

    public func settings() throws -> DocumentProfileSettings {
        try settings(in: readRoot())
    }

    /// Existing/replay requests without an explicit choice do no configuration I/O.
    public func resolve(
        explicit: DocumentFormattingProfile.Kind?, context: DocumentFormattingContext
    ) throws -> DocumentFormattingProfile? {
        if explicit == .inherit { return .inherit }
        if case .existingDocument = context, explicit == nil { return nil }
        let settings = try settings()
        let kind = explicit ?? settings.defaultProfile
        if kind == .inherit { return .inherit }
        guard let path = settings.officialSnapshot else { throw DocumentProfileStoreError.officialNotImported }
        let url = configURL.deletingLastPathComponent().appendingPathComponent(path)
        let profile = try JSONDecoder().decode(DocumentFormattingProfile.self, from: Data(contentsOf: url))
        try profile.validate()
        guard profile.kind == .official else { throw DocumentProfileStoreError.invalidOfficialSnapshot }
        return profile
    }

    public func setDefaultProfile(_ kind: DocumentFormattingProfile.Kind) throws {
        try updateDocument { $0["defaultProfile"] = kind.rawValue }
    }

    /// A fresh immutable snapshot is complete before its reference becomes visible.
    /// Importing changes only the reference, never the user's default selection.
    public func importOfficial(from templateURL: URL) throws {
        _ = try settings() // Refuse invalid configuration before any filesystem mutation.
        let profile = try DocumentFormattingProfile.importOfficial(from: templateURL)
        let encoder = JSONEncoder()
        encoder.outputFormatting = [.sortedKeys]
        let data = try encoder.encode(profile)
        let relative = "profiles/official-\(UUID().uuidString).json"
        let snapshot = configURL.deletingLastPathComponent().appendingPathComponent(relative)
        try FileManager.default.createDirectory(at: snapshot.deletingLastPathComponent(), withIntermediateDirectories: true)
        try data.write(to: snapshot, options: [.withoutOverwriting])
        do {
            try updateDocument { $0["officialSnapshot"] = relative }
        } catch {
            // Only our new unreferenced file is eligible for cleanup.
            try? FileManager.default.removeItem(at: snapshot)
            throw error
        }
    }

    private func readRoot() throws -> [String: Any] {
        let data: Data
        do { data = try Data(contentsOf: configURL) }
        catch let error as NSError where error.domain == NSCocoaErrorDomain && error.code == NSFileReadNoSuchFileError {
            return [:]
        }
        guard let root = try JSONSerialization.jsonObject(with: data, options: [.fragmentsAllowed]) as? [String: Any] else {
            throw DocumentProfileStoreError.invalidConfiguration("根節點必須是物件")
        }
        return root
    }

    private func settings(in root: [String: Any]) throws -> DocumentProfileSettings {
        let document: [String: Any]
        if let raw = root["document"] {
            guard let object = raw as? [String: Any] else {
                throw DocumentProfileStoreError.invalidConfiguration("document 必須是物件")
            }
            document = object
        } else { document = [:] }
        var kind: DocumentFormattingProfile.Kind = .inherit
        if let raw = document["defaultProfile"] {
            guard let value = raw as? String, let parsed = DocumentFormattingProfile.Kind(rawValue: value) else {
                throw DocumentProfileStoreError.invalidConfiguration("defaultProfile 必須是 inherit 或 official")
            }
            kind = parsed
        }
        var snapshot: String?
        if let raw = document["officialSnapshot"] {
            guard let value = raw as? String, !value.isEmpty, !value.hasPrefix("/"),
                  !value.split(separator: "/", omittingEmptySubsequences: false).contains(where: { $0 == ".." || $0 == "." || $0.isEmpty }) else {
                throw DocumentProfileStoreError.invalidConfiguration("officialSnapshot 必須是設定目錄內的相對檔案路徑")
            }
            snapshot = value
        }
        return DocumentProfileSettings(defaultProfile: kind, officialSnapshot: snapshot)
    }

    private func updateDocument(_ update: (inout [String: Any]) -> Void) throws {
        var root = try readRoot()
        _ = try settings(in: root)
        var document = root["document"] as? [String: Any] ?? [:]
        update(&document)
        root["document"] = document
        let bytes = try JSONSerialization.data(withJSONObject: root, options: [.prettyPrinted, .sortedKeys, .withoutEscapingSlashes])
        try FileManager.default.createDirectory(at: configURL.deletingLastPathComponent(), withIntermediateDirectories: true)
        try bytes.write(to: configURL, options: .atomic)
    }
}
