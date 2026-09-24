import Foundation

public struct DocumentProfileSettings: Equatable, Sendable {
    public let defaultProfile: DocumentFormattingProfile.Kind
    public let officialSnapshot: String?
}

public enum DocumentProfileStoreError: Error, Equatable, LocalizedError {
    case invalidConfiguration(String)
    case officialNotImported
    case invalidOfficialSnapshot
    case snapshotFileMissing(String)
    case configLockTimeout(String)

    public var errorDescription: String? {
        switch self {
        case .invalidConfiguration(let reason): return "文件格式設定無效：\(reason)"
        case .officialNotImported: return "尚未匯入 official 格式快照；請先執行 config document import-official。"
        case .invalidOfficialSnapshot: return "official 格式快照的 kind 必須是 official。"
        case .configLockTimeout(let path): return "無法在時限內取得設定檔鎖（另一個程序持有鎖）：\(path)"
        case .snapshotFileMissing(let path): return "official 格式快照檔不存在：\(path)；請重新執行 config document import-official。"
        }
    }
}

/// Application adapters opt in explicitly; generic document writers never access settings.
///
/// Trust boundary (PsychQuant/macdoc#194): the store lives in the invoking
/// user's own configuration directory (`~/.config/macdoc` by default) and
/// assumes that directory is single-user, not multi-tenant. Owner-only
/// permissions are defense in depth, not an isolation mechanism: the
/// profiles directory is created 0700, and every snapshot and config file
/// this store writes is created 0600 (independent of the process umask).
/// A newly created configuration directory is also 0700; an existing one
/// is left as the user set it.
public struct DocumentProfileStore: Sendable {
    public let configURL: URL

    /// How long a writer waits for the config lock, and how often it polls
    /// (PsychQuant/macdoc#204). Production uses the protocol defaults; the
    /// internal initializer exists so tests need not wait five seconds.
    let lockTimeout: TimeInterval
    let lockPollInterval: TimeInterval
    /// Test seam: runs in `resolve` after the snapshot reference is read and
    /// before the snapshot bytes are, to exercise that window.
    let afterSnapshotReferenceRead: (@Sendable () -> Void)?

    public init(configURL: URL) {
        self.init(configURL: configURL, lockTimeout: ConfigFileLock.defaultTimeout,
                  lockPollInterval: ConfigFileLock.defaultPollInterval)
    }

    internal init(configURL: URL, lockTimeout: TimeInterval, lockPollInterval: TimeInterval,
                  afterSnapshotReferenceRead: (@Sendable () -> Void)? = nil) {
        self.configURL = configURL
        self.lockTimeout = lockTimeout
        self.lockPollInterval = lockPollInterval
        self.afterSnapshotReferenceRead = afterSnapshotReferenceRead
    }

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
        let kind = try explicit ?? settings().defaultProfile
        if kind == .inherit { return .inherit }
        // The reference and the snapshot bytes are read under the config
        // lock, so an import + garbage collection cannot switch the
        // reference and delete the snapshot in between (PsychQuant/macdoc#194).
        return try withSnapshotReadLock { try officialSnapshotHoldingLock() }
    }

    private func officialSnapshotHoldingLock() throws -> DocumentFormattingProfile {
        guard let path = try settings().officialSnapshot else { throw DocumentProfileStoreError.officialNotImported }
        afterSnapshotReferenceRead?()
        let url = configURL.deletingLastPathComponent().appendingPathComponent(path)
        let data: Data
        do { data = try Data(contentsOf: url) }
        catch let error as NSError where error.domain == NSCocoaErrorDomain && error.code == NSFileReadNoSuchFileError {
            throw DocumentProfileStoreError.snapshotFileMissing(path)
        }
        let profile = try JSONDecoder().decode(DocumentFormattingProfile.self, from: data)
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
        // The snapshot and its reference are published under one lock hold,
        // so garbage collection never sees the new file unreferenced.
        try withConfigLock {
            try writeNewSnapshot(data, relativePath: relative)
            do {
                try updateDocumentHoldingLock { $0["officialSnapshot"] = relative }
            } catch {
                // Only our new unreferenced file is eligible for cleanup.
                try? FileManager.default.removeItem(at: snapshot)
                throw error
            }
        }
    }

    /// Writes a new immutable snapshot 0600 inside the 0700 profiles
    /// directory. Never overwrites: an existing file or symlink at the name
    /// fails the write (the `.withoutOverwriting` contract).
    internal func writeNewSnapshot(_ data: Data, relativePath relative: String) throws {
        let snapshot = configURL.deletingLastPathComponent().appendingPathComponent(relative)
        try Self.createOwnerOnlyDirectory(configURL.deletingLastPathComponent())
        let profiles = snapshot.deletingLastPathComponent()
        try Self.createOwnerOnlyDirectory(profiles)
        // The profiles directory belongs to this store alone; tighten one
        // created by an earlier version under a permissive umask.
        try FileManager.default.setAttributes([.posixPermissions: 0o700], ofItemAtPath: profiles.path)
        try Self.createOwnerOnlyFile(data, at: snapshot)
    }

    /// Lists the `profiles/official-*.json` snapshots that the current
    /// `officialSnapshot` does not reference, as config-relative paths, and
    /// deletes them unless `dryRun`. The referenced snapshot is never
    /// touched; other files, directories and symlinks in `profiles/` are
    /// ignored; an invalid configuration refuses before anything is deleted.
    /// Runs under the config lock, so a concurrent import's new snapshot is
    /// either already referenced or not yet written.
    public func garbageCollectOfficialSnapshots(dryRun: Bool) throws -> [String] {
        guard FileManager.default.fileExists(atPath: configURL.deletingLastPathComponent().path) else {
            return []
        }
        return try withConfigLock { try collectOfficialSnapshotsHoldingLock(dryRun: dryRun) }
    }

    private func collectOfficialSnapshotsHoldingLock(dryRun: Bool) throws -> [String] {
        let referenced = try settings().officialSnapshot
        let profiles = configURL.deletingLastPathComponent().appendingPathComponent("profiles")
        let names: [String]
        do { names = try FileManager.default.contentsOfDirectory(atPath: profiles.path) }
        catch let error as NSError where error.domain == NSCocoaErrorDomain && error.code == NSFileReadNoSuchFileError {
            return []
        }
        let unreferenced = try names.filter { name in
            guard name.hasPrefix("official-"), name.hasSuffix(".json"), "profiles/" + name != referenced else { return false }
            let attributes = try FileManager.default.attributesOfItem(atPath: profiles.appendingPathComponent(name).path)
            return attributes[.type] as? FileAttributeType == .typeRegular
        }.map { "profiles/" + $0 }.sorted()
        if !dryRun {
            for path in unreferenced {
                try FileManager.default.removeItem(at: configURL.deletingLastPathComponent().appendingPathComponent(path))
            }
        }
        return unreferenced
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

    /// Read → merge → atomic write of the `document` object, holding the
    /// cross-process config lock for the whole cycle (PsychQuant/macdoc#204).
    internal func updateDocument(_ update: (inout [String: Any]) -> Void) throws {
        try withConfigLock { try updateDocumentHoldingLock(update) }
    }

    /// Takes the config lock for a read without creating anything but the
    /// lock file. When the lock file does not exist and cannot be created
    /// (the configuration directory is missing or not writable), no writer
    /// can create or delete files there either, so the read runs unlocked.
    private func withSnapshotReadLock<T>(_ body: () throws -> T) throws -> T {
        let directory = configURL.deletingLastPathComponent().path
        if !FileManager.default.fileExists(atPath: configURL.path + ".lock"), access(directory, W_OK) != 0 {
            return try body()
        }
        return try ConfigFileLock.withLock(forConfigAt: configURL.path, pollInterval: lockPollInterval,
                                           timeout: lockTimeout, body)
    }

    /// Creates the config directory first — the lock file lives inside it —
    /// then runs `body` under `ConfigFileLock`.
    private func withConfigLock<T>(_ body: () throws -> T) throws -> T {
        try Self.createOwnerOnlyDirectory(configURL.deletingLastPathComponent())
        return try ConfigFileLock.withLock(forConfigAt: configURL.path, pollInterval: lockPollInterval,
                                           timeout: lockTimeout, body)
    }

    private func updateDocumentHoldingLock(_ update: (inout [String: Any]) -> Void) throws {
        var root = try readRoot()
        _ = try settings(in: root)
        var document = root["document"] as? [String: Any] ?? [:]
        update(&document)
        root["document"] = document
        let bytes = try JSONSerialization.data(withJSONObject: root, options: [.prettyPrinted, .sortedKeys, .withoutEscapingSlashes])
        try Self.createOwnerOnlyDirectory(configURL.deletingLastPathComponent())
        try Self.replaceOwnerOnlyFile(bytes, at: configURL)
    }

    /// Creates `url` (and missing parents) 0700; an existing directory is
    /// left as it is.
    private static func createOwnerOnlyDirectory(_ url: URL) throws {
        try FileManager.default.createDirectory(at: url, withIntermediateDirectories: true,
                                                attributes: [.posixPermissions: 0o700])
    }

    private static func posixError(_ code: Int32, _ url: URL) -> NSError {
        NSError(domain: NSPOSIXErrorDomain, code: Int(code), userInfo: [NSFilePathErrorKey: url.path])
    }

    /// Creates a new 0600 file (the mode is fixed with fchmod, so the umask
    /// cannot widen it) and fails if anything, including a symlink, already
    /// exists at `url`. A partially written file is removed.
    private static func createOwnerOnlyFile(_ data: Data, at url: URL) throws {
        let fd = open(url.path, O_WRONLY | O_CREAT | O_EXCL | O_NOFOLLOW | O_CLOEXEC, 0o600)
        guard fd >= 0 else {
            let code = errno
            if code == EEXIST { throw CocoaError(.fileWriteFileExists, userInfo: [NSFilePathErrorKey: url.path]) }
            throw posixError(code, url)
        }
        do {
            guard fchmod(fd, 0o600) == 0 else { throw posixError(errno, url) }
            try data.withUnsafeBytes { buffer in
                var offset = 0
                while offset < buffer.count {
                    let written = write(fd, buffer.baseAddress! + offset, buffer.count - offset)
                    if written < 0 {
                        if errno == EINTR { continue }
                        throw posixError(errno, url)
                    }
                    offset += written
                }
            }
            guard fsync(fd) == 0 else { throw posixError(errno, url) }
        } catch {
            close(fd)
            unlink(url.path)
            throw error
        }
        guard close(fd) == 0 else {
            let code = errno
            unlink(url.path)
            throw posixError(code, url)
        }
    }

    /// Atomically replaces `url` with a new 0600 file: write a sibling
    /// temporary file, then rename(2) it over the destination (which
    /// replaces a symlink there rather than following it).
    private static func replaceOwnerOnlyFile(_ data: Data, at url: URL) throws {
        let temporary = url.deletingLastPathComponent()
            .appendingPathComponent(".\(url.lastPathComponent).\(UUID().uuidString).tmp")
        try createOwnerOnlyFile(data, at: temporary)
        guard rename(temporary.path, url.path) == 0 else {
            let code = errno
            unlink(temporary.path)
            throw posixError(code, url)
        }
    }
}

/// Cross-process advisory lock for `config.json` (PsychQuant/macdoc#204).
///
/// The same file has another writer: pdf-to-latex-swift's `AIConfig.save`
/// (its `ConfigFileLock`) persists AI-CLI/OCR settings into
/// `~/.config/macdoc/config.json`, while `DocumentProfileStore` persists the
/// `document` object there. Both read the file, merge their own fields and
/// atomically replace it; without a shared lock one writer's replacement can
/// silently discard the other's update.
///
/// Protocol — it MUST stay identical in both implementations, or the lock
/// protects nothing:
/// - Lock file: `<config path>.lock`, beside the config file. The config
///   itself is replaced by rename (a new inode), so it cannot be the lock.
/// - The config directory is created before the lock is taken.
/// - `open(O_CREAT | O_RDWR | O_CLOEXEC, 0600)`, then `fchmod(fd, 0600)`
///   because the open mode is only a request under the umask.
/// - `flock(fd, LOCK_EX | LOCK_NB)` polled every 50 ms; give up with an error
///   after 5 s.
/// - Hold the lock across the whole read → merge → atomic write.
/// - Release with `flock(fd, LOCK_UN)`, then `close(fd)`.
/// - Never delete the lock file: a recreated one is a different inode, and
///   writers flocking different inodes exclude nothing.
enum ConfigFileLock {
    static let defaultPollInterval: TimeInterval = 0.05
    static let defaultTimeout: TimeInterval = 5.0

    static func withLock<T>(
        forConfigAt path: String,
        pollInterval: TimeInterval = ConfigFileLock.defaultPollInterval,
        timeout: TimeInterval = ConfigFileLock.defaultTimeout,
        _ body: () throws -> T
    ) throws -> T {
        let lockPath = path + ".lock"
        let fd = open(lockPath, O_CREAT | O_RDWR | O_CLOEXEC, 0o600)
        guard fd >= 0 else {
            throw NSError(domain: NSPOSIXErrorDomain, code: Int(errno), userInfo: [NSFilePathErrorKey: lockPath])
        }
        defer { close(fd) }
        _ = fchmod(fd, 0o600)

        let deadline = Date().addingTimeInterval(timeout)
        while flock(fd, LOCK_EX | LOCK_NB) != 0 {
            if Date() >= deadline { throw DocumentProfileStoreError.configLockTimeout(lockPath) }
            Thread.sleep(forTimeInterval: pollInterval)
        }
        defer { flock(fd, LOCK_UN) }

        return try body()
    }
}
