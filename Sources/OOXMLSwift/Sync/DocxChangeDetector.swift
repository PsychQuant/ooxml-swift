// DocxChangeDetector.swift
// word-aligned-state-sync Phase 3 tasks 4.5 + 4.6 — file watcher contract +
// Word lock-file interaction (`ooxml-word-sync` Requirements "File watcher
// contract" + "Word file-lock interaction").
//
// Polling by design: the spec forbids depending on fsevents/inotify
// (portability across supported macOS versions); mtime is the cheap
// first-line check and SHA-256 the authoritative one — an mtime bump with
// unchanged bytes (e.g. `touch`) must NOT trigger an import.

import Foundation

/// Detects real content changes of a docx between polls.
///
/// Value type holding the last-seen `(mtime, sha256)` baseline. `poll()`
/// advances the baseline whenever it reports `true`, so each content change
/// is reported exactly once.
public struct DocxChangeDetector {
    public let url: URL
    private var lastModificationDate: Date?
    private var lastContentHash: String

    public init(url: URL) throws {
        self.url = url
        self.lastModificationDate = Self.modificationDate(of: url)
        self.lastContentHash = SidecarStore.sha256Hex(of: try Data(contentsOf: url))
    }

    /// Returns `true` when the file's content actually changed since the
    /// last baseline. mtime fast-path: unchanged mtime → no hash computed.
    /// mtime changed but hash identical → baseline mtime updates, returns
    /// `false` (spec scenario "mtime-only change without content change is
    /// ignored").
    public mutating func poll() throws -> Bool {
        let mtime = Self.modificationDate(of: url)
        if mtime == lastModificationDate { return false }
        lastModificationDate = mtime

        let hash = SidecarStore.sha256Hex(of: try Data(contentsOf: url))
        if hash == lastContentHash { return false }
        lastContentHash = hash
        return true
    }

    private static func modificationDate(of url: URL) -> Date? {
        (try? FileManager.default.attributesOfItem(atPath: url.path))?[.modificationDate] as? Date
    }
}

/// Word owner-file ("lock file") detection.
public enum WordLock {

    /// Spec-literal owner-file path: `report.docx` → `~$report.docx` in the
    /// same directory.
    public static func lockFileURL(for docxURL: URL) -> URL {
        docxURL.deletingLastPathComponent()
            .appendingPathComponent("~$" + docxURL.lastPathComponent)
    }

    /// True when Word holds the docx open. Checks the spec-literal name and
    /// the historical Word naming variant that drops the first two filename
    /// characters (8.3-era legacy: `mydocument.docx` → `~$document.docx`).
    public static func isLockedByWord(_ docxURL: URL) -> Bool {
        let fm = FileManager.default
        if fm.fileExists(atPath: lockFileURL(for: docxURL).path) { return true }

        let name = docxURL.lastPathComponent
        if name.count > 2 {
            let minusTwo = docxURL.deletingLastPathComponent()
                .appendingPathComponent("~$" + String(name.dropFirst(2)))
            if fm.fileExists(atPath: minusTwo.path) { return true }
        }
        return false
    }
}
