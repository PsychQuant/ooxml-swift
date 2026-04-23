import Foundation

/// Wraps the unzip tempDir created by `DocxReader.read(from:)` so that
/// `DocxWriter.write(_:to:)` in overlay mode can preserve OOXML parts the
/// typed model does not manage (theme, webSettings, people, glossary, etc.).
///
/// Lifecycle: `cleanup()` deletes the underlying directory. `WordDocument.close()`
/// calls this. Forgetting to call `close()` leaks the tempDir until process exit
/// (macOS reclaims `/tmp` on reboot, so the leak is bounded).
///
/// Added in v0.12.0 to fix the lossy round-trip in PsychQuant/che-word-mcp#23.
internal struct PreservedArchive {
    /// Filesystem URL of the unzipped source archive directory.
    let tempDir: URL

    init(tempDir: URL) {
        self.tempDir = tempDir
    }

    /// Delete the tempDir. Errors are silently swallowed (matching the
    /// underlying `ZipHelper.cleanup` semantics — best-effort cleanup).
    func cleanup() {
        ZipHelper.cleanup(tempDir)
    }
}
