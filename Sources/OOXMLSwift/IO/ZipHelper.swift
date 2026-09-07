import Foundation
import ZIPFoundation

/// ZIP 壓縮/解壓縮工具
public struct ZipHelper {
    /// 解壓縮 ZIP 檔案到臨時目錄
    ///
    /// Refuses an archive that carries a symbolic-link entry, an entry whose
    /// path has a `..` component, an absolute entry path, or an empty or
    /// NUL-containing entry path (v3.7.0; the five are listed once more on
    /// `unzip(data:)`, which this calls). Either
    /// half on its own is enough to write outside the destination: a link to
    /// `.` inside the archive makes the kernel resolve later `..` components
    /// in place while ZIPFoundation's containment check collapses them
    /// lexically (verify R4 security: `word/a → .` then
    /// `word/a/a/../../../x` lands in the temporary directory's parent, and
    /// with enough components anywhere the process may write). A link is
    /// also followed by every read after it, so `word/alias.xml →
    /// document.xml` would be a second document. No Word output contains
    /// any of these (0 / 740 in the real corpus); refusing them here keeps
    /// one policy for the reader and the inspector.
    public static func unzip(_ url: URL) throws -> URL {
        try unzip(data: try Data(contentsOf: url))     // one read; everything below works on these bytes
    }

    /// The reader's extraction namespace under the temporary directory.
    public static let readerNamespace = "che-word-mcp"
    /// The inspector's (v3.7.0): the same policy, a different directory, so
    /// the reader's "did I leave an extraction directory behind" question
    /// keeps its answer when an inspection runs alongside (verify R5).
    public static let inspectorNamespace = "ooxml-swift-inspector"

    /// Extract a package whose bytes are already in hand. Refused before
    /// anything is written: an entry ZIPFoundation would extract as a
    /// symbolic link (`Entry.type == .symlink`, which is what it consults
    /// when extracting), an empty or NUL-containing path, an absolute path,
    /// a path with a `..` component. The policy pre-scan
    /// and the extraction see the **same immutable bytes**: the bytes are
    /// scanned as an in-memory archive and then written to a private,
    /// UUID-named copy inside the fresh temporary directory, which is what
    /// gets extracted (verify R5: pre-scanning one open of a URL and then
    /// extracting a second open let a source file swapped in between bring
    /// the symlink + `..` chain back). Nothing is written before the policy
    /// scan passes.
    /// Public entry for bytes already in hand (the reader's namespace).
    public static func unzip(data: Data) throws -> URL { try unzip(data: data, namespace: readerNamespace) }

    static func unzip(data: Data, namespace: String) throws -> URL {
        precondition(!namespace.isEmpty && !namespace.contains("/") && namespace != "." && namespace != "..", "namespace must be one path component")
        let archive = try Archive(data: data, accessMode: .read)
        for entry in archive {
            if entry.type == .symlink {
                throw WordError.invalidDocx("the package contains a symbolic-link entry (\(entry.path)); refusing to extract it.")
            }
            let path = entry.path
            if path.isEmpty || path.contains("\0") {
                throw WordError.invalidDocx("the package contains an entry with an empty or NUL-containing path; refusing to extract it.")
            }
            let components = path.split(separator: "/", omittingEmptySubsequences: false)
            if path.hasPrefix("/") || components.contains("..") {
                throw WordError.invalidDocx("the package contains an entry whose path leaves its own directory (\(path)); refusing to extract it.")
            }
        }
        let fm = FileManager.default
        // The namespace directory is owner-only from the moment it exists,
        // and it must be a real directory of ours — a pre-planted symbolic
        // link there would redirect every extraction (verify R6).
        let namespaceDir = fm.temporaryDirectory.appendingPathComponent(namespace)
        if let attrs = try? fm.attributesOfItem(atPath: namespaceDir.path) {
            guard (attrs[.type] as? FileAttributeType) == .typeDirectory else {     // attributesOfItem does not follow links
                throw WordError.invalidDocx("the extraction namespace \(namespaceDir.path) exists but is not a directory; refusing to extract.")
            }
            if let owner = (attrs[.ownerAccountID] as? NSNumber)?.uint32Value, owner != getuid() {
                throw WordError.invalidDocx("the extraction namespace \(namespaceDir.path) is owned by another user (uid \(owner)); refusing to extract.")
            }
            if (attrs[.posixPermissions] as? Int) != 0o700 {                  // created by an older version, or by hand
                try fm.setAttributes([.posixPermissions: 0o700], ofItemAtPath: namespaceDir.path)
            }
        } else {
            try fm.createDirectory(at: namespaceDir, withIntermediateDirectories: true, attributes: [.posixPermissions: 0o700])
        }
        let tempDir = namespaceDir.appendingPathComponent(UUID().uuidString)
        try fm.createDirectory(at: tempDir, withIntermediateDirectories: false, attributes: [.posixPermissions: 0o700])
        do {
            // A UUID name cannot collide with any entry the archive declares;
            // the copy is owner-only from creation.
            let privateCopy = tempDir.appendingPathComponent(UUID().uuidString + ".zip")
            guard fm.createFile(atPath: privateCopy.path, contents: data, attributes: [.posixPermissions: 0o600]) else {
                throw WordError.invalidDocx("could not write the package's private copy under \(tempDir.path)")
            }
            try fm.unzipItem(at: privateCopy, to: tempDir)
            try fm.removeItem(at: privateCopy)
            // ZIPFoundation applies the archive's own permission bits — setuid,
            // setgid, sticky, world-writable included (verify R5). Nothing in a
            // package needs any of that: owner-only, no special bits, on every
            // item including the root; a walk that cannot complete is an error,
            // not a partial result (verify R6).
            var walkError: Error?
            guard let walker = fm.enumerator(at: tempDir, includingPropertiesForKeys: [.isDirectoryKey], options: [], errorHandler: { _, error in walkError = error; return false }) else {
                throw WordError.invalidDocx("could not enumerate the extracted package under \(tempDir.path)")
            }
            for case let item as URL in walker {
                let isDirectory = try item.resourceValues(forKeys: [.isDirectoryKey]).isDirectory ?? false
                try fm.setAttributes([.posixPermissions: isDirectory ? 0o700 : 0o600], ofItemAtPath: item.path)
            }
            if let walkError { throw walkError }
            try fm.setAttributes([.posixPermissions: 0o700], ofItemAtPath: tempDir.path)
        } catch {
            try? fm.removeItem(at: tempDir)   // nothing to keep from a failed extraction
            throw error
        }

        return tempDir
    }

    /// 壓縮目錄內容為 ZIP 檔案（不包含目錄本身的路徑）
    public static func zip(_ directory: URL, to destination: URL) throws {
        let data = try zipToData(directory)

        if FileManager.default.fileExists(atPath: destination.path) {
            try FileManager.default.removeItem(at: destination)
        }
        try data.write(to: destination)
    }

    /// 壓縮目錄內容為 in-memory ZIP bytes（不寫入磁碟）
    public static func zipToData(_ directory: URL) throws -> Data {
        let archive: Archive
        do {
            archive = try Archive(accessMode: .create)
        } catch {
            throw WordError.zipError("無法建立 in-memory ZIP archive: \(error)")
        }

        let files = try getAllFiles(in: directory)

        for (relativePath, fileURL) in files {
            let fileData = try Data(contentsOf: fileURL)
            try archive.addEntry(
                with: relativePath,
                type: .file,
                uncompressedSize: Int64(fileData.count),
                compressionMethod: .deflate,
                provider: { position, size in
                    let startIndex = fileData.startIndex.advanced(by: Int(position))
                    let endIndex = startIndex.advanced(by: size)
                    return fileData.subdata(in: startIndex..<endIndex)
                }
            )
        }

        guard let data = archive.data else {
            throw WordError.zipError("in-memory ZIP archive 無 data 可讀")
        }
        return data
    }

    /// 取得目錄內所有檔案（回傳相對路徑和完整 URL 的配對）
    private static func getAllFiles(in directory: URL) throws -> [(String, URL)] {
        var result: [(String, URL)] = []
        let fileManager = FileManager.default

        // 使用 subpathsOfDirectory 取得所有子路徑
        let directoryPath = directory.path
        let subpaths = try fileManager.subpathsOfDirectory(atPath: directoryPath)

        for subpath in subpaths {
            let fullURL = directory.appendingPathComponent(subpath)
            var isDirectory: ObjCBool = false
            if fileManager.fileExists(atPath: fullURL.path, isDirectory: &isDirectory) {
                if !isDirectory.boolValue {
                    result.append((subpath, fullURL))
                }
            }
        }

        return result
    }

    /// 清理臨時目錄
    public static func cleanup(_ directory: URL) {
        try? FileManager.default.removeItem(at: directory)
    }
}
