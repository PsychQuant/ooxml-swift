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
    /// Public entry for bytes already in hand (the reader's namespace); the
    /// five refused entry kinds are the ones listed just above.
    public static func unzip(data: Data) throws -> URL { try unzip(data: data, namespace: readerNamespace) }

    /// The same, into a named namespace (one path component) under the
    /// temporary directory. Refuses the same five entry kinds.
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
        // The namespace directory is owner-only from the moment it exists, and
        // it must be a real directory of ours. The path is never trusted twice
        // (verify R7 codex R7-1: "check, then create" let another uid plant a
        // directory between the two): `mkdir` creates it atomically or reports
        // EEXIST, the directory is then OPENED without following a link, and it
        // is the open descriptor — not the path — whose type, owner and mode are
        // verified. Everything below is created relative to that descriptor.
        let namespaceDir = fm.temporaryDirectory.appendingPathComponent(namespace)
        let namespaceFD = try ownerOnlyDirectoryDescriptor(creatingIfAbsent: namespaceDir.path)
        defer { close(namespaceFD) }

        let uuid = UUID().uuidString
        let tempDir = namespaceDir.appendingPathComponent(uuid)
        // Cleanup is armed BEFORE the directory exists (codex R7-5): a creation
        // that fails after the directory appeared still leaves nothing behind.
        var succeeded = false
        // A failed extraction leaves a tree the archive may have made
        // unremovable (a directory entry stored 0500 or 0400 — verify R7 logic
        // N-L1-R7: `removeItem` cannot empty it, and the whole tree including
        // the private copy of the package stayed behind); every directory is
        // made ours again before removal, and a removal that still fails is
        // reported on stderr — it is the one thing this function cannot throw.
        defer { if !succeeded { removeTreeForcibly(tempDir) } }
        guard mkdirat(namespaceFD, uuid, 0o700) == 0 else {
            throw WordError.invalidDocx("could not create the extraction directory \(tempDir.path) (errno \(errno))")
        }
        let tempFD = try ownerOnlyDirectoryDescriptor(opening: uuid, relativeTo: namespaceFD, describedAs: tempDir.path)
        defer { close(tempFD) }

        // A UUID name cannot collide with any entry the archive declares; the
        // copy is created 0600 atomically (O_CREAT|O_EXCL with that mode — the
        // inode never exists with another mode, codex R7-6) and written through
        // the descriptor, then extracted from the path it now has.
        let copyName = UUID().uuidString + ".zip"
        let copyFD = openat(tempFD, copyName, O_WRONLY | O_CREAT | O_EXCL | O_NOFOLLOW | O_CLOEXEC, 0o600)
        guard copyFD >= 0 else {
            throw WordError.invalidDocx("could not create the package's private copy under \(tempDir.path) (errno \(errno))")
        }
        let copyHandle = FileHandle(fileDescriptor: copyFD, closeOnDealloc: true)
        do { try copyHandle.write(contentsOf: data); try copyHandle.close() }
        catch { throw WordError.invalidDocx("could not write the package's private copy under \(tempDir.path): \(error.localizedDescription)") }
        let privateCopy = tempDir.appendingPathComponent(copyName)
        // Every error from here on is ours to name (verify R7 logic N-L2-R7: a
        // raw POSIX error carried the temporary path to the caller).
        do { try fm.unzipItem(at: privateCopy, to: tempDir) }
        catch { throw WordError.invalidDocx("the package could not be extracted (\(describeWithoutPaths(error))); refusing it.") }
        do { try fm.removeItem(at: privateCopy) }
        catch { throw WordError.invalidDocx("the package's private copy could not be removed after extraction (\(describeWithoutPaths(error)))") }

        // ZIPFoundation applies the archive's own permission bits — setuid,
        // setgid, sticky, world-writable included (verify R5). Nothing in a
        // package needs any of that: owner-only, no special bits, on every
        // item including the root; a walk that cannot complete is an error,
        // not a partial result (verify R6), and an item whose kind cannot be
        // determined is an error too, not a file (codex R7-4). The root is set
        // before the walk (so a `./` entry cannot leave it unlistable) and
        // again after it.
        guard fchmod(tempFD, 0o700) == 0 else { throw WordError.invalidDocx("could not make \(tempDir.path) owner-only (errno \(errno))") }
        var walkError: Error?
        guard let walker = fm.enumerator(at: tempDir, includingPropertiesForKeys: [.isDirectoryKey], options: [], errorHandler: { _, error in walkError = error; return false }) else {
            throw WordError.invalidDocx("could not enumerate the extracted package under \(tempDir.path)")
        }
        for case let item as URL in walker {
            let relative = String(item.path.dropFirst(tempDir.path.count + 1))
            guard let isDirectory = try? item.resourceValues(forKeys: [.isDirectoryKey]).isDirectory else {
                throw WordError.invalidDocx("could not determine whether \(relative) in the package is a directory; refusing a partial permission reset.")
            }
            do { try fm.setAttributes([.posixPermissions: isDirectory ? 0o700 : 0o600], ofItemAtPath: item.path) }
            catch { throw WordError.invalidDocx("could not make \(relative) in the package owner-only (\(describeWithoutPaths(error)))") }
        }
        if let walkError { throw WordError.invalidDocx("could not enumerate the extracted package (\(describeWithoutPaths(walkError)))") }
        guard fchmod(tempFD, 0o700) == 0 else { throw WordError.invalidDocx("could not make \(tempDir.path) owner-only (errno \(errno))") }
        succeeded = true
        return tempDir
    }

    /// Make every directory under `root` ours again (0700, pre-order so the
    /// walk can descend) and remove the tree. Best effort by construction —
    /// the failure path cannot throw — so a tree that still cannot be removed
    /// is reported on stderr with its path, the one place a path is useful.
    static func removeTreeForcibly(_ root: URL) {
        let fm = FileManager.default
        chmod(root.path, 0o700)
        if let walker = fm.enumerator(at: root, includingPropertiesForKeys: [.isDirectoryKey], options: [], errorHandler: { _, _ in true }) {
            for case let item as URL in walker {
                var st = stat()
                guard lstat(item.path, &st) == 0 else { continue }
                chmod(item.path, (st.st_mode & S_IFMT) == S_IFDIR ? 0o700 : 0o600)
            }
        }
        do { try fm.removeItem(at: root) }
        catch { FileHandle.standardError.write(Data("ooxml-swift: could not remove the partial extraction at \(root.path) after a failure (\(error.localizedDescription))\n".utf8)) }
    }

    /// An error's description with no file-system path in it (the temporary
    /// directory is not the caller's business — #146).
    static func describeWithoutPaths(_ error: Error) -> String {
        let ns = error as NSError
        let text = ns.localizedFailureReason ?? ns.localizedDescription
        return text.split(separator: " ").filter { !$0.hasPrefix("/") && !$0.hasPrefix("“/") }.joined(separator: " ")
    }

    /// `mkdir(path, 0700)` — atomic create or `EEXIST` — then open the directory
    /// without following a link and verify the DESCRIPTOR: a directory, owned
    /// by this uid, mode 0700 (any other mode of our own directory is reset;
    /// anyone else's, or a file or link planted at the name, is refused).
    private static func ownerOnlyDirectoryDescriptor(creatingIfAbsent path: String) throws -> Int32 {
        if mkdir(path, 0o700) != 0, errno != EEXIST {
            throw WordError.invalidDocx("could not create the extraction namespace \(path) (errno \(errno))")
        }
        let fd = open(path, O_RDONLY | O_DIRECTORY | O_NOFOLLOW | O_CLOEXEC)
        guard fd >= 0 else {
            throw WordError.invalidDocx("the extraction namespace \(path) cannot be opened as a real directory (errno \(errno)); a symbolic link or a file planted there is refused.")
        }
        do { try verifyOwnerOnlyDirectory(fd, describedAs: path) } catch { close(fd); throw error }
        return fd
    }

    private static func ownerOnlyDirectoryDescriptor(opening name: String, relativeTo parentFD: Int32, describedAs path: String) throws -> Int32 {
        let fd = openat(parentFD, name, O_RDONLY | O_DIRECTORY | O_NOFOLLOW | O_CLOEXEC)
        guard fd >= 0 else { throw WordError.invalidDocx("could not open the extraction directory \(path) just created (errno \(errno))") }
        do { try verifyOwnerOnlyDirectory(fd, describedAs: path) } catch { close(fd); throw error }
        return fd
    }

    private static func verifyOwnerOnlyDirectory(_ fd: Int32, describedAs path: String) throws {
        var st = stat()
        guard fstat(fd, &st) == 0 else { throw WordError.invalidDocx("could not stat \(path) (errno \(errno))") }
        guard (st.st_mode & S_IFMT) == S_IFDIR else { throw WordError.invalidDocx("\(path) exists but is not a directory; refusing to extract.") }
        guard st.st_uid == getuid() else { throw WordError.invalidDocx("\(path) is owned by another user (uid \(st.st_uid)); refusing to extract.") }
        if (st.st_mode & 0o7777) != 0o700 {                                  // created by an older version, or by hand
            guard fchmod(fd, 0o700) == 0 else { throw WordError.invalidDocx("could not make \(path) owner-only (errno \(errno))") }
        }
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
