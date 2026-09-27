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
    public static func unzip(_ url: URL, limits: Limits = defaultLimits) throws -> URL {
        try unzip(data: try Data(contentsOf: url), limits: limits)     // one read; everything below works on these bytes
    }

    /// The reader's extraction namespace under the temporary directory.
    public static let readerNamespace = "che-word-mcp"
    /// The inspector's (v3.7.0): the same policy, a different directory, so
    /// the reader's "did I leave an extraction directory behind" question
    /// keeps its answer when an inspection runs alongside (verify R5).
    public static let inspectorNamespace = "ooxml-swift-inspector"

    /// How far a package may expand, and how large one part may be read into
    /// memory (PsychQuant/ooxml-swift#130). A 0.56 MB crafted package was
    /// measured expanding to 1.7 GB RSS over 39 s — a ~3000:1 amplification —
    /// against a library with no bound on either axis:
    ///
    ///   disk    the bytes an extraction writes out
    ///   memory  one part read into `Data` before it is parsed
    ///
    /// The corpus measurement behind these is small (738 real documents):
    /// largest single entry 17.5 MB, largest package 26.4 MB in total, highest
    /// compression ratio 117.8x (p99 17.9x, p99.9 32.2x). `maximumCompressionRatio`
    /// tracks that measurement directly — 500x sits 4.24x above the observed
    /// maximum, and this is the axis that actually stops the classic zip-bomb
    /// shape (a tiny file, extreme amplification): a real document's DEFLATE
    /// ratio has no reason to approach 500x regardless of how large real
    /// documents get, so tightening it further buys safety margin on
    /// amplification without touching the size axes below.
    ///
    /// `maximumEntryBytes` / `maximumTotalBytes` / `maximumPartBytes` (R2
    /// correction, 2026-09-28) do NOT track the corpus measurement the way the
    /// ratio does — 738 real documents are not 738 *large* real documents, and
    /// a corpus of small samples cannot bound how large a legitimate one gets.
    /// The concrete case this repo's own review named: a document or report with
    /// ~200 embedded 300-DPI A4 scans. A 300-DPI A4 page is ~2481×3508 px;
    /// stored as an unoptimized 24-bit bitmap that is ~26 MB per image and
    /// ~5.2 GB for 200 of them, but Word documents in practice carry compressed
    /// JPEG/PNG scans, not raw bitmaps — a realistic (if generous) per-image
    /// estimate for a high-quality 300-DPI scan is 5–10 MB, i.e. roughly 1–2 GB
    /// for 200 images. `maximumEntryBytes` covers the single-large-image case
    /// with room for the unoptimized-bitmap outlier (a single such image, not
    /// 200 of them, is the realistic single-entry ceiling); `maximumTotalBytes`
    /// covers the many-realistic-images case with about 2–4x headroom over
    /// that 1–2 GB estimate. Both remain hard, real bounds — nowhere near
    /// "effectively unlimited" — because the amplification defense lives on
    /// the ratio axis, not here: a package that is actually this large has to
    /// really contain that many actual bytes, which is a fixed cost to
    /// whoever sends it, not a multiplied one.
    ///
    /// Every default here is overridable per call via `limits:` — a caller
    /// with a genuinely larger corpus, or a workload the defaults are too
    /// tight for, does not need to fork this library to raise them.
    ///
    /// `maximumEntryCount` has no corpus measurement behind it either (real
    /// packages run to at most a few hundred parts); it exists for
    /// PsychQuant/ooxml-swift#107, where many tiny entries — each individually
    /// under every byte limit — could still exhaust memory via per-entry
    /// dictionary and string overhead in `RawPartChannel.readAllParts`.
    public struct Limits: Sendable, Equatable {
        public var maximumEntryBytes: Int64
        public var maximumTotalBytes: Int64
        public var maximumCompressionRatio: Double
        public var maximumPartBytes: Int64
        public var maximumEntryCount: Int
        public init(maximumEntryBytes: Int64 = 512 * 1024 * 1024,
                    maximumTotalBytes: Int64 = 4096 * 1024 * 1024,
                    maximumCompressionRatio: Double = 500,
                    maximumPartBytes: Int64 = 512 * 1024 * 1024,
                    maximumEntryCount: Int = 50_000) {
            self.maximumEntryBytes = maximumEntryBytes
            self.maximumTotalBytes = maximumTotalBytes
            self.maximumCompressionRatio = maximumCompressionRatio
            self.maximumPartBytes = maximumPartBytes
            self.maximumEntryCount = maximumEntryCount
        }
    }

    /// The limits used when a caller does not name its own. An immutable
    /// value passed to each call — never a shared mutable global: two
    /// concurrent extractions would otherwise be able to capture each other's
    /// scoped override, and a single document operation could run its
    /// extraction and its part reads under two different policies. A caller
    /// with a genuinely larger corpus passes its own `Limits` instead of
    /// discovering the ceiling as an unexplained refusal.
    public static let defaultLimits = Limits()

    static func describeBytes(_ n: Int64) -> String {
        n >= 1_048_576 ? String(format: "%.1f MB", Double(n) / 1_048_576)
                       : String(format: "%.1f KB", Double(n) / 1024)
    }

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
    public static func unzip(data: Data, limits: Limits = defaultLimits) throws -> URL {
        try unzip(data: data, namespace: readerNamespace, limits: limits)
    }

    /// The same, into a named namespace (one path component) under the
    /// temporary directory. Refuses the same five entry kinds.
    static func unzip(data: Data, namespace: String, limits: Limits = defaultLimits) throws -> URL {
        precondition(!namespace.isEmpty && !namespace.contains("/") && namespace != "." && namespace != "..", "namespace must be one path component")
        let archive = try Archive(data: data, accessMode: .read)
        var declaredTotal: Int64 = 0
        var entryCount = 0
        for entry in archive {
            if entry.type == .symlink {
                throw WordError.invalidDocx("the package contains a symbolic-link entry (\(displayName(entry.path))); refusing to extract it.")
            }
            let path = entry.path
            if path.isEmpty || path.contains("\0") {
                throw WordError.invalidDocx("the package contains an entry with an empty or NUL-containing path; refusing to extract it.")
            }
            let components = path.split(separator: "/", omittingEmptySubsequences: false)
            if path.hasPrefix("/") || components.contains("..") {
                throw WordError.invalidDocx("the package contains an entry whose path leaves its own directory (\(displayName(path))); refusing to extract it.")
            }
            // PsychQuant/ooxml-swift#145: the checks above are a hand-rolled
            // subset of `PathValidator.isSafeRelativeOOXMLPath` — the fuller
            // definition already used at the parse boundary for relationship
            // targets — and were the looser of the two: percent-encoded
            // traversal (`%2e%2e`, `%2f`, `%5c`), backslash-separated `..`,
            // control characters, and paths over 256 characters passed here
            // while `isSafeRelativeOOXMLPath` would refuse them. Neither gap
            // yields an escape on macOS today (ZIPFoundation does not decode
            // `%2e` or treat `\` as a path separator when it extracts), so
            // this is hardening rather than closing a hole — but one
            // definition of "safe relative path" replaces two, and it is now
            // the stricter one everywhere.
            guard isSafeRelativeOOXMLPath(path) else {
                throw WordError.invalidDocx("the package contains an entry whose path is not a safe relative OPC path (\(displayName(path))); refusing to extract it.")
            }
            entryCount += 1
            if entryCount > limits.maximumEntryCount {
                throw WordError.invalidDocx("the package declares more than \(limits.maximumEntryCount) entries; refusing to extract it.")
            }
            // Size policy, from the archive's own declarations. The central
            // directory can lie, so the extracted tree is measured again below
            // (#130): declared and actual are each refused on their own.
            // `clamping:`, not a plain `Int64(_:)` — these are UInt64 values
            // taken straight from the central directory, and a ZIP64 header
            // can declare up to `UInt64.max`, which the plain initializer
            // TRAPS on ("Not enough bits to represent the passed value").
            // Saturating instead lands an absurd declaration in the refusal
            // below, where it belongs.
            let declared = Int64(clamping: entry.uncompressedSize)
            if declared > limits.maximumEntryBytes {
                throw WordError.invalidDocx("the package declares an entry of \(describeBytes(declared)) (\(displayName(path))), over the \(describeBytes(limits.maximumEntryBytes)) limit; refusing to extract it.")
            }
            // `addingReportingOverflow`, not `+=`: Swift traps on integer
            // overflow, and a caller expressing "no limit" as
            // `maximumTotalBytes: .max` (the only spelling `Limits` offers)
            // lets two entries near `Int64.max` reach the addition. An
            // overflow IS "over the limit", so it is refused, not crashed on.
            let (sum, overflowed) = declaredTotal.addingReportingOverflow(declared)
            declaredTotal = overflowed ? Int64.max : sum
            if overflowed || declaredTotal > limits.maximumTotalBytes {
                throw WordError.invalidDocx("the package declares more than \(describeBytes(limits.maximumTotalBytes)) of content in total; refusing to extract it.")
            }
            let compressed = Int64(clamping: entry.compressedSize)
            if compressed > 0, Double(declared) / Double(compressed) > limits.maximumCompressionRatio {
                throw WordError.invalidDocx("the package declares an entry that expands \(Int(Double(declared) / Double(compressed)))x (\(displayName(path))), over the \(Int(limits.maximumCompressionRatio))x limit; refusing to extract it.")
            }
        }
        let fm = FileManager.default
        // Threat model (verify R8 security, measured): the descriptor checks
        // below defeat anything planted at the namespace name before or while
        // we create it, and anything a DIFFERENT non-root user could do —
        // PROVIDED the parent temporary directory is per-user 0700 (macOS's
        // `NSTemporaryDirectory()`, which ignores `TMPDIR` — measured) or
        // sticky: such a user cannot rename or replace our 0700 directory
        // there. Root is outside every model here. They do not defeat a
        // process running as the SAME user: extraction and every later read are path-based
        // (ZIPFoundation and the reader take paths; macOS does not resolve
        // `/dev/fd/N/child`, so there is no descriptor-relative extraction),
        // and a same-uid process can rename our directory away at any time —
        // it can also read the document directly, so that is outside the
        // model, not a gap in it.
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
            throw WordError.invalidDocx("could not create the extraction directory under the `\(namespace)` namespace (\(errnoText()))")
        }
        let tempFD = try ownerOnlyDirectoryDescriptor(opening: uuid, relativeTo: namespaceFD)
        defer { close(tempFD) }

        // A UUID name cannot collide with any entry the archive declares; the
        // copy is created 0600 atomically (O_CREAT|O_EXCL with that mode — the
        // inode never exists with another mode, codex R7-6) and written through
        // the descriptor, then extracted from the path it now has.
        let copyName = UUID().uuidString + ".zip"
        let copyFD = openat(tempFD, copyName, O_WRONLY | O_CREAT | O_EXCL | O_NOFOLLOW | O_CLOEXEC, 0o600)
        guard copyFD >= 0 else {
            throw WordError.invalidDocx("could not create the package's private copy (\(errnoText()))")
        }
        let copyHandle = FileHandle(fileDescriptor: copyFD, closeOnDealloc: true)
        do { try copyHandle.write(contentsOf: data); try copyHandle.close() }
        catch { throw WordError.invalidDocx("could not write the package's private copy (\(describeWithoutPaths(error)))") }
        let privateCopy = tempDir.appendingPathComponent(copyName)
        // Every error from here on is ours to name (verify R7 logic N-L2-R7: a
        // raw POSIX error carried the temporary path to the caller).
        //
        // Extraction is bounded AS IT RUNS, not only afterwards
        // (PsychQuant/ooxml-swift#157). The declared sizes were already
        // refused above; a central directory that UNDERSTATES what it writes
        // gets past that check, and without this, its bytes land on disk in
        // full before the post-extraction walk below ever measures them.
        // `ExtractionBudget` counts the bytes the decompressor actually
        // produces and cancels the moment the running total crosses the
        // limit — the check sits before each chunk reaches the consumer, so
        // the overshoot is at most one 16 KiB chunk, not the whole entry.
        let budget = ExtractionBudget(limit: limits.maximumTotalBytes)
        do { try fm.unzipItem(at: privateCopy, to: tempDir, progress: budget) }
        catch let error as Archive.ArchiveError where error == .cancelledOperation && budget.exceeded {
            throw WordError.invalidDocx("the package wrote more than \(describeBytes(limits.maximumTotalBytes)) while being extracted, whatever its central directory declared; refusing it.")
        }
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
        guard fchmod(tempFD, 0o700) == 0 else { throw WordError.invalidDocx("could not make the extraction directory owner-only (\(errnoText()))") }
        var walkError: Error?
        guard let walker = fm.enumerator(at: tempDir, includingPropertiesForKeys: [.isDirectoryKey], options: [], errorHandler: { _, error in walkError = error; return false }) else {
            throw WordError.invalidDocx("could not enumerate the extracted package")
        }
        let rootPrefixes = rootPrefixes(of: tempDir)                                // resolved once, not per item (verify R10 security N-S10-4)
        var actualTotal: Int64 = 0                                                 // the central directory can lie; measure what is on disk (#130) — defence in depth alongside the streaming budget above
        for case let item as URL in walker {
            let relative = displayName(relativeName(of: item, rootPrefixes: rootPrefixes))   // an entry name is attacker-controlled text
            // `lstat` + `fchmodat(AT_SYMLINK_NOFOLLOW)`: nothing here follows a
            // link (verify R8: `chmod` through a link reached outside the tree).
            // The archive cannot contain a link entry (refused before anything
            // is written); one that appears anyway is refused, not chmod-ed.
            var st = stat()
            guard lstat(item.path, &st) == 0 else {
                throw WordError.invalidDocx("could not stat \(relative) in the package (\(errnoText())); refusing a partial permission reset.")
            }
            let mode: mode_t
            switch st.st_mode & S_IFMT {
            case S_IFDIR: mode = 0o700
            case S_IFREG: mode = 0o600
            default: throw WordError.invalidDocx("\(relative) in the package is neither a file nor a directory (a link or special file appeared during extraction); refusing it.")
            }
            guard fchmodat(AT_FDCWD, item.path, mode, AT_SYMLINK_NOFOLLOW) == 0 else {
                throw WordError.invalidDocx("could not make \(relative) in the package owner-only (\(errnoText()))")
            }
            // What the archive DECLARED was checked before extraction; this is
            // what it actually wrote. The streaming budget above already stops
            // an understating archive mid-write; this is the second line —
            // measuring the finished tree — for whatever it lets through.
            if mode == 0o600 {
                let size = Int64(clamping: st.st_size)
                if size > limits.maximumEntryBytes {
                    throw WordError.invalidDocx("the package expanded \(relative) to \(describeBytes(size)), over the \(describeBytes(limits.maximumEntryBytes)) limit; refusing it.")
                }
                let (sum, overflowed) = actualTotal.addingReportingOverflow(size)
                actualTotal = overflowed ? Int64.max : sum
                if overflowed || actualTotal > limits.maximumTotalBytes {
                    throw WordError.invalidDocx("the package expanded to more than \(describeBytes(limits.maximumTotalBytes)) on disk; refusing it.")
                }
            }
        }
        if let walkError { throw WordError.invalidDocx("could not enumerate the extracted package (\(describeWithoutPaths(walkError)))") }
        guard fchmod(tempFD, 0o700) == 0 else { throw WordError.invalidDocx("could not make the extraction directory owner-only (\(errnoText()))") }
        succeeded = true
        return tempDir
    }

    /// Make every directory under `root` ours again (0700, pre-order so the
    /// walk can descend) and remove the tree. Best effort by construction —
    /// the failure path cannot throw — so a tree that still cannot be removed
    /// is reported on stderr with its path, the one place a path is useful.
    static func removeTreeForcibly(_ root: URL) {
        let fm = FileManager.default
        var rootStat = stat()
        guard lstat(root.path, &rootStat) == 0 else { return }       // never existed (creation failed) or already gone — nothing to report
        fchmodat(AT_FDCWD, root.path, 0o700, AT_SYMLINK_NOFOLLOW)
        if let walker = fm.enumerator(at: root, includingPropertiesForKeys: [.isDirectoryKey], options: [], errorHandler: { _, _ in true }) {
            for case let item as URL in walker {
                var st = stat()
                guard lstat(item.path, &st) == 0 else { continue }
                switch st.st_mode & S_IFMT {
                case S_IFDIR: fchmodat(AT_FDCWD, item.path, 0o700, AT_SYMLINK_NOFOLLOW)
                case S_IFREG: fchmodat(AT_FDCWD, item.path, 0o600, AT_SYMLINK_NOFOLLOW)
                default: break                                        // a link is removed with the tree, never followed (verify R8)
                }
            }
        }
        do { try fm.removeItem(at: root) }
        catch {
            // Already gone (a same-uid process renamed it away between our lstat
            // and this call — verify R8 security N-S8-3) is not a leftover.
            let ns = error as NSError
            let gone = (ns.domain == NSPOSIXErrorDomain && ns.code == Int(ENOENT)) || (ns.domain == NSCocoaErrorDomain && (ns.code == 4 || ns.code == 260))
                || ((ns.userInfo[NSUnderlyingErrorKey] as? NSError).map { $0.domain == NSPOSIXErrorDomain && $0.code == Int(ENOENT) } ?? false)
            if !gone { FileHandle.standardError.write(Data("ooxml-swift: could not remove the partial extraction at \(root.path) after a failure (\(describeWithoutPaths(error)))\n".utf8)) }
        }
    }

    /// Attacker-controlled text (an archive entry name) as it may appear in an
    /// error message: every control character escaped, at most 120 scalars of
    /// the name shown, and the rendered text capped at 480 characters (an
    /// escape expands a scalar up to six-fold — verify R9 logic NEW-R9-4).
    /// (Verify R8 security N-S8-1: a newline in an entry name forged a second
    /// line of output; ANSI sequences and 9 000-character names went straight
    /// into the message the consumer renders.)
    /// Whether `scalar` is shown as `\u{…}` instead of itself: a Unicode
    /// property, not a hand-written list — general category Cc (control), Cf
    /// (format), Zl / Zp (line / paragraph separator), or any
    /// Default_Ignorable_Code_Point. The R9 list (U+2028/2029, U+200B–200F,
    /// U+202A–202E, U+2066–2069, U+FEFF) missed U+00AD, U+061C, U+180E, U+FFF9,
    /// U+2060, U+3164, U+FE0F, U+E0001 … (verify R10 requirements N-R10-2,
    /// security N-S10-3): a criterion the reader can check against the UCD
    /// cannot drift from the list in someone's head.
    static func isEscapedInDisplay(_ scalar: Unicode.Scalar) -> Bool {
        switch scalar.properties.generalCategory {
        case .control, .format, .lineSeparator, .paragraphSeparator: return true
        default: return scalar.properties.isDefaultIgnorableCodePoint
        }
    }

    static func displayName(_ raw: String) -> String {
        var out = ""
        for scalar in raw.unicodeScalars.prefix(120) {
            switch scalar {
            case "\n": out += "\\n"
            case "\t": out += "\\t"
            case "\r": out += "\\r"
            // The escape character and the delimiter the messages wrap this text
            // in are themselves attacker-supplied when they appear in a name
            // (verify R11 security N-S11-1: an entry literally named `a\u{202E}b`
            // rendered identically to one containing a real U+202E, so the
            // rendering could not be read back; N-S11-3: an id containing a
            // backtick closed the code span and let prose escape it).
            case "\\": out += "\\\\"
            case "`": out += "\\u{60}"
            case _ where isEscapedInDisplay(scalar): out += String(format: "\\u{%02X}", scalar.value)
            default: out.unicodeScalars.append(scalar)
            }
        }
        if raw.unicodeScalars.count > 120 { out += "…(\(raw.unicodeScalars.count - 120) more characters)" }
        return out.count > 480 ? String(out.prefix(480)) + "…" : out
    }

    /// `item`'s path under `root`, both taken through `resolvingSymlinksInPath()`
    /// — the enumerator hands back `/private/var/…` while `temporaryDirectory`
    /// is `/var/…`, and Foundation's resolver maps BOTH to the `/var/…` form
    /// (verify R8 logic NEW-L1 named a non-existent item; R9 NEW-R9-2 found the
    /// fix comparing an unresolved item against resolved roots, never matching).
    static func relativeName(of item: URL, under root: URL) -> String {
        relativeName(of: item, rootPrefixes: rootPrefixes(of: root))
    }

    /// The spellings of `root` an enumerated item can start with, deduplicated
    /// — on the default macOS temporary directory the resolved and unresolved
    /// forms are the SAME string, and comparing both meant every item paid the
    /// identical failing comparison twice (verify R11 logic N-L11-2).
    ///
    /// Computed once per walk: resolving the root for every item made a
    /// 32 000-entry walk ~1.6× slower (verify R10 security N-S10-4). That
    /// hoisting is where the win came from — 183 → 15.6 ms per 20 000 items.
    static func rootPrefixes(of root: URL) -> [String] {
        let resolved = root.resolvingSymlinksInPath().path + "/", raw = root.path + "/"
        return resolved == raw ? [resolved] : [resolved, raw]
    }

    /// `item`'s name under a root whose prefixes are already computed.
    ///
    /// The unresolved path is tried first, but do not read that as a fast path
    /// that usually hits: on the default `TMPDIR` the enumerator hands back
    /// `/private/var/…` while the root resolves to `/var/…`, so it hits ZERO
    /// times in a real walk (verify R11 logic N-L11-2, measured over 55 items)
    /// and every item falls through to the resolver. It is kept because it IS
    /// the cheap answer wherever the two forms agree, and it costs one string
    /// comparison where they do not.
    static func relativeName(of item: URL, rootPrefixes: [String]) -> String {
        let unresolved = item.path
        for prefix in rootPrefixes where unresolved.hasPrefix(prefix) { return String(unresolved.dropFirst(prefix.count)) }
        let resolved = item.resolvingSymlinksInPath().path
        for prefix in rootPrefixes where resolved.hasPrefix(prefix) { return String(resolved.dropFirst(prefix.count)) }
        return item.lastPathComponent
    }

    /// Read one part of an extracted package, refusing one too large to hold.
    ///
    /// The extraction limits already cap any single entry, so a part that came
    /// through `unzip` cannot exceed `limits.maximumEntryBytes` — this is the
    /// second line, for a part reached by a path the caller supplied directly,
    /// and it is the bound on the MEMORY axis of #130: the original PoC's
    /// 1.7 GB RSS came from reading one expanded part into `Data` before
    /// parsing it, not from the extraction itself. Refusing before the read
    /// means the process never holds it.
    ///
    /// Boundary: when `lstat` fails, or reports something other than a
    /// regular file, this reads on rather than refusing — the caller's own
    /// existence/identity check is what decides whether to call this at all.
    static func readPart(at url: URL, describedAs name: String, limits: Limits = defaultLimits) throws -> Data {
        var st = stat()
        if lstat(url.path, &st) == 0, st.st_mode & S_IFMT == S_IFREG, Int64(clamping: st.st_size) > limits.maximumPartBytes {
            throw WordError.invalidDocx("\(displayName(name)) is \(describeBytes(Int64(clamping: st.st_size))), over the \(describeBytes(limits.maximumPartBytes)) limit for a single part; refusing to read it.")
        }
        return try Data(contentsOf: url)
    }

    /// `strerror(errno)` for the calling thread's last error.
    static func errnoText() -> String { String(cString: strerror(errno)) }

    /// A description built from the error's code, never from its text: an
    /// error's text can carry a file-system path, and the temporary
    /// directory is not the caller's business (#146; verify R8: a token
    /// filter kept quoted paths and dropped words that merely began with `/`).
    static func describeWithoutPaths(_ error: Error) -> String {
        let ns = error as NSError
        if ns.domain == NSPOSIXErrorDomain { return String(cString: strerror(Int32(ns.code))) }
        if let underlying = ns.userInfo[NSUnderlyingErrorKey] as? NSError, underlying.domain == NSPOSIXErrorDomain {
            return String(cString: strerror(Int32(underlying.code)))
        }
        if ns.domain == NSCocoaErrorDomain {
            switch ns.code {
            case 4, 260: return "no such file"
            case 513: return "permission denied"
            case 516: return "file exists"
            case 640: return "no space left"
            default: return "file error \(ns.code)"
            }
        }
        return "\(ns.domain) error \(ns.code)"
    }

    /// `mkdir(path, 0700)` — atomic create or `EEXIST` — then open the directory
    /// without following a link and verify the DESCRIPTOR: a directory, owned
    /// by this uid, mode 0700 (any other mode of our own directory is reset;
    /// anyone else's, or a file or link planted at the name, is refused).
    private static func ownerOnlyDirectoryDescriptor(creatingIfAbsent path: String) throws -> Int32 {
        let name = "the `" + (path as NSString).lastPathComponent + "` extraction namespace"
        if mkdir(path, 0o700) != 0, errno != EEXIST {
            throw WordError.invalidDocx("could not create \(name) (\(errnoText()))")
        }
        let fd = open(path, O_RDONLY | O_DIRECTORY | O_NOFOLLOW | O_CLOEXEC)
        guard fd >= 0 else {
            throw WordError.invalidDocx("\(name) cannot be opened as a real directory (\(errnoText())); a symbolic link or a file planted there is refused.")
        }
        do { try verifyOwnerOnlyDirectory(fd, describedAs: name) } catch { close(fd); throw error }
        return fd
    }

    private static func ownerOnlyDirectoryDescriptor(opening name: String, relativeTo parentFD: Int32) throws -> Int32 {
        let fd = openat(parentFD, name, O_RDONLY | O_DIRECTORY | O_NOFOLLOW | O_CLOEXEC)
        guard fd >= 0 else { throw WordError.invalidDocx("could not open the extraction directory just created (\(errnoText()))") }
        do { try verifyOwnerOnlyDirectory(fd, describedAs: "the extraction directory") } catch { close(fd); throw error }
        return fd
    }

    private static func verifyOwnerOnlyDirectory(_ fd: Int32, describedAs name: String) throws {
        var st = stat()
        guard fstat(fd, &st) == 0 else { throw WordError.invalidDocx("could not stat \(name) (\(errnoText()))") }
        guard (st.st_mode & S_IFMT) == S_IFDIR else { throw WordError.invalidDocx("\(name) exists but is not a directory; refusing to extract.") }
        guard st.st_uid == getuid() else { throw WordError.invalidDocx("\(name) is owned by another user (uid \(st.st_uid)); refusing to extract.") }
        if (st.st_mode & 0o7777) != 0o700 {                                  // created by an older version, or by hand
            guard fchmod(fd, 0o700) == 0 else { throw WordError.invalidDocx("could not make \(name) owner-only (\(errnoText()))") }
        }
    }

    /// 壓縮目錄內容為 ZIP 檔案（不包含目錄本身的路徑）
    ///
    /// Type-aware on purpose (PsychQuant/ooxml-swift#110). `fileExists(atPath:)`
    /// alone answers "is something here", not "is it a file this call may
    /// replace" — and `destination` naming an existing DIRECTORY previously
    /// fell into that same branch, so `removeItem(at:)` deleted the whole tree
    /// before writing the zip in its place. `destination` names a file to
    /// produce, never a tree to clear, so a directory there is refused
    /// regardless of what the tree contains.
    public static func zip(_ directory: URL, to destination: URL) throws {
        let data = try zipToData(directory)

        var isDirectory: ObjCBool = false
        let exists = FileManager.default.fileExists(atPath: destination.path, isDirectory: &isDirectory)
        if exists, isDirectory.boolValue {
            throw WordError.zipError("the destination `\(destination.lastPathComponent)` is an existing directory, not a file to produce; refusing to delete it.")
        }
        if exists {
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

/// Counts what an extraction actually writes, and stops it when that crosses a
/// limit (PsychQuant/ooxml-swift#157).
///
/// ZIPFoundation reports progress per entry: `unzipItem` builds its own total
/// from the archive's DECLARED sizes and aggregates per-entry children with
/// `addChild`. Observing the parent is therefore useless for the case that
/// matters — an archive that understates its entries makes each child saturate
/// against its own declared total, so the parent's count stops rising exactly
/// when the overflow happens (measured: a child at 4 194 304 actual bytes
/// against a declared 1 024, with the parent stuck at 1 024 and its
/// `fractionCompleted` pinned to 1.0).
///
/// The CHILD's `completedUnitCount` is the real figure and is not clamped, so
/// this intercepts the children as they are attached and observes each one.
/// Cancellation then propagates parent → child, and ZIPFoundation's
/// `readCompressed` checks `isCancelled` BEFORE handing each chunk to the
/// consumer — so the file on disk overshoots by at most one
/// `defaultReadChunkSize` (16 KiB).
///
/// **The interception rides on an implementation detail**: that ZIPFoundation
/// attaches its per-entry progress with `addChild(_:withPendingUnitCount:)`.
/// Foundation has another way to attach a child —
/// `Progress(totalUnitCount:parent:pendingUnitCount:)` — which does NOT call
/// it. If a future version switches, nothing is captured, the total stays
/// zero, and the limit disappears with no error at all — the same
/// silent-failure shape as the `Int64` trap this feature's first round
/// shipped. `testTheExtractionBudgetActuallySeesTheEntries` asserts the
/// captured count matches the archive's file entries, so that upgrade turns a
/// test red instead of turning the bound off.
final class ExtractionBudget: Progress, @unchecked Sendable {
    private let limit: Int64
    private var observations: [NSKeyValueObservation] = []
    private var children: [Progress] = []
    private var totals: [ObjectIdentifier: Int64] = [:]
    private(set) var exceeded = false
    /// How many per-entry children were seen — the mechanism check above.
    private(set) var capturedChildren = 0

    init(limit: Int64) {
        self.limit = limit
        super.init(parent: nil, userInfo: nil)
    }

    override func addChild(_ child: Progress, withPendingUnitCount inUnitCount: Int64) {
        capturedChildren += 1
        children.append(child)
        // A directory entry is reported by ASSIGNING 1, not by adding bytes, so
        // it contributes at most one unit each — below any limit worth setting.
        observations.append(child.observe(\.completedUnitCount, options: [.new]) { [weak self] observed, _ in
            guard let self else { return }
            self.totals[ObjectIdentifier(observed)] = observed.completedUnitCount
            let written = self.totals.values.reduce(0, +)
            if written > self.limit, !self.exceeded {
                self.exceeded = true
                // Cancel each entry's own progress, not just this parent.
                // `Progress.cancel()` on the parent does NOT set `isCancelled`
                // on a child attached through the overridden `addChild` —
                // measured `child.isCancelled == false` immediately after —
                // and ZIPFoundation checks the CHILD. Cancelling them directly
                // is what actually stops the next chunk.
                self.cancel()
                for child in self.children { child.cancel() }
            }
        })
        super.addChild(child, withPendingUnitCount: inUnitCount)
    }
}
