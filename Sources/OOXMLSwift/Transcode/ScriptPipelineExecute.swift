import Foundation

/// Result of executing a `.mdocx.swift` rebuild script.
///
/// Both optional fields follow one rule: **absence means the thing did not
/// happen**, never "happened and the answer was no".
///
/// - `verified` is nil when verification did not run, which is NOT the same
///   as "ran and failed". A caller that only inspects `brokenParts` must not
///   be able to read an unverified run as a clean one.
/// - `written` is nil when nothing was published. Naming the requested output
///   path unconditionally would be a positive false claim — worse than a
///   missing signal — because that path may still hold its original bytes.
public struct ScriptExecuteResult: Sendable {
    /// Path the rebuilt docx was published to; nil when nothing was published.
    public let written: String?
    /// Stage-B verdict when verification was requested; nil when it was not.
    public let verified: Bool?
    /// Part paths that failed byte equality (empty unless verified == false).
    public let brokenParts: [String]

    public init(written: String?, verified: Bool?, brokenParts: [String]) {
        self.written = written
        self.verified = verified
        self.brokenParts = brokenParts
    }
}

public enum ScriptPipelineError: LocalizedError {
    case fileNotFound(String)
    case outputExists(String)
    case outputIsDirectory(String)

    public var errorDescription: String? {
        switch self {
        case .fileNotFound(let path):
            return "找不到輸入檔案: \(path)"
        case .outputExists(let path):
            // Deliberately carries no advice about HOW to permit overwriting:
            // the CLI would say `--force` and the MCP tool would say the
            // `overwrite` argument. Face-specific wording belongs on the face.
            return "輸出檔案已存在: \(path)"
        case .outputIsDirectory(let path):
            // Says DIRECTORY, and offers no overwrite advice on purpose. The
            // old message called this an existing 檔案, which steered the
            // operator into passing the very flag that destroyed the tree.
            return "輸出路徑是一個目錄，不是檔案: \(path)"
        }
    }
}

/// `.mdocx.swift` script → rebuilt docx. When `verifyAgainst` names a
/// reference docx, the rebuilt XML part set is compared for Stage-B byte
/// equality and the verdict (with broken part paths) rides the result.
///
/// This is the single execution entry point for the script pipeline: both
/// the `macdoc word render` CLI command and che-word-mcp's `execute_script`
/// tool call it rather than reimplementing the orchestration. Keeping one
/// implementation is what makes the two faces agree by construction instead
/// of by convention — and that is why the overwrite gate below lives HERE
/// rather than on either wrapper. A guard bolted onto one face is, by
/// definition, outside the parity the shared implementation buys; that is
/// exactly how the MCP face came to have no overwrite protection at all.
///
/// Ordering contract: the reference is read and pinned in memory BEFORE the
/// output is written. This (a) makes `outputPath == verifyAgainst` compare
/// against the PRE-write reference bytes instead of the run's own output
/// (which would be a guaranteed false-positive `verified: true`), and
/// (b) surfaces a mistyped reference path before any write side effect.
/// The staging write below independently protects (a) as well; both are kept
/// so the invariant does not rest on a single mechanism.
///
/// Publication contract: the rebuilt package is written to a staging path in
/// the output's own directory and moved onto `outputPath` only once the
/// verdict passes, or when no verification was asked for. A failing verdict
/// therefore leaves `outputPath` exactly as it was found — holding its former
/// contents, or holding nothing. Same-directory staging keeps that final move
/// atomic; staging under the system temporary directory would frequently
/// cross a filesystem boundary and degrade the move into a copy.
///
/// - Parameter overwrite: permits replacing an existing file at `outputPath`.
///   Defaults to refusing, so a caller who forgets it is protected rather
///   than destructive. Note that supplying one path as both `outputPath` and
///   `verifyAgainst` necessarily targets an existing file and therefore
///   requires this to be set.
public func scriptPipelineExecute(
    scriptPath: String,
    outputPath: String,
    verifyAgainst: String? = nil,
    overwrite: Bool = false
) throws -> ScriptExecuteResult {
    let fm = FileManager.default
    let scriptURL = URL(fileURLWithPath: scriptPath)
    let outputURL = URL(fileURLWithPath: outputPath)

    // --- Guards, cheapest first. Nothing below here has a side effect until
    // the staging write, and none of these guards costs a parse or a replay.
    guard fm.fileExists(atPath: scriptURL.path) else {
        throw ScriptPipelineError.fileNotFound(scriptPath)
    }
    if let referencePath = verifyAgainst,
       !fm.fileExists(atPath: URL(fileURLWithPath: referencePath).path) {
        throw ScriptPipelineError.fileNotFound(referencePath)
    }
    // Type-aware on purpose. `fileExists(atPath:)` alone answers "is something
    // here", and both this gate and the publish below then treat that as "a
    // file is here" — so a directory was refused with a message calling it a
    // 檔案, and `overwrite` went on to replace the whole tree with the rebuilt
    // docx, reporting success. Measured against the 0.6.0 release too, so this
    // long predates the staging rework. The two-argument form is already the
    // idiom elsewhere in this package (see ZipHelper).
    var outputIsDirectory: ObjCBool = false
    let outputPresent = fm.fileExists(atPath: outputURL.path, isDirectory: &outputIsDirectory)
    if outputPresent, outputIsDirectory.boolValue {
        // Refused regardless of `overwrite`: that flag means "replace the
        // existing FILE", and no reading of it authorises deleting a directory.
        throw ScriptPipelineError.outputIsDirectory(outputPath)
    }
    guard overwrite || !outputPresent else {
        throw ScriptPipelineError.outputExists(outputPath)
    }

    let source = try String(contentsOf: scriptURL, encoding: .utf8)
    let log = try ScriptImporter.parse(source: source)

    // Pin the reference BEFORE any write (see ordering contract above).
    var reference: [String: Data]?
    if let referencePath = verifyAgainst {
        reference = try RawPartChannel.readAllParts(
            from: URL(fileURLWithPath: referencePath))
    }

    var document = WordDocument.emptyAuthoringDocument()
    try document.apply(operations: log.entries.map(\.op))

    // Staging lives beside the output so the publish below is a rename, not
    // a cross-device copy. Removed on every path — including the failing
    // verdict, which is the whole point of staging in the first place.
    let stagingURL = outputURL
        .deletingLastPathComponent()
        .appendingPathComponent(
            ".\(outputURL.lastPathComponent).mdocx-staging-\(UUID().uuidString)")
    defer { try? fm.removeItem(at: stagingURL) }
    try document.writeAuthoringPackage(to: stagingURL)

    // The branch is on `overwrite`, NOT on whether the output currently
    // exists. Branching on existence re-opens the gate at write time: if the
    // gate passed because nothing was there, and a file appeared during the
    // parse/replay/stage window, an existence-branch would take the
    // `replaceItemAt` path — which succeeds unconditionally on a plain file —
    // and silently destroy it despite the caller having asked us not to
    // overwrite anything. Measured: replaceItemAt replaced a concurrently
    // created file without error. Keying on `overwrite` means the refusing
    // caller always takes `moveItem`, which fails when the destination exists,
    // so the concurrent file survives.
    func publish() throws {
        // Re-checked at write time: the gate above runs before the replay, so a
        // directory created during that window would otherwise reach
        // replaceItemAt — which accepts a directory destination and succeeds.
        var isDirNow: ObjCBool = false
        if fm.fileExists(atPath: outputURL.path, isDirectory: &isDirNow), isDirNow.boolValue {
            throw ScriptPipelineError.outputIsDirectory(outputPath)
        }
        if overwrite, fm.fileExists(atPath: outputURL.path) {
            _ = try fm.replaceItemAt(outputURL, withItemAt: stagingURL,
                                     backupItemName: nil, options: [])
        } else {
            try fm.moveItem(at: stagingURL, to: outputURL)
        }
    }

    guard let reference else {
        try publish()
        return ScriptExecuteResult(written: outputPath, verified: nil, brokenParts: [])
    }

    let rebuilt = try RawPartChannel.readAllParts(from: stagingURL)
    let broken = PartFidelity.compareParts(reference: reference, rebuilt: rebuilt)
        .filter { !$0.isEqual }
        .map(\.partPath)
        .sorted()

    guard broken.isEmpty else {
        // Publish nothing. `written` is nil rather than the requested path,
        // because that path was not written.
        return ScriptExecuteResult(written: nil, verified: false, brokenParts: broken)
    }
    try publish()
    return ScriptExecuteResult(written: outputPath, verified: true, brokenParts: [])
}
