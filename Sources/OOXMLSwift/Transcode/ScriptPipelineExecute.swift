import Foundation

/// Result of executing a `.mdocx.swift` rebuild script.
///
/// `verified` is deliberately optional rather than defaulting to `false`:
/// absence means verification did not run, which is NOT the same as "ran
/// and failed". A caller that only inspects `brokenParts` must not be able
/// to read an unverified run as a clean one.
public struct ScriptExecuteResult: Sendable {
    /// Path the rebuilt docx was written to.
    public let written: String
    /// Stage-B verdict when verification was requested; nil when it was not.
    public let verified: Bool?
    /// Part paths that failed byte equality (empty unless verified == false).
    public let brokenParts: [String]

    public init(written: String, verified: Bool?, brokenParts: [String]) {
        self.written = written
        self.verified = verified
        self.brokenParts = brokenParts
    }
}

public enum ScriptPipelineError: LocalizedError {
    case fileNotFound(String)

    public var errorDescription: String? {
        switch self {
        case .fileNotFound(let path):
            return "找不到輸入檔案: \(path)"
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
/// of by convention.
///
/// Ordering contract: the reference is read and pinned in memory BEFORE the
/// output is written. This (a) makes `outputPath == verifyAgainst` compare
/// against the PRE-write reference bytes instead of the run's own output
/// (which would be a guaranteed false-positive `verified: true`), and
/// (b) surfaces a mistyped reference path before any write side effect.
public func scriptPipelineExecute(
    scriptPath: String,
    outputPath: String,
    verifyAgainst: String? = nil
) throws -> ScriptExecuteResult {
    let scriptURL = URL(fileURLWithPath: scriptPath)
    guard FileManager.default.fileExists(atPath: scriptURL.path) else {
        throw ScriptPipelineError.fileNotFound(scriptPath)
    }
    let source = try String(contentsOf: scriptURL, encoding: .utf8)
    let log = try ScriptImporter.parse(source: source)

    // Pin the reference BEFORE any write (see ordering contract above).
    var reference: [String: Data]?
    if let referencePath = verifyAgainst {
        let referenceURL = URL(fileURLWithPath: referencePath)
        guard FileManager.default.fileExists(atPath: referenceURL.path) else {
            throw ScriptPipelineError.fileNotFound(referencePath)
        }
        reference = try RawPartChannel.readAllParts(from: referenceURL)
    }

    var document = WordDocument.emptyAuthoringDocument()
    try document.apply(operations: log.entries.map(\.op))
    let outputURL = URL(fileURLWithPath: outputPath)
    try document.writeAuthoringPackage(to: outputURL)

    guard let reference else {
        return ScriptExecuteResult(written: outputPath, verified: nil, brokenParts: [])
    }
    let rebuilt = try RawPartChannel.readAllParts(from: outputURL)
    let broken = PartFidelity.compareParts(reference: reference, rebuilt: rebuilt)
        .filter { !$0.isEqual }
        .map(\.partPath)
        .sorted()
    return ScriptExecuteResult(
        written: outputPath, verified: broken.isEmpty, brokenParts: broken)
}
