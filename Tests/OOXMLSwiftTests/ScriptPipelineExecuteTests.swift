import XCTest
@testable import OOXMLSwift

/// `ooxml-script-transcode` — script execution orchestration as a shared
/// transcode entry point (Spectra change `script-pipeline-surface`, tasks
/// 1.1 / 1.2).
///
/// These tests pin two requirements:
///
/// - "Script execution orchestration is a shared transcode entry point":
///   a caller holding only OOXMLSwift can execute a rebuild script. The
///   verdict is ABSENT — not `false` — when no reference was supplied, so
///   that silence can never be read as a passing verification.
/// - "The reference document is read before any output is written": the
///   reference is pinned in memory ahead of the write, so passing one path
///   as both output and reference compares against the PRE-write bytes
///   rather than against the run's own output.
final class ScriptPipelineExecuteTests: XCTestCase {

    // MARK: - Fixtures

    private func tempDir() throws -> URL {
        let dir = URL(fileURLWithPath: NSTemporaryDirectory())
            .appendingPathComponent("script-pipeline-execute-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        addTeardownBlock { try? FileManager.default.removeItem(at: dir) }
        return dir
    }

    /// An authoring log of one paragraph — the DSL-spellable subset, so the
    /// exported script exercises the typed channel rather than raw ops only.
    private func log(text: String, paraId: String) -> OperationLog {
        var l = OperationLog()
        l.append(
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: text, styleId: nil, paraId: paraId)),
            source: .swift)
        return l
    }

    /// Materialise a docx from a log, the same way execution will rebuild it.
    private func writeDocx(from l: OperationLog, to url: URL) throws {
        var document = WordDocument.emptyAuthoringDocument()
        try document.apply(operations: l.entries.map(\.op))
        try document.writeAuthoringPackage(to: url)
    }

    private func writeScript(from l: OperationLog, to url: URL) throws {
        try ScriptExporter.exportSwift(log: l)
            .write(to: url, atomically: true, encoding: String.Encoding.utf8)
    }

    // MARK: - Requirement: shared transcode entry point

    /// No reference supplied → the verdict must be absent. An empty
    /// broken-parts list alongside a `false`/`true` verdict would let a
    /// caller that checks only `brokenParts.isEmpty` read "not checked" as
    /// "checked and clean".
    func testExecutionWithoutReferenceReportsNoVerdict() throws {
        let dir = try tempDir()
        let script = dir.appendingPathComponent("doc.mdocx.swift")
        let output = dir.appendingPathComponent("out.docx")
        try writeScript(from: log(text: "第一段", paraId: "P1"), to: script)

        let result = try scriptPipelineExecute(
            scriptPath: script.path, outputPath: output.path)

        XCTAssertTrue(FileManager.default.fileExists(atPath: output.path),
                      "the rebuilt document must be written")
        XCTAssertEqual(result.written, output.path)
        XCTAssertNil(result.verified,
                     "verdict must be ABSENT when no reference was supplied")
    }

    /// Reference that matches → passing verdict, no differing parts.
    func testExecutionWithMatchingReferenceReportsByteEquality() throws {
        let dir = try tempDir()
        let l = log(text: "第一段", paraId: "P1")
        let reference = dir.appendingPathComponent("reference.docx")
        let script = dir.appendingPathComponent("doc.mdocx.swift")
        let output = dir.appendingPathComponent("out.docx")
        try writeDocx(from: l, to: reference)
        try writeScript(from: l, to: script)

        let result = try scriptPipelineExecute(
            scriptPath: script.path, outputPath: output.path,
            verifyAgainst: reference.path)

        XCTAssertEqual(result.verified, true, "replay of the same log must be byte-equal")
        XCTAssertEqual(result.brokenParts, [], "no part may differ")
    }

    /// Reference that diverges → failing verdict naming the differing parts.
    func testExecutionWithDivergingReferenceNamesBrokenParts() throws {
        let dir = try tempDir()
        let reference = dir.appendingPathComponent("reference.docx")
        let script = dir.appendingPathComponent("doc.mdocx.swift")
        let output = dir.appendingPathComponent("out.docx")
        try writeDocx(from: log(text: "原始內容", paraId: "P1"), to: reference)
        try writeScript(from: log(text: "改過的內容", paraId: "P1"), to: script)

        let result = try scriptPipelineExecute(
            scriptPath: script.path, outputPath: output.path,
            verifyAgainst: reference.path)

        XCTAssertEqual(result.verified, false, "diverging content must fail verification")
        XCTAssertFalse(result.brokenParts.isEmpty, "the differing parts must be named")
    }

    // MARK: - Requirement: reference read before any output is written

    /// Output path == reference path. Reading the reference AFTER the write
    /// would compare the rebuilt document against itself and manufacture a
    /// passing verdict. The verdict must reflect the file's PRE-write bytes.
    func testSamePathAsOutputAndReferenceDoesNotSelfCompare() throws {
        let dir = try tempDir()
        let shared = dir.appendingPathComponent("shared.docx")
        let script = dir.appendingPathComponent("doc.mdocx.swift")
        // The file on disk says one thing; the script rebuilds another.
        try writeDocx(from: log(text: "磁碟上的原始內容", paraId: "P1"), to: shared)
        try writeScript(from: log(text: "腳本重建的不同內容", paraId: "P1"), to: script)

        // Overwrite must be requested explicitly: this pattern necessarily
        // targets a file that already exists, so the gate applies to it on
        // the same terms as any other pre-existing output.
        let result = try scriptPipelineExecute(
            scriptPath: script.path, outputPath: shared.path,
            verifyAgainst: shared.path, overwrite: true)

        XCTAssertEqual(result.verified, false,
                       "must compare against the PRE-write bytes, not the run's own output")
        XCTAssertFalse(result.brokenParts.isEmpty,
                       "the parts that differ from the pre-write reference must be named")
    }

    // MARK: - Requirement: failed verification publishes nothing
    //
    // Spectra change `script-pipeline-failure-contract`, task 1.1.
    // These pin the non-destructive guarantee: a run whose verification
    // fails must leave the output path exactly as it found it. The old
    // shape wrote the rebuilt package to the output path FIRST and read it
    // back to compare, so a failing verdict arrived only after the original
    // had already been destroyed.

    /// No file at the output path + failing verification → nothing appears.
    /// Also asserts no staging artefact is abandoned in the output directory.
    func testFailedVerificationCreatesNoFileWhenNoneExisted() throws {
        let dir = try tempDir()
        let reference = dir.appendingPathComponent("reference.docx")
        let script = dir.appendingPathComponent("doc.mdocx.swift")
        let output = dir.appendingPathComponent("out.docx")
        try writeDocx(from: log(text: "原始內容", paraId: "P1"), to: reference)
        try writeScript(from: log(text: "改過的內容", paraId: "P1"), to: script)

        let result = try scriptPipelineExecute(
            scriptPath: script.path, outputPath: output.path,
            verifyAgainst: reference.path)

        XCTAssertEqual(result.verified, false, "diverging content must fail verification")
        XCTAssertFalse(FileManager.default.fileExists(atPath: output.path),
                       "a failed verification must not leave a document behind")

        let residue = try FileManager.default
            .contentsOfDirectory(atPath: dir.path)
            .filter { $0 != "reference.docx" && $0 != "doc.mdocx.swift" }
        XCTAssertEqual(residue, [],
                       "no staging artefact may be abandoned in the output directory")
        XCTAssertNil(result.written,
                     "nothing was published, so no written path may be named")
    }

    /// A file already at the output path + failing verification → that file
    /// is byte-identical afterwards.
    func testFailedVerificationLeavesExistingOutputUnmodified() throws {
        let dir = try tempDir()
        let reference = dir.appendingPathComponent("reference.docx")
        let script = dir.appendingPathComponent("doc.mdocx.swift")
        let output = dir.appendingPathComponent("out.docx")
        try writeDocx(from: log(text: "原始內容", paraId: "P1"), to: reference)
        try writeDocx(from: log(text: "先前就在輸出路徑上的文件", paraId: "P9"), to: output)
        try writeScript(from: log(text: "改過的內容", paraId: "P1"), to: script)

        let before = try Data(contentsOf: output)

        let result = try scriptPipelineExecute(
            scriptPath: script.path, outputPath: output.path,
            verifyAgainst: reference.path, overwrite: true)

        XCTAssertEqual(result.verified, false, "diverging content must fail verification")
        XCTAssertEqual(try Data(contentsOf: output), before,
                       "a failed verification must leave the pre-existing file untouched")
        XCTAssertNil(result.written,
                     "nothing was published, so no written path may be named")
    }

    // MARK: - Requirement: existing output is refused before the replay

    /// The gate lives in the entry point, defaults to refusing, and fires
    /// before the script is parsed — proven here by handing it a script that
    /// cannot parse: the overwrite refusal must win over the parse failure.
    func testExistingOutputIsRefusedBeforeTheScriptIsParsed() throws {
        let dir = try tempDir()
        let script = dir.appendingPathComponent("broken.mdocx.swift")
        let output = dir.appendingPathComponent("out.docx")
        try "this is not a valid mdocx script".write(
            to: script, atomically: true, encoding: .utf8)
        try writeDocx(from: log(text: "先前就在輸出路徑上的文件", paraId: "P9"), to: output)
        let before = try Data(contentsOf: output)

        XCTAssertThrowsError(
            try scriptPipelineExecute(
                scriptPath: script.path, outputPath: output.path)
        ) { error in
            guard case ScriptPipelineError.outputExists = error else {
                return XCTFail(
                    "the overwrite refusal must precede the parse, got \(type(of: error)): \(error)")
            }
        }
        XCTAssertEqual(try Data(contentsOf: output), before,
                       "the refused run must leave the existing file untouched")
    }

    /// Same path as output and reference, without overwrite permission.
    func testSamePathAsOutputAndReferenceWithoutOverwriteIsRefused() throws {
        let dir = try tempDir()
        let shared = dir.appendingPathComponent("shared.docx")
        let script = dir.appendingPathComponent("doc.mdocx.swift")
        try writeDocx(from: log(text: "磁碟上的原始內容", paraId: "P1"), to: shared)
        try writeScript(from: log(text: "腳本重建的不同內容", paraId: "P1"), to: script)
        let before = try Data(contentsOf: shared)

        XCTAssertThrowsError(
            try scriptPipelineExecute(
                scriptPath: script.path, outputPath: shared.path,
                verifyAgainst: shared.path))
        XCTAssertEqual(try Data(contentsOf: shared), before,
                       "the refused run must leave the file untouched")
    }

    /// A mistyped reference path is surfaced before any write side effect —
    /// the second reason the ordering contract exists.
    func testMissingReferenceIsRefusedBeforeAnyWrite() throws {
        let dir = try tempDir()
        let script = dir.appendingPathComponent("doc.mdocx.swift")
        let output = dir.appendingPathComponent("out.docx")
        try writeScript(from: log(text: "第一段", paraId: "P1"), to: script)

        XCTAssertThrowsError(
            try scriptPipelineExecute(
                scriptPath: script.path, outputPath: output.path,
                verifyAgainst: dir.appendingPathComponent("nope.docx").path))

        XCTAssertFalse(FileManager.default.fileExists(atPath: output.path),
                       "no output may be written when the reference is missing")
    }

    /// An unparseable script surfaces the transcoder's location-bearing
    /// reason rather than a generic failure.
    func testUnparseableScriptSurfacesTranscodeError() throws {
        let dir = try tempDir()
        let script = dir.appendingPathComponent("broken.mdocx.swift")
        let output = dir.appendingPathComponent("out.docx")
        try "this is not a valid mdocx script".write(
            to: script, atomically: true, encoding: .utf8)

        XCTAssertThrowsError(
            try scriptPipelineExecute(
                scriptPath: script.path, outputPath: output.path)
        ) { error in
            XCTAssertTrue(error is TranscodeError,
                          "parse failure must surface as TranscodeError, got \(type(of: error))")
        }
    }
}
