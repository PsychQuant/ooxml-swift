import XCTest
@testable import OOXMLSwift

/// Phase-aware parameterized test runner for the `.mdocx` fixture corpus
/// (Spectra change `mdocx-fixture-corpus`, spec
/// `openspec/specs/mdocx-fixture-corpus/spec.md`, Requirement
/// "Phase-aware parameterized test runner contract").
///
/// ## Phase model
///
/// The runner walks every directory under
/// `Tests/OOXMLSwiftTests/Fixtures/mdocx/` (excluding `_normalizer/` and any
/// directory beginning with `_`) and runs assertions per fixture, gated by a
/// single Boolean `activatePhaseB`:
///
/// - **Phase A** (always-on, ships *active* in this change):
///   1. Directory layout pattern `<NN>[<letter>]-<short-slug>`.
///   2. Minimum file set (`<slug>.mdocx.swift`, `<slug>.docx`,
///      `<slug>.normalized.docx`, `README.md`) — slug derived from dir name.
///   3. `<slug>.mdocx.swift` parses as well-formed Swift source against the
///      `WordDSLSwift` module — see "Compile-pass design" below.
///   4. Corpus completeness: every Requirement 1-15 from `mdocx-grammar` has
///      ≥ 1 fixture matching its `NN` prefix.
///
/// - **Phase B** (gated by `activatePhaseB == true`, default `false`):
///   5. Execute `<slug>.mdocx.swift` through `WordDSLSwift` to produce a docx.
///   6. Pipe through `MdocxFixtureNormalizer.normalize(...)`.
///   7. Byte-equal compare against `<slug>.normalized.docx`.
///   8. If `<slug>.oplog.jsonl` present → byte-equal op log compare.
///   9. If `<slug>.snapshot.json` present (Requirement 14 only) → byte-equal
///      snapshot compare.
///   10. If `<slug>.expected-source.mdocx.swift` present (Requirement 15
///       only) → reverse-CLI diff against canonicalized source.
///
/// ## Activating Phase B
///
/// The Spectra change `word-aligned-state-sync` Phase 4 (which lands the full
/// `WordDSLSwift` module implementation) flips `activatePhaseB` from `false`
/// to `true`. That change MUST do nothing else to this file — Phase B logic
/// already lives here, ready to fire. The activating change SHALL also
/// confirm fixture `01-dual-extension-recognition` passes Phase B before
/// flipping the flag.
///
/// ## Compile-pass design (Phase A assertion 3)
///
/// Phase A uses a lightweight tokenization-based "looks like Swift" check:
/// the file must be non-empty, contain an `import` keyword, and contain at
/// least one balanced brace pair. This is **Option B** from the task spec —
/// the simpler approach that defers real `swiftc` invocation to follow-up
/// work. The proper Phase B (or follow-up) approach is Option A2: an
/// auxiliary SwiftPM target wired in `Package.swift` that includes every
/// `.mdocx.swift` and depends on `WordDSLSwift`, asserting that auxiliary
/// target builds. When fixture authors write malformed Swift, the WordDSLSwift
/// module test target (or the activating change's Phase B execution) will
/// catch it.
///
/// TODO(post-mdocx-fixture-corpus): Wire Option A2 — add a SwiftPM auxiliary
/// "MdocxFixturesCompileCheck" target whose source files are the
/// `.mdocx.swift` files and whose dependency is `WordDSLSwift`. Replace the
/// tokenization heuristic in `assertMdocxSwiftCompiles(...)` with a build
/// assertion against that target.
///
/// ## Convention references
///
/// - `Tests/OOXMLSwiftTests/TreeRoundTripGoldenTests.swift` — `#filePath`
///   resolution for fixture roots in SwiftPM test runs.
/// - `Tests/OOXMLSwiftTests/Fixtures/mdocx/_normalizer/word-default-theme.xml`
///   — default theme bytes consumed by `MdocxFixtureNormalizer` in Phase B.
final class MdocxFixtureCorpusTests: XCTestCase {

    // MARK: - Phase gate

    /// Master gate for Phase B. Flipped to `true` by the change that lands
    /// the full `WordDSLSwift` implementation (`word-aligned-state-sync`
    /// Phase 4). Do NOT flip this in any other change.
    static let activatePhaseB: Bool = false

    // MARK: - Configuration

    /// Range of `mdocx-grammar` Requirement numbers the corpus MUST cover.
    /// Sourced from `openspec/specs/mdocx-grammar/spec.md` which defines
    /// 15 numbered Requirements.
    private static let requiredRequirementRange: ClosedRange<Int> = 1...15

    /// Requirements that legitimately use `<slug>.snapshot.json` (per spec
    /// Requirement "Optional file additions for special Requirements").
    private static let snapshotEligibleRequirements: Set<Int> = [14]

    /// Requirements that legitimately use `<slug>.expected-source.mdocx.swift`.
    private static let reverseSourceEligibleRequirements: Set<Int> = [15]

    /// Requirements that legitimately use `<slug>.oplog.jsonl`.
    /// Per spec, Requirement 7 (component envelope), 9 (define-on-first-use
    /// ordering), and 14 (atomic three-file save — the oplog is one of the
    /// three files atomically written) constrain op-log shape.
    private static let oplogEligibleRequirements: Set<Int> = [7, 9, 14]

    // MARK: - Fixture root resolution

    /// Resolve the `Fixtures/mdocx/` directory relative to this source file.
    /// `#filePath` is the robust SwiftPM test convention used elsewhere in
    /// this target (see `TreeRoundTripGoldenTests.swift`).
    private static var fixturesRoot: URL {
        URL(fileURLWithPath: #filePath)
            .deletingLastPathComponent()  // OOXMLSwiftTests/
            .appendingPathComponent("Fixtures")
            .appendingPathComponent("mdocx")
    }

    // MARK: - Top-level test entry

    /// Single XCTest method that runs all parameterized fixtures via
    /// `XCTContext.runActivity`. Each fixture becomes a named activity so
    /// failures report per-fixture in Xcode.
    func testAllFixtures() throws {
        let fm = FileManager.default
        let root = Self.fixturesRoot

        // Gracefully handle absent root: corpus completeness will still fail
        // (no Requirement covered) but layout/file-set checks must not crash.
        var fixtureDirs: [URL] = []
        if fm.fileExists(atPath: root.path) {
            fixtureDirs = try Self.discoverFixtureDirectories(at: root)
        }

        // Track Requirement coverage as we walk fixtures.
        var coveredRequirements: Set<Int> = []

        for fixtureDir in fixtureDirs {
            let dirName = fixtureDir.lastPathComponent
            XCTContext.runActivity(named: "fixture: \(dirName)") { _ in
                do {
                    let parsed = try Self.parseDirectoryName(dirName)
                    coveredRequirements.insert(parsed.requirementNumber)

                    // Phase A assertions
                    try Self.assertMinimumFileSet(
                        fixtureDir: fixtureDir,
                        slug: parsed.slug
                    )
                    try Self.assertOptionalFilesAreEligible(
                        fixtureDir: fixtureDir,
                        slug: parsed.slug,
                        requirementNumber: parsed.requirementNumber
                    )
                    try Self.assertMdocxSwiftCompiles(
                        fixtureDir: fixtureDir,
                        slug: parsed.slug
                    )

                    // Phase B assertions (gated)
                    if Self.activatePhaseB {
                        try Self.runPhaseB(
                            fixtureDir: fixtureDir,
                            slug: parsed.slug,
                            requirementNumber: parsed.requirementNumber
                        )
                    }
                } catch let error as FixtureFailure {
                    XCTFail(error.message)
                } catch {
                    XCTFail(Self.formatFailure(
                        fixtureDir: dirName,
                        phase: .phaseA,
                        assertion: "unexpected error",
                        details: "\(error)"
                    ))
                }
            }
        }

        // Phase A assertion 4: corpus completeness.
        Self.assertCorpusCompleteness(coveredRequirements: coveredRequirements)
    }

    // MARK: - Discovery

    /// List immediate subdirectories under `root`, excluding any name that
    /// starts with `_` (e.g., `_normalizer/`). Sorted for deterministic
    /// iteration order.
    private static func discoverFixtureDirectories(at root: URL) throws -> [URL] {
        let fm = FileManager.default
        let contents = try fm.contentsOfDirectory(
            at: root,
            includingPropertiesForKeys: [.isDirectoryKey],
            options: [.skipsHiddenFiles]
        )
        return contents
            .filter { url in
                let name = url.lastPathComponent
                guard !name.hasPrefix("_") else { return false }
                var isDir: ObjCBool = false
                let exists = fm.fileExists(atPath: url.path, isDirectory: &isDir)
                return exists && isDir.boolValue
            }
            .sorted { $0.lastPathComponent < $1.lastPathComponent }
    }

    // MARK: - Phase A assertions

    /// Parse a fixture directory name into `(NN, optional letter, slug)`.
    /// Pattern: `<NN>[<letter>]-<short-slug>` where NN is two-digit
    /// zero-padded.
    private static func parseDirectoryName(_ name: String) throws -> ParsedDirectoryName {
        // Reject obvious malformed cases up front for clearer messages.
        guard !name.isEmpty else {
            throw FixtureFailure(formatFailure(
                fixtureDir: name,
                phase: .phaseA,
                assertion: "layout pattern",
                details: "directory name is empty"
            ))
        }
        // Layout regex: ^(\d{2})([a-z])?-([a-z0-9]+(?:-[a-z0-9]+)*)$
        // Disallow uppercase, spaces, or punctuation other than hyphen separators.
        let pattern = #"^(\d{2})([a-z])?-([a-z0-9]+(?:-[a-z0-9]+)*)$"#
        guard let regex = try? NSRegularExpression(pattern: pattern) else {
            throw FixtureFailure(formatFailure(
                fixtureDir: name,
                phase: .phaseA,
                assertion: "layout pattern",
                details: "internal: regex compile failed"
            ))
        }
        let nsName = name as NSString
        let range = NSRange(location: 0, length: nsName.length)
        guard let match = regex.firstMatch(in: name, range: range),
              match.numberOfRanges >= 4
        else {
            throw FixtureFailure(formatFailure(
                fixtureDir: name,
                phase: .phaseA,
                assertion: "layout pattern",
                details: "directory name '\(name)' does not match `<NN>[<letter>]-<short-slug>` "
                    + "(NN must be two-digit zero-padded; slug must be lowercase kebab-case)"
            ))
        }
        let nnString = nsName.substring(with: match.range(at: 1))
        guard let requirementNumber = Int(nnString) else {
            throw FixtureFailure(formatFailure(
                fixtureDir: name,
                phase: .phaseA,
                assertion: "layout pattern",
                details: "could not parse '\(nnString)' as Requirement number"
            ))
        }
        let letterRange = match.range(at: 2)
        let letter: String? = letterRange.location == NSNotFound
            ? nil
            : nsName.substring(with: letterRange)
        let slug = nsName.substring(with: match.range(at: 3))
        return ParsedDirectoryName(
            directoryName: name,
            requirementNumber: requirementNumber,
            letter: letter,
            slug: slug
        )
    }

    /// Assert the four-file minimum set is present.
    private static func assertMinimumFileSet(fixtureDir: URL, slug: String) throws {
        let fm = FileManager.default
        let required = [
            "\(slug).mdocx.swift",
            "\(slug).docx",
            "\(slug).normalized.docx",
            "README.md"
        ]
        var missing: [String] = []
        for filename in required {
            let path = fixtureDir.appendingPathComponent(filename).path
            if !fm.fileExists(atPath: path) {
                missing.append(filename)
            }
        }
        guard missing.isEmpty else {
            throw FixtureFailure(formatFailure(
                fixtureDir: fixtureDir.lastPathComponent,
                phase: .phaseA,
                assertion: "file set",
                details: "missing required file(s): \(missing.joined(separator: ", "))",
                paths: missing.map { fixtureDir.appendingPathComponent($0).path }
            ))
        }
    }

    /// Assert that any optional files present are eligible for the covered
    /// Requirement. Per spec Requirement "Optional file additions for special
    /// Requirements", an irrelevant fixture containing e.g. `snapshot.json`
    /// SHALL fail with a message naming which Requirements legitimately use
    /// each optional file.
    private static func assertOptionalFilesAreEligible(
        fixtureDir: URL,
        slug: String,
        requirementNumber: Int
    ) throws {
        let fm = FileManager.default
        let snapshotPath = fixtureDir.appendingPathComponent("\(slug).snapshot.json")
        if fm.fileExists(atPath: snapshotPath.path),
           !snapshotEligibleRequirements.contains(requirementNumber) {
            throw FixtureFailure(formatFailure(
                fixtureDir: fixtureDir.lastPathComponent,
                phase: .phaseA,
                assertion: "optional file eligibility",
                details: "Requirement \(requirementNumber) does not use snapshot assertions; "
                    + "snapshot.json is legitimate only for Requirement(s): "
                    + snapshotEligibleRequirements.sorted().map(String.init).joined(separator: ", "),
                paths: [snapshotPath.path]
            ))
        }
        let reverseSourcePath = fixtureDir.appendingPathComponent("\(slug).expected-source.mdocx.swift")
        if fm.fileExists(atPath: reverseSourcePath.path),
           !reverseSourceEligibleRequirements.contains(requirementNumber) {
            throw FixtureFailure(formatFailure(
                fixtureDir: fixtureDir.lastPathComponent,
                phase: .phaseA,
                assertion: "optional file eligibility",
                details: "Requirement \(requirementNumber) does not use reverse-source assertions; "
                    + "expected-source.mdocx.swift is legitimate only for Requirement(s): "
                    + reverseSourceEligibleRequirements.sorted().map(String.init).joined(separator: ", "),
                paths: [reverseSourcePath.path]
            ))
        }
        let oplogPath = fixtureDir.appendingPathComponent("\(slug).oplog.jsonl")
        if fm.fileExists(atPath: oplogPath.path),
           !oplogEligibleRequirements.contains(requirementNumber) {
            throw FixtureFailure(formatFailure(
                fixtureDir: fixtureDir.lastPathComponent,
                phase: .phaseA,
                assertion: "optional file eligibility",
                details: "Requirement \(requirementNumber) does not use op-log assertions; "
                    + "oplog.jsonl is legitimate only for Requirement(s): "
                    + oplogEligibleRequirements.sorted().map(String.init).joined(separator: ", "),
                paths: [oplogPath.path]
            ))
        }
    }

    /// Phase A "compile-pass" check using Option B: lightweight tokenization
    /// that asserts the file is non-empty, contains an `import` keyword, and
    /// has at least one balanced brace pair. See class header "Compile-pass
    /// design" for why this is sufficient at Phase A and the TODO for the
    /// proper Option A2 approach.
    private static func assertMdocxSwiftCompiles(fixtureDir: URL, slug: String) throws {
        let scriptURL = fixtureDir.appendingPathComponent("\(slug).mdocx.swift")
        let data: Data
        do {
            data = try Data(contentsOf: scriptURL)
        } catch {
            throw FixtureFailure(formatFailure(
                fixtureDir: fixtureDir.lastPathComponent,
                phase: .phaseA,
                assertion: "Swift compile-pass",
                details: "could not read \(scriptURL.lastPathComponent): \(error)",
                paths: [scriptURL.path]
            ))
        }
        guard !data.isEmpty else {
            throw FixtureFailure(formatFailure(
                fixtureDir: fixtureDir.lastPathComponent,
                phase: .phaseA,
                assertion: "Swift compile-pass",
                details: "\(scriptURL.lastPathComponent) is empty",
                paths: [scriptURL.path]
            ))
        }
        guard let source = String(data: data, encoding: .utf8) else {
            throw FixtureFailure(formatFailure(
                fixtureDir: fixtureDir.lastPathComponent,
                phase: .phaseA,
                assertion: "Swift compile-pass",
                details: "\(scriptURL.lastPathComponent) is not valid UTF-8",
                paths: [scriptURL.path]
            ))
        }
        let stripped = stripSwiftCommentsAndStrings(source)
        guard stripped.range(of: #"\bimport\b"#, options: .regularExpression) != nil else {
            throw FixtureFailure(formatFailure(
                fixtureDir: fixtureDir.lastPathComponent,
                phase: .phaseA,
                assertion: "Swift compile-pass",
                details: "\(scriptURL.lastPathComponent) contains no `import` declaration "
                    + "(every .mdocx.swift fixture must `import WordDSLSwift`)",
                paths: [scriptURL.path]
            ))
        }
        let openBraces = stripped.filter { $0 == "{" }.count
        let closeBraces = stripped.filter { $0 == "}" }.count
        guard openBraces > 0, openBraces == closeBraces else {
            throw FixtureFailure(formatFailure(
                fixtureDir: fixtureDir.lastPathComponent,
                phase: .phaseA,
                assertion: "Swift compile-pass",
                details: "\(scriptURL.lastPathComponent) has unbalanced braces "
                    + "(open=\(openBraces), close=\(closeBraces))",
                paths: [scriptURL.path]
            ))
        }
    }

    /// Strip Swift line comments, block comments, and string literals from
    /// `source` so brace-counting and keyword-detection in
    /// `assertMdocxSwiftCompiles` aren't fooled by braces or `import` words
    /// embedded in comments / strings.
    ///
    /// Intentionally minimal — does not handle every Swift edge case
    /// (e.g., raw string literals `#"..."#`, multiline strings `"""..."""`,
    /// nested block comments). Sufficient for the heuristic; real
    /// compile-pass is Option A2 (see TODO above).
    private static func stripSwiftCommentsAndStrings(_ source: String) -> String {
        enum Mode { case code, lineComment, blockComment, string }
        var mode: Mode = .code
        var out = ""
        out.reserveCapacity(source.count)
        let chars = Array(source)
        var i = 0
        while i < chars.count {
            let c = chars[i]
            switch mode {
            case .code:
                if c == "/", i + 1 < chars.count, chars[i + 1] == "/" {
                    mode = .lineComment
                    i += 2
                } else if c == "/", i + 1 < chars.count, chars[i + 1] == "*" {
                    mode = .blockComment
                    i += 2
                } else if c == "\"" {
                    mode = .string
                    i += 1
                } else {
                    out.append(c)
                    i += 1
                }
            case .lineComment:
                if c == "\n" {
                    mode = .code
                    out.append(c)
                }
                i += 1
            case .blockComment:
                if c == "*", i + 1 < chars.count, chars[i + 1] == "/" {
                    mode = .code
                    i += 2
                } else {
                    i += 1
                }
            case .string:
                if c == "\\", i + 1 < chars.count {
                    // Skip escaped char (covers \", \\, etc.)
                    i += 2
                } else if c == "\"" {
                    mode = .code
                    i += 1
                } else {
                    i += 1
                }
            }
        }
        return out
    }

    /// Assert every Requirement in `requiredRequirementRange` is covered by
    /// at least one fixture (matched by `NN` prefix).
    private static func assertCorpusCompleteness(coveredRequirements: Set<Int>) {
        let uncovered = requiredRequirementRange.filter { !coveredRequirements.contains($0) }
        guard uncovered.isEmpty else {
            let list = uncovered.map { String(format: "%02d", $0) }.joined(separator: ", ")
            XCTFail(formatFailure(
                fixtureDir: "<corpus>",
                phase: .phaseA,
                assertion: "corpus completeness",
                details: "the following mdocx-grammar Requirements have no fixture "
                    + "(no directory under Fixtures/mdocx/ with matching NN prefix): \(list). "
                    + "Per spec `mdocx-fixture-corpus`, every Requirement 1-15 from "
                    + "`mdocx-grammar` MUST have ≥ 1 fixture. See: "
                    + "openspec/specs/mdocx-fixture-corpus/spec.md "
                    + "Requirement 'Corpus completeness contract'."
            ))
            return
        }
    }

    // MARK: - Phase B assertions (gated)

    /// Execute the Phase B assertion chain for a single fixture. Only invoked
    /// when `activatePhaseB == true`. The activating change MUST land
    /// `MdocxFixtureNormalizer` (sibling task 1.2 in this same Spectra
    /// change) and the full `WordDSLSwift` execution surface
    /// (`word-aligned-state-sync` Phase 4) before flipping the flag.
    private static func runPhaseB(
        fixtureDir: URL,
        slug: String,
        requirementNumber: Int
    ) throws {
        let scriptURL = fixtureDir.appendingPathComponent("\(slug).mdocx.swift")
        let normalizedGoldenURL = fixtureDir.appendingPathComponent("\(slug).normalized.docx")
        let oplogURL = fixtureDir.appendingPathComponent("\(slug).oplog.jsonl")
        let snapshotURL = fixtureDir.appendingPathComponent("\(slug).snapshot.json")
        let reverseSourceURL = fixtureDir.appendingPathComponent("\(slug).expected-source.mdocx.swift")

        // 5. Execute the .mdocx.swift script through WordDSLSwift to produce
        //    an output docx. The script runner is provided as a subprocess:
        //    `swift run --package-path <ooxml-swift> WordDSLSwiftScriptRunner
        //    <scriptURL> <workDir>` writes its outputs (`out.docx`, optional
        //    `out.oplog.jsonl`, optional `out.snapshot.json`) into a temp
        //    work directory we control. Both the runner executable and its
        //    output contract land alongside `word-aligned-state-sync` Phase 4.
        let workDir = makeTempDirectory()
        defer { try? FileManager.default.removeItem(at: workDir) }
        let executionResult: ScriptExecutionResult
        do {
            executionResult = try runMdocxScript(scriptURL: scriptURL, workDir: workDir)
        } catch {
            throw FixtureFailure(formatFailure(
                fixtureDir: fixtureDir.lastPathComponent,
                phase: .phaseB,
                assertion: "script execution",
                details: "running \(scriptURL.lastPathComponent) failed: \(error)",
                paths: [scriptURL.path]
            ))
        }

        // 6. Pipe through MdocxFixtureNormalizer.
        let defaultThemeURL = fixturesRoot
            .appendingPathComponent("_normalizer")
            .appendingPathComponent("word-default-theme.xml")
        let defaultThemeBytes: Data
        do {
            defaultThemeBytes = try Data(contentsOf: defaultThemeURL)
        } catch {
            throw FixtureFailure(formatFailure(
                fixtureDir: fixtureDir.lastPathComponent,
                phase: .phaseB,
                assertion: "normalizer setup",
                details: "could not read default-theme reference: \(error)",
                paths: [defaultThemeURL.path]
            ))
        }
        let normalizedOutput: Data
        do {
            normalizedOutput = try MdocxFixtureNormalizer.normalize(
                docxBytes: executionResult.docxBytes,
                defaultThemeBytes: defaultThemeBytes
            )
        } catch {
            throw FixtureFailure(formatFailure(
                fixtureDir: fixtureDir.lastPathComponent,
                phase: .phaseB,
                assertion: "normalizer execution",
                details: "MdocxFixtureNormalizer.normalize threw: \(error)"
            ))
        }

        // 7. Byte-equal compare against <slug>.normalized.docx.
        let goldenBytes: Data
        do {
            goldenBytes = try Data(contentsOf: normalizedGoldenURL)
        } catch {
            throw FixtureFailure(formatFailure(
                fixtureDir: fixtureDir.lastPathComponent,
                phase: .phaseB,
                assertion: "golden read",
                details: "could not read \(normalizedGoldenURL.lastPathComponent): \(error)",
                paths: [normalizedGoldenURL.path]
            ))
        }
        guard normalizedOutput == goldenBytes else {
            throw FixtureFailure(formatFailure(
                fixtureDir: fixtureDir.lastPathComponent,
                phase: .phaseB,
                assertion: "docx byte-diff",
                details: "normalized output (\(normalizedOutput.count) bytes) does not match "
                    + "\(normalizedGoldenURL.lastPathComponent) (\(goldenBytes.count) bytes)",
                paths: [normalizedGoldenURL.path]
            ))
        }

        // 8. If <slug>.oplog.jsonl present → byte-equal op log compare.
        let fm = FileManager.default
        if fm.fileExists(atPath: oplogURL.path) {
            let expectedOplog = try Data(contentsOf: oplogURL)
            guard executionResult.opLogBytes == expectedOplog else {
                throw FixtureFailure(formatFailure(
                    fixtureDir: fixtureDir.lastPathComponent,
                    phase: .phaseB,
                    assertion: "op-log byte-diff",
                    details: "produced op log (\(executionResult.opLogBytes.count) bytes) "
                        + "does not match \(oplogURL.lastPathComponent) "
                        + "(\(expectedOplog.count) bytes)",
                    paths: [oplogURL.path]
                ))
            }
        }

        // 9. If <slug>.snapshot.json present (Requirement 14 only) → byte-equal
        //    snapshot compare.
        if fm.fileExists(atPath: snapshotURL.path) {
            // Eligibility was already enforced in Phase A; defensive double-check.
            guard snapshotEligibleRequirements.contains(requirementNumber) else {
                throw FixtureFailure(formatFailure(
                    fixtureDir: fixtureDir.lastPathComponent,
                    phase: .phaseB,
                    assertion: "snapshot eligibility",
                    details: "snapshot.json present but Requirement \(requirementNumber) is not eligible",
                    paths: [snapshotURL.path]
                ))
            }
            let expectedSnapshot = try Data(contentsOf: snapshotURL)
            guard executionResult.snapshotBytes == expectedSnapshot else {
                throw FixtureFailure(formatFailure(
                    fixtureDir: fixtureDir.lastPathComponent,
                    phase: .phaseB,
                    assertion: "snapshot byte-diff",
                    details: "produced snapshot (\(executionResult.snapshotBytes?.count ?? 0) bytes) "
                        + "does not match \(snapshotURL.lastPathComponent) "
                        + "(\(expectedSnapshot.count) bytes)",
                    paths: [snapshotURL.path]
                ))
            }
        }

        // 10. If <slug>.expected-source.mdocx.swift present (Requirement 15
        //     only) → reverse-CLI diff against canonicalized source.
        if fm.fileExists(atPath: reverseSourceURL.path) {
            guard reverseSourceEligibleRequirements.contains(requirementNumber) else {
                throw FixtureFailure(formatFailure(
                    fixtureDir: fixtureDir.lastPathComponent,
                    phase: .phaseB,
                    assertion: "reverse-source eligibility",
                    details: "expected-source.mdocx.swift present but Requirement "
                        + "\(requirementNumber) is not eligible",
                    paths: [reverseSourceURL.path]
                ))
            }
            let expectedReverseSource = try String(contentsOf: reverseSourceURL, encoding: .utf8)
            let reversedSource: String
            do {
                reversedSource = try runMacdocWordReverse(docxBytes: normalizedOutput)
            } catch {
                throw FixtureFailure(formatFailure(
                    fixtureDir: fixtureDir.lastPathComponent,
                    phase: .phaseB,
                    assertion: "reverse-CLI execution",
                    details: "`macdoc word reverse` failed: \(error)"
                ))
            }
            let canonicalProduced = canonicalizeMdocxSource(reversedSource)
            let canonicalExpected = canonicalizeMdocxSource(expectedReverseSource)
            guard canonicalProduced == canonicalExpected else {
                throw FixtureFailure(formatFailure(
                    fixtureDir: fixtureDir.lastPathComponent,
                    phase: .phaseB,
                    assertion: "reverse-source diff",
                    details: "canonicalized reverse source does not match "
                        + "\(reverseSourceURL.lastPathComponent)",
                    paths: [reverseSourceURL.path]
                ))
            }
        }
    }

    /// Canonicalize a `.mdocx.swift` source for reverse-direction comparison:
    /// collapse whitespace runs and normalize parameter order. Per spec:
    /// "after canonicalization (whitespace + parameter order normalization)".
    /// Parameter-order normalization is delegated to a future helper; the
    /// minimal implementation here normalizes whitespace, which catches the
    /// common case while leaving room for the parameter-order pass to land
    /// alongside the activating change.
    private static func canonicalizeMdocxSource(_ source: String) -> String {
        let lines = source.components(separatedBy: .newlines)
        let trimmed = lines.map { $0.trimmingCharacters(in: .whitespaces) }
        let nonEmpty = trimmed.filter { !$0.isEmpty }
        return nonEmpty.joined(separator: "\n")
    }

    // MARK: - Failure formatting

    private enum Phase: String {
        case phaseA = "Phase A"
        case phaseB = "Phase B"
    }

    /// Format a failure message containing fixture name, phase, assertion
    /// type, paths, and the README pointer.
    private static func formatFailure(
        fixtureDir: String,
        phase: Phase,
        assertion: String,
        details: String,
        paths: [String] = []
    ) -> String {
        var lines: [String] = []
        lines.append("[mdocx-fixture-corpus] \(phase.rawValue) — \(assertion) failed")
        lines.append("Fixture: \(fixtureDir)")
        lines.append("Details: \(details)")
        if !paths.isEmpty {
            lines.append("Paths:")
            for p in paths {
                lines.append("  - \(p)")
            }
        }
        if fixtureDir != "<corpus>" {
            lines.append("See: Tests/OOXMLSwiftTests/Fixtures/mdocx/\(fixtureDir)/README.md")
        }
        return lines.joined(separator: "\n")
    }

    // MARK: - Supporting types

    private struct ParsedDirectoryName {
        let directoryName: String
        let requirementNumber: Int
        let letter: String?
        let slug: String
    }

    /// Internal failure marker — its `message` is the fully formatted
    /// `formatFailure(...)` output. Throwing this preserves the rich message
    /// while letting the top-level loop catch and report per-fixture.
    private struct FixtureFailure: Error {
        let message: String
        init(_ message: String) { self.message = message }
    }

    /// Result of executing a `.mdocx.swift` script: the produced docx bytes,
    /// plus optional op-log and snapshot bytes when the script's component /
    /// snapshot facets emitted them. The runner subprocess writes these to
    /// fixed filenames in `workDir`.
    private struct ScriptExecutionResult {
        let docxBytes: Data
        let opLogBytes: Data
        let snapshotBytes: Data?
    }

    // MARK: - Phase B subprocess helpers

    /// Create a unique temporary directory for one fixture's Phase B run.
    /// Caller is responsible for `removeItem(at:)` cleanup.
    private static func makeTempDirectory() -> URL {
        let url = URL(fileURLWithPath: NSTemporaryDirectory())
            .appendingPathComponent("mdocx-fixture-corpus-\(UUID().uuidString)", isDirectory: true)
        try? FileManager.default.createDirectory(at: url, withIntermediateDirectories: true)
        return url
    }

    /// Invoke the `WordDSLSwiftScriptRunner` subprocess to execute a fixture's
    /// `.mdocx.swift` and read back the produced artifacts. The runner is
    /// landed by `word-aligned-state-sync` Phase 4 alongside the
    /// `activatePhaseB` flip; it writes:
    ///   - `<workDir>/out.docx` (always)
    ///   - `<workDir>/out.oplog.jsonl` (always)
    ///   - `<workDir>/out.snapshot.json` (only when the script emits one)
    private static func runMdocxScript(scriptURL: URL, workDir: URL) throws -> ScriptExecutionResult {
        try runProcess(
            executable: "/usr/bin/env",
            arguments: [
                "swift", "run",
                "--package-path", packageRoot().path,
                "WordDSLSwiftScriptRunner",
                scriptURL.path,
                workDir.path
            ]
        )
        let docxURL = workDir.appendingPathComponent("out.docx")
        let oplogURL = workDir.appendingPathComponent("out.oplog.jsonl")
        let snapshotURL = workDir.appendingPathComponent("out.snapshot.json")
        let docxBytes = try Data(contentsOf: docxURL)
        let opLogBytes = try Data(contentsOf: oplogURL)
        let fm = FileManager.default
        let snapshotBytes: Data? = fm.fileExists(atPath: snapshotURL.path)
            ? try Data(contentsOf: snapshotURL)
            : nil
        return ScriptExecutionResult(
            docxBytes: docxBytes,
            opLogBytes: opLogBytes,
            snapshotBytes: snapshotBytes
        )
    }

    /// Invoke `macdoc word reverse` against a docx and return the reversed
    /// `.mdocx.swift` source (Requirement 15).
    private static func runMacdocWordReverse(docxBytes: Data) throws -> String {
        let workDir = makeTempDirectory()
        defer { try? FileManager.default.removeItem(at: workDir) }
        let docxURL = workDir.appendingPathComponent("input.docx")
        let outURL = workDir.appendingPathComponent("reversed.mdocx.swift")
        try docxBytes.write(to: docxURL)
        try runProcess(
            executable: "/usr/bin/env",
            arguments: [
                "macdoc", "word", "reverse",
                docxURL.path,
                "--output", outURL.path
            ]
        )
        return try String(contentsOf: outURL, encoding: .utf8)
    }

    /// Run a subprocess and throw if its exit status is non-zero.
    @discardableResult
    private static func runProcess(executable: String, arguments: [String]) throws -> String {
        let process = Process()
        process.executableURL = URL(fileURLWithPath: executable)
        process.arguments = arguments
        let stdoutPipe = Pipe()
        let stderrPipe = Pipe()
        process.standardOutput = stdoutPipe
        process.standardError = stderrPipe
        try process.run()
        process.waitUntilExit()
        let stdoutData = stdoutPipe.fileHandleForReading.readDataToEndOfFile()
        let stderrData = stderrPipe.fileHandleForReading.readDataToEndOfFile()
        let stdout = String(data: stdoutData, encoding: .utf8) ?? ""
        let stderr = String(data: stderrData, encoding: .utf8) ?? ""
        guard process.terminationStatus == 0 else {
            throw NSError(
                domain: "MdocxFixtureCorpusTests",
                code: Int(process.terminationStatus),
                userInfo: [
                    NSLocalizedDescriptionKey: """
                        subprocess `\(executable) \(arguments.joined(separator: " "))` \
                        exited with status \(process.terminationStatus).
                        stdout: \(stdout)
                        stderr: \(stderr)
                        """
                ]
            )
        }
        return stdout
    }

    /// Resolve the `ooxml-swift` package root from `#filePath`. The runner
    /// subprocess is invoked with `--package-path <packageRoot>`.
    private static func packageRoot() -> URL {
        URL(fileURLWithPath: #filePath)
            .deletingLastPathComponent()  // OOXMLSwiftTests/
            .deletingLastPathComponent()  // Tests/
            .deletingLastPathComponent()  // <package root>
    }
}
