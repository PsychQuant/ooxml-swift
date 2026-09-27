import XCTest
import Foundation

/// #162 — source guard against host-absolute fixture paths.
///
/// PR #119 (for #118 / PsychQuant/macdoc#172) moved the REC, document and JPA
/// fixtures onto `TemplateFixtureGate` / `MACDOC_TEMPLATE_DIR`, but the
/// planned source guard — a test that fails when a test file reintroduces a
/// host-absolute or parent-repo fixture path — was never added. Without it
/// the pattern comes back unnoticed: two instances existed on `main` at
/// v3.8.0 (`V1BenchmarkTests` hard-coding `~/Downloads/*.docx`,
/// `Issue66HeaderVMLProbeTests` using its own `OOXML_LOCAL_DOCUMENT_FIXTURE`
/// instead of the shared gate) — both migrated onto `TemplateFixtureGate` in
/// the same change that adds this guard.
///
/// This test scans `Tests/`, `Sources/` and `scripts/` (the last does not
/// currently exist in this repo, but is scanned if it ever does) for any
/// `.swift` file containing a literal path shaped like a macOS or Linux
/// per-user home directory (see `forbiddenPrefixes` below for the exact two
/// prefixes). A fixture path belongs behind `MACDOC_TEMPLATE_DIR` (see
/// `Tests/OOXMLSwiftTests/Helpers/TemplateFixtureGate.swift`), never baked
/// into source.
final class HostAbsoluteFixturePathGuardTests: XCTestCase {

    /// Built via concatenation, not as contiguous string literals — so this
    /// guard's own source text does not itself contain the patterns it
    /// hunts for, and would not flag itself if some future refactor folds
    /// the self-exclusion check below away.
    private static let forbiddenPrefixes: [String] = [
        "/" + "Users" + "/",
        "/" + "home" + "/",
    ]

    /// Files that are allowed to contain a forbidden prefix because the
    /// string is a FICTITIOUS example (attacker-controlled input asserted
    /// against in a security test), not a real host-specific fixture path.
    ///
    /// This is a closed, documented allowlist — not a general escape hatch.
    /// Adding an entry requires writing down why the match is not a real
    /// fixture path, the same way the two entries below do.
    private static let allowlist: [String: String] = [
        "PathTraversalSecurityTests.swift":
            "asserts that a fictitious attacker-controlled home-directory-shaped "
                + "path (a made-up username, an SSH key under it) is REJECTED by "
                + "isSafeRelativeOOXMLPath() — the string exists to be tested "
                + "AGAINST as a negative fixture, not to point at a real path",
        "HostAbsoluteFixturePathGuardTests.swift":
            "this file: a defensive backstop in case a future edit spells out a "
                + "path-shaped example in a doc comment or allowlist description "
                + "(as this dictionary's own values must, to explain WHY each "
                + "entry is allowlisted) — forbiddenPrefixes itself stays "
                + "concatenation-built so it does not need this backstop",
    ]

    private enum GuardError: Error, CustomStringConvertible {
        case repoRootNotFound(startedFrom: String)
        var description: String {
            switch self {
            case .repoRootNotFound(let startedFrom):
                // #162 (system-prompt spec): failing to find the repo root
                // MUST fail the test, not skip it — a skip here would be a
                // guard that silently guards nothing.
                return "could not locate repo root (directory containing Package.swift) "
                    + "walking up from \(startedFrom)"
            }
        }
    }

    private func findRepoRoot() throws -> URL {
        var dir = URL(fileURLWithPath: #filePath).deletingLastPathComponent()
        for _ in 0..<12 {
            if FileManager.default.fileExists(atPath: dir.appendingPathComponent("Package.swift").path) {
                return dir
            }
            let parent = dir.deletingLastPathComponent()
            if parent.path == dir.path { break }
            dir = parent
        }
        throw GuardError.repoRootNotFound(startedFrom: #filePath)
    }

    func testNoHostAbsoluteFixturePathsUnderTestsSourcesOrScripts() throws {
        let repoRoot = try findRepoRoot()
        let searchDirNames = ["Tests", "Sources", "scripts"]

        var violations: [String] = []
        for dirName in searchDirNames {
            let dirURL = repoRoot.appendingPathComponent(dirName)
            guard FileManager.default.fileExists(atPath: dirURL.path) else { continue }
            guard let enumerator = FileManager.default.enumerator(
                at: dirURL, includingPropertiesForKeys: nil
            ) else { continue }

            for case let fileURL as URL in enumerator {
                guard fileURL.pathExtension == "swift" else { continue }
                let fileName = fileURL.lastPathComponent
                if Self.allowlist[fileName] != nil { continue }
                guard let contents = try? String(contentsOf: fileURL, encoding: .utf8) else { continue }

                for (lineIndex, line) in contents.components(separatedBy: "\n").enumerated() {
                    for prefix in Self.forbiddenPrefixes {
                        if line.contains(prefix) {
                            violations.append(
                                "\(dirName)/\(fileURL.lastPathComponent):\(lineIndex + 1): "
                                    + "contains host-absolute path prefix \"\(prefix)\" — "
                                    + "fixture paths belong behind MACDOC_TEMPLATE_DIR "
                                    + "(TemplateFixtureGate), not baked into source")
                        }
                    }
                }
            }
        }

        XCTAssertTrue(violations.isEmpty,
            "host-absolute fixture paths found:\n" + violations.joined(separator: "\n"))
    }
}
