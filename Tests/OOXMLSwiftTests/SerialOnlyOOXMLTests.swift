import XCTest
@testable import OOXMLSwift

/// Phase B of `che-word-mcp-insert-crash-autosave-fix` (closes #41 prerequisite).
///
/// Spec: `openspec/changes/che-word-mcp-insert-crash-autosave-fix/specs/ooxml-content-insertion-primitives/spec.md`
/// Requirement: "ooxml-swift IO layer is fully serial; parallel primitives forbidden"
///
/// Rationale: libxml2-backed `XMLElement` is not thread-safe; `recover_from_autosave`
/// requires deterministic parsing. Regression-as-test ensures parallel primitives
/// never reappear under `Sources/OOXMLSwift/IO/`.
final class SerialOnlyOOXMLTests: XCTestCase {

    func testNoParallelPrimitivesInOOXMLIO() throws {
        // Resolve project-relative path to OOXML IO sources.
        let ioDir = URL(fileURLWithPath: #filePath)
            .deletingLastPathComponent()  // Tests/OOXMLSwiftTests/
            .deletingLastPathComponent()  // Tests/
            .deletingLastPathComponent()  // ooxml-swift/
            .appendingPathComponent("Sources/OOXMLSwift/IO")

        guard FileManager.default.fileExists(atPath: ioDir.path) else {
            throw XCTSkip("IO source dir not found at \(ioDir.path) (unexpected layout)")
        }

        // Forbidden primitives per the spec ADDED requirement.
        let forbidden = [
            "concurrentPerform",
            "withTaskGroup",
            "withThrowingTaskGroup",
            "DispatchQueue.global",
            "DispatchQueue.main.async",
            "Task.detached"
        ]

        // Walk the IO directory and grep each file.
        let files = try FileManager.default.contentsOfDirectory(
            at: ioDir, includingPropertiesForKeys: nil
        ).filter { $0.pathExtension == "swift" }

        var violations: [String] = []
        for file in files {
            let contents = (try? String(contentsOf: file, encoding: .utf8)) ?? ""
            let lines = contents.components(separatedBy: "\n")
            for (lineIdx, line) in lines.enumerated() {
                // Strip line comments to avoid false positives in doc-comments.
                let codeOnly: String = {
                    if let commentRange = line.range(of: "//") {
                        return String(line[..<commentRange.lowerBound])
                    }
                    return line
                }()
                for pattern in forbidden where codeOnly.contains(pattern) {
                    violations.append("\(file.lastPathComponent):\(lineIdx + 1) — '\(pattern)' in: \(line.trimmingCharacters(in: .whitespaces))")
                }
            }
        }

        XCTAssertTrue(violations.isEmpty,
                      "Parallel primitives forbidden under Sources/OOXMLSwift/IO/. Violations:\n" +
                      violations.joined(separator: "\n"))
    }
}
