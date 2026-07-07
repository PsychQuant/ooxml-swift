// TemplateFixtureGate.swift
// format-alignment-engine Phase A task 1.5 — env-gate for private real-world
// template fixtures (`format-alignment-pipeline`, «Template fixture policy»;
// Decision 5). Real templates live outside version control under
// `$MACDOC_TEMPLATE_DIR`; tests resolve them here and XCTSkip when absent, so
// CI (which lacks the private documents) stays green. Mirrors the inline
// `OOXML_LOCAL_THESIS_FIXTURE` guard (Issue66HeaderVMLProbeTests) as a shared,
// reusable helper.

import XCTest
import Foundation

enum TemplateFixtureGate {

    /// The real-world baseline template named in the change proposal — the
    /// measured baseline (`90_template_ja.docx`) that motivated the dual-track
    /// design.
    static var baselineTemplateName: String { "90_template_ja.docx" }

    /// Resolves a real-template fixture `name` under the gate directory
    /// (`dirOverride` when given — for tests injecting a temp dir without
    /// mutating process env — otherwise `$MACDOC_TEMPLATE_DIR`). Throws
    /// `XCTSkip` when the gate is unset or the file is missing, so the calling
    /// test skips loudly rather than failing on CI.
    static func requireTemplate(_ name: String, dirOverride: String? = nil) throws -> URL {
        guard let dir = dirOverride ?? ProcessInfo.processInfo.environment["MACDOC_TEMPLATE_DIR"] else {
            throw XCTSkip("set MACDOC_TEMPLATE_DIR to a directory of real .docx templates to run this test")
        }
        let url = URL(fileURLWithPath: dir).appendingPathComponent(name)
        guard FileManager.default.fileExists(atPath: url.path) else {
            throw XCTSkip("template '\(name)' not found under \(dir) — gate present but file absent")
        }
        return url
    }
}
