// TemplateFixtureGate.swift
// format-alignment-engine Phase A task 1.5 — env-gate for private real-world
// template fixtures (`format-alignment-pipeline`, «Template fixture policy»;
// Decision 5). Real templates live outside version control under
// `$MACDOC_TEMPLATE_DIR`; tests resolve them here and XCTSkip when absent, so
// CI (which lacks the private documents) stays green. Also used by
// `V1BenchmarkTests` and `Issue66HeaderVMLProbeTests` (#162) — both
// previously bypassed this shared gate (host-absolute home-directory paths
// hard-coded in source, and a standalone `OOXML_LOCAL_DOCUMENT_FIXTURE` env
// var, respectively).

import XCTest
import Foundation

enum TemplateFixtureGate {

    /// The real-world baseline template named in the change proposal — the
    /// measured baseline (`90_template_ja.docx`) that motivated the dual-track
    /// design.
    static var baselineTemplateName: String { "90_template_ja.docx" }

    /// Private real-world REC form used by the raw-channel slot E2E matrix.
    /// The filename is stable, but the document lives outside version control
    /// and is supplied through the same MACDOC_TEMPLATE_DIR gate.
    static var recFixtureName: String {
        "REC-O-01-新案審查送件核對單-20250220公告.docx"
    }

    /// Private document corpus fixture used by the content-equality matrix.
    static var documentFixtureName: String { "document-fixture.docx" }

    /// #162 — `V1BenchmarkTests` 6.6/6.6-attr candidate fixtures, largest
    /// first. Neither name is special; they're simply the two real-world
    /// documents recorded in `docs/benchmarks-word-aligned-state-sync.md`
    /// when those benchmarks were last run.
    static var benchmarkFixtureNames: [String] {
        ["20260505v.docx", "ETST.docx"]
    }

    /// Resolves a real-template fixture `name` under the gate directory
    /// (`dirOverride` when given — for tests injecting a temp dir without
    /// mutating process env — otherwise `$MACDOC_TEMPLATE_DIR`). Throws
    /// `XCTSkip` when the gate is unset or the file is missing, so the calling
    /// test skips loudly rather than failing on CI.
    static func missingGateMessage(for name: String) -> String {
        "set MACDOC_TEMPLATE_DIR to a directory containing '\(name)' to run this test"
    }

    static func missingFixtureMessage(for name: String) -> String {
        "fixture '\(name)' not found under MACDOC_TEMPLATE_DIR"
    }

    static func requireTemplate(_ name: String, dirOverride: String? = nil) throws -> URL {
        guard let dir = dirOverride ?? ProcessInfo.processInfo.environment["MACDOC_TEMPLATE_DIR"] else {
            throw XCTSkip(missingGateMessage(for: name))
        }
        let url = URL(fileURLWithPath: dir).appendingPathComponent(name)
        guard FileManager.default.fileExists(atPath: url.path) else {
            throw XCTSkip(missingFixtureMessage(for: name))
        }
        return url
    }

    /// Resolves the first of `names` that exists under the gate directory —
    /// for callers with a primary fixture plus fallbacks (#162,
    /// `V1BenchmarkTests`' "largest real-world fixture, else a smaller one"
    /// shape). `XCTSkip`s (gate unset, or none of `names` present) exactly
    /// like `requireTemplate`.
    static func requireFirstAvailable(_ names: [String], dirOverride: String? = nil) throws -> URL {
        guard let dir = dirOverride ?? ProcessInfo.processInfo.environment["MACDOC_TEMPLATE_DIR"] else {
            throw XCTSkip(missingGateMessage(for: names.first ?? "<fixture>"))
        }
        for name in names {
            let url = URL(fileURLWithPath: dir).appendingPathComponent(name)
            if FileManager.default.fileExists(atPath: url.path) {
                return url
            }
        }
        throw XCTSkip("none of \(names) found under MACDOC_TEMPLATE_DIR")
    }
}
