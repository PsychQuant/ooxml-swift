import XCTest
import Foundation
@testable import WordDSLSwift
import OOXMLSwift

/// word-aligned-state-sync Phase 4 task 5.3 — the WordDSLSwift result-builder
/// runtime: op emission semantics + `save(to:)` atomic three-file write.
/// Compiling this file IS part of the verification: the inline DSL bodies
/// below use exactly the fixture corpus syntax (02a / 05), proving the
/// canonical grammar subset builds under the real Swift compiler.
final class WordDSLRuntimeTests: XCTestCase {

    // MARK: - Fixture 02a syntax compiles + emission semantics

    func testPlainStringFixtureSyntaxCompilesAndEmits() throws {
        // Body mirrors Fixtures/mdocx/02a-plain-string verbatim.
        let document = WordDocument {
            Section(id: "main") {
                Paragraph(id: "p1") {
                    "本章探討"
                    "意識本質"
                    "的議題。"
                }
            }
        }
        let log = try document.buildLog()

        XCTAssertEqual(log.entries.count, 1)
        guard case .appendParagraph(let container, let p) = log.entries[0].op else {
            return XCTFail("expected appendParagraph")
        }
        XCTAssertNil(container)
        XCTAssertEqual(p.paraId, "p1")
        XCTAssertEqual(p.text, "本章探討意識本質的議題。",
                       "implicit String literals join in declaration order")
    }

    // MARK: - Fixture 05 semantics: define-on-first-use

    func testStyleDefineOnFirstUse() throws {
        let document = WordDocument {
            Section(id: "main") {
                Paragraph(id: "h1", style: .heading1) { "Title" }
                Paragraph(id: "h2", style: .heading1) { "Subtitle" }
            }
        }
        let log = try document.buildLog()

        let defineCount = log.entries.filter {
            if case .defineStyle = $0.op { return true } else { return false }
        }.count
        XCTAssertEqual(defineCount, 1,
                       "two references to the same WordStyle emit exactly one defineStyle")
        guard case .defineStyle(let payload) = log.entries[0].op else {
            return XCTFail("defineStyle must precede the first referencing paragraph")
        }
        XCTAssertEqual(payload.styleId, "Heading1")
        guard case .appendParagraph(_, let p1) = log.entries[1].op else {
            return XCTFail("expected appendParagraph after defineStyle")
        }
        XCTAssertEqual(p1.styleId, "Heading1")
    }

    // MARK: - Formatted runs + atoms

    func testFormattedRunsEmitSetRuns() throws {
        let document = WordDocument {
            Section(id: "main") {
                Paragraph(id: "p1") {
                    "本章探討"
                    Run("意識本質", bold: true, italic: true, color: "663300")
                    "的議題。"
                }
            }
        }
        let log = try document.buildLog()

        XCTAssertEqual(log.entries.count, 2, "appendParagraph + setRuns")
        guard case .setRuns(let target, let runs) = log.entries[1].op else {
            return XCTFail("expected setRuns")
        }
        XCTAssertEqual(target.raw, "w14:paraId=p1")
        XCTAssertEqual(runs.count, 3)
        XCTAssertNil(runs[0].bold)
        XCTAssertEqual(runs[1].bold, true)
        XCTAssertEqual(runs[1].color, "663300")
    }

    func testAtomsEmitParagraphTargetedOps() throws {
        let document = WordDocument {
            Section(id: "main") {
                Paragraph(id: "p1") {
                    "Header"
                    Tab()
                    Break()
                }
            }
        }
        let log = try document.buildLog()
        XCTAssertEqual(log.entries.count, 3, "appendParagraph + insertTab + insertBreak")
        guard case .insertTab(let t) = log.entries[1].op else {
            return XCTFail("expected insertTab")
        }
        XCTAssertEqual(t.raw, "w14:paraId=p1")
    }

    // MARK: - save(to:) three-file write (mdocx-grammar requirement)

    private func scratchDir() throws -> URL {
        let dir = FileManager.default.temporaryDirectory
            .appendingPathComponent("dsl-save-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        return dir
    }

    func testEmptyDocumentSaveProducesValidDocx() throws {
        // Spec scenario: WordDocument { }.save(to:) → valid docx with the
        // four mandatory parts, readable back.
        let dir = try scratchDir()
        defer { try? FileManager.default.removeItem(at: dir) }
        let url = dir.appendingPathComponent("empty.docx")

        try WordDocument { }.save(to: url)

        XCTAssertTrue(FileManager.default.fileExists(atPath: url.path))
        XCTAssertTrue(FileManager.default.fileExists(
            atPath: SidecarStore.oplogURL(for: url).path))
        XCTAssertTrue(FileManager.default.fileExists(
            atPath: SidecarStore.snapshotURL(for: url).path))

        let reread = try DocxReader.read(from: url)
        XCTAssertTrue(reread.body.children.isEmpty, "empty document body")
    }

    func testScriptBuiltDocxRoundTripsWithParaIdPreserved() throws {
        let dir = try scratchDir()
        defer { try? FileManager.default.removeItem(at: dir) }
        let url = dir.appendingPathComponent("script.docx")

        let document = WordDocument {
            Section(id: "main") {
                Paragraph(id: "ch1-intro") { "本章探討" }
                Paragraph(id: "h1", style: .heading1) { "Title" }
            }
        }
        try document.save(to: url)

        let reread = try DocxReader.read(from: url, wireTreeBackedViews: true)
        XCTAssertEqual(reread.body.children.count, 2)
        if case .paragraph(let p) = reread.body.children.first {
            XCTAssertEqual(p.text, "本章探討")
            XCTAssertEqual(p.elementID?.raw, "w14:paraId=ch1-intro",
                           "explicit DSL id must persist as w14:paraId in the docx")
        } else {
            XCTFail("expected paragraph")
        }
        // Op log sidecar restores the authored history.
        let restored = try SidecarStore.loadLog(alongside: url)
        XCTAssertEqual(restored?.entries.isEmpty, false)
    }

    func testSaveRefusesWhileWordHoldsLock() throws {
        let dir = try scratchDir()
        defer { try? FileManager.default.removeItem(at: dir) }
        let url = dir.appendingPathComponent("locked.docx")
        try Data("lock".utf8).write(to: dir.appendingPathComponent("~$locked.docx"))

        XCTAssertThrowsError(try WordDocument { }.save(to: url)) { error in
            guard case SyncError.fileLockedByWord = error else {
                return XCTFail("expected fileLockedByWord, got \(error)")
            }
        }
        XCTAssertFalse(FileManager.default.fileExists(atPath: url.path),
                       "no partial output while locked")
    }

    func testSaveFailureRollsBackAllThreeFiles() throws {
        // Spec scenario "failure during second-file write rolls back":
        // pre-create the oplog sidecar path as a DIRECTORY so SidecarStore's
        // file write fails after the docx write succeeded.
        let dir = try scratchDir()
        defer { try? FileManager.default.removeItem(at: dir) }
        let url = dir.appendingPathComponent("rollback.docx")
        try FileManager.default.createDirectory(
            at: SidecarStore.oplogURL(for: url), withIntermediateDirectories: true)

        XCTAssertThrowsError(
            try WordDocument {
                Section(id: "main") { Paragraph(id: "p1") { "x" } }
            }.save(to: url))

        XCTAssertFalse(FileManager.default.fileExists(atPath: url.path),
                       "docx must be rolled back when a sidecar write fails")
        XCTAssertFalse(FileManager.default.fileExists(
            atPath: SidecarStore.snapshotURL(for: url).path))
    }

    // MARK: - Real fixture file parses through the transcoder importer

    func testRealFixture02aParsesThroughImporter() throws {
        let fixtureURL = URL(fileURLWithPath: #filePath)
            .deletingLastPathComponent()            // WordDSLSwiftTests
            .deletingLastPathComponent()            // Tests
            .appendingPathComponent("OOXMLSwiftTests/Fixtures/mdocx/02a-plain-string/plain-string.mdocx.swift")
        guard FileManager.default.fileExists(atPath: fixtureURL.path) else {
            throw XCTSkip("fixture corpus not present at \(fixtureURL.path)")
        }
        let source = try String(contentsOf: fixtureURL, encoding: .utf8)
        let log = try ScriptImporter.parse(source: source)

        XCTAssertEqual(log.entries.count, 1)
        guard case .appendParagraph(_, let p) = log.entries[0].op else {
            return XCTFail("expected appendParagraph from the real fixture")
        }
        XCTAssertEqual(p.paraId, "p1")
        XCTAssertEqual(p.text, "本章探討意識本質的議題。")

        // Cross-check: the fixture's semantics == the inline DSL's semantics.
        let inline = try WordDocument {
            Section(id: "main") {
                Paragraph(id: "p1") {
                    "本章探討"
                    "意識本質"
                    "的議題。"
                }
            }
        }.buildLog()
        XCTAssertEqual(inline.entries[0].op, log.entries[0].op,
                       "fixture parse and DSL execution agree")
    }
}

extension WordDSLRuntimeTests {

    /// 5.3 "round-trip through Word save preserves structural equivalence" —
    /// live Microsoft Word opens the script-built docx, saves without edits,
    /// and ooxml-swift re-reads it. Gated behind RUN_WORD_INTEGRATION=1.
    func testScriptBuiltDocxSurvivesLiveWordResave() throws {
        guard ProcessInfo.processInfo.environment["RUN_WORD_INTEGRATION"] == "1" else {
            throw XCTSkip("live Word integration gated behind RUN_WORD_INTEGRATION=1")
        }
        guard FileManager.default.fileExists(atPath: "/Applications/Microsoft Word.app") else {
            throw XCTSkip("Microsoft Word not installed")
        }
        let dir = FileManager.default.homeDirectoryForCurrentUser
            .appendingPathComponent(".cache/mdocx-word-resave-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: dir) }
        let url = dir.appendingPathComponent("script-output.docx")

        try WordDocument {
            Section(id: "main") {
                Paragraph(id: "ch1-intro") { "本章探討" }
                Paragraph(id: "h1", style: .heading1) { "Title" }
            }
        }.save(to: url)

        let script = """
        tell application "Microsoft Word"
            open POSIX file "\(url.path)"
            save active document
            close active document saving no
        end tell
        """
        let process = Process()
        process.executableURL = URL(fileURLWithPath: "/usr/bin/osascript")
        process.arguments = ["-e", script]
        let pipe = Pipe(); process.standardOutput = pipe; process.standardError = pipe
        try process.run(); process.waitUntilExit()
        guard process.terminationStatus == 0 else {
            let out = String(decoding: pipe.fileHandleForReading.readDataToEndOfFile(), as: UTF8.self)
            throw XCTSkip("osascript could not drive Word: \(out)")
        }

        // Word rewrote the package (rsids, settings, theme, …) — structural
        // equivalence of the typed views is the acceptance bar, and no
        // Word-side rejection dialog is implied by the successful save.
        let reread = try DocxReader.read(from: url, wireTreeBackedViews: true)
        XCTAssertEqual(reread.body.children.count, 2,
                       "both paragraphs survive a live Word re-save")
        if case .paragraph(let p) = reread.body.children.first {
            XCTAssertEqual(p.text, "本章探討")
        } else { XCTFail("paragraph 1 lost") }
        if case .paragraph(let p2) = reread.body.children.dropFirst().first {
            XCTAssertEqual(p2.text, "Title")
        } else { XCTFail("paragraph 2 lost") }
    }
}

// MARK: - 5.5 structural coverage (component envelope + compile-only types)

/// Fixture 07's component definition, verbatim shape.
private struct Summary: WordComponent {
    let id: String
    let body: () -> WordDSLSwift.Paragraph
    init(id: String, @WordBuilder body: @escaping () -> WordDSLSwift.Paragraph) {
        self.id = id
        self.body = body
    }
}

extension WordStyle {
    fileprivate static let summaryFrame = WordStyle(styleId: "SummaryFrame")
}

final class StructuralCoverageTests: XCTestCase {

    func testComponentEnvelopeEmission() throws {
        // Fixture 07 syntax, verbatim.
        let document = WordDocument {
            Section(id: "main") {
                Summary(id: "ch1-summary") {
                    Paragraph(id: "sum-frame", style: .summaryFrame) {
                        "note text"
                    }
                }
            }
        }
        let log = try document.buildLog()

        guard case .beginComponent(let type, let cid) = log.entries[0].op else {
            return XCTFail("first op must be beginComponent, got \(log.entries[0].op)")
        }
        XCTAssertEqual(type, "Summary")
        XCTAssertEqual(cid.raw, "ch1-summary")
        guard case .endComponent(let eid) = log.entries.last?.op else {
            return XCTFail("last op must be endComponent")
        }
        XCTAssertEqual(eid.raw, "ch1-summary")
        // Envelope contains defineStyle + appendParagraph for the body.
        XCTAssertTrue(log.entries.contains {
            if case .appendParagraph(_, let p) = $0.op { return p.paraId == "sum-frame" }
            return false
        })
    }

    func testComponentEnvelopeProducesNoOOXMLArtifact() throws {
        // mdocx-grammar scenario "component metadata produces no OOXML artifact".
        let dir = FileManager.default.temporaryDirectory
            .appendingPathComponent("component-save-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: dir) }
        let url = dir.appendingPathComponent("component.docx")

        try WordDocument {
            Section(id: "main") {
                Summary(id: "ch1-summary") {
                    Paragraph(id: "sum-frame") { "note text" }
                }
            }
        }.save(to: url)

        let reread = try DocxReader.read(from: url)
        XCTAssertEqual(reread.body.children.count, 1, "one paragraph, no component artifact")
        if case .paragraph(let p) = reread.body.children.first {
            XCTAssertEqual(p.text, "note text")
        }
    }

    func testComponentRoundTripsThroughTranscoderAsDSLForm() throws {
        let document = WordDocument {
            Section(id: "main") {
                Summary(id: "ch1-summary") {
                    Paragraph(id: "sum-frame") { "note text" }
                }
            }
        }
        let log = try document.buildLog()
        let script = ScriptExporter.exportSwift(log: log)

        XCTAssertTrue(script.contains("Summary(id: \"ch1-summary\") {"),
                      "exporter reconstructs the component invocation, got:\n\(script)")
        XCTAssertFalse(script.contains("@op {\"op_type\":\"beginComponent\""),
                       "envelope must be DSL-form, not raw escape")

        let reconstructed = try ScriptImporter.parse(source: script)
        XCTAssertEqual(reconstructed.entries.count, log.entries.count)
        for (a, b) in zip(log.entries, reconstructed.entries) {
            XCTAssertEqual(a.op, b.op)
        }
    }

    func testTableAndBookmarkAndHyperlinkTypesCompile() throws {
        // Fixtures 10a / 12 / 13a syntax — COMPILING these bodies is the
        // v0.34 acceptance for the structural types; emission channels are
        // loud-throw pending taxonomy/reducer increments.
        let tableDoc = WordDocument {
            Section(id: "main") {
                Table(id: "tbl1") {
                    TableRow(id: "tbl1-r0") {
                        TableCell(id: "tbl1-r0-c0") {
                            Paragraph(id: "tbl1-r0-c0-p0") { "cell content" }
                        }
                    }
                }
            }
        }
        XCTAssertThrowsError(try tableDoc.buildLog()) { error in
            guard case DSLEmissionError.unsupportedElement(let element, _) = error else {
                return XCTFail("expected loud DSLEmissionError, got \(error)")
            }
            XCTAssertTrue(element.contains("Table"))
        }

        let bookmarkDoc = WordDocument {
            Section(id: "main") {
                Paragraph(id: "p1") {
                    Bookmark(id: "intro_text") { "本章探討" }
                }
            }
        }
        XCTAssertThrowsError(try bookmarkDoc.buildLog())

        let hyperlinkDoc = WordDocument {
            Section(id: "main") {
                Paragraph(id: "p1") {
                    "see "
                    Hyperlink(to: .anchor("ch1-intro")) { "Chapter 1" }
                }
            }
        }
        XCTAssertThrowsError(try hyperlinkDoc.buildLog())
    }

    func testRealFixture10aAnd13aSyntaxParseability() throws {
        // The real fixture FILES for table/bookmark use containers outside
        // the importer's v0.34 canonical subset — importer must reject them
        // loudly (not silently mis-parse). DSL-form import for these rides
        // with the taxonomy increment.
        let base = URL(fileURLWithPath: #filePath)
            .deletingLastPathComponent().deletingLastPathComponent()
            .appendingPathComponent("OOXMLSwiftTests/Fixtures/mdocx")
        let table = base.appendingPathComponent("10a-table-1x1/table-1x1.mdocx.swift")
        guard FileManager.default.fileExists(atPath: table.path) else {
            throw XCTSkip("fixture corpus not present")
        }
        let source = try String(contentsOf: table, encoding: .utf8)
        XCTAssertThrowsError(try ScriptImporter.parse(source: source)) { error in
            guard case TranscodeError.unsupportedSyntax = error else {
                return XCTFail("expected unsupportedSyntax for out-of-subset container")
            }
        }
    }
}
