// ContentSlotOpLevelTests.swift
// word-canonical-forms Phase 3 task 3.2(b) — op-level slot substitution.
//
// The Phase-D slot mechanism (`template-content-slots`) only parameterized
// DSL-spellable paragraphs (`Paragraph(id){text}`). Real templates like
// 90_template_ja carry *formatted* paragraphs that ride the raw `// @op`
// escape, with their visible text in a single-run `setRuns` op. This extends
// slots to those paragraphs: the exporter emits a `// @slot <name> <paraId>`
// directive + a `makeDocument` parameter; the importer substitutes the run
// text at the call site while keeping every formatting attribute byte-equal.

import XCTest
@testable import OOXMLSwift

final class ContentSlotOpLevelTests: XCTestCase {

    /// A 90_template_ja-shaped formatted paragraph: rich pPr (first-line indent
    /// in chars + a paragraph-mark run) forces the raw `// @op` escape, and the
    /// visible text lives in a single richly-formatted run.
    private func makeFormattedReference() throws -> (parts: [String: Data], log: OperationLog) {
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "", paraId: "P1",
                indentFirstLine: 180, indentFirstLineChars: 100,
                paragraphMarkRun: RunPayload(
                    text: "", fontAscii: "Times New Roman", sizeHalfPoints: 36))),
            .setRuns(target: ElementID(rawString: "w14:paraId=P1"), runs: [RunPayload(
                text: "原文の見出し", bold: true, fontEastAsia: "ＭＳ ゴシック",
                sizeHalfPoints: 36)]),
        ])
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("oplevel-ref-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try doc.writeAuthoringPackage(to: url)
        let parts = try RawPartChannel.readAllParts(from: url)
        let result = try ReverseExtractor.reverse(parts: parts)
        // Precondition: the formatted paragraph must NOT be DSL-spellable, so
        // this test actually exercises the op-level path (not the DSL path).
        let script = ScriptExporter.exportSwift(log: result.log)
        XCTAssertTrue(script.contains("// @op"),
                      "formatted paragraph must ride the raw escape (op-level slot territory)")
        XCTAssertFalse(script.contains("Paragraph(id: \"P1\") {"),
                       "formatted paragraph must NOT get a DSL Paragraph block")
        return (parts, result.log)
    }

    private func execute(script: String) throws -> [String: Data] {
        let log = try ScriptImporter.parse(source: script)
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: log.entries.map(\.op))
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("oplevel-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try doc.writeAuthoringPackage(to: url)
        return try RawPartChannel.readAllParts(from: url)
    }

    /// The op-level slot emits a `// @slot` directive and a function parameter;
    /// with the DEFAULT call-site argument (the extracted run text) it rebuilds
    /// the reference byte-equal — the slot changes nothing until substituted.
    func testOpLevelSlotDefaultRebuildsByteEqual() throws {
        let (reference, log) = try makeFormattedReference()
        let script = try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "heading", paraId: "P1"),
        ])
        XCTAssertTrue(script.contains("func makeDocument("), "parameterized form expected")
        XCTAssertTrue(script.contains("heading: String"), "slot parameter expected")
        XCTAssertTrue(script.contains("// @slot heading P1"),
                      "op-level slot must emit a // @slot directive")
        XCTAssertTrue(script.contains("heading: \"原文の見出し\""),
                      "call-site default must carry the extracted run text")

        let rebuilt = try execute(script: script)
        XCTAssertTrue(PartFidelity.stageB(reference: reference, rebuilt: rebuilt),
                      "default argument must reproduce the reference byte-equal")
    }

    /// Substituting a NEW call-site value replaces ONLY the run text; every
    /// formatting attribute (pPr indent, rFonts, bold, sz) stays intact, and
    /// non-document parts remain byte-equal to the reference.
    func testOpLevelSlotSubstitutesRunTextKeepingFormatting() throws {
        let (reference, log) = try makeFormattedReference()
        var script = try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "heading", paraId: "P1"),
        ])
        script = script.replacingOccurrences(
            of: "    heading: \"原文の見出し\"", with: "    heading: \"新しい見出し\"")
        XCTAssertFalse(script.contains("heading: \"原文の見出し\""),
                       "old call-site value must be replaced")

        let rebuilt = try execute(script: script)
        let docXML = String(decoding: rebuilt["word/document.xml"] ?? Data(), as: UTF8.self)

        // New text present, old text gone.
        XCTAssertTrue(docXML.contains("新しい見出し"), "slot must carry new run text")
        XCTAssertFalse(docXML.contains("原文の見出し"), "old run text must be gone")

        // Formatting intact: eastAsia font + bold + the first-line-char indent.
        XCTAssertTrue(docXML.contains("ＭＳ ゴシック"), "run rFonts must survive")
        XCTAssertTrue(docXML.contains("<w:b/>"), "run bold must survive")
        XCTAssertTrue(docXML.contains("w:firstLineChars=\"100\""), "pPr indent must survive")

        // Non-document parts untouched.
        for (path, bytes) in reference where path != "word/document.xml" {
            XCTAssertEqual(rebuilt[path], bytes, "non-slot part \(path) must stay as extracted")
        }
    }

    /// A formatted paragraph with NO substitutable text target (e.g. inline
    /// markers only, no single-run setRuns) fails loudly in strict mode.
    func testOpLevelSlotWithoutTextTargetFailsLoudly() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: [
            // Two runs: no single unambiguous substitution target.
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "", paraId: "P1", indentFirstLineChars: 100,
                paragraphMarkRun: RunPayload(text: "", sizeHalfPoints: 36))),
            .setRuns(target: ElementID(rawString: "w14:paraId=P1"), runs: [
                RunPayload(text: "前半"), RunPayload(text: "後半", bold: true),
            ]),
        ])
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("oplevel-multi-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try doc.writeAuthoringPackage(to: url)
        let parts = try RawPartChannel.readAllParts(from: url)
        let log = try ReverseExtractor.reverse(parts: parts).log

        XCTAssertThrowsError(try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "heading", paraId: "P1"),
        ])) { error in
            guard case TranscodeError.slotDesignationFailure(let name, _) = error else {
                return XCTFail("expected slotDesignationFailure, got \(error)")
            }
            XCTAssertEqual(name, "heading")
        }
    }
}
