// ContentSlotOpLevelTests.swift
// word-canonical-forms Phase 3 task 3.2(b) — op-level slot substitution.
//
// The Phase-D slot mechanism (`template-content-slots`) only parameterized
// DSL-spellable paragraphs (`Paragraph(id){text}`). Real templates like
// 90_template_ja carry *formatted* paragraphs that ride the raw `// @op`
// escape, with their visible text in a `setRuns` op. This extends slots to
// both single- and multi-run paragraphs: the exporter emits a
// `// @slot <name> <paraId>` directive + a `makeDocument` parameter; the
// importer substitutes text at the call site while keeping every run and
// formatting attribute intact.

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

    /// A multi-run slot keeps its original run partition byte-equal when the
    /// default binding is used. When callers supply new text, only run text is
    /// rewritten: the first non-whitespace run is the deterministic carrier,
    /// and every run's formatting payload stays intact.
    func testOpLevelMultiRunSlotPreservesFormatting() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "", paraId: "P1", indentFirstLineChars: 100,
                paragraphMarkRun: RunPayload(text: "", sizeHalfPoints: 36))),
            .setRuns(target: ElementID(rawString: "w14:paraId=P1"), runs: [
                RunPayload(text: "　", preserveSpace: true),
                RunPayload(text: "原文標題", bold: true, fontEastAsia: "ＭＳ ゴシック"),
                RunPayload(text: "（説明）", color: "FF0000", sizeHalfPoints: 18),
            ]),
        ])
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("oplevel-multi-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try doc.writeAuthoringPackage(to: url)
        let reference = try RawPartChannel.readAllParts(from: url)
        let log = try ReverseExtractor.reverse(parts: reference).log

        var script = try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "heading", paraId: "P1"),
        ])
        XCTAssertTrue(script.contains("heading: \"　原文標題（説明）\""),
                      "the slot default must be the visible text joined in run order")
        XCTAssertTrue(PartFidelity.stageB(reference: reference,
                                          rebuilt: try execute(script: script)),
                      "the default binding must not repartition runs")

        script = script.replacingOccurrences(
            of: "    heading: \"　原文標題（説明）\"",
            with: "    heading: \"新しい標題\"")
        let substituted = try ScriptImporter.parse(source: script)
        let runs = try XCTUnwrap(substituted.entries.compactMap { entry -> [RunPayload]? in
            guard case .setRuns(let target, let runs) = entry.op,
                  target.raw == "w14:paraId=P1" else { return nil }
            return runs
        }.first)
        XCTAssertEqual(runs.map(\.text), ["", "新しい標題", ""],
                       "new text must use the first non-whitespace carrier run")

        let originalRuns = try XCTUnwrap(log.entries.compactMap { entry -> [RunPayload]? in
            guard case .setRuns(let target, let runs) = entry.op,
                  target.raw == "w14:paraId=P1" else { return nil }
            return runs
        }.first)
        XCTAssertEqual(runs.map { var run = $0; run.text = ""; return run },
                       originalRuns.map { var run = $0; run.text = ""; return run },
                       "slot substitution must preserve every run formatting field")

        let rebuilt = try execute(script: script)
        let documentXML = String(decoding: rebuilt["word/document.xml"] ?? Data(), as: UTF8.self)
        XCTAssertTrue(documentXML.contains("新しい標題"))
        XCTAssertFalse(documentXML.contains("原文標題"))
        XCTAssertFalse(documentXML.contains("（説明）"))
        for (path, bytes) in reference where path != "word/document.xml" {
            XCTAssertEqual(rebuilt[path], bytes, "non-slot part \(path) must stay byte-equal")
        }
    }

    /// Strict mode still rejects a formatted paragraph when neither setRuns
    /// nor appendParagraph carries any text; multi-run support must not turn
    /// an empty target into a silent no-op slot.
    func testOpLevelSlotWithoutTextTargetFailsLoudly() throws {
        var log = OperationLog()
        log.append(.appendParagraph(in: nil, paragraph: ParagraphPayload(
            text: "", paraId: "P1", indentFirstLineChars: 100,
            paragraphMarkRun: RunPayload(text: "", sizeHalfPoints: 36))), source: .word)
        log.append(.setRuns(target: ElementID(rawString: "w14:paraId=P1"), runs: [
            RunPayload(text: ""), RunPayload(text: "", bold: true),
        ]), source: .word)

        XCTAssertThrowsError(try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "heading", paraId: "P1"),
        ])) { error in
            guard case TranscodeError.slotDesignationFailure(let name, _) = error else {
                return XCTFail("expected slotDesignationFailure, got \(error)")
            }
            XCTAssertEqual(name, "heading")
        }
    }

    /// Multiple setRuns operations for one paragraph are sequential writes;
    /// only the final occurrence is the visible state and therefore the slot
    /// default and substitution target. Earlier history must remain unchanged.
    func testOpLevelSlotTargetsOnlyFinalSetRunsOccurrence() throws {
        var log = OperationLog()
        log.append(.appendParagraph(in: nil, paragraph: ParagraphPayload(
            text: "", paraId: "P1", indentFirstLineChars: 100,
            paragraphMarkRun: RunPayload(text: "", sizeHalfPoints: 36))), source: .word)
        log.append(.setRuns(target: ElementID(rawString: "w14:paraId=P1"), runs: [
            RunPayload(text: "歷史文字", italic: true),
        ]), source: .word)
        log.append(.setRuns(target: ElementID(rawString: "w14:paraId=P1"), runs: [
            RunPayload(text: "目前"), RunPayload(text: "內容", bold: true),
        ]), source: .word)

        var script = try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "heading", paraId: "P1"),
        ])
        XCTAssertTrue(script.contains("heading: \"目前內容\""),
                      "the final setRuns operation must supply the default")
        XCTAssertEqual(try ScriptImporter.parse(source: script).entries.map(\.op),
                       log.entries.map(\.op),
                       "default binding must preserve all operation history")

        script = script.replacingOccurrences(
            of: "    heading: \"目前內容\"",
            with: "    heading: \"替換內容\"")
        let parsed = try ScriptImporter.parse(source: script)
        let occurrences = parsed.entries.compactMap { entry -> [RunPayload]? in
            guard case .setRuns(let target, let runs) = entry.op,
                  target.raw == "w14:paraId=P1" else { return nil }
            return runs
        }
        XCTAssertEqual(occurrences.map { $0.map(\.text) },
                       [["歷史文字"], ["替換內容", ""]],
                       "only the final setRuns occurrence may be substituted")
        XCTAssertTrue(occurrences[0][0].italic == true,
                      "earlier operation payload must remain untouched")
        XCTAssertTrue(occurrences[1][1].bold == true,
                      "final run formatting must remain intact")
    }

    /// A replacement with XML boundary whitespace must opt the carrier text
    /// into xml:space preservation even when the original run did not need it.
    func testOpLevelSlotPreservesBoundaryWhitespace() throws {
        var log = OperationLog()
        log.append(.appendParagraph(in: nil, paragraph: ParagraphPayload(
            text: "", paraId: "P1", indentFirstLineChars: 100,
            paragraphMarkRun: RunPayload(text: "", sizeHalfPoints: 36))), source: .word)
        log.append(.setRuns(target: ElementID(rawString: "w14:paraId=P1"), runs: [
            RunPayload(text: "Original", bold: true),
        ]), source: .word)

        let template = try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "heading", paraId: "P1"),
        ])
        for value in [" padded ", "\tpadded", "padded\n", "\r\npadded"] {
            let script = template.replacingOccurrences(
                of: "    heading: \"Original\"",
                with: "    heading: \(ScriptExporter.quote(value))")
            let parsed = try ScriptImporter.parse(source: script)
            let run = try XCTUnwrap(parsed.entries.compactMap { entry -> RunPayload? in
                guard case .setRuns(let target, let runs) = entry.op,
                      target.raw == "w14:paraId=P1" else { return nil }
                return runs.first
            }.first)
            XCTAssertEqual(run.text, value)
            XCTAssertEqual(run.preserveSpace, true,
                           "boundary whitespace requires xml:space=preserve: \(value.debugDescription)")
        }
    }
}
