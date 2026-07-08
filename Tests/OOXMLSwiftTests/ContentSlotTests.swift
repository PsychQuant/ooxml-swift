// ContentSlotTests.swift
// format-alignment-engine Phase D tasks 4.1 (slot designation → parameterized
// script) + 4.2 (no-designation invariant) — `template-content-slots`
// capability: strict mode, explicit designation only, Swift function
// parameters per design Q2.

import XCTest
@testable import OOXMLSwift

final class ContentSlotTests: XCTestCase {

    /// Builds the reference package and returns (parts, reverse result).
    private func makeReference() throws -> (parts: [String: Data], log: OperationLog) {
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "原文のタイトル", styleId: "Title", paraId: "P1")),
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "原文の本文です。", paraId: "P2")),
            .setSectionProperties(at: nil, section: SectionPayload(
                pageWidth: 11906, pageHeight: 16838, marginTop: 1985, marginRight: 1701,
                marginBottom: 1701, marginLeft: 1701, columnCount: 2, columnSpace: 425)),
        ])
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("slot-ref-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try doc.writeAuthoringPackage(to: url)
        let parts = try RawPartChannel.readAllParts(from: url)
        let result = try ReverseExtractor.reverse(parts: parts)
        return (parts, result.log)
    }

    private func execute(script: String) throws -> [String: Data] {
        let log = try ScriptImporter.parse(source: script)
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: log.entries.map(\.op))
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("slot-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try doc.writeAuthoringPackage(to: url)
        return try RawPartChannel.readAllParts(from: url)
    }

    // MARK: - Task 4.1: spec scenario "title and body slots"

    /// Slotted script with the DEFAULT call-site arguments (the extracted
    /// content) reproduces the reference byte-equal — slots change nothing
    /// until the caller substitutes values.
    func testSlottedScriptWithDefaultsRebuildsByteEqual() throws {
        let (reference, log) = try makeReference()
        let script = try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "title", paraId: "P1"),
            SlotDesignation(name: "body", paraId: "P2"),
        ])
        XCTAssertTrue(script.contains("func makeDocument("), "parameterized form expected")
        XCTAssertTrue(script.contains("title: String,"), "slot parameter expected")

        let rebuilt = try execute(script: script)
        XCTAssertTrue(PartFidelity.stageB(reference: reference, rebuilt: rebuilt),
                      "default arguments must reproduce the reference byte-equal")
    }

    /// Executing with NEW content puts the new text in the designated
    /// positions while formatting (styles, sections, sibling parts) stays
    /// as extracted.
    func testSlottedScriptWithNewContentKeepsFormatting() throws {
        let (reference, log) = try makeReference()
        var script = try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "title", paraId: "P1"),
            SlotDesignation(name: "body", paraId: "P2"),
        ])
        // Caller provides new content at the call site (Swift argument form).
        script = script.replacingOccurrences(
            of: "    title: \"原文のタイトル\",", with: "    title: \"新計畫紀錄\",")
        script = script.replacingOccurrences(
            of: "    body: \"原文の本文です。\"", with: "    body: \"新しい本文内容。\"")
        XCTAssertFalse(script.contains("原文のタイトル"), "old title must be substituted")

        let rebuilt = try execute(script: script)

        // Slot positions carry the new content.
        let docXML = String(decoding: rebuilt["word/document.xml"] ?? Data(), as: UTF8.self)
        XCTAssertTrue(docXML.contains("新計畫紀錄"), "title slot must carry new text")
        XCTAssertTrue(docXML.contains("新しい本文内容。"), "body slot must carry new text")
        XCTAssertFalse(docXML.contains("原文のタイトル"))

        // Formatting intact: Title style + two-column section survive.
        XCTAssertTrue(docXML.contains("w:val=\"Title\""), "pStyle must survive")
        XCTAssertTrue(docXML.contains("w:num=\"2\""), "two-column sectPr must survive")

        // Every non-document part is untouched (byte-equal to the reference).
        for (path, bytes) in reference where path != "word/document.xml" {
            XCTAssertEqual(rebuilt[path], bytes, "non-slot part \(path) must stay as extracted")
        }
    }

    /// Strict mode: designating a nonexistent paragraph fails loudly.
    func testUnknownParaIdFailsLoudly() throws {
        let (_, log) = try makeReference()
        XCTAssertThrowsError(try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "title", paraId: "NOPE"),
        ])) { error in
            guard case TranscodeError.slotDesignationFailure(let name, _) = error else {
                return XCTFail("expected slotDesignationFailure, got \(error)")
            }
            XCTAssertEqual(name, "title")
        }
    }

    /// Strict mode: invalid slot names and duplicates fail loudly.
    func testInvalidSlotNamesFailLoudly() throws {
        let (_, log) = try makeReference()
        XCTAssertThrowsError(try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "Title", paraId: "P1"),   // uppercase start
        ]))
        XCTAssertThrowsError(try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "t", paraId: "P1"),
            SlotDesignation(name: "t", paraId: "P2"),        // duplicate name
        ]))
    }

    // MARK: - Task 4.2: no-designation invariant

    /// A script produced without any slot designation is the canonical form
    /// and reproduces the reference byte-equal (Stage B) — no substitution
    /// points.
    func testNoDesignationReproducesByteEqual() throws {
        let (reference, log) = try makeReference()
        let script = try ScriptExporter.exportSwift(log: log, slots: [])
        XCTAssertFalse(script.contains("makeDocument"), "no-slot script is the canonical form")

        let rebuilt = try execute(script: script)
        XCTAssertTrue(PartFidelity.stageB(reference: reference, rebuilt: rebuilt),
                      "no-designation script must rebuild Stage B byte-equal")
    }
}
