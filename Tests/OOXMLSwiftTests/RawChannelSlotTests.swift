// RawChannelSlotTests.swift
// raw-channel-slot-support — slots on documents whose word/document.xml rides
// the raw channel (PsychQuant/macdoc#171).
//
// A table-bearing document fails the DSL upgrade trial, so its entire
// document.xml enters the log as one `.carryPart` op with no paragraph-level
// ops. These tests pin the third slot form: designation falls back to
// scanning the carried XML for `w14:paraId`, the script carries a
// `// @slot-raw <name> <paraId>` directive, substitution is run-level surgery
// at import time, and an all-default execution leaves the part untouched
// (identity shortcut) so Stage B byte-equality holds unchanged.

import XCTest
@testable import OOXMLSwift

final class RawChannelSlotTests: XCTestCase {

    /// Table-bearing document.xml: the `<w:tbl>` forces the raw channel; two
    /// body paragraphs carry paraIds. P2 has two runs with distinct rPr — the
    /// longer run ("申請修正第") is dominant; the short underlined run ("1")
    /// exercises the collapse-to-dominant-run semantic.
    private static let tableDocumentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="AAAA1111"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>表單標題</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tr><w:tc><w:p w14:paraId="CCCC3333"><w:r><w:t>儲存格</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p w14:paraId="BBBB2222"><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri"/></w:rPr><w:t>申請修正第</w:t></w:r><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>1</w:t></w:r><w:r><w:t>次</w:t></w:r></w:p><w:sectPr/></w:body></w:document>
        """

    /// Minimal complete package with the table-bearing document.xml swapped
    /// in. The base package supplies content types / rels so the rebuilt
    /// package round-trips; carryPart replay never parses the swapped XML, so
    /// byte-fidelity is exercised end to end.
    private func makeTableReference(documentXML: String = tableDocumentXML) throws -> [String: Data] {
        var base = WordDocument.emptyAuthoringDocument()
        try base.apply(operations: [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(text: "seed", paraId: "SEED0001")),
        ])
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("rawslot-ref-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try base.writeAuthoringPackage(to: url)
        var parts = try RawPartChannel.readAllParts(from: url)
        parts["word/document.xml"] = Data(documentXML.utf8)
        return parts
    }

    private func reverseExpectingRaw(_ parts: [String: Data]) throws -> OperationLog {
        let result = try ReverseExtractor.reverse(parts: parts)
        XCTAssertFalse(result.dslParts.contains("word/document.xml"),
                       "table-bearing document.xml must ride the raw channel")
        return result.log
    }

    private func execute(script: String) throws -> [String: Data] {
        let log = try ScriptImporter.parse(source: script)
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(log: log)
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("rawslot-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try doc.writeAuthoringPackage(to: url)
        return try RawPartChannel.readAllParts(from: url)
    }

    /// Excises the `<w:p …paraId…>…</w:p>` fragment of the designated
    /// paragraph so reference and output can be compared "everything except
    /// the designated paragraph". The collapse-to-dominant-run semantic
    /// changes run structure INSIDE the paragraph by design, so the honest
    /// invariant is byte-identity of the remainders.
    private func excisingParagraph(_ xml: String, paraId: String) -> String {
        guard let range = RawChannelSlotSurgery.paragraphFragmentRange(paraId: paraId, in: xml) else {
            return xml
        }
        var out = xml
        out.removeSubrange(range)
        return out
    }

    // MARK: - Designation (tasks 1.1–1.3, 2.1)

    /// Slot designation on a raw-channel document succeeds, emits the
    /// `// @slot-raw` directive, and the call-site default carries the
    /// paragraph's concatenated text.
    func testRawChannelSlotExportEmitsDirectiveAndDefault() throws {
        let log = try reverseExpectingRaw(try makeTableReference())
        let script = try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "amendment", paraId: "BBBB2222"),
        ])
        XCTAssertTrue(script.contains("// @slot-raw amendment BBBB2222"),
                      "raw-channel slot must emit a // @slot-raw directive")
        XCTAssertTrue(script.contains("func makeDocument("), "parameterized form expected")
        XCTAssertTrue(script.contains("amendment: String"), "slot parameter expected")
        XCTAssertTrue(script.contains("amendment: \"申請修正第1次\""),
                      "call-site default must carry the paragraph's concatenated text")
    }

    /// A paraId absent from both the DSL log and the raw part refuses with a
    /// reason naming both searched domains.
    func testRawChannelSlotUnknownParaIdNamesBothDomains() throws {
        let log = try reverseExpectingRaw(try makeTableReference())
        XCTAssertThrowsError(try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "ghost", paraId: "DEAD0000"),
        ])) { error in
            guard case TranscodeError.slotDesignationFailure(_, let reason) = error else {
                return XCTFail("expected slotDesignationFailure, got \(error)")
            }
            XCTAssertTrue(reason.contains("log"), "reason must name the DSL log domain: \(reason)")
            XCTAssertTrue(reason.contains("raw"), "reason must name the raw part domain: \(reason)")
        }
    }

    /// A paraId occurring more than once in the carried XML refuses loudly —
    /// no first-match guessing on an official form.
    func testRawChannelSlotDuplicateParaIdRefuses() throws {
        let duplicated = Self.tableDocumentXML.replacingOccurrences(
            of: "w14:paraId=\"AAAA1111\"", with: "w14:paraId=\"BBBB2222\"")
        let log = try reverseExpectingRaw(try makeTableReference(documentXML: duplicated))
        XCTAssertThrowsError(try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "amendment", paraId: "BBBB2222"),
        ])) { error in
            guard case TranscodeError.slotDesignationFailure(_, let reason) = error else {
                return XCTFail("expected slotDesignationFailure, got \(error)")
            }
            XCTAssertTrue(reason.contains("2"), "reason must name the occurrence count: \(reason)")
        }
    }

    // MARK: - Execution (tasks 2.2, 3.1–3.2)

    /// All-default execution leaves the carried part untouched (identity
    /// shortcut) — Stage B byte-equality holds with zero verification changes.
    func testRawChannelSlotDefaultRebuildsByteEqual() throws {
        let reference = try makeTableReference()
        let log = try reverseExpectingRaw(reference)
        let script = try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "amendment", paraId: "BBBB2222"),
        ])
        let rebuilt = try execute(script: script)
        XCTAssertTrue(PartFidelity.stageB(reference: reference, rebuilt: rebuilt),
                      "default argument must reproduce the reference byte-equal")
    }

    /// A substituted value changes only the designated paragraph: pPr is
    /// preserved, the runs collapse to one run carrying the dominant run's
    /// rPr, and stripping text makes the part byte-identical again.
    func testRawChannelSlotSubstitutesOnlyDesignatedParagraph() throws {
        let reference = try makeTableReference()
        let log = try reverseExpectingRaw(reference)
        var script = try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "amendment", paraId: "BBBB2222"),
        ])
        script = script.replacingOccurrences(
            of: "amendment: \"申請修正第1次\"",
            with: "amendment: \"申請修正第2次\"")
        let rebuilt = try execute(script: script)

        let refXML = String(data: reference["word/document.xml"]!, encoding: .utf8)!
        let outXML = String(data: rebuilt["word/document.xml"]!, encoding: .utf8)!
        XCTAssertTrue(outXML.contains("申請修正第2次"), "substituted text must be present")
        XCTAssertFalse(outXML.contains("<w:u w:val=\"single\"/>"),
                       "short run's rPr must not survive the collapse")
        XCTAssertTrue(outXML.contains("<w:jc w:val=\"both\"/>"), "pPr must be preserved")
        XCTAssertTrue(outXML.contains("<w:rFonts w:ascii=\"Calibri\"/>"),
                      "dominant run's rPr must be applied to the collapsed run")
        // Everything outside the designated paragraph is untouched: the
        // reference with BBBB2222's paragraph excised must appear verbatim.
        XCTAssertTrue(outXML.contains("<w:p w14:paraId=\"AAAA1111\">"),
                      "sibling paragraph opening must be untouched")
        XCTAssertTrue(outXML.contains("表單標題"), "sibling paragraph text must be untouched")
        XCTAssertTrue(outXML.contains("<w:tbl>"), "table must be untouched")
        // Other parts byte-identical.
        for (path, bytes) in reference where path != "word/document.xml" {
            XCTAssertEqual(rebuilt[path], bytes, "non-document part \(path) must be byte-identical")
        }
        // Excise comparison: byte-identical outside the designated paragraph.
        XCTAssertEqual(excisingParagraph(refXML, paraId: "BBBB2222"),
                       excisingParagraph(outXML, paraId: "BBBB2222"),
                       "everything outside the designated paragraph must be byte-identical")
    }

    // MARK: - End-to-end on the REC-O-01 fixture (tasks 4.1–4.2)

    private func recFixtureURL() throws -> URL {
        // packages/ooxml-swift/Tests/OOXMLSwiftTests/ → repo root /test-files/
        var dir = URL(fileURLWithPath: #filePath).deletingLastPathComponent()
        for _ in 0..<4 { dir.deleteLastPathComponent() }
        let url = dir.appendingPathComponent("test-files/REC-O-01-新案審查送件核對單-20250220公告.docx")
        guard FileManager.default.fileExists(atPath: url.path) else {
            throw XCTSkip("REC-O-01 fixture not present at \(url.path)")
        }
        return url
    }

    /// The real table-bearing official form: designation succeeds and the
    /// all-default execution passes the byte-equal acceptance unchanged.
    func testRECFixtureDefaultReplayByteEqual() throws {
        let reference = try RawPartChannel.readAllParts(from: try recFixtureURL())
        let result = try ReverseExtractor.reverse(parts: reference)
        XCTAssertFalse(result.dslParts.contains("word/document.xml"),
                       "REC-O-01 must ride the raw channel")
        let xml = String(data: reference["word/document.xml"]!, encoding: .utf8)!
        guard let idRange = xml.range(of: #"w14:paraId=""#),
              let idEnd = xml[idRange.upperBound...].firstIndex(of: "\"") else {
            return XCTFail("fixture must carry at least one w14:paraId")
        }
        let paraId = String(xml[idRange.upperBound..<idEnd])
        let script = try ScriptExporter.exportSwift(log: result.log, slots: [
            SlotDesignation(name: "field", paraId: paraId),
        ])
        XCTAssertTrue(script.contains("// @slot-raw field \(paraId)"))
        let rebuilt = try execute(script: script)
        XCTAssertTrue(PartFidelity.stageB(reference: reference, rebuilt: rebuilt),
                      "REC-O-01 all-default replay must be byte-equal")
    }

    /// Substituting a new value on the real form changes exactly the
    /// designated paragraph (strip-text comparison) and carries the new text.
    func testRECFixtureSubstitutionChangesOnlyDesignatedParagraph() throws {
        let reference = try RawPartChannel.readAllParts(from: try recFixtureURL())
        let result = try ReverseExtractor.reverse(parts: reference)
        let xml = String(data: reference["word/document.xml"]!, encoding: .utf8)!
        guard let idRange = xml.range(of: #"w14:paraId=""#),
              let idEnd = xml[idRange.upperBound...].firstIndex(of: "\"") else {
            return XCTFail("fixture must carry at least one w14:paraId")
        }
        let paraId = String(xml[idRange.upperBound..<idEnd])
        var script = try ScriptExporter.exportSwift(log: result.log, slots: [
            SlotDesignation(name: "field", paraId: paraId),
        ])
        guard let defaultRange = script.range(of: #"field: ""#),
              let defaultEnd = script[defaultRange.upperBound...].firstIndex(of: "\"") else {
            return XCTFail("script must carry a call-site default for the slot")
        }
        let defaultValue = String(script[defaultRange.upperBound..<defaultEnd])
        script = script.replacingOccurrences(
            of: "field: \"\(defaultValue)\"", with: "field: \"覆核完成\"")
        let rebuilt = try execute(script: script)
        let outXML = String(data: rebuilt["word/document.xml"]!, encoding: .utf8)!
        XCTAssertTrue(outXML.contains("覆核完成"), "substituted text must be present")
        XCTAssertEqual(excisingParagraph(xml, paraId: paraId),
                       excisingParagraph(outXML, paraId: paraId),
                       "everything outside the designated paragraph must be byte-identical")
        for (path, bytes) in reference where path != "word/document.xml" {
            XCTAssertEqual(rebuilt[path], bytes, "non-document part \(path) must be byte-identical")
        }
    }
}

extension RawChannelSlotTests {

    /// A malformed `// @slot-raw` line (missing paraId) is skipped exactly
    /// like the existing `// @slot` pre-pass — parsing succeeds and the
    /// execution stays byte-equal.
    func testMalformedRawSlotDirectiveIsIgnored() throws {
        let reference = try makeTableReference()
        let log = try reverseExpectingRaw(reference)
        var script = try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "amendment", paraId: "BBBB2222"),
        ])
        script = script.replacingOccurrences(
            of: "// @slot-raw amendment BBBB2222",
            with: "// @slot-raw malformed-only-one-token")
        let rebuilt = try execute(script: script)
        XCTAssertTrue(PartFidelity.stageB(reference: reference, rebuilt: rebuilt),
                      "malformed directive must be ignored; replay stays byte-equal")
    }
}
