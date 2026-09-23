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
        guard case .unique(let span) = RawChannelSlotSurgery.locate(paraId: paraId, in: xml) else {
            return xml
        }
        var out = xml
        out.removeSubrange(span.range)
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

    /// A malformed `// @slot-raw` line (missing paraId) FAILS LOUDLY at
    /// parse: a mangled directive would leave its makeDocument parameter
    /// unconsumed and render an unfilled form byte-equal to a correct
    /// all-default run — the fail-silent class this feature refuses
    /// (verify round 2, N1; supersedes the round-1 silent-skip pin).
    func testMalformedRawSlotDirectiveFailsLoudly() throws {
        let reference = try makeTableReference()
        let log = try reverseExpectingRaw(reference)
        var script = try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "amendment", paraId: "BBBB2222"),
        ])
        script = script.replacingOccurrences(
            of: "// @slot-raw amendment BBBB2222",
            with: "// @slot-raw malformed-only-one-token")
        XCTAssertThrowsError(try ScriptImporter.parse(source: script)) { error in
            guard case TranscodeError.unsupportedSyntax(_, _, let reason) = error else {
                return XCTFail("expected unsupportedSyntax, got \(error)")
            }
            XCTAssertTrue(reason.contains("@slot-raw"), "reason names the directive: \(reason)")
        }
    }

    /// Overlapping designated paragraphs (outer + one nested in its
    /// w:txbxContent) refuse at import — substituting one would silently
    /// discard the other, order-dependently (verify round 2, Codex HIGH).
    func testOverlappingNestedSlotDesignationsRefuse() throws {
        let xml = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:v="urn:schemas-microsoft-com:vml"><w:body><w:p w14:paraId="ZZZZ0001"><w:r><w:t>outer</w:t></w:r><w:r><w:pict><v:shape><v:textbox><w:txbxContent><w:p w14:paraId="AAAA0001"><w:r><w:t>inner</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p><w:tbl><w:tblPr/><w:tr><w:tc><w:p w14:paraId="CELL0001"><w:r><w:t>格</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>
            """
        let reference = try makeTableReference(documentXML: xml)
        let log = try reverseExpectingRaw(reference)
        var script = try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "outer", paraId: "ZZZZ0001"),
            SlotDesignation(name: "inner", paraId: "AAAA0001"),
        ])
        script = script.replacingOccurrences(
            of: "inner: \"inner\"", with: "inner: \"new\"")
        XCTAssertThrowsError(try execute(script: script)) { error in
            guard case TranscodeError.rawSlotExecutionFailure(_, let reason) = error else {
                return XCTFail("expected rawSlotExecutionFailure, got \(error)")
            }
            XCTAssertTrue(reason.contains("overlap"), "reason names the overlap: \(reason)")
        }
    }

    /// A `// @slot-raw` directive in a script with NO word/document.xml
    /// carryPart refuses — never a silent no-op (verify round 2, Codex HIGH).
    func testRawDirectiveWithoutDocumentCarryRefuses() throws {
        let script = """
            // @slot-raw ghost DEAD0000
            import WordDSLSwift

            let document = WordDocument {
                Section(id: "main") {
                }
            }
            """
        XCTAssertThrowsError(try ScriptImporter.parse(source: script)) { error in
            guard case TranscodeError.rawSlotExecutionFailure(_, let reason) = error else {
                return XCTFail("expected rawSlotExecutionFailure, got \(error)")
            }
            XCTAssertTrue(reason.contains("no word/document.xml"),
                          "reason names the missing carry: \(reason)")
        }
    }

    /// A paragraph whose only formatted run lives inside an inline wrapper
    /// (w:hyperlink) still contributes its rPr to the collapsed run
    /// (verify round 2, N2 — fidelity regression fix).
    func testHyperlinkWrappedRunContributesDominantRPr() throws {
        let xml = Self.tableDocumentXML.replacingOccurrences(
            of: "<w:p w14:paraId=\"AAAA1111\"><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>表單標題</w:t></w:r></w:p>",
            with: "<w:p w14:paraId=\"AAAA1111\"><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:hyperlink><w:r><w:rPr><w:b/></w:rPr><w:t>表單標題</w:t></w:r></w:hyperlink></w:p>")
        let reference = try makeTableReference(documentXML: xml)
        let log = try reverseExpectingRaw(reference)
        var script = try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "title", paraId: "AAAA1111"),
        ])
        script = script.replacingOccurrences(
            of: "title: \"表單標題\"", with: "title: \"新標題\"")
        let rebuilt = try execute(script: script)
        let outXML = String(data: rebuilt["word/document.xml"]!, encoding: .utf8)!
        XCTAssertTrue(outXML.contains("新標題"))
        XCTAssertTrue(outXML.contains("<w:b/>"),
                      "the hyperlink-wrapped run's rPr must survive as the dominant rPr")
    }
}

// MARK: - Verify round 1 hardening (structure-aware locator, fail-loud import)

extension RawChannelSlotTests {

    private func assertWellFormed(_ data: Data, _ message: String,
                                  file: StaticString = #filePath, line: UInt = #line) {
        let parser = XMLParser(data: data)
        XCTAssertTrue(parser.parse(),
                      "\(message) — XML must be well-formed; parser error: \(String(describing: parser.parserError))",
                      file: file, line: line)
    }

    /// A designated paragraph containing a text box: `w:p` nests through
    /// `w:txbxContent`, so depth-aware close matching is required. Surgery
    /// must replace the WHOLE outer paragraph and stay well-formed.
    func testTextboxNestedParagraphSubstitutionStaysWellFormed() throws {
        let xml = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:v="urn:schemas-microsoft-com:vml"><w:body><w:p w14:paraId="TBOX0001"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>外層文字</w:t></w:r><w:r><w:pict><v:shape><v:textbox><w:txbxContent><w:p w14:paraId="INNER001"><w:r><w:t>盒內文字</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p><w:tbl><w:tblPr/><w:tr><w:tc><w:p w14:paraId="CELL0001"><w:r><w:t>格</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>
            """
        let reference = try makeTableReference(documentXML: xml)
        let log = try reverseExpectingRaw(reference)
        var script = try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "outer", paraId: "TBOX0001"),
        ])
        script = script.replacingOccurrences(
            of: "outer: \"外層文字盒內文字\"", with: "outer: \"換掉了\"")
        let rebuilt = try execute(script: script)
        let out = rebuilt["word/document.xml"]!
        assertWellFormed(out, "textbox-nested substitution")
        let outXML = String(data: out, encoding: .utf8)!
        XCTAssertTrue(outXML.contains("換掉了"))
        XCTAssertFalse(outXML.contains("txbxContent"),
                       "the whole outer paragraph (including the text box) must be replaced")
        XCTAssertTrue(outXML.contains("<w:tbl>"), "the sibling table must be untouched")
        XCTAssertTrue(outXML.contains("CELL0001"), "table cell paragraph must survive")
    }

    /// Word writes `w14:paraId` on `<w:tr>` too (REC-O-01: 20 of 115).
    /// Designating a row-owned id must refuse naming the carrier — never
    /// corrupt the table.
    func testTableRowParaIdRefusesNamingCarrier() throws {
        let xml = Self.tableDocumentXML.replacingOccurrences(
            of: "<w:tr><w:tc>", with: "<w:tr w14:paraId=\"TROW0001\"><w:tc>")
        let log = try reverseExpectingRaw(try makeTableReference(documentXML: xml))
        XCTAssertThrowsError(try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "row", paraId: "TROW0001"),
        ])) { error in
            guard case TranscodeError.slotDesignationFailure(_, let reason) = error else {
                return XCTFail("expected slotDesignationFailure, got \(error)")
            }
            XCTAssertTrue(reason.contains("w:tr"), "reason must name the actual carrier: \(reason)")
        }
    }

    /// The literal token inside `<w:t>` text content is never a designation
    /// anchor — absence refusal, not a ghost match.
    func testParaIdOnlyInTextContentRefuses() throws {
        let xml = Self.tableDocumentXML.replacingOccurrences(
            of: "<w:t>表單標題</w:t>", with: "<w:t>w14:paraId=\"FAKE0001\" 假錨</w:t>")
        let log = try reverseExpectingRaw(try makeTableReference(documentXML: xml))
        XCTAssertThrowsError(try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "ghost", paraId: "FAKE0001"),
        ])) { error in
            guard case TranscodeError.slotDesignationFailure(_, let reason) = error else {
                return XCTFail("expected slotDesignationFailure, got \(error)")
            }
            XCTAssertTrue(reason.contains("log") && reason.contains("raw"),
                          "text content never anchors — absence refusal: \(reason)")
        }
    }

    /// A slot value containing the literal `w14:paraId="…"` token of ANOTHER
    /// slot must not redirect that slot's surgery (attribute-position
    /// anchoring, not token search).
    func testQuoteForgingValueCannotRedirectLaterSlot() throws {
        let reference = try makeTableReference()
        let log = try reverseExpectingRaw(reference)
        var script = try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "title", paraId: "AAAA1111"),
            SlotDesignation(name: "amendment", paraId: "BBBB2222"),
        ])
        script = script.replacingOccurrences(
            of: "title: \"表單標題\"",
            with: "title: \"w14:paraId=\\\"BBBB2222\\\" 偽造\"")
        script = script.replacingOccurrences(
            of: "amendment: \"申請修正第1次\"", with: "amendment: \"申請修正第3次\"")
        let rebuilt = try execute(script: script)
        let out = rebuilt["word/document.xml"]!
        assertWellFormed(out, "quote-forging value")
        let outXML = String(data: out, encoding: .utf8)!
        XCTAssertTrue(outXML.contains("申請修正第3次"),
                      "the later slot must land in its own paragraph")
        XCTAssertTrue(outXML.contains("<w:jc w:val=\"both\"/>"),
                      "the later slot's paragraph pPr must be preserved")
        XCTAssertTrue(outXML.contains("偽造"), "the forged text lands as inert text")
    }

    /// `w:pPrChange` nests a `w:pPr` inside the paragraph's `w:pPr`
    /// (track-changes documents). The preserved pPr block must be the whole
    /// outer block — depth-aware, not first `</w:pPr>`.
    func testPPrChangeNestedPPrPreservedAndWellFormed() throws {
        let xml = Self.tableDocumentXML.replacingOccurrences(
            of: "<w:pPr><w:jc w:val=\"both\"/></w:pPr>",
            with: "<w:pPr><w:jc w:val=\"both\"/><w:pPrChange w:id=\"1\"><w:pPr><w:jc w:val=\"center\"/></w:pPrChange></w:pPr>")
            .replacingOccurrences(
                of: "<w:pPrChange w:id=\"1\"><w:pPr><w:jc w:val=\"center\"/></w:pPrChange>",
                with: "<w:pPrChange w:id=\"1\"><w:pPr><w:jc w:val=\"center\"/></w:pPr></w:pPrChange>")
        let reference = try makeTableReference(documentXML: xml)
        let log = try reverseExpectingRaw(reference)
        var script = try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "amendment", paraId: "BBBB2222"),
        ])
        script = script.replacingOccurrences(
            of: "amendment: \"申請修正第1次\"", with: "amendment: \"申請修正第4次\"")
        let rebuilt = try execute(script: script)
        let out = rebuilt["word/document.xml"]!
        assertWellFormed(out, "pPrChange-nested substitution")
        let outXML = String(data: out, encoding: .utf8)!
        XCTAssertTrue(outXML.contains("<w:pPrChange w:id=\"1\">"),
                      "the whole outer pPr block including pPrChange must be preserved")
        XCTAssertTrue(outXML.contains("申請修正第4次"))
    }

    /// Values containing characters forbidden by XML 1.0 are refused, not
    /// written into the part.
    func testControlCharacterValueRefuses() throws {
        let reference = try makeTableReference()
        let log = try reverseExpectingRaw(reference)
        var script = try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "amendment", paraId: "BBBB2222"),
        ])
        script = script.replacingOccurrences(
            of: "amendment: \"申請修正第1次\"", with: "amendment: \"bad\u{0}value\"")
        XCTAssertThrowsError(try execute(script: script)) { error in
            guard case TranscodeError.rawSlotExecutionFailure = error else {
                return XCTFail("expected rawSlotExecutionFailure, got \(error)")
            }
        }
    }

    /// Numeric character references decode into the exported default, and the
    /// identity shortcut fires for the decoded value.
    func testNumericCharacterReferenceDefaultAndIdentity() throws {
        let xml = Self.tableDocumentXML.replacingOccurrences(
            of: "<w:t>表單標題</w:t>", with: "<w:t>&#x41;&#66;</w:t>")
        let reference = try makeTableReference(documentXML: xml)
        let log = try reverseExpectingRaw(reference)
        let script = try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "title", paraId: "AAAA1111"),
        ])
        XCTAssertTrue(script.contains("title: \"AB\""),
                      "numeric character references must decode into the default")
        let rebuilt = try execute(script: script)
        XCTAssertTrue(PartFidelity.stageB(reference: reference, rebuilt: rebuilt),
                      "decoded default must take the identity shortcut — byte-equal replay")
    }

    /// A collected directive whose paraId no longer resolves fails loudly at
    /// import — a stale directive must never silently render an unfilled form.
    func testImportStaleDirectiveFailsLoudly() throws {
        let reference = try makeTableReference()
        let log = try reverseExpectingRaw(reference)
        var script = try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "amendment", paraId: "BBBB2222"),
        ])
        script = script.replacingOccurrences(
            of: "// @slot-raw amendment BBBB2222",
            with: "// @slot-raw amendment DEAD0000")
        XCTAssertThrowsError(try execute(script: script)) { error in
            guard case TranscodeError.rawSlotExecutionFailure(_, let reason) = error else {
                return XCTFail("expected rawSlotExecutionFailure, got \(error)")
            }
            XCTAssertTrue(reason.contains("DEAD0000"), "reason must name the paraId: \(reason)")
        }
    }

    /// Duplicate paraId introduced AFTER export (hand-edited script) fails
    /// loudly at import — the guards re-apply at execution time.
    func testImportDuplicateParaIdFailsLoudly() throws {
        let reference = try makeTableReference()
        let log = try reverseExpectingRaw(reference)
        var script = try ScriptExporter.exportSwift(log: log, slots: [
            SlotDesignation(name: "amendment", paraId: "BBBB2222"),
        ])
        // Duplicate the paraId inside the carried XML (JSON-escaped in @op).
        script = script.replacingOccurrences(
            of: "w14:paraId=\\\"AAAA1111\\\"", with: "w14:paraId=\\\"BBBB2222\\\"")
        script = script.replacingOccurrences(
            of: "amendment: \"申請修正第1次\"", with: "amendment: \"值\"")
        XCTAssertThrowsError(try execute(script: script)) { error in
            guard case TranscodeError.rawSlotExecutionFailure(_, let reason) = error else {
                return XCTFail("expected rawSlotExecutionFailure, got \(error)")
            }
            XCTAssertTrue(reason.contains("2"), "reason must name the count: \(reason)")
        }
    }
}

extension RawChannelSlotTests {

    /// Verify-round-1 regression pin: sweep EVERY unique paraId in the real
    /// REC-O-01 fixture through the substitution path. Round 1 measured
    /// 88 ok / 19 silent table corruption / 2 malformed with the lexical
    /// locator. The structure-aware contract: every substitution that
    /// succeeds yields WELL-FORMED XML, and every non-`<w:p>` id refuses at
    /// designation. Zero corruption, zero silent failure.
    func testRECFixtureFullParaIdSweepZeroCorruption() throws {
        let reference = try RawPartChannel.readAllParts(from: try recFixtureURL())
        let result = try ReverseExtractor.reverse(parts: reference)
        let xml = String(data: reference["word/document.xml"]!, encoding: .utf8)!
        var ids: [String] = []
        var seen = Set<String>()
        var search = xml.startIndex
        while let r = xml.range(of: "w14:paraId=\"", range: search..<xml.endIndex) {
            guard let end = xml[r.upperBound...].firstIndex(of: "\"") else { break }
            let id = String(xml[r.upperBound..<end])
            if seen.insert(id).inserted { ids.append(id) }
            search = xml.index(after: end)
        }
        XCTAssertGreaterThan(ids.count, 100, "fixture should carry 100+ unique paraIds")
        var ok = 0, refused = 0
        for id in ids {
            let script: String
            do {
                script = try ScriptExporter.exportSwift(log: result.log, slots: [
                    SlotDesignation(name: "field", paraId: id)])
            } catch {
                guard case TranscodeError.slotDesignationFailure = error else {
                    return XCTFail("id \(id): unexpected designation error \(error)")
                }
                refused += 1
                continue
            }
            guard let dr = script.range(of: "field: \""),
                  let dEnd = script[dr.upperBound...].firstIndex(of: "\"") else {
                return XCTFail("id \(id): no call-site default")
            }
            let def = String(script[dr.upperBound..<dEnd])
            let mutated = script.replacingOccurrences(
                of: "field: \"\(def)\"", with: "field: \"SWEEPVALUE\"")
            let log2 = try ScriptImporter.parse(source: mutated)
            for entry in log2.entries {
                if case .carryPart(let p, let x) = entry.op, p == "word/document.xml" {
                    let parser = XMLParser(data: Data(x.utf8))
                    XCTAssertTrue(parser.parse(),
                                  "id \(id): substituted part must be well-formed")
                    ok += 1
                }
            }
        }
        XCTAssertEqual(ok + refused, ids.count, "every id accounted for")
        XCTAssertGreaterThan(refused, 0,
                             "fixture carries w:tr-owned ids — some refusals expected")
    }
}
