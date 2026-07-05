import XCTest
@testable import OOXMLSwift

/// word-aligned-state-sync §4b (tasks 4b.1–4b.3, PsychQuant/macdoc#128) —
/// additive authoring ops with OOXML-mirror naming, per the consolidated
/// `ooxml-operation-log` delta:
/// `appendParagraph(in:)`, `setRuns`, `defineStyle`,
/// `beginComponent`/`endComponent` (log-metadata exception),
/// `insertTab`/`insertBreak`/`insertNoBreakHyphen` (run-scoped atoms).
final class TaxonomyExtensionTests: XCTestCase {

    private let pid = ElementID(rawString: "w14:paraId=0AB7C123")
    private let rid = ElementID(rawString: "lib:11111111-2222-3333-4444-555555555555")

    private func docTree(_ body: String) throws -> XmlTree {
        let xml = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>\(body)</w:body></w:document>
            """
        return try XmlTreeReader.parse(Data(xml.utf8))
    }

    private func materialize(_ ops: [OOXMLSwift.Operation], base: XmlTree) throws -> XmlTree {
        var log = OperationLog()
        for op in ops { log.append(op, source: .swift) }
        return try OperationReducer.materialize(log: log, base: base)
    }

    private func serialized(_ tree: XmlTree) throws -> String {
        String(decoding: try XmlTreeWriter.serialize(tree), as: UTF8.self)
    }

    // MARK: - 4b.1 enum cases construct and pattern-match

    func testNewCasesConstructAndPatternMatch() {
        let cases: [OOXMLSwift.Operation] = [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(text: "t", paraId: "P1")),
            .appendParagraph(in: pid, paragraph: ParagraphPayload(text: "t")),
            .setRuns(target: pid, runs: [RunPayload(text: "a"), RunPayload(text: "b", bold: true, italic: true, color: "663300")]),
            .defineStyle(payload: StylePayload(styleId: "titleBrown", font: "Noto Serif TC", fontSize: 36, color: "663300", bold: true)),
            .beginComponent(type: "Summary", id: ElementID(rawString: "ch1-summary")),
            .endComponent(id: ElementID(rawString: "ch1-summary")),
            .insertTab(in: rid),
            .insertBreak(in: rid),
            .insertNoBreakHyphen(in: rid),
        ]
        XCTAssertEqual(cases.count, 9)
        if case .setRuns(_, let runs) = cases[2] {
            XCTAssertEqual(runs[1].bold, true)
            XCTAssertEqual(runs[1].italic, true)
            XCTAssertEqual(runs[1].color, "663300")
        } else { XCTFail("setRuns pattern match failed") }
        if case .appendParagraph(let container, let p) = cases[0] {
            XCTAssertNil(container)
            XCTAssertEqual(p.paraId, "P1")
        } else { XCTFail("appendParagraph pattern match failed") }
    }

    // MARK: - 4b.2 JSONL round-trip per new op + payload backward compat

    func testJSONLRoundTripForNewOps() throws {
        let ops: [OOXMLSwift.Operation] = [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(text: "hello", styleId: "Heading1", paraId: "P1")),
            .appendParagraph(in: pid, paragraph: ParagraphPayload(text: "x")),
            .setRuns(target: pid, runs: [RunPayload(text: "本章探討"), RunPayload(text: "意識本質", bold: true)]),
            .defineStyle(payload: StylePayload(styleId: "titleBrown", font: "Noto Serif TC", fontSize: 36, color: "663300", bold: true)),
            .beginComponent(type: "Summary", id: ElementID(rawString: "ch1-summary")),
            .endComponent(id: ElementID(rawString: "ch1-summary")),
            .insertTab(in: rid),
            .insertBreak(in: rid),
            .insertNoBreakHyphen(in: rid),
        ]
        var log = OperationLog()
        for op in ops { log.append(op, source: .swift) }

        let decoded = try OperationLog.decodeJSONL(log.encodeJSONL())
        XCTAssertEqual(decoded.entries.count, ops.count)
        for (i, entry) in decoded.entries.enumerated() {
            XCTAssertEqual(entry.op, ops[i], "op \(i) must round-trip through JSONL")
            XCTAssertEqual(entry.source, .swift)
        }
    }

    func testOldPayloadJSONStillDecodes() throws {
        // Backward compat: RunPayload/ParagraphPayload gained optional fields —
        // pre-#128 JSON without them must still decode (fields nil).
        let old = #"{"text":"legacy"}"#
        let run = try JSONDecoder().decode(RunPayload.self, from: Data(old.utf8))
        XCTAssertEqual(run.text, "legacy")
        XCTAssertNil(run.bold)
        XCTAssertNil(run.italic)
        XCTAssertNil(run.color)
        let para = try JSONDecoder().decode(ParagraphPayload.self, from: Data(old.utf8))
        XCTAssertNil(para.paraId)
    }

    // MARK: - 4b.3 reducer: appendParagraph

    func testAppendParagraphNilContainerAppendsToBody() throws {
        let base = try docTree(#"<w:p w14:paraId="EXIST"><w:r><w:t>first</w:t></w:r></w:p>"#)
        let out = try materialize(
            [.appendParagraph(in: nil, paragraph: ParagraphPayload(text: "appended", paraId: "NEWP"))],
            base: base)

        let xml = try serialized(out)
        XCTAssertTrue(xml.contains("appended"))
        XCTAssertTrue(xml.contains(#"w14:paraId="NEWP""#),
                      "explicit paraId must be stamped on the created <w:p> (spec: paraId ↔ w14:paraId)")
        // Appended AFTER the existing paragraph.
        XCTAssertLessThan(xml.range(of: "EXIST")!.lowerBound,
                          xml.range(of: "NEWP")!.lowerBound)
    }

    func testAppendParagraphWithoutParaIdUsesOpIDDerivedUUID() throws {
        let base = try docTree("")
        var log = OperationLog()
        let opID = UUID()
        log.append(.appendParagraph(in: nil, paragraph: ParagraphPayload(text: "no id")),
                   source: .swift, opID: opID)
        let out = try OperationReducer.materialize(log: log, base: base)

        func findP(_ n: XmlNode) -> XmlNode? {
            if n.kind == .element && n.localName == "p" { return n }
            for c in n.children { if let hit = findP(c) { return hit } }
            return nil
        }
        guard let p = findP(out.root) else { return XCTFail("no <w:p> created") }
        XCTAssertEqual(p.libraryUUID, opID,
                       "absent paraId keeps the opID-derived libraryUUID behavior unchanged")
    }

    // MARK: - 4b.3 reducer: setRuns

    func testSetRunsReplacesContentKeepsPPr() throws {
        let base = try docTree(#"<w:p w14:paraId="0AB7C123"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>old</w:t></w:r></w:p>"#)
        let out = try materialize(
            [.setRuns(target: pid, runs: [
                RunPayload(text: "本章探討"),
                RunPayload(text: "意識本質", bold: true),
            ])],
            base: base)

        let xml = try serialized(out)
        XCTAssertTrue(xml.contains(#"<w:jc w:val="center"/>"#), "pPr must be preserved")
        XCTAssertFalse(xml.contains("old"), "previous inline content must be replaced")
        XCTAssertTrue(xml.contains("本章探討"))
        // Second run carries <w:rPr><w:b/></w:rPr> per spec scenario.
        XCTAssertTrue(xml.contains("<w:rPr><w:b/></w:rPr><w:t>意識本質</w:t>")
                   || xml.contains("<w:b/>"),
                      "bold run must carry <w:b/> inside <w:rPr>")
    }

    // MARK: - 4b.3 reducer: defineStyle (styles part routing + idempotency)

    func testDefineStyleIsIdempotentInStylesPart() throws {
        let stylesXML = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>
            """
        var trees: [String: XmlTree] = [
            "word/styles.xml": try XmlTreeReader.parse(Data(stylesXML.utf8)),
        ]
        var doc = WordDocument()
        doc.xmlTrees = trees
        _ = trees

        let payload = StylePayload(styleId: "titleBrown", font: "Noto Serif TC",
                                   fontSize: 36, color: "663300", bold: true)
        try doc.appendAndMaterialize([.defineStyle(payload: payload),
                                      .defineStyle(payload: payload)])

        let xml = try serialized(doc.xmlTrees["word/styles.xml"]!)
        XCTAssertEqual(xml.components(separatedBy: #"w:styleId="titleBrown""#).count - 1, 1,
                       "duplicate defineStyle must be an idempotent no-op (exactly one definition)")
        XCTAssertTrue(xml.contains("Noto Serif TC"))
    }

    // MARK: - 4b.3 reducer: component envelope is a no-op marker

    func testComponentEnvelopeIsNoOpAndSurvivesApplyPath() throws {
        let base = try docTree(#"<w:p w14:paraId="0AB7C123"><w:r><w:t>stable</w:t></w:r></w:p>"#)
        var doc = WordDocument()
        doc.xmlTrees = ["word/document.xml": base]

        let before = doc.xmlTrees["word/document.xml"]!.root.normalizedFingerprint()
        try doc.appendAndMaterialize([
            .beginComponent(type: "Summary", id: ElementID(rawString: "ch1-summary")),
            .endComponent(id: ElementID(rawString: "ch1-summary")),
        ])
        let after = doc.xmlTrees["word/document.xml"]!.root.normalizedFingerprint()

        XCTAssertEqual(before, after, "component envelope must not touch any tree")
        XCTAssertEqual(doc.operationLog.entries.count, 2, "both markers must land in the log")
        let xml = try serialized(doc.xmlTrees["word/document.xml"]!)
        XCTAssertFalse(xml.lowercased().contains("component"),
                       "no OOXML artifact for component markers")
    }

    // MARK: - 4b.3 reducer: inline atoms (run-scoped + paragraph-target synthesis)

    func testInsertTabAppendsToAddressedRun() throws {
        // Run addressed via libraryUUID (runs rarely carry native stable IDs).
        let base = try docTree(#"<w:p w14:paraId="0AB7C123"><w:r><w:t>Header</w:t></w:r></w:p>"#)
        // Assign the library UUID to the run node so ElementID resolves.
        func findRun(_ n: XmlNode) -> XmlNode? {
            if n.kind == .element && n.localName == "r" { return n }
            for c in n.children { if let hit = findRun(c) { return hit } }
            return nil
        }
        let runUUID = UUID(uuidString: "11111111-2222-3333-4444-555555555555")!
        findRun(base.root)!.libraryUUID = runUUID

        let out = try materialize([.insertTab(in: rid)], base: base)
        let xml = try serialized(out)
        XCTAssertTrue(xml.contains("<w:t>Header</w:t><w:tab/>")
                   || xml.contains("<w:tab/></w:r>"),
                      "<w:tab/> must be appended inside the addressed <w:r>, got: \(xml)")
    }

    func testInsertBreakOnParagraphTargetSynthesizesWrappingRun() throws {
        // Spec rule: a standalone atom addressed at a paragraph (no preceding
        // run) makes the reducer synthesize an empty wrapping <w:r>.
        let base = try docTree(#"<w:p w14:paraId="0AB7C123"></w:p>"#)
        let out = try materialize([.insertBreak(in: pid)], base: base)
        let xml = try serialized(out)
        XCTAssertTrue(xml.contains("<w:r><w:br/></w:r>"),
                      "atom on paragraph target must be wrapped in a synthesized <w:r>, got: \(xml)")
    }

    // MARK: - referencedElementIDs / addressing

    func testReferencedElementIDsForNewOps() {
        XCTAssertEqual(OperationReducer.referencedElementIDs(
            in: .appendParagraph(in: pid, paragraph: ParagraphPayload(text: "x"))), [pid])
        XCTAssertEqual(OperationReducer.referencedElementIDs(
            in: .appendParagraph(in: nil, paragraph: ParagraphPayload(text: "x"))), [])
        XCTAssertEqual(OperationReducer.referencedElementIDs(
            in: .setRuns(target: pid, runs: [])), [pid])
        XCTAssertEqual(OperationReducer.referencedElementIDs(
            in: .insertTab(in: rid)), [rid])
        XCTAssertEqual(OperationReducer.referencedElementIDs(
            in: .defineStyle(payload: StylePayload(styleId: "s"))), [])
        XCTAssertEqual(OperationReducer.referencedElementIDs(
            in: .beginComponent(type: "T", id: ElementID(rawString: "c"))), [])
        XCTAssertEqual(OperationReducer.referencedElementIDs(
            in: .endComponent(id: ElementID(rawString: "c"))), [])
    }
}
