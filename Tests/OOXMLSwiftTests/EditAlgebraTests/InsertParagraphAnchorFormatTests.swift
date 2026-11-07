// InsertParagraphAnchorFormatTests.swift
// PsychQuant/ooxml-swift#234 — `inheritAnchorFormat` on OOXMLEdit.insertParagraph
// and insertParagraphBefore.
//
// The rules follow macdoc `word form apply --insert-after` (PsychQuant/macdoc#249):
//   - paragraph format: the anchor's own `w:pPr` without `w:numPr`, `w:sectPr`,
//     `w:pPrChange` and `w:pageBreakBefore`; its paragraph-mark `w:rPr` without
//     `w:ins` / `w:del` / `w:moveFrom` / `w:moveTo` / `w:rPrChange`
//   - text format: the `w:rPr` of the anchor's first own run with direct `w:t`
//     text, without `w:rPrChange`
//   - section break: after the anchor, the anchor's `w:sectPr` moves to the new
//     paragraph; before the anchor nothing moves

import XCTest
@testable import OOXMLSwift

final class InsertParagraphAnchorFormatTests: XCTestCase {

    // MARK: - Fixtures

    private static let namespaces =
        #"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "# +
        #"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" "# +
        #"xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" "# +
        #"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#

    private static let anchorID = ElementID(rawString: "w14:paraId=0000A001")

    /// A formatted anchor carrying every element the copy must drop, an empty
    /// first run, and a second run that is the first one with text.
    private static let formattedAnchor =
        #"<w:p w14:paraId="0000A001"><w:pPr>"# +
        #"<w:pStyle w:val="FormLabel"/><w:keepNext/><w:pageBreakBefore/>"# +
        #"<w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr>"# +
        #"<w:spacing w:before="120"/><w:ind w:left="480"/><w:jc w:val="both"/>"# +
        #"<w:rPr><w:ins w:id="7" w:author="A" w:date="2026-01-01T00:00:00Z"/>"# +
        #"<w:rFonts w:eastAsia="KaiTi"/><w:sz w:val="24"/></w:rPr>"# +
        #"<w:pPrChange w:id="8" w:author="A" w:date="2026-01-01T00:00:00Z"><w:pPr/></w:pPrChange>"# +
        #"</w:pPr>"# +
        #"<w:r><w:rPr><w:b/></w:rPr><w:t></w:t></w:r>"# +
        #"<w:r><w:rPr><w:rFonts w:eastAsia="KaiTi"/><w:sz w:val="28"/>"# +
        #"<w:rPrChange w:id="9" w:author="A" w:date="2026-01-01T00:00:00Z"><w:rPr/></w:rPrChange>"# +
        #"</w:rPr><w:t>anchor</w:t></w:r></w:p>"#

    /// An anchor that ends section 1, followed by a paragraph in section 2.
    private static let sectionEndingAnchor =
        #"<w:p w14:paraId="0000A001"><w:pPr><w:jc w:val="center"/>"# +
        #"<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr>"# +
        #"<w:r><w:t>end of section 1</w:t></w:r></w:p>"# +
        #"<w:p w14:paraId="0000A002"><w:r><w:t>section 2</w:t></w:r></w:p>"# +
        #"<w:sectPr><w:pgSz w:w="16838" w:h="11906"/></w:sectPr>"#

    private func documentXML(_ body: String) -> Data {
        Data(("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
              "<w:document \(Self.namespaces)><w:body>\(body)</w:body></w:document>").utf8)
    }

    private func makeDoc(_ body: String) throws -> WordDocument {
        var doc = WordDocument()
        doc.xmlTrees["word/document.xml"] = try XmlTreeReader.parse(documentXML(body))
        return doc
    }

    // MARK: - Tree helpers

    private func body(_ doc: WordDocument) -> XmlNode? {
        doc.xmlTrees["word/document.xml"].flatMap { child($0.root, "body") }
    }

    private func bodyParagraphs(_ doc: WordDocument) -> [XmlNode] {
        body(doc)?.children.filter { $0.kind == .element && $0.localName == "p" } ?? []
    }

    private func child(_ node: XmlNode?, _ name: String) -> XmlNode? {
        node?.children.first { $0.kind == .element && $0.localName == name }
    }

    private func childNames(_ node: XmlNode?) -> [String] {
        node?.children.filter { $0.kind == .element }.map(\.localName) ?? []
    }

    private func val(_ node: XmlNode?) -> String? {
        node?.attributeValue(prefix: "w", localName: "val")
    }

    private func text(of paragraph: XmlNode) -> String {
        func collect(_ node: XmlNode) -> String {
            if node.kind == .text { return node.textContent }
            return node.children.map(collect).joined()
        }
        return collect(paragraph)
    }

    private func serialized(_ doc: WordDocument) throws -> String {
        let tree = try XCTUnwrap(doc.xmlTrees["word/document.xml"])
        return String(decoding: try XmlTreeWriter.serialize(tree), as: UTF8.self)
    }

    private func occurrences(of needle: String, in haystack: String) -> Int {
        haystack.components(separatedBy: needle).count - 1
    }

    private func insertAfter(_ content: String = "new", styleId: String? = nil,
                             in body: String) throws -> WordDocument {
        try makeDoc(body).apply(OOXMLEdit.insertParagraph(
            after: Self.anchorID, content: content, styleId: styleId, inheritAnchorFormat: true))
    }

    // MARK: - Paragraph format

    func testAfterCopiesAnchorParagraphPropertiesWithoutOneOffStructure() throws {
        let result = try insertAfter(in: Self.formattedAnchor)
        let paragraphs = bodyParagraphs(result)
        XCTAssertEqual(paragraphs.count, 2)
        let pPr = child(paragraphs[1], "pPr")
        XCTAssertEqual(childNames(pPr), ["pStyle", "keepNext", "spacing", "ind", "jc", "rPr"],
                       "numPr, pageBreakBefore and pPrChange are dropped; the rest is kept in order")
        XCTAssertEqual(val(child(pPr, "pStyle")), "FormLabel")
        XCTAssertEqual(child(pPr, "ind")?.attributeValue(prefix: "w", localName: "left"), "480")
    }

    func testAfterDropsRevisionMarksFromParagraphMarkRunProperties() throws {
        let result = try insertAfter(in: Self.formattedAnchor)
        let markRPr = child(child(bodyParagraphs(result)[1], "pPr"), "rPr")
        XCTAssertEqual(childNames(markRPr), ["rFonts", "sz"])
        XCTAssertEqual(val(child(markRPr, "sz")), "24")
    }

    func testPPrStrippedToNothingIsOmitted() throws {
        let anchor = #"<w:p w14:paraId="0000A001"><w:pPr><w:numPr><w:ilvl w:val="0"/>"# +
            #"<w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>item</w:t></w:r></w:p>"#
        let result = try insertAfter(in: anchor)
        XCTAssertNil(child(bodyParagraphs(result)[1], "pPr"))
    }

    func testAnchorWithoutParagraphPropertiesGivesNoPPr() throws {
        let anchor = #"<w:p w14:paraId="0000A001"><w:r><w:rPr><w:i/></w:rPr><w:t>x</w:t></w:r></w:p>"#
        let inserted = bodyParagraphs(try insertAfter(in: anchor))[1]
        XCTAssertNil(child(inserted, "pPr"))
        XCTAssertEqual(childNames(child(child(inserted, "r"), "rPr")), ["i"])
    }

    func testCopiedPropertiesDoNotDuplicateLibraryIdentities() throws {
        let doc = try makeDoc(Self.formattedAnchor)
        let anchor = try XCTUnwrap(bodyParagraphs(doc).first)
        let jc = try XCTUnwrap(child(child(anchor, "pPr"), "jc"))
        let jcID = UUID()
        jc.libraryUUID = jcID
        let result = try doc.apply(OOXMLEdit.insertParagraph(
            after: Self.anchorID, content: "new", styleId: nil, inheritAnchorFormat: true))
        func count(_ node: XmlNode) -> Int {
            (node.libraryUUID == jcID ? 1 : 0) + node.children.map(count).reduce(0, +)
        }
        let root = try XCTUnwrap(result.xmlTrees["word/document.xml"]?.root)
        XCTAssertEqual(count(root), 1, "a copied element must not reuse the anchor's identity")
    }

    // MARK: - Text format

    func testAfterCopiesFirstTextRunPropertiesWithoutRPrChange() throws {
        let inserted = bodyParagraphs(try insertAfter(in: Self.formattedAnchor))[1]
        let run = try XCTUnwrap(child(inserted, "r"))
        XCTAssertEqual(childNames(run), ["rPr", "t"])
        XCTAssertEqual(childNames(child(run, "rPr")), ["rFonts", "sz"],
                       "the empty first run is skipped; rPrChange is dropped")
        XCTAssertEqual(val(child(child(run, "rPr"), "sz")), "28")
        XCTAssertEqual(text(of: inserted), "new")
    }

    func testFirstTextRunIgnoresRunsOfNestedParagraphsAndCountsHyperlinkRuns() throws {
        // The text-box run comes first in document order but belongs to the
        // nested paragraph; the hyperlink run belongs to the anchor.
        let anchor = #"<w:p w14:paraId="0000A001">"# +
            #"<w:r><w:pict><w:txbxContent><w:p><w:r><w:rPr><w:i/></w:rPr><w:t>box</w:t></w:r></w:p>"# +
            #"</w:txbxContent></w:pict></w:r>"# +
            #"<w:hyperlink r:id="rId5"><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>link</w:t></w:r></w:hyperlink>"# +
            #"</w:p>"#
        let inserted = bodyParagraphs(try insertAfter(in: anchor))[1]
        XCTAssertEqual(childNames(child(child(inserted, "r"), "rPr")), ["u"])
    }

    func testFirstTextRunIsNotAMathRun() throws {
        // verify R1 #1: OMML's m:r / m:t / m:rPr share local names with
        // WordprocessingML's. An anchor that opens with inline math must not
        // hand the math run's m:rPr to the new w:r (schema-invalid output).
        let anchor = #"<w:p w14:paraId="0000A001"><m:oMath><m:r>"# +
            #"<m:rPr><m:sty m:val="p"/></m:rPr>"# +
            #"<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>"# +
            #"<m:t>x</m:t></m:r></m:oMath>"# +
            #"<w:r><w:rPr><w:b/></w:rPr><w:t> is a variable</w:t></w:r></w:p>"#
        let inserted = bodyParagraphs(try insertAfter(in: anchor))[1]
        let run = try XCTUnwrap(child(inserted, "r"))
        XCTAssertEqual(childNames(child(run, "rPr")), ["b"],
                       "the first WordprocessingML run with text, not the math run")
        let xml = try serialized(try insertAfter(in: anchor))
        XCTAssertEqual(occurrences(of: "<m:rPr>", in: xml), 1, "no math rPr is copied")
    }

    func testFlagKeepsAnExplicitParaIdInBothDirections() throws {
        for op in [
            OOXMLSwift.Operation.insertParagraphAfter(
                after: Self.anchorID,
                paragraph: ParagraphPayload(text: "x", paraId: "0000B0B0", inheritAnchorFormat: true)),
            OOXMLSwift.Operation.insertParagraphBefore(
                before: Self.anchorID,
                paragraph: ParagraphPayload(text: "x", paraId: "0000B0B0", inheritAnchorFormat: true)),
        ] {
            var log = OperationLog()
            log.append(op, source: .swift)
            let tree = try OperationReducer.materialize(
                log: log, base: try XmlTreeReader.parse(documentXML(Self.formattedAnchor)))
            let inserted = try XCTUnwrap((child(tree.root, "body")?.children ?? []).first {
                $0.attributeValue(prefix: "w14", localName: "paraId") == "0000B0B0"
            }, "the new paragraph is addressable by its paraId")
            XCTAssertEqual(childNames(child(inserted, "pPr")),
                           ["pStyle", "keepNext", "spacing", "ind", "jc", "rPr"])
        }
    }

    func testAnchorWithoutTextRunGivesUnformattedRun() throws {
        let anchor = #"<w:p w14:paraId="0000A001"><w:pPr><w:jc w:val="center"/></w:pPr></w:p>"#
        let inserted = bodyParagraphs(try insertAfter(in: anchor))[1]
        XCTAssertEqual(childNames(child(inserted, "pPr")), ["jc"])
        XCTAssertEqual(childNames(child(inserted, "r")), ["t"])
    }

    // MARK: - Section break

    func testAfterMovesAnchorSectionBreakToInsertedParagraph() throws {
        let result = try insertAfter(in: Self.sectionEndingAnchor)
        let paragraphs = bodyParagraphs(result)
        XCTAssertEqual(paragraphs.count, 3)
        XCTAssertEqual(childNames(child(paragraphs[0], "pPr")), ["jc"],
                       "the anchor no longer ends the section")
        let insertedPPr = child(paragraphs[1], "pPr")
        XCTAssertEqual(childNames(insertedPPr), ["jc", "sectPr"])
        XCTAssertEqual(child(child(insertedPPr, "sectPr"), "pgSz")?
            .attributeValue(prefix: "w", localName: "w"), "11906")
        XCTAssertEqual(text(of: paragraphs[2]), "section 2")
    }

    func testTwoInsertsAfterSameAnchorKeepOneSectionBreakAtSectionEnd() throws {
        let doc = try makeDoc(Self.sectionEndingAnchor)
        let first = try doc.apply(OOXMLEdit.insertParagraph(
            after: Self.anchorID, content: "A", styleId: nil, inheritAnchorFormat: true))
        let second = try first.apply(OOXMLEdit.insertParagraph(
            after: Self.anchorID, content: "B", styleId: nil, inheritAnchorFormat: true))
        let paragraphs = bodyParagraphs(second)
        XCTAssertEqual(paragraphs.map { text(of: $0) },
                       ["end of section 1", "B", "A", "section 2"])
        XCTAssertEqual(paragraphs.map { child(child($0, "pPr"), "sectPr") != nil },
                       [false, false, true, false],
                       "the break stays on the last paragraph of section 1")
    }

    func testBeforeNeitherCopiesNorMovesSectionBreak() throws {
        let result = try makeDoc(Self.sectionEndingAnchor).apply(OOXMLEdit.insertParagraphBefore(
            before: Self.anchorID, content: "new", styleId: nil, inheritAnchorFormat: true))
        let paragraphs = bodyParagraphs(result)
        XCTAssertEqual(text(of: paragraphs[0]), "new")
        XCTAssertEqual(childNames(child(paragraphs[0], "pPr")), ["jc"])
        XCTAssertEqual(childNames(child(paragraphs[1], "pPr")), ["jc", "sectPr"],
                       "the anchor keeps its section break")
    }

    func testBeforeCopiesParagraphAndTextFormat() throws {
        let result = try makeDoc(Self.formattedAnchor).apply(OOXMLEdit.insertParagraphBefore(
            before: Self.anchorID, content: "new", styleId: nil, inheritAnchorFormat: true))
        let inserted = bodyParagraphs(result)[0]
        XCTAssertEqual(childNames(child(inserted, "pPr")),
                       ["pStyle", "keepNext", "spacing", "ind", "jc", "rPr"])
        XCTAssertEqual(childNames(child(child(inserted, "r"), "rPr")), ["rFonts", "sz"])
    }

    func testUndoRestoresAnchorSectionBreak() throws {
        // Structural ops are undone as a batch envelope (the reducer only
        // undoes setText / setParagraphStyle individually); undo replays the
        // history without the batch, so the anchor's moved sectPr comes back.
        let base = try XmlTreeReader.parse(documentXML(Self.sectionEndingAnchor))
        let batchID = UUID()
        var log = OperationLog()
        log.append(.batchBegin(label: "insert"), source: .swift, opID: batchID)
        log.append(.insertParagraphAfter(
            after: Self.anchorID,
            paragraph: ParagraphPayload(text: "new", inheritAnchorFormat: true)), source: .swift)
        log.append(.batchEnd, source: .swift)
        let applied = try OperationReducer.materialize(log: log, base: base)
        XCTAssertNil(child(child((child(applied.root, "body")?.children ?? [])
            .first { $0.localName == "p" }, "pPr"), "sectPr"),
                     "precondition: the insert moved the break off the anchor")
        log.append(.undo(targetOpID: batchID), source: .swift)
        let tree = try OperationReducer.materialize(log: log, base: base)
        let paragraphs = (child(tree.root, "body")?.children ?? [])
            .filter { $0.kind == .element && $0.localName == "p" }
        XCTAssertEqual(paragraphs.count, 2)
        XCTAssertEqual(childNames(child(paragraphs[0], "pPr")), ["jc", "sectPr"])
    }

    // MARK: - styleId

    func testExplicitStyleIdReplacesCopiedPStyle() throws {
        let pPr = child(bodyParagraphs(try insertAfter(styleId: "Body", in: Self.formattedAnchor))[1], "pPr")
        XCTAssertEqual(childNames(pPr), ["pStyle", "keepNext", "spacing", "ind", "jc", "rPr"])
        XCTAssertEqual(val(child(pPr, "pStyle")), "Body")
    }

    func testExplicitStyleIdIsAddedFirstWhenAnchorHasNone() throws {
        let anchor = #"<w:p w14:paraId="0000A001"><w:pPr><w:jc w:val="center"/></w:pPr>"# +
            #"<w:r><w:t>x</w:t></w:r></w:p>"#
        let pPr = child(bodyParagraphs(try insertAfter(styleId: "Body", in: anchor))[1], "pPr")
        XCTAssertEqual(childNames(pPr), ["pStyle", "jc"])
        XCTAssertEqual(val(child(pPr, "pStyle")), "Body")
    }

    // MARK: - Default behaviour

    func testDefaultLeavesInsertedParagraphUnformatted() throws {
        let result = try makeDoc(Self.sectionEndingAnchor).apply(OOXMLEdit.insertParagraph(
            after: Self.anchorID, content: "new", styleId: nil))
        let paragraphs = bodyParagraphs(result)
        XCTAssertNil(child(paragraphs[1], "pPr"))
        XCTAssertEqual(childNames(child(paragraphs[1], "r")), ["t"])
        XCTAssertEqual(childNames(child(paragraphs[0], "pPr")), ["jc", "sectPr"])
    }

    func testOperationsCarryFlagOnlyWhenTrue() throws {
        let on = try OOXMLEdit.insertParagraph(
            after: Self.anchorID, content: "x", styleId: "S", inheritAnchorFormat: true).operations()
        XCTAssertEqual(on, [.insertParagraphAfter(
            after: Self.anchorID,
            paragraph: ParagraphPayload(text: "x", styleId: "S", inheritAnchorFormat: true))])
        let off = try OOXMLEdit.insertParagraphBefore(
            before: Self.anchorID, content: "x", styleId: "S").operations()
        XCTAssertEqual(off, [.insertParagraphBefore(
            before: Self.anchorID, paragraph: ParagraphPayload(text: "x", styleId: "S"))])
    }

    // MARK: - Refusals

    func testNonParagraphAnchorIsRejected() throws {
        let doc = try makeDoc(#"<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#)
        let table = try XCTUnwrap(body(doc)?.children.first { $0.localName == "tbl" })
        let tableID = UUID()
        table.libraryUUID = tableID
        XCTAssertThrowsError(try doc.apply(OOXMLEdit.insertParagraph(
            after: ElementID(libraryUUID: tableID), content: "x", styleId: nil,
            inheritAnchorFormat: true)))
    }

    private func assertMalformed(_ op: OOXMLSwift.Operation, file: StaticString = #filePath, line: UInt = #line) throws {
        let base = try XmlTreeReader.parse(documentXML(Self.formattedAnchor))
        var log = OperationLog()
        log.append(op, source: .swift)
        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: base), file: file, line: line) { error in
            guard case ReducerError.malformedOp = error else {
                return XCTFail("expected malformedOp, got \(error)", file: file, line: line)
            }
        }
    }

    func testFlagWithExplicitParagraphFieldsIsRejected() throws {
        try assertMalformed(.insertParagraphAfter(
            after: Self.anchorID,
            paragraph: ParagraphPayload(text: "x", alignment: "center", inheritAnchorFormat: true)))
        try assertMalformed(.insertParagraphBefore(
            before: Self.anchorID,
            paragraph: ParagraphPayload(text: "x", numId: 3, inheritAnchorFormat: true)))
    }

    func testAppendParagraphWithFlagIsRejected() throws {
        try assertMalformed(.appendParagraph(
            in: nil, paragraph: ParagraphPayload(text: "x", inheritAnchorFormat: true)))
    }

    // MARK: - Serialization and log compatibility

    func testSerializedOutputCarriesTheStrippedCopy() throws {
        let xml = try serialized(try insertAfter(in: Self.formattedAnchor))
        XCTAssertEqual(occurrences(of: "<w:numPr>", in: xml), 1, "only the anchor keeps its numbering")
        XCTAssertEqual(occurrences(of: "<w:pageBreakBefore/>", in: xml), 1)
        XCTAssertEqual(occurrences(of: "<w:pPrChange", in: xml), 1)
        XCTAssertEqual(occurrences(of: "<w:rPrChange", in: xml), 1)
        XCTAssertEqual(occurrences(of: "<w:ins ", in: xml), 1)
        XCTAssertEqual(occurrences(of: #"<w:pStyle w:val="FormLabel"/>"#, in: xml), 2)
        XCTAssertNoThrow(try XmlTreeReader.parse(Data(xml.utf8)), "output stays well-formed")
    }

    func testJSONLRoundTripsTheFlag() throws {
        var log = OperationLog()
        let op = OOXMLSwift.Operation.insertParagraphAfter(
            after: Self.anchorID, paragraph: ParagraphPayload(text: "x", inheritAnchorFormat: true))
        log.append(op, source: .swift)
        let decoded = try OperationLog.decodeJSONL(log.encodeJSONL())
        XCTAssertEqual(decoded.entries.map(\.op), [op])
    }

    func testJSONLWithoutTheFlagIsUnchanged() throws {
        var log = OperationLog()
        log.append(.insertParagraphAfter(
            after: Self.anchorID, paragraph: ParagraphPayload(text: "x", styleId: "S")), source: .swift)
        let line = String(decoding: log.encodeJSONL(), as: UTF8.self)
        XCTAssertFalse(line.contains("inheritAnchorFormat"),
                       "logs written without the option keep their existing bytes")
    }

    func testLegacyJSONLLineDecodesWithoutTheFlagAndReencodesToTheSameBytes() throws {
        // verify R1 #2: a line written before the field existed.
        let line = #"{"op_id":"33333333-3333-4333-8333-333333333333","ts":"2026-07-01T00:00:00Z","source":"swift","op_type":"insertParagraphAfter","after":"w14:paraId=0000A001","paragraph":{"styleId":"Body","text":"hi"}}"#
        let decoded = try OperationLog.decodeJSONL(Data((line + "\n").utf8))
        guard case .insertParagraphAfter(let after, let payload) = try XCTUnwrap(decoded.entries.first).op else {
            return XCTFail("expected insertParagraphAfter")
        }
        XCTAssertEqual(after, Self.anchorID)
        XCTAssertNil(payload.inheritAnchorFormat)
        XCTAssertEqual(payload, ParagraphPayload(text: "hi", styleId: "Body"))
        XCTAssertEqual(String(decoding: decoded.encodeJSONL(), as: UTF8.self), line + "\n",
                       "an existing log keeps its bytes")
    }

    func testScriptExportKeepsTheFlagOnAnAppendParagraph() throws {
        // verify R1 #5: the DSL spelling has no place for the flag, so the
        // exporter must use the op escape instead of dropping it.
        var log = OperationLog()
        log.append(.appendParagraph(in: nil, paragraph: ParagraphPayload(
            text: "x", paraId: "0000C0C0", inheritAnchorFormat: true)), source: .swift)
        let script = ScriptExporter.exportSwift(log: log)
        XCTAssertFalse(script.contains(#"Paragraph(id: "0000C0C0")"#))
        XCTAssertTrue(script.contains("inheritAnchorFormat"))
    }

    func testApplyKeepsTheDocumentPartTreeAuthoritative() throws {
        let result = try insertAfter(in: Self.formattedAnchor)
        XCTAssertTrue(result.treeFreshParts.contains("word/document.xml"),
                      "the insert goes through the op log, not a typed mutation")
    }
}
