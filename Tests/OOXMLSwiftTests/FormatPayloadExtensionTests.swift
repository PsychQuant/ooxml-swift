// FormatPayloadExtensionTests.swift
// format-alignment-engine Phase B task 2.1 — additive payload extensions for
// five-layer extraction (`ooxml-operation-log`, «Format-payload additive
// extensions»; Decision 3). RunPayload gains fonts/size/underline/vertAlign,
// ParagraphPayload gains spacing/indent/alignment/numPr, SectionPayload is
// new (with the `setSectionProperties` op). All fields optional per the #128
// additive-only wire discipline; names never collide with the envelope keys
// op_id / ts / source / op_type.

import XCTest
@testable import OOXMLSwift

final class FormatPayloadExtensionTests: XCTestCase {

    // MARK: - Wire compat (spec scenarios)

    /// Spec scenario: old sidecar decodes under extended payloads.
    func testOldSidecarLinesDecodeWithNewFieldsAbsent() throws {
        // v1.0.x-form lines: payloads carry only the pre-extension fields.
        let lines = """
            {"op_id":"11111111-1111-4111-8111-111111111111","ts":"2026-07-01T00:00:00Z","source":"swift","op_type":"appendParagraph","in":null,"paragraph":{"paraId":"P1","styleId":"Body","text":"hi"}}
            {"op_id":"22222222-2222-4222-8222-222222222222","ts":"2026-07-01T00:00:01Z","source":"swift","op_type":"setRuns","target":"w14:paraId=P1","runs":[{"bold":true,"text":"hi"}]}
            """
        let log = try OperationLog.decodeJSONL(Data((lines + "\n").utf8))
        XCTAssertEqual(log.entries.count, 2)

        guard case .appendParagraph(_, let para) = log.entries[0].op else {
            return XCTFail("expected appendParagraph, got \(log.entries[0].op)")
        }
        XCTAssertEqual(para.text, "hi")
        XCTAssertNil(para.alignment)
        XCTAssertNil(para.spacingBefore)
        XCTAssertNil(para.indentLeft)
        XCTAssertNil(para.numId)

        guard case .setRuns(_, let runs) = log.entries[1].op else {
            return XCTFail("expected setRuns, got \(log.entries[1].op)")
        }
        XCTAssertEqual(runs.count, 1)
        XCTAssertEqual(runs[0].bold, true)
        XCTAssertNil(runs[0].fontAscii)
        XCTAssertNil(runs[0].fontEastAsia)
        XCTAssertNil(runs[0].sizeHalfPoints)
        XCTAssertNil(runs[0].underline)
        XCTAssertNil(runs[0].vertAlign)

        // Replay is unchanged: the old-form ops still apply cleanly.
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: log.entries.map(\.op))
    }

    /// Spec scenario: extended fields round-trip the wire (eastAsia font +
    /// size 21 half-points per the spec example values).
    func testExtendedRunPayloadRoundTripsFieldForField() throws {
        let run = RunPayload(
            text: "こんにちは",
            bold: true,
            color: "FF0000",
            fontAscii: "Times New Roman",
            fontEastAsia: "ＭＳ ゴシック",
            sizeHalfPoints: 21,
            underline: "single",
            vertAlign: "superscript")
        var log = OperationLog()
        log.append(.setRuns(target: ElementID(rawString: "w14:paraId=P1"), runs: [run]),
                   source: .swift)

        let decoded = try OperationLog.decodeJSONL(log.encodeJSONL())
        guard case .setRuns(_, let runs) = decoded.entries[0].op else {
            return XCTFail("expected setRuns")
        }
        XCTAssertEqual(runs, [run], "extended RunPayload must round-trip field-for-field")
        XCTAssertEqual(log.encodeJSONL(), decoded.encodeJSONL(), "byte-equal re-encode")
    }

    func testExtendedParagraphPayloadRoundTripsFieldForField() throws {
        let para = ParagraphPayload(
            text: "段落",
            styleId: "Body",
            paraId: "P9",
            alignment: "both",
            spacingBefore: 120,
            spacingAfter: 240,
            spacingLine: 360,
            spacingLineRule: "auto",
            indentLeft: 720,
            indentRight: 360,
            indentFirstLine: 420,
            indentHanging: nil,
            numId: 3,
            numLevel: 1)
        var log = OperationLog()
        log.append(.appendParagraph(in: nil, paragraph: para), source: .swift)

        let decoded = try OperationLog.decodeJSONL(log.encodeJSONL())
        guard case .appendParagraph(_, let got) = decoded.entries[0].op else {
            return XCTFail("expected appendParagraph")
        }
        XCTAssertEqual(got, para, "extended ParagraphPayload must round-trip field-for-field")
    }

    func testSetSectionPropertiesRoundTripsFieldForField() throws {
        let section = SectionPayload(
            pageWidth: 11906,
            pageHeight: 16838,
            orientation: nil,
            marginTop: 1985,
            marginRight: 1701,
            marginBottom: 1701,
            marginLeft: 1701,
            marginHeader: 851,
            marginFooter: 992,
            marginGutter: 0,
            columnCount: 2,
            columnSpace: 425,
            headerReferences: [HeaderFooterReference(type: "default", relationshipId: "rId4")],
            footerReferences: [])
        var log = OperationLog()
        log.append(.setSectionProperties(at: nil, section: section), source: .swift)
        log.append(.setSectionProperties(at: ElementID(rawString: "w14:paraId=P2"), section: section),
                   source: .swift)

        let decoded = try OperationLog.decodeJSONL(log.encodeJSONL())
        guard case .setSectionProperties(let at0, let got0) = decoded.entries[0].op else {
            return XCTFail("expected setSectionProperties")
        }
        XCTAssertNil(at0)
        XCTAssertEqual(got0, section)
        guard case .setSectionProperties(let at1, _) = decoded.entries[1].op else {
            return XCTFail("expected setSectionProperties")
        }
        XCTAssertEqual(at1?.raw, "w14:paraId=P2")
        XCTAssertEqual(log.encodeJSONL(), decoded.encodeJSONL(), "byte-equal re-encode")
    }

    // MARK: - Reducer stamping

    private func documentRoot(_ doc: WordDocument) throws -> XmlNode {
        try XCTUnwrap(doc.xmlTrees["word/document.xml"]).root
    }

    private func child(_ node: XmlNode, _ localName: String) -> XmlNode? {
        node.children.first { $0.kind == .element && $0.localName == localName }
    }

    private func attr(_ node: XmlNode, _ localName: String) -> String? {
        node.attributes.first { $0.localName == localName }?.value
    }

    func testAppendParagraphStampsPPrFieldsInSchemaOrder() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: [.appendParagraph(in: nil, paragraph: ParagraphPayload(
            text: "t", styleId: "Body", paraId: "P1",
            alignment: "center",
            spacingBefore: 100, spacingAfter: 200, spacingLine: 240, spacingLineRule: "auto",
            indentLeft: 720, indentFirstLine: 420,
            numId: 5, numLevel: 0))])

        let body = try XCTUnwrap(child(try documentRoot(doc), "body"))
        let p = try XCTUnwrap(child(body, "p"))
        let pPr = try XCTUnwrap(child(p, "pPr"), "pPr must exist")

        let names = pPr.children.filter { $0.kind == .element }.map(\.localName)
        XCTAssertEqual(names, ["pStyle", "numPr", "spacing", "ind", "jc"],
                       "pPr children must follow CT_PPr schema order")

        let numPr = try XCTUnwrap(child(pPr, "numPr"))
        XCTAssertEqual(numPr.children.filter { $0.kind == .element }.map(\.localName),
                       ["ilvl", "numId"])
        XCTAssertEqual(attr(try XCTUnwrap(child(numPr, "ilvl")), "val"), "0")
        XCTAssertEqual(attr(try XCTUnwrap(child(numPr, "numId")), "val"), "5")

        let spacing = try XCTUnwrap(child(pPr, "spacing"))
        XCTAssertEqual(attr(spacing, "before"), "100")
        XCTAssertEqual(attr(spacing, "after"), "200")
        XCTAssertEqual(attr(spacing, "line"), "240")
        XCTAssertEqual(attr(spacing, "lineRule"), "auto")

        let ind = try XCTUnwrap(child(pPr, "ind"))
        XCTAssertEqual(attr(ind, "left"), "720")
        XCTAssertEqual(attr(ind, "firstLine"), "420")
        XCTAssertNil(attr(ind, "right"))

        XCTAssertEqual(attr(try XCTUnwrap(child(pPr, "jc")), "val"), "center")
    }

    func testSetRunsStampsExtendedRPrInSchemaOrder() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(text: "", paraId: "P1")),
            .setRuns(target: ElementID(rawString: "w14:paraId=P1"), runs: [RunPayload(
                text: "字",
                bold: true,
                color: "0000FF",
                fontAscii: "Times New Roman",
                fontEastAsia: "ＭＳ ゴシック",
                sizeHalfPoints: 21,
                underline: "single",
                vertAlign: "superscript")]),
        ])

        let body = try XCTUnwrap(child(try documentRoot(doc), "body"))
        let p = try XCTUnwrap(child(body, "p"))
        let r = try XCTUnwrap(child(p, "r"))
        let rPr = try XCTUnwrap(child(r, "rPr"))

        let names = rPr.children.filter { $0.kind == .element }.map(\.localName)
        XCTAssertEqual(names, ["rFonts", "b", "color", "sz", "u", "vertAlign"],
                       "rPr children must follow CT_RPr schema order")

        let rFonts = try XCTUnwrap(child(rPr, "rFonts"))
        XCTAssertEqual(attr(rFonts, "ascii"), "Times New Roman")
        XCTAssertEqual(attr(rFonts, "eastAsia"), "ＭＳ ゴシック")
        XCTAssertEqual(attr(try XCTUnwrap(child(rPr, "sz")), "val"), "21")
        XCTAssertEqual(attr(try XCTUnwrap(child(rPr, "u")), "val"), "single")
        XCTAssertEqual(attr(try XCTUnwrap(child(rPr, "vertAlign")), "val"), "superscript")
    }

    func testSetSectionPropertiesStampsTrailingBodySectPr() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(text: "t", paraId: "P1")),
            .setSectionProperties(at: nil, section: SectionPayload(
                pageWidth: 11906, pageHeight: 16838,
                marginTop: 1985, marginRight: 1701, marginBottom: 1701, marginLeft: 1701,
                marginHeader: 851, marginFooter: 992, marginGutter: 0,
                columnCount: 2, columnSpace: 425)),
        ])

        let body = try XCTUnwrap(child(try documentRoot(doc), "body"))
        // sectPr must be the LAST body child.
        let last = try XCTUnwrap(body.children.last { $0.kind == .element })
        XCTAssertEqual(last.localName, "sectPr")

        let names = last.children.filter { $0.kind == .element }.map(\.localName)
        XCTAssertEqual(names, ["pgSz", "pgMar", "cols"],
                       "sectPr children must follow CT_SectPr schema order")
        let pgSz = try XCTUnwrap(child(last, "pgSz"))
        XCTAssertEqual(attr(pgSz, "w"), "11906")
        XCTAssertEqual(attr(pgSz, "h"), "16838")
        let cols = try XCTUnwrap(child(last, "cols"))
        XCTAssertEqual(attr(cols, "num"), "2")
        XCTAssertEqual(attr(cols, "space"), "425")
    }

    func testSetSectionPropertiesStampsMidBodySectPrInsidePPr() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "end of section one", styleId: "Body", paraId: "P1")),
            .setSectionProperties(at: ElementID(rawString: "w14:paraId=P1"), section: SectionPayload(
                pageWidth: 11906, pageHeight: 16838, columnSpace: 708)),
        ])

        let body = try XCTUnwrap(child(try documentRoot(doc), "body"))
        let p = try XCTUnwrap(child(body, "p"))
        let pPr = try XCTUnwrap(child(p, "pPr"), "mid-body sectPr requires a pPr host")
        // sectPr must be the LAST pPr child (CT_PPr places sectPr at the end).
        let last = try XCTUnwrap(pPr.children.last { $0.kind == .element })
        XCTAssertEqual(last.localName, "sectPr")
        XCTAssertNotNil(child(last, "pgSz"))
        let cols = try XCTUnwrap(child(last, "cols"))
        XCTAssertEqual(attr(cols, "space"), "708")
        XCTAssertNil(attr(cols, "num"))

        // The trailing paragraph content is untouched.
        XCTAssertNotNil(child(p, "r"))
        // pStyle stays first.
        XCTAssertEqual(pPr.children.first { $0.kind == .element }?.localName, "pStyle")
    }

    /// A trailing setSectionProperties(at: nil) after paragraphs keeps the
    /// sectPr last even when more paragraphs are appended afterwards
    /// (appendParagraph inserts before the trailing sectPr).
    func testAppendAfterTrailingSectPrKeepsSectPrLast() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(text: "a", paraId: "P1")),
            .setSectionProperties(at: nil, section: SectionPayload(pageWidth: 100, pageHeight: 200)),
            .appendParagraph(in: nil, paragraph: ParagraphPayload(text: "b", paraId: "P2")),
        ])
        let body = try XCTUnwrap(child(try documentRoot(doc), "body"))
        let names = body.children.filter { $0.kind == .element }.map(\.localName)
        XCTAssertEqual(names, ["p", "p", "sectPr"])
    }
}
