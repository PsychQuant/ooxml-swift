// UpgradeClassGuardTests.swift
// format-alignment-engine Phase C task 3.3 — the raw→DSL upgrade-class
// regression pin (`format-alignment-pipeline`, «Dual-track acceptance»;
// Decision 2). One parameterized test enumerates every content class that
// has been upgraded from the raw channel to the typed DSL channel and
// asserts Stage B stays green when that class is exercised. A class that
// regresses to non-byte-equal fails HERE, naming the class.

import XCTest
@testable import OOXMLSwift

final class UpgradeClassGuardTests: XCTestCase {

    /// Every shipped upgrade class with a representative op sequence that
    /// exercises it. Grows as later phases upgrade more classes.
    private static let upgradeClasses: [(name: String, ops: [OOXMLSwift.Operation])] = [
        ("text+style (wass baseline)", [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "見出しと本文", styleId: "Heading1", paraId: "P1")),
        ]),
        ("run-formatting (2.2)", [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(text: "", paraId: "P1")),
            .setRuns(target: ElementID(rawString: "w14:paraId=P1"), runs: [
                RunPayload(text: "ゴシック", bold: true, italic: true, color: "336699",
                           fontAscii: "Times New Roman", fontEastAsia: "ＭＳ ゴシック",
                           sizeHalfPoints: 21, underline: "single", vertAlign: "superscript"),
            ]),
        ]),
        ("paragraph-formatting (2.3)", [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "整形段落", paraId: "P1", alignment: "center",
                spacingBefore: 100, spacingAfter: 200, spacingLine: 240,
                spacingLineRule: "auto", indentLeft: 720, indentHanging: 360,
                numId: 2, numLevel: 0)),
        ]),
        ("sections (2.4)", [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(text: "第一節", paraId: "P1")),
            .setSectionProperties(at: ElementID(rawString: "w14:paraId=P1"), section: SectionPayload(
                pageWidth: 11906, pageHeight: 16838, columnSpace: 708)),
            .appendParagraph(in: nil, paragraph: ParagraphPayload(text: "第二節", paraId: "P2")),
            .setSectionProperties(at: nil, section: SectionPayload(
                pageWidth: 11906, pageHeight: 16838, marginTop: 1985, marginRight: 1701,
                marginBottom: 1701, marginLeft: 1701, columnCount: 2, columnSpace: 425)),
        ]),
        ("tables (2.5)", [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(text: "表", paraId: "P1")),
            .appendTable(in: nil, table: TablePayload(rows: 2, columns: 2, cells: [
                ["a", "b"], ["c", "d"],
            ])),
        ]),
        ("document-root (wcf 2.1)", [
            .setDocumentRoot(attributes: [
                RootAttribute(prefix: "xmlns", localName: "w", value: "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
                RootAttribute(prefix: "xmlns", localName: "w14", value: "http://schemas.microsoft.com/office/word/2010/wordml"),
                RootAttribute(prefix: "xmlns", localName: "mc", value: "http://schemas.openxmlformats.org/markup-compatibility/2006"),
                RootAttribute(prefix: "mc", localName: "Ignorable", value: "w14"),
            ]),
            .appendParagraph(in: nil, paragraph: ParagraphPayload(text: "根", paraId: "P1")),
        ]),
        ("rsid + textId (wcf 2.2)", [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "", paraId: "1FE40057", textId: "740FF560",
                rsidR: "00040150", rsidRPr: "007F04E2", rsidRDefault: "00C74BEF", rsidP: "00F32D54")),
            .setRuns(target: ElementID(rawString: "w14:paraId=1FE40057"),
                     runs: [RunPayload(text: "本文", rsidRPr: "007F04E2")]),
        ]),
        ("xml:space preserve (wcf 2.3)", [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(text: "", paraId: "P1")),
            .setRuns(target: ElementID(rawString: "w14:paraId=P1"),
                     runs: [RunPayload(text: " spaced ", preserveSpace: true)]),
        ]),
        ("inline markers (wcf 2.4)", [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(text: "", paraId: "P1")),
            .setParagraphContent(target: ElementID(rawString: "w14:paraId=P1"), items: [
                .marker(InlineMarker(localName: "bookmarkStart", attributes: [
                    RootAttribute(prefix: "w", localName: "id", value: "0"),
                    RootAttribute(prefix: "w", localName: "name", value: "_Hlk96608833")])),
                .marker(InlineMarker(localName: "proofErr", attributes: [
                    RootAttribute(prefix: "w", localName: "type", value: "gramStart")])),
                .run(RunPayload(text: "本文")),
                .marker(InlineMarker(localName: "bookmarkEnd", attributes: [
                    RootAttribute(prefix: "w", localName: "id", value: "0")])),
            ]),
        ]),
        ("pPr/rPr + rFonts long-tail + docGrid + prolog (wcf 3.1)", [
            .setDocumentProlog(prolog: "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"),
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "", paraId: "P1",
                indentFirstLine: 180, indentFirstLineChars: 100,
                paragraphMarkRun: RunPayload(
                    text: "", fontAscii: "Times New Roman", sizeHalfPoints: 36,
                    fontHAnsi: "Times New Roman", sizeCsHalfPoints: 36))),
            .setRuns(target: ElementID(rawString: "w14:paraId=P1"), runs: [RunPayload(
                text: "ゴシック", fontEastAsia: "ＭＳ ゴシック", sizeHalfPoints: 21,
                fontHAnsi: "ＭＳ ゴシック", fontHint: "eastAsia",
                boldCs: true, italicCs: true, sizeCsHalfPoints: 21)]),
            .setSectionProperties(at: nil, section: SectionPayload(
                pageWidth: 11906, pageHeight: 16838,
                marginTop: 1985, columnSpace: 425,
                rsidR: "0081382F", rsidRPr: "006C42D6", rsidSect: "00BC5145",
                docGridType: "linesAndChars", docGridLinePitch: 286,
                pageSizeCode: 9, sectionType: "continuous")),
        ]),
    ]

    /// Parameterized guard: every upgrade class must (1) actually upgrade
    /// (document.xml on the DSL channel) and (2) hold Stage B byte equality
    /// through the full script round-trip.
    func testEveryUpgradeClassKeepsStageBGreen() throws {
        for upgradeClass in Self.upgradeClasses {
            // Build the reference package for this class.
            var doc = WordDocument.emptyAuthoringDocument()
            try doc.apply(operations: upgradeClass.ops)
            let refURL = FileManager.default.temporaryDirectory
                .appendingPathComponent("guard-\(UUID().uuidString).docx")
            defer { try? FileManager.default.removeItem(at: refURL) }
            try doc.writeAuthoringPackage(to: refURL)
            let reference = try RawPartChannel.readAllParts(from: refURL)

            // Reverse with upgrades → script → execute → compare.
            let result = try ReverseExtractor.reverse(parts: reference)
            XCTAssertTrue(
                result.dslParts.contains("word/document.xml"),
                "[\(upgradeClass.name)] regressed: no longer upgrades to the DSL channel "
                + "(rawReasons: \(result.rawReasons["word/document.xml"] ?? "-"))")

            let script = ScriptExporter.exportSwift(log: result.log)
            let parsed = try ScriptImporter.parse(source: script)
            var rebuiltDoc = WordDocument.emptyAuthoringDocument()
            try rebuiltDoc.apply(operations: parsed.entries.map(\.op))
            let outURL = FileManager.default.temporaryDirectory
                .appendingPathComponent("guard-out-\(UUID().uuidString).docx")
            defer { try? FileManager.default.removeItem(at: outURL) }
            try rebuiltDoc.writeAuthoringPackage(to: outURL)
            let rebuilt = try RawPartChannel.readAllParts(from: outURL)

            let verdicts = PartFidelity.compareParts(reference: reference, rebuilt: rebuilt)
            let broken = verdicts.filter { !$0.isEqual }
            XCTAssertTrue(
                broken.isEmpty,
                "[\(upgradeClass.name)] Stage B regressed on: "
                + broken.map { "\($0.partPath) (\($0.status))" }.joined(separator: ", "))
        }
    }
}
