// DocxFormPublicAPITests.swift — the DocxForm surface a client package sees
// (#238). This target imports DocxForm without @testable, so a test here that
// stops compiling is an API break for the callers outside this package:
// macdoc's `word form` command and che-word-mcp's form tools.

import Foundation
import XCTest
import DocxForm

final class DocxFormPublicAPITests: XCTestCase {
    private static let namespaces = #"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#

    /// A two-column table (label, field) followed by a checkbox line.
    private let form: [UInt8] = Array((
        #"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#
        + "<w:document \(DocxFormPublicAPITests.namespaces)><w:body>"
        + #"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr>"#
        + #"<w:tc><w:p w14:paraId="0000A001"><w:r><w:t>Name</w:t></w:r></w:p></w:tc>"#
        + #"<w:tc><w:p w14:paraId="0000A002"><w:r><w:t>Applicant</w:t></w:r></w:p></w:tc>"#
        + "</w:tr></w:tbl>"
        + "<w:p w14:paraId=\"0000A003\"><w:r><w:t>\u{25A1} Agree</w:t></w:r></w:p>"
        + #"<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>"#
        + "</w:body></w:document>").utf8)

    func testClassifyReportsKindsAndGridPositions() throws {
        let paragraphs = try FormClassifier.classify(documentXML: form, numberingXML: nil, withCells: true)
        XCTAssertEqual(paragraphs.map(\.paraId), ["0000A001", "0000A002", "0000A003"])
        XCTAssertEqual(paragraphs.map(\.kind), [.tableField, .tableField, .checkboxLine])
        XCTAssertEqual(paragraphs.map { $0.cell?.description }, ["t1:r1c1", "t1:r1c2", nil])
        XCTAssertEqual(ParagraphKind(rawValue: "list-checkbox"), .listCheckbox)
        XCTAssertTrue(FormClassifier.containsBoxGlyph("\u{2610}"))
    }

    func testFieldPlanSetsTheNamedSize() throws {
        let size = try FieldSize.parse(points: "12")
        XCTAssertEqual(size, FieldSize(halfPoints: 24))
        let plan = FormPlan(fields: ["0000A002"], fieldSize: size)
        XCTAssertTrue(plan.hasOperations)
        let result = try plan.apply(toDocumentXML: form)
        XCTAssertEqual(result.fieldSize, size)
        XCTAssertFalse(result.fieldSizeWasDerived)
        XCTAssertTrue(String(decoding: result.documentXML, as: UTF8.self).contains(#"w:val="24""#))
    }

    func testTableDefaultsReportCountsTheFillCells() throws {
        let defaults = try TableDefaults.parse("0000A001=col:2")
        XCTAssertEqual(defaults.anchor, "0000A001")
        let result = try FormPlan(fieldSize: FieldSize(halfPoints: 24), tableDefaults: [defaults])
            .apply(toDocumentXML: form)
        XCTAssertEqual(result.tableDefaultsReports,
                       [TableDefaultsReport(anchor: "0000A001", fillCells: 1, partialCells: [])])
    }

    func testRefusalsAreFormErrors() {
        XCTAssertThrowsError(try FormPlan(fields: ["FFFFFFFF"]).apply(toDocumentXML: form)) { error in
            XCTAssertTrue(error is FormError, "\(error)")
        }
        XCTAssertThrowsError(try TableDefaults.parse("0000A001=col:0")) { error in
            XCTAssertTrue(error is FormError, "\(error)")
        }
    }

    /// Touches every DocxForm name that macdoc's `word form` command and its
    /// CLI tests use (macdoc 41b6ef38), so a name that stops being public fails
    /// here instead of in macdoc (verify round 1).
    func testEveryNameMacdocsWordFormUsesIsPublic() throws {
        let plan = FormPlan(fields: ["0000A002"], fieldSize: FieldSize(halfPoints: 24), checkboxFont: "MS Gothic",
                            tickList: [], insertions: [FormInsertion(anchor: "0000A003", text: "Note")],
                            removals: [], removedTables: [], truncateFrom: nil, tableDefaults: [])
        let result: FormResult = try plan.apply(toDocumentXML: form)
        XCTAssertEqual(result.skippedCheckboxRuns, 0)
        XCTAssertFalse(result.addedFinalParagraph)
        let skips: [CleanupSkip] = result.skippedCleanups
        XCTAssertEqual(skips.map(\.paragraphId), [])
        XCTAssertEqual(skips.flatMap(\.reasons), [])
        XCTAssertEqual(CleanupSkip.Reason.allCases.map(\.elementName), ["w:tbl", "w:sdt", "w:customXml"])
        XCTAssertEqual(FieldSize(halfPoints: 21).pointsDescription, "10.5 pt")

        let report = try FormPlan(fieldSize: FieldSize(halfPoints: 24),
                                  tableDefaults: [TableDefaults.parse("0000A001=col:2")])
            .apply(toDocumentXML: form).tableDefaultsReports[0]
        XCTAssertEqual(report.partialCells, [])
        XCTAssertEqual(report.skippedCells, [])

        XCTAssertEqual(try FormClassifier.kinds(of: ["0000A002", "0000A003"], documentXML: form, numberingXML: nil),
                       ["0000A002": .tableField, "0000A003": .checkboxLine])
        XCTAssertFalse(FormClassifier.isReadableNumbering(Array("not xml".utf8)))
        let paragraphs = try FormClassifier.classify(documentXML: form, numberingXML: nil)
        XCTAssertEqual(paragraphs.last?.text, "\u{25A1} Agree")

        let mergedTable: [UInt8] = Array((
            "<w:document \(DocxFormPublicAPITests.namespaces)><w:body><w:tbl><w:tblPr/><w:tblGrid/><w:tr>"
            + #"<w:tc><w:tcPr><w:hMerge w:val="restart"/></w:tcPr><w:p w14:paraId="0000B001"/></w:tc>"#
            + "</w:tr></w:tbl></w:body></w:document>").utf8)
        let cell = try FormClassifier.classify(documentXML: mergedTable, numberingXML: nil, withCells: true).first?.cell
        guard case .unreadableTable(let table)? = cell else { return XCTFail("\(String(describing: cell))") }
        XCTAssertEqual(table, 1)
        XCTAssertEqual(cell?.description, "t1:?")

        XCTAssertThrowsError(try FormClassifier.classify(documentXML: Array("<w:document><w:body>".utf8),
                                                         numberingXML: nil)) {
            XCTAssertTrue($0 is XMLScanError, "\($0)")
        }
        XCTAssertFalse(FormError.malformedPart("x").description.isEmpty)
    }

    func testMissingDocumentPartIsAPackageError() throws {
        let dir = FileManager.default.temporaryDirectory
            .appendingPathComponent("DocxFormPublicAPITests-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: dir) }
        let input = dir.appendingPathComponent("in.docx")
        try DocxPackage.write([DocxPackage.Entry(path: "word/styles.xml", data: Data("<w:styles/>".utf8))], to: input)
        XCTAssertEqual(try DocxPackage.read(from: input).map(\.path), ["word/styles.xml"])
        XCTAssertEqual(DocxPackage.numberingPart, "word/numbering.xml")
        XCTAssertThrowsError(try DocxPackage.rewriteDocument(input: input, output: dir.appendingPathComponent("out.docx"),
                                                             force: false) { d, _ in d }) {
            XCTAssertEqual($0 as? DocxPackageError, .missingPart(DocxPackage.documentPart))
        }
    }

    func testRewriteReplacesOnlyTheDocumentPart() throws {
        let dir = FileManager.default.temporaryDirectory
            .appendingPathComponent("DocxFormPublicAPITests-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: dir) }
        let input = dir.appendingPathComponent("in.docx")
        let output = dir.appendingPathComponent("out.docx")
        let styles = Data("<w:styles/>".utf8)
        try DocxPackage.write([DocxPackage.Entry(path: DocxPackage.documentPart, data: Data(form)),
                               DocxPackage.Entry(path: "word/styles.xml", data: styles)], to: input)

        try DocxPackage.rewriteDocument(input: input, output: output, force: false) { document, _ in
            Data(try FormPlan(fields: ["0000A002"], fieldSize: FieldSize(halfPoints: 24))
                .apply(toDocumentXML: Array(document)).documentXML)
        }

        let parts = try DocxPackage.readParts(from: output, paths: ["word/styles.xml", DocxPackage.documentPart])
        XCTAssertEqual(parts.map(\.path), [DocxPackage.documentPart, "word/styles.xml"])
        XCTAssertEqual(parts.last?.data, styles)
        XCTAssertNotEqual(parts.first?.data, Data(form))
        XCTAssertThrowsError(try DocxPackage.rewriteDocument(input: input, output: output, force: false) { d, _ in d }) {
            XCTAssertEqual($0 as? DocxPackageError, .outputExists(output.path))
        }
    }
}
