import Foundation
import XCTest
@testable import OOXMLSwift

/// Regression suite for PsychQuant/che-word-mcp#224, sub-problem (b):
/// `<w:trPr><w:jc>` (row-level horizontal justification) was silently
/// dropped on every save — `TableRowProperties` had no field to hold it at
/// all — and any `<w:trHeight>` that originally had no `w:hRule` attribute
/// gained a manufactured `w:hRule="auto"` out of nowhere. Because
/// `DocxWriter` re-serializes the whole `document.xml` from the typed model
/// once any table mutation marks it dirty, both defects hit every row in
/// the file, not just the one being edited (the issue's "18 of 18 untouched
/// rows" symptom).
final class TableRowJustificationAndHeightRuleFidelityTests: XCTestCase {

    // MARK: - Unit-level: writer shape

    func testRowJustificationWritesJcElement() {
        var rowProps = TableRowProperties()
        rowProps.justification = .center
        let row = TableRow(cells: [TableCell(text: "J")], properties: rowProps)
        let xml = row.toXML()
        XCTAssertTrue(xml.contains("<w:trPr>"))
        XCTAssertTrue(xml.contains("<w:jc w:val=\"center\"/>"))
    }

    /// che-word-mcp#224(b): a `<w:trHeight>` that never carried `w:hRule`
    /// must not gain one on save. `heightRule == nil` now means "omit the
    /// attribute", not "assume auto" — OOXML's own default for `w:hRule`
    /// is already "auto" (ECMA-376 §17.4.88), so this changes zero
    /// documents' rendered behavior while stopping the manufactured
    /// attribute from showing up as a spurious diff on unrelated rows.
    func testRowHeightWithoutExplicitRuleOmitsHRuleAttribute() {
        var rowProps = TableRowProperties()
        rowProps.height = 303
        let row = TableRow(cells: [TableCell(text: "H")], properties: rowProps)
        let xml = row.toXML()
        XCTAssertTrue(xml.contains("<w:trHeight w:val=\"303\"/>"), xml)
        XCTAssertFalse(xml.contains("w:hRule"), xml)
    }

    /// An explicitly-set `.auto` must still be written out — only *absence*
    /// of a rule is suppressed, not the `.auto` case itself. This keeps the
    /// nil-vs-.auto distinction meaningful for any caller that cares.
    func testRowHeightWithExplicitAutoRuleStillWritesAttribute() {
        var rowProps = TableRowProperties()
        rowProps.height = 303
        rowProps.heightRule = .auto
        let row = TableRow(cells: [TableCell(text: "H")], properties: rowProps)
        let xml = row.toXML()
        XCTAssertTrue(xml.contains("<w:trHeight w:val=\"303\" w:hRule=\"auto\"/>"), xml)
    }

    func testRowHeightWithExactRuleUnaffected() {
        var rowProps = TableRowProperties()
        rowProps.height = 400
        rowProps.heightRule = .exact
        let row = TableRow(cells: [TableCell(text: "H")], properties: rowProps)
        let xml = row.toXML()
        XCTAssertTrue(xml.contains("<w:trHeight w:val=\"400\" w:hRule=\"exact\"/>"), xml)
    }

    /// While auditing `<w:trPr>` for other unmodeled children (che-word-mcp#224(b)
    /// asked us to check), `<w:hidden/>` (ECMA-376 §17.4.36) turned
    /// out to be the exact same shape as the already-modeled `tblHeader`/
    /// `cantSplit` flags — a bare boolean, trivial and low-risk to add
    /// alongside the `jc` fix rather than leaving it for a future pass.
    func testRowHiddenFlagWritesHiddenElement() {
        var rowProps = TableRowProperties()
        rowProps.hidden = true
        let row = TableRow(cells: [TableCell(text: "H")], properties: rowProps)
        let xml = row.toXML()
        XCTAssertTrue(xml.contains("<w:trPr>"))
        XCTAssertTrue(xml.contains("<w:hidden/>"))
    }

    /// Round-trip companion to the unit test above: row 0 carries
    /// `<w:hidden/>` (injected into the seed, since the typed writer had no
    /// field to produce it from pre-fix); row 1 is the only row edited.
    /// Row 0's flag must survive.
    func testRowHiddenFlagSurvivesEditingAnotherRow() throws {
        let directory = FileManager.default.temporaryDirectory
            .appendingPathComponent("row-hidden-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: directory, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: directory) }

        var seed = WordDocument()
        seed.appendTable(Table(rows: [
            TableRow(cells: [TableCell(text: "H0")]),
            TableRow(cells: [TableCell(text: "H1")]),
        ]))
        let seedURL = directory.appendingPathComponent("seed.docx")
        try DocxWriter.write(seed, to: seedURL)

        let extracted = try ZipHelper.unzip(seedURL)
        defer { ZipHelper.cleanup(extracted) }
        let documentURL = extracted.appendingPathComponent("word/document.xml")
        var xml = try String(contentsOf: documentURL, encoding: .utf8)

        guard let anchorRange = xml.range(of: ">H0<"),
              let trRange = xml.range(
                of: "<w:tr>", options: .backwards, range: xml.startIndex..<anchorRange.lowerBound)
        else {
            XCTFail("could not locate row 0's opening <w:tr> tag")
            return
        }
        xml.insert(contentsOf: "<w:trPr><w:hidden/></w:trPr>", at: trRange.upperBound)
        try xml.write(to: documentURL, atomically: true, encoding: .utf8)

        let inputURL = directory.appendingPathComponent("input.docx")
        try ZipHelper.zip(extracted, to: inputURL)

        var document = try DocxReader.read(from: inputURL)
        defer { document.close() }
        XCTAssertTrue(document.getTables()[0].rows[0].properties.hidden)

        try document.updateCell(tableIndex: 0, row: 1, col: 0, text: "H1-edited")

        let outputURL = directory.appendingPathComponent("output.docx")
        try DocxWriter.write(document, to: outputURL)
        let outputPackage = try ZipHelper.unzip(outputURL)
        defer { ZipHelper.cleanup(outputPackage) }
        let outputXML = try String(
            contentsOf: outputPackage.appendingPathComponent("word/document.xml"),
            encoding: .utf8)

        XCTAssertTrue(outputXML.contains("<w:hidden/>"))
        XCTAssertTrue(outputXML.contains(">H1-edited<"))
    }

    // MARK: - Round-trip: the scenario from the issue

    /// Builds a 3-row, 1-column table where every row's `<w:trPr>` carries
    /// `<w:trHeight w:val="303"/><w:jc w:val="center"/>` — no `w:hRule`, a
    /// row-level `jc` the typed model previously had no field for at all.
    /// We edit only the middle row and assert *every* row's trPr (including
    /// the two untouched ones, and the edited row's own trPr) still has the
    /// justification and still lacks a manufactured `w:hRule`.
    func testEditingOneRowPreservesJustificationAndOmittedHeightRuleOnAllRows() throws {
        let directory = FileManager.default.temporaryDirectory
            .appendingPathComponent("row-jc-hrule-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: directory, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: directory) }

        var seed = WordDocument()
        seed.appendTable(Table(rows: [
            TableRow(cells: [TableCell(text: "R0")]),
            TableRow(cells: [TableCell(text: "R1")]),
            TableRow(cells: [TableCell(text: "R2")]),
        ]))
        let seedURL = directory.appendingPathComponent("seed.docx")
        try DocxWriter.write(seed, to: seedURL)

        let extracted = try ZipHelper.unzip(seedURL)
        defer { ZipHelper.cleanup(extracted) }
        let documentURL = extracted.appendingPathComponent("word/document.xml")
        var xml = try String(contentsOf: documentURL, encoding: .utf8)

        // `<w:tr>` (no attributes, since none of the seeded rows set
        // isHeader/height/cantSplit) is the row-opening tag and nothing
        // else, so a global replace safely injects the same trPr into all
        // three rows in one shot.
        XCTAssertEqual(xml.components(separatedBy: "<w:tr>").count - 1, 3)
        xml = xml.replacingOccurrences(
            of: "<w:tr>",
            with: "<w:tr><w:trPr><w:trHeight w:val=\"303\"/><w:jc w:val=\"center\"/></w:trPr>")
        try xml.write(to: documentURL, atomically: true, encoding: .utf8)

        let inputURL = directory.appendingPathComponent("input.docx")
        try ZipHelper.zip(extracted, to: inputURL)

        var document = try DocxReader.read(from: inputURL)
        defer { document.close() }

        // Sanity check the read side before exercising the "survives an
        // edit" question, so a failure below is unambiguous about which
        // half of the pipeline is at fault.
        let seededRows = document.getTables()[0].rows
        for row in seededRows {
            XCTAssertEqual(row.properties.justification, .center)
            XCTAssertNil(row.properties.heightRule)
        }

        try document.updateCell(tableIndex: 0, row: 1, col: 0, text: "R1-edited")

        let outputURL = directory.appendingPathComponent("output.docx")
        try DocxWriter.write(document, to: outputURL)
        let outputPackage = try ZipHelper.unzip(outputURL)
        defer { ZipHelper.cleanup(outputPackage) }
        let outputXML = try String(
            contentsOf: outputPackage.appendingPathComponent("word/document.xml"),
            encoding: .utf8)

        XCTAssertEqual(
            outputXML.components(separatedBy: "<w:jc w:val=\"center\"/>").count - 1,
            3,
            "all three rows' justification must survive, including the untouched ones")
        XCTAssertFalse(
            outputXML.contains("w:hRule"),
            "no row originally had w:hRule; none may gain one from an unrelated edit")
        XCTAssertEqual(
            outputXML.components(separatedBy: "<w:trHeight w:val=\"303\"/>").count - 1,
            3)
        XCTAssertTrue(outputXML.contains(">R0<"))
        XCTAssertTrue(outputXML.contains(">R1-edited<"))
        XCTAssertTrue(outputXML.contains(">R2<"))
    }
}
