import XCTest
@testable import OOXMLSwift

/// PsychQuant/ooxml-swift#155 — a cell holding a nested table must not gain a
/// paragraph every time the document is re-serialised from the typed model.
///
/// Measured on real documents before the fix (v3.7.0): the same file, one
/// typed-dirty operation per generation, went 9 -> 14 -> 19 -> 24 cell
/// paragraphs. Unbounded, and the count of affected cells equalled the
/// document's nested-table count exactly.
final class Issue155NestedTableCellTests: XCTestCase {

    /// A cell whose content is: one paragraph, a nested table, and the trailing
    /// paragraph OOXML requires after it — the shape a real document has.
    private func cellWithNestedTable() -> TableCell {
        var inner = Table()
        var innerRow = TableRow()
        var innerCell = TableCell()
        innerCell.paragraphs = [Paragraph(text: "inner")]
        innerRow.cells = [innerCell]
        inner.rows = [innerRow]

        var cell = TableCell()
        cell.paragraphs = [Paragraph(text: "before"), Paragraph()]   // the trailing empty one
        cell.nestedTables = [inner]
        return cell
    }

    private func paragraphCount(inCellXML xml: String) -> Int {
        // count <w:p> openers that belong to THIS cell, not the nested table's
        var depth = 0, count = 0
        var rest = Substring(xml)
        while let i = rest.firstIndex(of: "<") {
            rest = rest[rest.index(after: i)...]
            if rest.hasPrefix("w:tbl>") || rest.hasPrefix("w:tbl ") { depth += 1 }
            else if rest.hasPrefix("/w:tbl>") { depth -= 1 }
            else if depth == 0, rest.hasPrefix("w:p>") || rest.hasPrefix("w:p ") || rest.hasPrefix("w:p/>") { count += 1 }
        }
        return count
    }

    /// The defect, stated as a fixed point: emitting a cell, reading the result
    /// back and emitting it again must produce the same number of paragraphs.
    func testReEmittingACellWithANestedTableReachesAFixedPoint() throws {
        let cell = cellWithNestedTable()
        let gen1 = cell.toXML()
        let n1 = paragraphCount(inCellXML: gen1)

        // Re-read gen1 as a tree-backed cell and emit again — the generation
        // loop the real corpus went through.
        let doc = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:tbl><w:tr>\(gen1)</w:tr></w:tbl></w:body></w:document>"
        let tree = try XmlTreeReader.parse(Data(doc.utf8))
        let body = try XCTUnwrap(tree.root.children.first { $0.localName == "body" })
        let tbl = try XCTUnwrap(body.children.first { $0.localName == "tbl" })
        let tr = try XCTUnwrap(tbl.children.first { $0.localName == "tr" })
        let tc = try XCTUnwrap(tr.children.first { $0.localName == "tc" })
        let gen2 = TableCell(xmlNode: tc).toXML()
        let n2 = paragraphCount(inCellXML: gen2)

        XCTAssertEqual(n2, n1, "a cell with a nested table must not grow a paragraph per save (was \(n1) -> \(n2))")
    }

    /// The rule the trailing paragraph exists for still holds: after a nested
    /// table, the cell ends with a paragraph.
    func testACellStillEndsWithAParagraphAfterItsNestedTable() {
        let xml = cellWithNestedTable().toXML()
        let afterLastTable = xml.range(of: "</w:tbl>", options: .backwards).map { String(xml[$0.upperBound...]) } ?? ""
        XCTAssertTrue(afterLastTable.contains("<w:p"), "a paragraph follows the nested table: \(afterLastTable)")
    }

    /// The paragraph that follows the nested table is the cell's own trailing
    /// paragraph, carried across — not a blank one substituted for it. Holding
    /// back only EMPTY trailing paragraphs was tried first and left a real form
    /// gaining a paragraph per nested-table cell on its first save.
    func testTheTrailingParagraphKeepsItsContent() {
        var inner = Table()
        var innerRow = TableRow()
        var innerCell = TableCell()
        innerCell.paragraphs = [Paragraph(text: "inner")]
        innerRow.cells = [innerCell]
        inner.rows = [innerRow]

        var cell = TableCell()
        cell.paragraphs = [Paragraph(text: "before"), Paragraph(text: "after the table")]
        cell.nestedTables = [inner]

        let xml = cell.toXML()
        let tail = xml.range(of: "</w:tbl>", options: .backwards).map { String(xml[$0.upperBound...]) } ?? ""
        XCTAssertTrue(tail.contains("after the table"),
                      "the cell's own trailing paragraph follows the table: \(tail)")
        XCTAssertEqual(paragraphCount(inCellXML: xml), 2, "no paragraph is added")
    }

    /// A cell with no nested table is untouched by any of this.
    func testACellWithoutANestedTableIsUnchanged() {
        var cell = TableCell()
        cell.paragraphs = [Paragraph(text: "only")]
        XCTAssertEqual(paragraphCount(inCellXML: cell.toXML()), 1)
    }

    // MARK: - Real documents (gated)

    /// The generation loop the defect was measured with, over a directory of
    /// real documents. Off by default — point `OOXML_CORPUS_DIR` at a folder of
    /// .docx files to run it. Measured on a real 2-nested-table government form:
    /// 12 -> 14 -> 14 -> 14 before this fix's final form, 12 -> 12 -> 12 -> 12 after.
    func testRealDocumentsSurviveRepeatedSaves() throws {
        guard let dir = ProcessInfo.processInfo.environment["OOXML_CORPUS_DIR"] else {
            throw XCTSkip("set OOXML_CORPUS_DIR to a folder of .docx files to run this")
        }
        let root = URL(fileURLWithPath: NSString(string: dir).expandingTildeInPath)
        let files = ((try? FileManager.default.contentsOfDirectory(at: root, includingPropertiesForKeys: nil)) ?? [])
            .filter { $0.pathExtension == "docx" && !$0.lastPathComponent.hasPrefix("~$") }
            .sorted { $0.lastPathComponent < $1.lastPathComponent }

        var examined = 0
        for f in files {
            guard let doc = try? DocxReader.read(from: f), nestedTableCells(doc).cells > 0 else { continue }
            examined += 1

            var counts: [Int] = []
            var cur = f
            for gen in 0..<4 {
                var d = try DocxReader.read(from: cur)
                counts.append(nestedTableCells(d).paragraphs)
                if gen == 3 { break }
                d.markTypedDirty("word/document.xml")           // the trigger: any typed-dirty op
                let out = URL(fileURLWithPath: NSTemporaryDirectory())
                    .appendingPathComponent("issue155-\(examined)-\(gen).docx")
                try DocxWriter.write(d, to: out)
                cur = out
            }
            XCTAssertEqual(Set(counts).count, 1,
                           "\(f.lastPathComponent): cell paragraphs moved across saves: \(counts)")
        }
        print("issue155: examined \(examined) document(s) holding nested tables")
    }

    /// Cells that hold a nested table, and how many paragraphs they hold in total.
    private func nestedTableCells(_ doc: WordDocument) -> (cells: Int, paragraphs: Int) {
        var pending: [Table] = doc.body.tables
        for child in doc.body.children { if case .table(let t) = child { pending.append(t) } }
        var cells = 0, paragraphs = 0
        while let t = pending.popLast() {
            for row in t.rows {
                for cell in row.cells {
                    if !cell.nestedTables.isEmpty { cells += 1; paragraphs += cell.paragraphs.count }
                    pending.append(contentsOf: cell.nestedTables)
                }
            }
        }
        return (cells, paragraphs)
    }
}
