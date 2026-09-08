import XCTest
@testable import OOXMLSwift

/// PsychQuant/ooxml-swift#155 — a cell holding a nested table must survive
/// re-serialisation with its children in the same ORDER and the same COUNT.
///
/// The count alone is not enough, and that is the whole lesson of this issue:
/// the first fix attempt held the last paragraph back, which kept every count
/// stable while moving user content across the table. A corpus measurement of
/// "12 -> 12 -> 12" was green for that broken fix.
final class Issue155NestedTableCellTests: XCTestCase {

    private let innerTable = "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid>"
        + "<w:tr><w:tc><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"

    /// Parse `<w:tc>` content into a tree-backed cell, the way a read document
    /// produces one.
    private func cell(holding inner: String) throws -> TableCell {
        let doc = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            + "<w:body><w:tbl><w:tr><w:tc>\(inner)</w:tc></w:tr></w:tbl></w:body></w:document>"
        let tree = try XmlTreeReader.parse(Data(doc.utf8))
        let body = try XCTUnwrap(tree.root.children.first { $0.localName == "body" })
        let tbl = try XCTUnwrap(body.children.first { $0.localName == "tbl" })
        let tr = try XCTUnwrap(tbl.children.first { $0.localName == "tr" })
        let tc = try XCTUnwrap(tr.children.first { $0.localName == "tc" })
        return TableCell(xmlNode: tc)
    }

    private func para(_ t: String) -> String { "<w:p><w:r><w:t>\(t)</w:t></w:r></w:p>" }

    /// The cell's own block children, in emitted order: "p:<text>" / "tbl".
    private func blockOrder(_ xml: String) -> [String] {
        var out: [String] = []
        var depth = 0
        var rest = Substring(xml)
        var pendingParagraph = false
        var text = ""
        while let i = rest.firstIndex(of: "<") {
            rest = rest[rest.index(after: i)...]
            if rest.hasPrefix("w:tbl>") || rest.hasPrefix("w:tbl ") {
                if depth == 0 { out.append("tbl") }
                depth += 1
            } else if rest.hasPrefix("/w:tbl>") {
                depth -= 1
            } else if depth == 0, rest.hasPrefix("w:p>") || rest.hasPrefix("w:p ") {
                pendingParagraph = true; text = ""
            } else if depth == 0, rest.hasPrefix("w:p/>") {
                out.append("p:")
            } else if depth == 0, pendingParagraph, rest.hasPrefix("w:t>") || rest.hasPrefix("w:t ") {
                if let close = rest.firstIndex(of: ">"), let end = rest.range(of: "</w:t>") {
                    text += rest[rest.index(after: close)..<end.lowerBound]
                }
            } else if depth == 0, rest.hasPrefix("/w:p>"), pendingParagraph {
                out.append("p:\(text)"); pendingParagraph = false
            }
        }
        return out
    }

    // MARK: - The reported defect

    /// The growth: emit, read back, emit again must give the same cell.
    func testReEmittingACellWithANestedTableIsAFixedPoint() throws {
        let source = para("before") + innerTable + "<w:p/>"
        let gen1 = try cell(holding: source).toXML()
        let gen2 = try cell(holding: String(gen1.dropFirst("<w:tc>".count).dropLast("</w:tc>".count))).toXML()
        XCTAssertEqual(gen2, gen1, "a cell with a nested table must re-emit identically")
        XCTAssertEqual(blockOrder(gen1), ["p:before", "tbl", "p:"], "and in the source order")
    }

    // MARK: - Order, which the count cannot see

    /// A cell whose LAST block child is a table. Holding back the last
    /// paragraph reversed this into `table, before`.
    func testACellEndingWithATableKeepsItsParagraphFirst() throws {
        let xml = try cell(holding: para("before") + innerTable).toXML()
        XCTAssertEqual(blockOrder(xml), ["p:before", "tbl"],
                       "the paragraph came before the table in the source and must stay there")
    }

    /// ONE nested table with two paragraphs after it — an ordinary cell.
    /// Both the original flattening and the hold-back attempt moved `B`.
    func testOneNestedTableWithTwoTrailingParagraphs() throws {
        let xml = try cell(holding: para("A") + innerTable + para("B") + para("C")).toXML()
        XCTAssertEqual(blockOrder(xml), ["p:A", "tbl", "p:B", "p:C"],
                       "one nested table is already enough to move a paragraph if order is inferred")
    }

    /// Several nested tables with paragraphs between them — documented as
    /// unfixable while the emit flattened; the ordered walk handles it.
    func testParagraphsBetweenSeveralNestedTables() throws {
        let xml = try cell(holding: para("A") + innerTable + para("B") + innerTable + para("C")).toXML()
        XCTAssertEqual(blockOrder(xml), ["p:A", "tbl", "p:B", "tbl", "p:C"])
    }

    // MARK: - Unchanged behaviour

    func testACellWithoutANestedTableIsUnchanged() throws {
        XCTAssertEqual(blockOrder(try cell(holding: para("only")).toXML()), ["p:only"])
    }

    /// An empty tree-backed cell still emits the one paragraph a cell requires.
    func testAnEmptyCellStillEmitsAParagraph() throws {
        XCTAssertEqual(blockOrder(try cell(holding: "").toXML()), ["p:"])
    }

    /// Detached cells have no recorded order, so they keep the pre-#155 shape:
    /// paragraphs, tables, then an added trailing paragraph.
    func testADetachedCellKeepsItsPreviousShape() {
        var inner = Table()
        var row = TableRow(); var innerCell = TableCell()
        innerCell.paragraphs = [Paragraph(text: "inner")]
        row.cells = [innerCell]; inner.rows = [row]

        var c = TableCell()
        c.paragraphs = [Paragraph(text: "A")]
        c.nestedTables = [inner]
        XCTAssertEqual(blockOrder(c.toXML()), ["p:A", "tbl", "p:"])
    }

    // MARK: - Real documents (gated)

    /// The generation loop over real documents, comparing the EMITTED CELL XML
    /// rather than a paragraph count — a count is stable under exactly the
    /// corruption this issue is about. Point `OOXML_CORPUS_DIR` at a folder of
    /// .docx files to run it.
    func testRealDocumentsReEmitIdentically() throws {
        guard let dir = ProcessInfo.processInfo.environment["OOXML_CORPUS_DIR"] else {
            throw XCTSkip("set OOXML_CORPUS_DIR to a folder of .docx files to run this")
        }
        let root = URL(fileURLWithPath: NSString(string: dir).expandingTildeInPath)
        let files = ((try? FileManager.default.contentsOfDirectory(at: root, includingPropertiesForKeys: nil)) ?? [])
            .filter { $0.pathExtension == "docx" && !$0.lastPathComponent.hasPrefix("~$") }
            .sorted { $0.lastPathComponent < $1.lastPathComponent }

        var examined = 0
        for f in files {
            guard let doc = try? DocxReader.read(from: f), !nestedTableCellXML(doc).isEmpty else { continue }
            examined += 1

            var snapshots: [[String]] = []
            var cur = f
            for gen in 0..<4 {
                var d = try DocxReader.read(from: cur)
                snapshots.append(nestedTableCellXML(d))
                if gen == 3 { break }
                d.markTypedDirty("word/document.xml")       // the trigger: any typed-dirty op
                let out = URL(fileURLWithPath: NSTemporaryDirectory())
                    .appendingPathComponent("issue155-\(examined)-\(gen).docx")
                try DocxWriter.write(d, to: out)
                cur = out
            }
            for (i, snap) in snapshots.enumerated().dropFirst() {
                XCTAssertEqual(snap, snapshots[0],
                               "\(f.lastPathComponent): nested-table cells changed at generation \(i)")
            }
        }
        XCTAssertGreaterThan(examined, 0,
                             "no document in \(root.path) holds a nested table — this test proved nothing")
        print("issue155: examined \(examined) document(s) holding nested tables")
    }

    /// Emitted XML of every cell that holds a nested table.
    private func nestedTableCellXML(_ doc: WordDocument) -> [String] {
        var pending: [Table] = doc.body.tables
        for child in doc.body.children { if case .table(let t) = child { pending.append(t) } }
        var out: [String] = []
        while let t = pending.popLast() {
            for row in t.rows {
                for cell in row.cells {
                    if !cell.nestedTables.isEmpty { out.append(cell.toXML()) }
                    pending.append(contentsOf: cell.nestedTables)
                }
            }
        }
        return out
    }
}
