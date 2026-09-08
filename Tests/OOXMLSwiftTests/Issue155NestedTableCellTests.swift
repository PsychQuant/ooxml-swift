import XCTest
import ZIPFoundation
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
    ///
    /// The required trailing paragraph IS added here — the source cell does not
    /// end with one, and `<w:tc>` must. What must not happen is `before` moving
    /// across the table, which is what the first position asserts. 3.7.0
    /// appended one here too; the difference is that it also moved `before`.
    func testACellEndingWithATableKeepsItsParagraphFirst() throws {
        let xml = try cell(holding: para("before") + innerTable).toXML()
        XCTAssertEqual(blockOrder(xml), ["p:before", "tbl", "p:"],
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

    /// A cell whose ONLY child is a nested table still ends with a paragraph.
    ///
    /// Counting emitted BLOCKS rather than paragraphs for that guarantee made
    /// this cell come out as `<w:tc><w:tbl/></w:tc>` — no paragraph at all,
    /// where 3.7.0 emitted two. Word treats such a cell as damaged.
    func testACellHoldingOnlyANestedTableStillEndsWithAParagraph() throws {
        XCTAssertEqual(blockOrder(try cell(holding: innerTable).toXML()), ["tbl", "p:"])
    }

    /// …and adding it does not restart the growth: the added paragraph is
    /// recorded on the next read, so the following save emits it rather than
    /// appending another.
    func testTheAddedTrailingParagraphIsAFixedPoint() throws {
        var body = innerTable
        var seen: [[String]] = []
        for _ in 0..<3 {
            let xml = try cell(holding: body).toXML()
            seen.append(blockOrder(xml))
            body = String(xml.dropFirst("<w:tc>".count).dropLast("</w:tc>".count))
        }
        XCTAssertEqual(Set(seen.map { $0.joined(separator: ",") }).count, 1,
                       "the cell must stop changing after the required paragraph is added: \(seen)")
        XCTAssertEqual(seen[0], ["tbl", "p:"])
    }

    /// The reader's two views of the same cell must agree. Guarding the
    /// "a cell holds at least one paragraph" append with `blocks.isEmpty` left
    /// `paragraphs` holding a paragraph that `_blocks` did not know about.
    func testTheReadersParagraphListAndBlockListAgree() throws {
        let c = try cell(holding: innerTable)
        // tree-backed: both views derive from the node, so they cannot disagree
        XCTAssertEqual(c.paragraphs.count, 0)
        XCTAssertEqual(c.nestedTables.count, 1)

        // a caller-built cell records no order at all
        var detached = TableCell()
        detached.paragraphs = []
        detached.nestedTables = []
        XCTAssertNil(detached._blockOrder, "a caller-built cell records no order")
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

    // MARK: - Through the reader (not gated)

    /// Every test above builds a tree-backed cell with `TableCell(xmlNode:)`.
    /// `DocxReader.read` produces DETACHED cells — the other branch entirely —
    /// so none of them covered the path real documents take. Making that branch
    /// dead code left the whole 1567-test suite green.
    private func readCell(bodyXML: String, inHeader: Bool = false) throws -> (WordDocument, TableCell) {
        let body = inHeader ? "<w:p/>" : bodyXML
        let doc = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            + "<w:body>\(body)<w:sectPr>\(inHeader ? "<w:headerReference xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId9\" w:type=\"default\"/>" : "")</w:sectPr></w:body></w:document>"
        let url = URL(fileURLWithPath: NSTemporaryDirectory()).appendingPathComponent("i155-\(UUID().uuidString).docx")
        try writeDocx(document: doc, header: inHeader ? headerPart(bodyXML) : nil, to: url)
        defer { try? FileManager.default.removeItem(at: url) }
        let read = try DocxReader.read(from: url)
        let cell = try XCTUnwrap(firstNestedTableCell(read, inHeader: inHeader), "no nested-table cell found")
        return (read, cell)
    }

    private func headerPart(_ inner: String) -> String {
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        + "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\(inner)</w:hdr>"
    }

    private func firstNestedTableCell(_ doc: WordDocument, inHeader: Bool) -> TableCell? {
        var tables: [Table] = []
        let children: [BodyChild] = inHeader ? (doc.headers.first?.bodyChildren ?? []) : doc.body.children
        if !inHeader { tables = doc.body.tables }
        for ch in children { if case .table(let t) = ch { tables.append(t) } }
        while let t = tables.popLast() {
            for row in t.rows {
                for cell in row.cells {
                    if !cell.nestedTables.isEmpty { return cell }
                    tables.append(contentsOf: cell.nestedTables)
                }
            }
        }
        return nil
    }

    private func writeDocx(document: String, header: String?, to url: URL) throws {
        let arch = try Archive(accessMode: .create)
        func add(_ path: String, _ text: String) throws {
            let d = Data(text.utf8)
            try arch.addEntry(with: path, type: .file, uncompressedSize: Int64(d.count),
                              provider: { pos, size in d.subdata(in: Int(pos)..<Int(pos) + size) })
        }
        var overrides = "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
        var rels = "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
        if header != nil {
            overrides += "<Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>"
        }
        try add("[Content_Types].xml", "<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>\(overrides)</Types>")
        try add("_rels/.rels", "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\(rels)</Relationships>")
        try add("word/document.xml", document)
        if let header {
            try add("word/header1.xml", header)
            try add("word/_rels/document.xml.rels", "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId9\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/></Relationships>")
        }
        try (arch.data ?? Data()).write(to: url)
    }

    private func outerTable(holding inner: String) -> String {
        "<w:tbl><w:tr><w:tc>\(inner)</w:tc></w:tr></w:tbl>"
    }

    /// The reader's own path: order survives a real read.
    func testAReadDocumentKeepsItsCellOrder() throws {
        let (_, cell) = try readCell(bodyXML: outerTable(holding: para("A") + innerTable + para("B") + para("C")))
        XCTAssertEqual(blockOrder(cell.toXML()), ["p:A", "tbl", "p:B", "p:C"])
    }

    /// The same cell inside a HEADER. `DocxReader` runs a hyperlink-id rewrite
    /// over every header, footer, footnote and endnote, and that walk writes
    /// each cell's paragraphs back through the setter. A record that the setter
    /// destroyed therefore never reached any of them — the fix applied to the
    /// body only, and no test could see it.
    func testAHeaderCellKeepsItsOrderToo() throws {
        let (_, cell) = try readCell(bodyXML: outerTable(holding: para("A") + innerTable + para("B")), inHeader: true)
        XCTAssertEqual(blockOrder(cell.toXML()), ["p:A", "tbl", "p:B"])
    }

    /// Editing a paragraph IN PLACE must not lose the order. Swift routes
    /// `transform(&cell.paragraphs[i])` through the setter, so a record the
    /// setter discarded was lost to a no-op — and to `Document`'s accept/reject
    /// revision walk, which writes every cell of a table back.
    func testEditingAParagraphInPlaceKeepsTheOrder() throws {
        let (_, original) = try readCell(bodyXML: outerTable(holding: para("A") + innerTable + para("B")))
        var cell = original
        XCTAssertEqual(blockOrder(cell.toXML()), ["p:A", "tbl", "p:B"], "precondition")

        func noop(_ p: inout Paragraph) { _ = p }
        noop(&cell.paragraphs[0])
        XCTAssertEqual(blockOrder(cell.toXML()), ["p:A", "tbl", "p:B"], "a no-op in-place edit must change nothing")

        cell.paragraphs[1] = Paragraph(text: "B edited")
        XCTAssertEqual(blockOrder(cell.toXML()), ["p:A", "tbl", "p:B edited"],
                       "replacing one paragraph keeps the interleaving")
    }

    /// A write that changes HOW MANY blocks the cell holds invalidates the
    /// record, and the cell falls back to the historical shape rather than
    /// emitting against a stale map.
    func testAddingAParagraphFallsBackRatherThanMisplacing() throws {
        let (_, original) = try readCell(bodyXML: outerTable(holding: para("A") + innerTable + para("B")))
        var cell = original
        cell.paragraphs.append(Paragraph(text: "C"))
        let order = blockOrder(cell.toXML())
        XCTAssertEqual(order.filter { $0.hasPrefix("p:") }.count, 4, "three paragraphs plus the required trailing one")
        XCTAssertTrue(order.contains("tbl"))
    }

    // MARK: - Real documents (gated)

    /// The generation loop over real documents, comparing each generation
    /// against **the order in the source file's own `word/document.xml`**.
    ///
    /// An earlier version compared generation N against generation 0 of its own
    /// output. That detects INSTABILITY, not infidelity: an emit that moves
    /// content once and then stays put is stable, so it passed. A mutation
    /// recording the shape as "every table, then every paragraph" survived it.
    /// This is the same failure the paragraph COUNT had, one level up, in the
    /// test written to replace the count.
    ///
    /// Point `OOXML_CORPUS_DIR` at a folder of .docx files to run it.
    func testRealDocumentsReEmitIdentically() throws {
        guard let dir = ProcessInfo.processInfo.environment["OOXML_CORPUS_DIR"] else {
            throw XCTSkip("set OOXML_CORPUS_DIR to a folder of .docx files to run this")
        }
        let root = URL(fileURLWithPath: NSString(string: dir).expandingTildeInPath)
        let files = ((try? FileManager.default.contentsOfDirectory(at: root, includingPropertiesForKeys: nil)) ?? [])
            .filter { $0.pathExtension == "docx" && !$0.lastPathComponent.hasPrefix("~$") }
            .sorted { $0.lastPathComponent < $1.lastPathComponent }

        var examined = 0, cellsChecked = 0
        for f in files {
            guard let sourceOrders = try? nestedTableCellOrdersInSource(of: f), !sourceOrders.isEmpty else { continue }
            examined += 1

            var cur = f
            for gen in 0..<4 {
                var d = try DocxReader.read(from: cur)
                let emitted = nestedTableCellXML(d).map { blockOrder($0) }
                XCTAssertEqual(emitted, sourceOrders,
                               "\(f.lastPathComponent) gen \(gen): emitted cell order differs from the source file")
                cellsChecked += emitted.count
                if gen == 3 { break }
                d.markTypedDirty("word/document.xml")       // the trigger: any typed-dirty op
                let out = URL(fileURLWithPath: NSTemporaryDirectory())
                    .appendingPathComponent("issue155-\(examined)-\(gen).docx")
                try DocxWriter.write(d, to: out)
                cur = out
            }
        }
        XCTAssertGreaterThan(examined, 0,
                             "no document in \(root.path) holds a nested table — this test proved nothing")
        XCTAssertGreaterThan(cellsChecked, 0, "no cell was compared")
        print("issue155: \(examined) document(s), \(cellsChecked) cell comparisons against the source")
    }

    /// The p/tbl order of every nested-table cell, read from the package's own
    /// `word/document.xml` bytes — the ground truth this library is supposed to
    /// preserve, independent of anything the typed model does.
    private func nestedTableCellOrdersInSource(of url: URL) throws -> [[String]] {
        let archive = try Archive(url: url, accessMode: .read)
        guard let entry = archive["word/document.xml"] else { return [] }
        var bytes = Data()
        _ = try archive.extract(entry, skipCRC32: true) { bytes.append($0) }
        let xml = try XMLDocument(data: bytes, options: [])
        var out: [[String]] = []
        func walk(_ el: XMLElement) {
            if el.name == "w:tc" {
                let kids = (el.children ?? []).compactMap { $0 as? XMLElement }
                if kids.contains(where: { $0.name == "w:tbl" }) {
                    out.append(kids.compactMap { k in
                        switch k.name {
                        case "w:p":   return "p:" + textOf(k)
                        case "w:tbl": return "tbl"
                        default:      return nil
                        }
                    })
                }
            }
            for child in (el.children ?? []).compactMap({ $0 as? XMLElement }) { walk(child) }
        }
        if let rootEl = xml.rootElement() { walk(rootEl) }
        return out
    }

    /// Concatenated `<w:t>` text of a paragraph, matching what `blockOrder` reads.
    private func textOf(_ paragraph: XMLElement) -> String {
        var text = ""
        func walk(_ el: XMLElement) {
            if el.name == "w:t" { text += el.stringValue ?? "" }
            for child in (el.children ?? []).compactMap({ $0 as? XMLElement }) { walk(child) }
        }
        walk(paragraph)
        return text
    }

    /// Emitted XML of every cell that holds a nested table, in DOCUMENT ORDER.
    ///
    /// The walk order matters: this list is compared element-wise against the
    /// same cells read from the source XML, so a stack-order walk (`popLast()`)
    /// silently compared cell 1 against cell 4.
    private func nestedTableCellXML(_ doc: WordDocument) -> [String] {
        var out: [String] = []
        func visit(_ t: Table) {
            for row in t.rows {
                for cell in row.cells {
                    if !cell.nestedTables.isEmpty { out.append(cell.toXML()) }
                    for nested in cell.nestedTables { visit(nested) }
                }
            }
        }
        for child in doc.body.children { if case .table(let t) = child { visit(t) } }
        for t in doc.body.tables where !doc.body.children.contains(where: {
            if case .table(let c) = $0 { return c == t } else { return false }
        }) { visit(t) }
        return out
    }
}
