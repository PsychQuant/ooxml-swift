import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Regression tests for [PsychQuant/ooxml-swift#195](https://github.com/PsychQuant/ooxml-swift/issues/195).
///
/// The issue reports that a document with two levels of nested tables
/// crashes (debug: `EXC_BAD_ACCESS`, backtrace `Table.toXML()` →
/// `TableRow.toXML()` → `TableCell.toXML()`) or hangs (release) as soon as
/// an unrelated typed edit forces `word/document.xml` through the legacy
/// typed-writer re-serialization path (`DocxWriter.writeDocument` →
/// `Body.children` → `Table.toXML()`).
///
/// Investigation (see the accompanying fix report) could not reproduce an
/// actual crash for exactly two levels of nesting through any of
/// `DocxReader.read`'s entry points on this branch — both the detached
/// (default) and `wireTreeBackedViews: true` reads, single and repeated
/// saves, all serialize correctly, and `DocxReader.parseTable`'s
/// `MAX_TABLE_NEST_DEPTH = 5` guard runs unconditionally before any
/// tree-backed wiring, so a document can never be *read* with pathological
/// table-nesting depth through that entry point either way.
///
/// The underlying mechanism the issue describes IS real and WAS reproduced
/// here, just via a different entry point: `Table.toXML()` →
/// `TableRow.toXML()` → `TableCell.toXML()` had NO recursion-depth guard of
/// their own — unlike `DocxReader.parseTable` (`MAX_TABLE_NEST_DEPTH`) and
/// `XmlTreeReader`'s `maxElementDepth` (added specifically to fix an
/// earlier "unbounded recursion was an uncatchable stack overflow" finding
/// in this exact codebase). A `Table` built purely through the public
/// struct API — `TableCell.nestedTables` is a plain array setter with no
/// depth validation — or reached via `resyncBodyFromDocumentTree()` (which
/// re-types `Table(xmlNode:)` values straight from a live tree, bypassing
/// `parseTable` entirely) carries arbitrary nesting depth into `toXML()`,
/// which recurses exactly as deep as the data goes. Before the fix,
/// `testAPIConstructedExcessiveDepthDoesNotCrashAndTruncates`'s 300-level
/// table reliably reproduced `EXC_BAD_ACCESS` at roughly 62-64 levels of
/// real recursion on a full 8MB debug-build thread stack — each level
/// through this specific call chain costs far more stack than a typical
/// Swift function (`TableProperties.toXML`'s slot-computation machinery
/// alone runs once per level). This test suite closes that gap with a
/// matching depth guard (`Table.maxSerializationDepth`), and locks down
/// round-trip correctness for legitimate 2- and 3-level nested tables so
/// the fix cannot regress into the very crash it closes.
final class Issue195TwoLevelNestedTableCrashTests: XCTestCase {

    // MARK: - Fixture: outer table → nested 1×2 table → doubly-nested 1×1 table

    /// `<w:tbl>` (outer, 1×1) containing in its single cell:
    ///   "Outer" paragraph, then `<w:tbl>` (middle, 1×2) whose first cell has
    ///   "Middle1" + `<w:tbl>` (innermost, 1×1) with "Innermost", and whose
    ///   second cell has "Middle2".
    private static let documentXMLTemplate = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
        <w:p><w:r><w:t>Unrelated</w:t></w:r></w:p>\
        <w:tbl><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>\
        <w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>\
        <w:p><w:r><w:t>Outer</w:t></w:r></w:p>\
        <w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>\
        <w:tr>\
        <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>\
        <w:p><w:r><w:t>Middle1</w:t></w:r></w:p>\
        <w:tbl><w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>\
        <w:tr><w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr>\
        <w:p><w:r><w:t>Innermost</w:t></w:r></w:p>\
        </w:tc></w:tr></w:tbl>\
        <w:p/>\
        </w:tc>\
        <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>\
        <w:p><w:r><w:t>Middle2</w:t></w:r></w:p>\
        </w:tc>\
        </w:tr></w:tbl>\
        <w:p/>\
        </w:tc></w:tr></w:tbl>\
        <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>\
        </w:body></w:document>
        """

    private func buildFixture(documentXML: String) throws -> URL {
        let staging = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue195-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: staging, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: staging) }

        func write(_ content: String, to relativePath: String) throws {
            let url = staging.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(), withIntermediateDirectories: true)
            try content.write(to: url, atomically: true, encoding: .utf8)
        }

        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="xml" ContentType="application/xml"/>
                <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>
            """, to: "[Content_Types].xml")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """, to: "_rels/.rels")
        try write(documentXML, to: "word/document.xml")

        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue195-\(UUID().uuidString).docx")
        do {
            let archive = try Archive(url: docxURL, accessMode: .create)
            let base = staging.resolvingSymlinksInPath().path
            let enumerator = FileManager.default.enumerator(
                at: staging, includingPropertiesForKeys: [.isDirectoryKey])!
            for case let fileURL as URL in enumerator {
                let isDir = (try? fileURL.resourceValues(forKeys: [.isDirectoryKey]))?.isDirectory ?? false
                if isDir { continue }
                let entry = String(fileURL.resolvingSymlinksInPath().path.dropFirst(base.count + 1))
                try archive.addEntry(with: entry, fileURL: fileURL, compressionMethod: .deflate)
            }
        } catch {
            try? FileManager.default.removeItem(at: docxURL)
            throw error
        }
        return docxURL
    }

    private func extractDocumentXML(from docxURL: URL) throws -> Data {
        let archive = try Archive(url: docxURL, accessMode: .read)
        let entry = try XCTUnwrap(archive["word/document.xml"])
        var data = Data()
        _ = try archive.extract(entry) { data.append($0) }
        return data
    }

    /// Reproduces the issue's own minimal scenario: read → unrelated typed
    /// edit (`insertParagraph`, which forces `word/document.xml` off the
    /// tree-fresh path per `Document.insertParagraph` → `markTypedDirty`) →
    /// save. Detached (default) read.
    func testReadUnrelatedEditSaveDoesNotCrashOrHang() throws {
        let url = try buildFixture(documentXML: Self.documentXMLTemplate)
        defer { try? FileManager.default.removeItem(at: url) }

        var doc = try DocxReader.read(from: url)
        defer { doc.close() }

        doc.insertParagraph(Paragraph(text: "inserted"), at: 0)

        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue195-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }

        try DocxWriter.write(doc, to: outURL)

        let xml = try extractDocumentXML(from: outURL)
        let text = String(data: xml, encoding: .utf8)!
        XCTAssertTrue(text.contains("inserted"))
        XCTAssertTrue(text.contains("Outer"))
        XCTAssertTrue(text.contains("Middle1"))
        XCTAssertTrue(text.contains("Middle2"))
        XCTAssertTrue(text.contains("Innermost"))
        // Exactly three <w:tbl> elements: outer, middle, innermost.
        XCTAssertEqual(text.components(separatedBy: "<w:tbl>").count - 1, 3)
    }

    /// Same fixture and edit as above, but with `wireTreeBackedViews: true`
    /// (che-word-mcp's `open_document` op-log path) — the body-level Table
    /// becomes tree-backed (`xmlNode != nil`), so `rows`/`cells`/`nestedTables`
    /// walk `xmlNode.children` at `toXML()` time instead of reading detached
    /// stored arrays.
    func testReadTreeBackedUnrelatedEditSaveDoesNotCrashOrHang() throws {
        let url = try buildFixture(documentXML: Self.documentXMLTemplate)
        defer { try? FileManager.default.removeItem(at: url) }

        var doc = try DocxReader.read(from: url, wireTreeBackedViews: true)
        defer { doc.close() }

        doc.insertParagraph(Paragraph(text: "inserted"), at: 0)

        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue195-treebacked-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }

        try DocxWriter.write(doc, to: outURL)

        let xml = try extractDocumentXML(from: outURL)
        let text = String(data: xml, encoding: .utf8)!
        XCTAssertTrue(text.contains("inserted"))
        XCTAssertTrue(text.contains("Outer"))
        XCTAssertTrue(text.contains("Middle1"))
        XCTAssertTrue(text.contains("Middle2"))
        XCTAssertTrue(text.contains("Innermost"))
        XCTAssertEqual(text.components(separatedBy: "<w:tbl>").count - 1, 3)
    }

    /// Repeated read → unrelated edit → save generations (mirrors #155's own
    /// regression style): table count must stay flat and the paragraph
    /// count must only grow by the one paragraph each generation actually
    /// inserts — no runaway growth, no crash, across five generations.
    func testRepeatedSaveTwoLevelNestingDoesNotGrowOrCrash() throws {
        var url = try buildFixture(documentXML: Self.documentXMLTemplate)
        defer { try? FileManager.default.removeItem(at: url) }

        var tblCounts: [Int] = []
        var pCounts: [Int] = []
        for generation in 0..<5 {
            var doc = try DocxReader.read(from: url, wireTreeBackedViews: true)
            doc.insertParagraph(Paragraph(text: "gen\(generation)"), at: 0)

            let outURL = FileManager.default.temporaryDirectory
                .appendingPathComponent("issue195-rep-gen\(generation)-\(UUID().uuidString).docx")
            try DocxWriter.write(doc, to: outURL)
            doc.close()
            try? FileManager.default.removeItem(at: url)
            url = outURL

            let xml = try extractDocumentXML(from: url)
            let text = String(data: xml, encoding: .utf8)!
            tblCounts.append(text.components(separatedBy: "<w:tbl>").count - 1)
            pCounts.append(text.components(separatedBy: "<w:p").count - 1)
        }
        XCTAssertEqual(Set(tblCounts), [3], "table count must stay flat across generations — got \(tblCounts)")
        for i in 1..<pCounts.count {
            XCTAssertEqual(pCounts[i], pCounts[i - 1] + 1,
                           "paragraph count must grow by exactly 1 per generation (the inserted paragraph) — got \(pCounts)")
        }
    }

    /// Three-level nesting (one level deeper than the issue's minimal
    /// repro): outer → middle → inner → innermost. Guards against a fix
    /// that only special-cases exactly two levels. Content is verified
    /// directly against the saved XML bytes (not `getAllParagraphs()`,
    /// which has its own, unrelated gap: `collectAllParagraphs`'s `.table`
    /// case only walks a table's own cells and does not recurse into
    /// `cell.nestedTables` — a pre-existing enumeration limitation in the
    /// same family as che-word-mcp#188, out of scope for this crash fix).
    func testThreeLevelNestingReadEditSaveReadBackPreservesContent() throws {
        let threeLevelXML = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
            <w:p><w:r><w:t>Unrelated</w:t></w:r></w:p>\
            <w:tbl><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>\
            <w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>\
            <w:p><w:r><w:t>L1</w:t></w:r></w:p>\
            <w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>\
            <w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>\
            <w:p><w:r><w:t>L2</w:t></w:r></w:p>\
            <w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>\
            <w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>\
            <w:p><w:r><w:t>L3</w:t></w:r></w:p>\
            <w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>\
            <w:tr><w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr>\
            <w:p><w:r><w:t>L4</w:t></w:r></w:p>\
            </w:tc></w:tr></w:tbl>\
            <w:p/>\
            </w:tc></w:tr></w:tbl>\
            <w:p/>\
            </w:tc></w:tr></w:tbl>\
            <w:p/>\
            </w:tc></w:tr></w:tbl>\
            <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>\
            </w:body></w:document>
            """
        for wireTreeBackedViews in [false, true] {
            let url = try buildFixture(documentXML: threeLevelXML)
            defer { try? FileManager.default.removeItem(at: url) }

            var doc = try DocxReader.read(from: url, wireTreeBackedViews: wireTreeBackedViews)
            defer { doc.close() }

            doc.insertParagraph(Paragraph(text: "inserted"), at: 0)

            let outURL = FileManager.default.temporaryDirectory
                .appendingPathComponent("issue195-3level-out-\(UUID().uuidString).docx")
            defer { try? FileManager.default.removeItem(at: outURL) }
            try DocxWriter.write(doc, to: outURL)

            let xml = try extractDocumentXML(from: outURL)
            let text = String(data: xml, encoding: .utf8)!
            XCTAssertTrue(text.contains("inserted"), "wireTreeBackedViews=\(wireTreeBackedViews) — got: \(text)")
            XCTAssertTrue(text.contains("L1"), "wireTreeBackedViews=\(wireTreeBackedViews) — got: \(text)")
            XCTAssertTrue(text.contains("L2"), "wireTreeBackedViews=\(wireTreeBackedViews) — got: \(text)")
            XCTAssertTrue(text.contains("L3"), "wireTreeBackedViews=\(wireTreeBackedViews) — got: \(text)")
            XCTAssertTrue(text.contains("L4"), "wireTreeBackedViews=\(wireTreeBackedViews) — got: \(text)")
            XCTAssertEqual(text.components(separatedBy: "<w:tbl>").count - 1, 4,
                           "all four nesting levels must round-trip (wireTreeBackedViews=\(wireTreeBackedViews)) — got: \(text)")

            // Read the saved file back through the reader too, confirming
            // it is well-formed and re-parses without error (a genuinely
            // corrupted / runaway-recursion save would not).
            var readBack = try DocxReader.read(from: outURL)
            readBack.close()
        }
    }

    /// White-box: a table built purely through the public struct API (no
    /// XML / `DocxReader` involved at all) has NO depth validation on the
    /// way in — `TableCell.nestedTables` is a plain array setter. This is
    /// the path `DocxReader.parseTable`'s `MAX_TABLE_NEST_DEPTH` guard
    /// cannot protect (it only runs during XML parsing), and the one
    /// `resyncBodyFromDocumentTree()` also bypasses (it constructs
    /// `Table(xmlNode:)` values straight from a live tree with no depth
    /// check either). Before the fix, `table.toXML()` here recursed exactly
    /// as deep as the constructed structure — 5000 levels reliably
    /// overflows the stack. After the fix it stops well short of that and
    /// returns a well-formed (truncated) result instead of crashing.
    func testAPIConstructedExcessiveDepthDoesNotCrashAndTruncates() throws {
        // Constructing and dropping a deeply-nested VALUE-TYPE structure is
        // its own (separate, Swift-runtime-level) concern independent of
        // this fix — array/struct deinit for a 5000-level chain is also
        // recursive. 300 levels is comfortably past
        // `Table.maxSerializationDepth` (proving the guard actually
        // truncates) while keeping construction/deinit cheap and unrelated
        // to what this test is verifying.
        let depth = 300
        var innermost = Table(rows: [TableRow(cells: [TableCell(text: "leaf")])])
        for i in 0..<depth {
            var cell = TableCell(paragraphs: [Paragraph(text: "L\(i)")])
            cell.nestedTables = [innermost]
            innermost = Table(rows: [TableRow(cells: [cell])])
        }

        let xml = innermost.toXML()
        XCTAssertTrue(xml.hasPrefix("<w:tbl>"))
        XCTAssertTrue(xml.hasSuffix("</w:tbl>"))

        // Recursion must stop near Table.maxSerializationDepth, not carry
        // on for all 300 levels.
        let tblOpens = xml.components(separatedBy: "<w:tbl>").count - 1
        XCTAssertGreaterThan(tblOpens, 1, "sanity floor: some legitimate nesting should still be emitted")
        XCTAssertLessThanOrEqual(
            tblOpens, Table.maxSerializationDepth + 1,
            "expected recursion to stop at maxSerializationDepth, got \(tblOpens) <w:tbl> opens")
    }

    /// Tables within the depth-guard bound must be completely unaffected —
    /// `Table.maxSerializationDepth` sits far above anything this suite's
    /// 2–4 level fixtures reach, so this is redundant with the tests above
    /// in content, but pins the boundary behavior explicitly: exactly N
    /// levels in, exactly N `<w:tbl>` elements out.
    func testAPIConstructedNestingWithinGuardRoundTripsExactly() throws {
        let depth = 10
        var innermost = Table(rows: [TableRow(cells: [TableCell(text: "leaf")])])
        for i in 0..<depth {
            var cell = TableCell(paragraphs: [Paragraph(text: "L\(i)")])
            cell.nestedTables = [innermost]
            innermost = Table(rows: [TableRow(cells: [cell])])
        }

        let xml = innermost.toXML()
        let tblOpens = xml.components(separatedBy: "<w:tbl>").count - 1
        XCTAssertEqual(tblOpens, depth + 1, "expected all \(depth + 1) tables (10 wrappers + 1 leaf) to round-trip — got \(tblOpens)")
        XCTAssertTrue(xml.contains("leaf"))
        for i in 0..<depth {
            XCTAssertTrue(xml.contains("L\(i)"), "missing L\(i) — got: \(xml)")
        }
    }
}
