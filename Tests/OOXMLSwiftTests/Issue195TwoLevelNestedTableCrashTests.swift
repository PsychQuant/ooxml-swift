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

    // MARK: - R2: cooperative-pool-sized (512KB) thread regression tests
    //
    // R1's fix (a depth guard) did not address the coordinator's actual
    // reproduction: che-word-mcp's real crash comes from `update_cell` (not
    // `insert_paragraph`) on a table whose cell already has source-loaded
    // (position-marked) paragraph content, saved from a thread with a
    // small (~512KB) stack — Swift concurrency's cooperative-pool worker
    // threads, not the 8MB main thread R1's measurements used. At that
    // stack size, just 3 levels of real nesting overflowed the stack
    // BEFORE any of R1's depth guard logic even mattered (528-544KB
    // threshold vs 512KB available). These tests pin the ACTUAL repro
    // shape (`updateCell` on a table read normally, full
    // `DocxWriter.write`, an explicit small thread stack) so a future
    // regression in the `toXMLSortedByPosition` / `TableCell.toXML`
    // frame-splitting fixes below shows up as a real crash, not just a
    // slow test.

    /// Runs `body` on a `Thread` with the given stack size (rounded down
    /// to a page-aligned value `Thread.stackSize` accepts) and asserts it
    /// completes within `timeout` seconds — a crash inside `body` kills
    /// the whole test PROCESS (matching the debug-build `EXC_BAD_ACCESS`
    /// this issue is about), which `swift test` reports as the test
    /// executable exiting with an unexpected signal, not a normal
    /// assertion failure. There is no in-process way to catch that and
    /// convert it into a soft `XCTFail` — the crash IS the evidence.
    private func runOnStackConstrainedThread(
        stackKB: Int, timeout: TimeInterval = 30, _ body: @escaping () throws -> Void
    ) throws {
        let sem = DispatchSemaphore(value: 0)
        var caught: Error?
        let thread = Thread {
            do { try body() } catch { caught = error }
            sem.signal()
        }
        thread.stackSize = stackKB * 1024
        thread.start()
        let waitResult = sem.wait(timeout: .now() + timeout)
        XCTAssertEqual(waitResult, .success, "stack=\(stackKB)KB: timed out (hang)")
        if let caught { throw caught }
    }

    /// The coordinator's confirmed repro, reproduced directly: read the
    /// issue's own 3-level fixture, `update_cell`-equivalent
    /// (`Document.updateCell`) on the OUTER cell only — nothing about the
    /// nested tables is touched — then save on a thread sized like Swift
    /// concurrency's cooperative pool. Before the R2 fix this crashed
    /// (SIGBUS) at the assumed production stack size (512KB); the R1
    /// depth-20 guard alone did nothing for this case since only 3 real
    /// levels of recursion are involved.
    func testUpdateCellOuterCellSmallStackSaveDoesNotCrash() throws {
        let url = try buildFixture(documentXML: Self.documentXMLTemplate)
        defer { try? FileManager.default.removeItem(at: url) }

        var doc = try DocxReader.read(from: url)
        defer { doc.close() }
        try doc.updateCell(tableIndex: 0, row: 0, col: 0, text: "updated-outer-cell")

        guard case .table(let table) = doc.body.children[1] else {
            XCTFail("expected table at body.children[1]")
            return
        }
        XCTAssertTrue(
            table.rows[0].cells[0].paragraphs[0].hasSourcePositionedChildren,
            "sanity check: the updated cell's paragraph must route through " +
            "toXMLSortedByPosition() for this test to exercise the real repro")

        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue195-r2-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }

        try runOnStackConstrainedThread(stackKB: 512) {
            try DocxWriter.write(doc, to: outURL)
        }

        let xml = try extractDocumentXML(from: outURL)
        let text = String(data: xml, encoding: .utf8)!
        XCTAssertTrue(text.contains("updated-outer-cell"))
        XCTAssertTrue(text.contains("Middle1"))
        XCTAssertTrue(text.contains("Innermost"))
    }

    /// Builds a document with `levels` levels of table-in-single-cell
    /// nesting (each cell one row/one col), a real paragraph of text at
    /// each level. Mirrors what `updateCell` needs: normal parsed
    /// paragraphs (not API-built), so every run gets a source `position`
    /// (`DocxReader`'s `childPosition` starts at 1) and every cell
    /// paragraph routes through `Paragraph.toXMLSortedByPosition()`.
    private func makeNLevelNestedTableXML(levels: Int) -> String {
        var open = ""
        var close = ""
        for i in 0..<levels {
            open += "<w:tbl><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr>"
            open += "<w:p><w:r><w:t>L\(i)</w:t></w:r></w:p>"
            close = "<w:p/></w:tc></w:tr></w:tbl>" + close
        }
        return """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
            <w:p><w:r><w:t>Unrelated</w:t></w:r></w:p>\
            \(open)<w:p><w:r><w:t>Leaf</w:t></w:r></w:p>\(close)\
            <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>\
            </w:body></w:document>
            """
    }

    /// Five levels of nesting — `DocxReader.parseTable`'s own
    /// `MAX_TABLE_NEST_DEPTH` ceiling, i.e. the deepest a document can
    /// ever legitimately arrive through the standard reader — `updateCell`
    /// on the outer cell, full `DocxWriter.write` on a 512KB-stack thread.
    /// Measured threshold before this passed: ~536KB (R1 only); ~212KB
    /// after the R2 frame-splitting fixes — comfortable margin under the
    /// assumed 512KB production stack.
    func testFiveLevelNestingUpdateCellSmallStackSaveDoesNotCrash() throws {
        let xml = makeNLevelNestedTableXML(levels: 5)
        let url = try buildFixture(documentXML: xml)
        defer { try? FileManager.default.removeItem(at: url) }

        var doc = try DocxReader.read(from: url)
        defer { doc.close() }
        try doc.updateCell(tableIndex: 0, row: 0, col: 0, text: "updated-outer-cell")

        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue195-r2n5-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }

        try runOnStackConstrainedThread(stackKB: 512) {
            try DocxWriter.write(doc, to: outURL)
        }

        let xmlOut = try extractDocumentXML(from: outURL)
        let text = String(data: xmlOut, encoding: .utf8)!
        XCTAssertTrue(text.contains("updated-outer-cell"))
        XCTAssertTrue(text.contains("Leaf"))
        XCTAssertEqual(text.components(separatedBy: "<w:tbl>").count - 1, 5)
    }

    /// API-constructed variant, bypassing `DocxReader.parseTable`'s
    /// `MAX_TABLE_NEST_DEPTH = 5` entirely (same technique as
    /// `testAPIConstructedExcessiveDepthDoesNotCrashAndTruncates`) to reach
    /// exactly `Table.maxSerializationDepth` (10) — the deepest a
    /// pathological, API-constructed or tree-mutated document can still
    /// reach before the R2-recalibrated guard truncates it. Each level's
    /// cell paragraph gets a `Run` with `position = 1` set explicitly —
    /// this is what forces `Paragraph.toXML()` onto the expensive
    /// `toXMLSortedByPosition()` path (`hasSourcePositionedChildren` checks
    /// `run.position ?? 0 > 0`); an API-built paragraph otherwise defaults
    /// to `position == 0` and takes the cheaper `toXMLLegacy()` path,
    /// understating the worst case. Measured threshold: ~292KB — comfortable
    /// (>40%) margin under the assumed 512KB production stack. See
    /// `Table.maxSerializationDepth`'s doc comment for the full
    /// depth-vs-threshold table.
    func testAPIConstructedGuardBoundaryDepthSmallStackSaveDoesNotCrash() throws {
        let levels = Table.maxSerializationDepth

        func positionedParagraph(_ text: String) -> Paragraph {
            var run = Run(text: text)
            run.position = 1
            return Paragraph(runs: [run])
        }

        var innermost = Table(rows: [TableRow(cells: [
            TableCell(paragraphs: [positionedParagraph("leaf")])
        ])])
        for i in 0..<levels {
            var cell = TableCell(paragraphs: [positionedParagraph("L\(i)")])
            cell.nestedTables = [innermost]
            innermost = Table(rows: [TableRow(cells: [cell])])
        }

        var doc = WordDocument()
        doc.body.children = [.paragraph(positionedParagraph("Unrelated")), .table(innermost)]

        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue195-r2api-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }

        try runOnStackConstrainedThread(stackKB: 512) {
            try DocxWriter.write(doc, to: outURL)
        }

        let xml = try extractDocumentXML(from: outURL)
        let text = String(data: xml, encoding: .utf8)!
        // All `levels + 1` tables (the leaf plus every wrapper) must
        // survive un-truncated: `levels == maxSerializationDepth` sits
        // exactly at the guard boundary, not past it.
        XCTAssertEqual(text.components(separatedBy: "<w:tbl>").count - 1, levels + 1)
        XCTAssertTrue(text.contains("leaf"))
    }
}
