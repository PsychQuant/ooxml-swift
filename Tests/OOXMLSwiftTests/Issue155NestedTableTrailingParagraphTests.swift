import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Regression tests for [PsychQuant/ooxml-swift#155](https://github.com/PsychQuant/ooxml-swift/issues/155).
///
/// `TableCell.toXML()` unconditionally appended a fresh trailing `<w:p/>`
/// whenever the cell held a nested table — the comment said "if missing"
/// but the code never checked. Because `DocxReader` merges every direct
/// `<w:p>` child of a cell into the SAME `paragraphs` array regardless of
/// whether it appeared before or after the nested `<w:tbl>`, the
/// OOXML-required trailing paragraph a source document already carries gets
/// re-classified as an ordinary LEADING paragraph on parse, and the writer
/// appends ANOTHER trailing paragraph on top — +1 paragraph per
/// nested-table cell, per save, unbounded.
///
/// After this fix: if the cell's last paragraph is run-less, `toXML()`
/// reuses it as the trailing paragraph (moves it after the nested tables)
/// instead of leaving it in place AND synthesizing a new one — repeated
/// saves are idempotent.
final class Issue155NestedTableTrailingParagraphTests: XCTestCase {

    // MARK: - Section A: unit-level, no docx I/O

    /// Mirrors exactly what `DocxReader` produces for a cell whose source
    /// XML is `<w:p>Outer</w:p><w:tbl>…</w:tbl><w:p/>` — one real paragraph,
    /// one nested table, and the already-present OOXML-required trailing
    /// paragraph (empty, no runs) merged into the same `paragraphs` array.
    private func makeCellWithNestedTableAndTrailingPlaceholder() -> TableCell {
        var cell = TableCell(paragraphs: [Paragraph(text: "Outer"), Paragraph()])
        cell.nestedTables = [Table(rows: [
            TableRow(cells: [TableCell(text: "Inner")])
        ])]
        return cell
    }

    func testToXMLDoesNotDuplicateTrailingParagraph() {
        let cell = makeCellWithNestedTableAndTrailingPlaceholder()
        let xml = cell.toXML()
        let pCount = xml.components(separatedBy: "<w:p").count - 1
        // Exactly 2 <w:p...> in the OUTER cell: "Outer" + the one trailing
        // placeholder — NOT 3 (which is what the pre-fix code produced by
        // leaving the placeholder in place AND appending a fresh one).
        // (The nested table's own "Inner" cell paragraph is a separate <w:p>
        // inside <w:tbl>, so this counts only elements textually preceding
        // the nested table's opening tag to isolate the outer cell's own count.)
        let beforeNestedTable = xml.components(separatedBy: "<w:tbl>").first ?? xml
        let afterNestedTable = xml.components(separatedBy: "</w:tbl>").last ?? ""
        let outerPCount = (beforeNestedTable.components(separatedBy: "<w:p").count - 1)
            + (afterNestedTable.components(separatedBy: "<w:p").count - 1)
        XCTAssertEqual(outerPCount, 2, "got: \(xml)")
        _ = pCount
    }

    func testToXMLEndsWithATrailingParagraphAfterTheNestedTable() {
        let cell = makeCellWithNestedTableAndTrailingPlaceholder()
        let xml = cell.toXML()
        XCTAssertTrue(xml.hasSuffix("<w:p/></w:tc>") || xml.hasSuffix("<w:p></w:p></w:tc>"),
                      "cell must still end with a paragraph after the nested table — got: \(xml)")
        XCTAssertTrue(xml.contains(">Outer<"), "the real paragraph must survive — got: \(xml)")
    }

    /// A cell with a nested table but NO pre-existing trailing placeholder
    /// (e.g. built directly via the API) still gets exactly one synthesized
    /// trailing paragraph — no regression on the case the original code did
    /// handle correctly.
    func testToXMLSynthesizesTrailingParagraphWhenNoneExists() {
        var cell = TableCell(paragraphs: [Paragraph(text: "Outer")])
        cell.nestedTables = [Table(rows: [TableRow(cells: [TableCell(text: "Inner")])])]
        let xml = cell.toXML()
        let beforeNestedTable = xml.components(separatedBy: "<w:tbl>").first ?? xml
        let afterNestedTable = xml.components(separatedBy: "</w:tbl>").last ?? ""
        let outerPCount = (beforeNestedTable.components(separatedBy: "<w:p").count - 1)
            + (afterNestedTable.components(separatedBy: "<w:p").count - 1)
        XCTAssertEqual(outerPCount, 2, "1 real + 1 synthesized trailing — got: \(xml)")
    }

    // MARK: - Section A2: the reused placeholder's OWN content/format
    // (independent review finding F5)
    //
    // The tests above only assert PARAGRAPH COUNT stays bounded. Count
    // staying flat is also what a mutation that always synthesizes a fresh
    // `Paragraph()` (never reusing the real one) would produce — one
    // paragraph is dropped from `bodyParagraphs`, one fresh one is
    // appended, so the count is unchanged either way. These tests instead
    // assert the SPECIFIC paragraph object survives: its own `pPr`
    // (alignment) and, in the docx-round-trip case below, its `w14:paraId`.

    /// Trailing placeholder with non-default `properties` (alignment) — the
    /// shape che-word-mcp#224(a) already fixed elsewhere for the "empty but
    /// formatted" paragraph case.
    private func makeCellWithFormattedTrailingPlaceholder() -> TableCell {
        var placeholder = Paragraph()
        placeholder.properties.alignment = .center
        var cell = TableCell(paragraphs: [Paragraph(text: "Outer"), placeholder])
        cell.nestedTables = [Table(rows: [
            TableRow(cells: [TableCell(text: "Inner")])
        ])]
        return cell
    }

    func testToXMLPreservesTrailingPlaceholderFormatting() throws {
        let cell = makeCellWithFormattedTrailingPlaceholder()
        let xml = cell.toXML()
        // The trailing paragraph (after the nested table) must carry the
        // alignment — not a bare `<w:p/>` that silently dropped it.
        let afterNestedTable = xml.components(separatedBy: "</w:tbl>").last ?? ""
        XCTAssertTrue(afterNestedTable.contains("w:jc") && afterNestedTable.contains("center"),
                     "the reused trailing paragraph must keep its own pPr — got: \(xml)")
        XCTAssertFalse(afterNestedTable.hasPrefix("<w:p/>"),
                      "a bare self-closing <w:p/> means the real placeholder was discarded and a fresh default one synthesized instead — got: \(xml)")
    }

    // MARK: - Section B: full-docx regression — the issue's own guard test

    private static let documentXMLTemplate = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
        <w:p><w:r><w:t>Unrelated</w:t></w:r></w:p>\
        <w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>\
        <w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>\
        <w:p><w:r><w:t>Outer</w:t></w:r></w:p>\
        <w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>\
        <w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>\
        <w:p><w:r><w:t>Inner</w:t></w:r></w:p>\
        </w:tc></w:tr></w:tbl>\
        <w:p/>\
        </w:tc></w:tr></w:tbl>\
        <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>\
        </w:body></w:document>
        """

    private func buildFixture(documentXML: String) throws -> URL {
        let staging = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue155-staging-\(UUID().uuidString)")
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
            .appendingPathComponent("issue155-\(UUID().uuidString).docx")
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

    /// Counts `<w:p` element occurrences (`<w:p>`, `<w:p/>`, or
    /// `<w:p w14:paraId="...">`) in `s` — precise counting: a naive
    /// `components(separatedBy: "<w:p")` also matches `<w:pPr>`, `<w:pStyle>`,
    /// etc. (any element whose name starts with `p`), over-counting the
    /// moment a counted region contains a paragraph WITH properties.
    private func countParagraphElements(in s: String) -> Int {
        var count = 0
        var searchRange = s.startIndex..<s.endIndex
        while let match = s.range(of: "<w:p", range: searchRange) {
            let next = s.index(match.upperBound, offsetBy: 0)
            if next < s.endIndex, [">", "/", " "].contains(String(s[next])) {
                count += 1
            }
            searchRange = match.upperBound..<s.endIndex
        }
        return count
    }

    /// Counts the OUTER cell's own direct `<w:p` occurrences, excluding
    /// whatever is inside the nested `<w:tbl>…</w:tbl>`. There are exactly
    /// two `<w:tbl>` elements in this fixture (outer, inner): the outer
    /// cell's LEADING paragraphs sit between the first `<w:tbl>` (outer's
    /// own open tag) and the second `<w:tbl>` (inner's open tag); its
    /// TRAILING paragraph(s) sit between the first `</w:tbl>` (inner's close)
    /// and the second `</w:tbl>` (outer's close). Text before/after those two
    /// tables (the unrelated top-level paragraph, `<w:sectPr>`, …) is
    /// deliberately excluded.
    private func outerCellParagraphCount(in documentXML: Data) throws -> Int {
        let xml = String(data: documentXML, encoding: .utf8)!
        let openParts = xml.components(separatedBy: "<w:tbl>")
        let closeParts = xml.components(separatedBy: "</w:tbl>")
        guard openParts.count >= 3, closeParts.count >= 3 else {
            throw NSError(domain: "test", code: 0, userInfo: [
                NSLocalizedDescriptionKey: "expected exactly 2 <w:tbl> elements; got \(openParts.count - 1)"])
        }
        let leading = openParts[1]   // between outer <w:tbl> and inner <w:tbl>
        let trailing = closeParts[1] // between inner </w:tbl> and outer </w:tbl>
        return countParagraphElements(in: leading) + countParagraphElements(in: trailing)
    }

    /// The issue's own suggested guard: save the same document three times
    /// (each through a genuine typed-dirty edit elsewhere, so the cell's
    /// `toXML()` path actually runs) and assert the paragraph count does
    /// not move.
    func testRepeatedSaveDoesNotGrowTrailingParagraphCount() throws {
        var url = try buildFixture(documentXML: Self.documentXMLTemplate)
        defer { try? FileManager.default.removeItem(at: url) }

        var counts: [Int] = []
        for generation in 0..<3 {
            var doc = try DocxReader.read(from: url)
            // A typed-dirty edit unrelated to the nested-table cell — this is
            // what forces `word/document.xml` through the lossy typed
            // re-serialization path (#155's mechanism requires a real
            // typed-dirty save, not a byte-identical no-op save). Updates the
            // SAME top-level paragraph each generation (rather than
            // inserting a new one) so the body's own paragraph count stays
            // constant and only the nested-table cell's count is under test.
            try doc.updateParagraph(at: 0, text: "unrelated edit \(generation)")

            let outURL = FileManager.default.temporaryDirectory
                .appendingPathComponent("issue155-gen\(generation)-\(UUID().uuidString).docx")
            try DocxWriter.write(doc, to: outURL)
            doc.close()
            try? FileManager.default.removeItem(at: url)
            url = outURL

            let xml = try extractDocumentXML(from: url)
            counts.append(try outerCellParagraphCount(in: xml))
        }

        // The issue's own suggested guard: the count must not move across
        // repeated saves. Pre-fix this measured [3, 4, 5] (+1 per
        // generation, unbounded).
        XCTAssertEqual(counts.count, 3)
        XCTAssertTrue(counts.allSatisfy { $0 >= 2 },
                     "sanity floor: outer cell must have at least its real paragraph + 1 trailing — got \(counts)")
        XCTAssertEqual(Set(counts).count, 1,
                       "outer cell paragraph count must stay bounded across repeated saves — got \(counts)")
    }

    /// Same fixture shape as `documentXMLTemplate`, but the trailing
    /// placeholder carries a `pPr` (centered alignment) instead of being
    /// bare `<w:p/>` — the "empty but formatted" shape #182's
    /// che-word-mcp#224(a) note already treats as a real case, not a
    /// hypothetical one.
    private static let documentXMLWithFormattedPlaceholder = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" \
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>\
        <w:p><w:r><w:t>Unrelated</w:t></w:r></w:p>\
        <w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>\
        <w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>\
        <w:p><w:r><w:t>Outer</w:t></w:r></w:p>\
        <w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>\
        <w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>\
        <w:p><w:r><w:t>Inner</w:t></w:r></w:p>\
        </w:tc></w:tr></w:tbl>\
        <w:p w14:paraId="7F00BA55"><w:pPr><w:jc w:val="center"/></w:pPr></w:p>\
        </w:tc></w:tr></w:tbl>\
        <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>\
        </w:body></w:document>
        """

    /// Review finding F5: the count staying flat is not proof the SAME
    /// placeholder survived — a mutation that always drops the real one and
    /// synthesizes `Paragraph()` fresh would also keep the count flat (one
    /// removed from `bodyParagraphs`, one added back). This asserts the
    /// reused paragraph's own `pPr` (and `w14:paraId`, which only a REUSED
    /// object — not a freshly synthesized `Paragraph()` — can carry) survive
    /// an unrelated typed-dirty save.
    func testRepeatedSavePreservesTrailingPlaceholderOwnFormatting() throws {
        let url = try buildFixture(documentXML: Self.documentXMLWithFormattedPlaceholder)
        defer { try? FileManager.default.removeItem(at: url) }
        var doc = try DocxReader.read(from: url)
        defer { doc.close() }

        try doc.updateParagraph(at: 0, text: "unrelated edit")

        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue155-formatted-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(doc, to: outURL)

        let xml = try extractDocumentXML(from: outURL)
        let outerCount = try outerCellParagraphCount(in: xml)
        XCTAssertEqual(outerCount, 2, "count must still be bounded — got \(outerCount)")

        let closeParts = (String(data: xml, encoding: .utf8)!).components(separatedBy: "</w:tbl>")
        let trailing = closeParts.count >= 3 ? closeParts[1] : "" // between inner </w:tbl> and outer </w:tbl>
        XCTAssertTrue(trailing.contains("w:jc") && trailing.contains("center"),
                     "the trailing paragraph's own pPr (centered alignment) must survive — got: \(trailing)")
        XCTAssertTrue(trailing.contains("7F00BA55"),
                     "the trailing paragraph must be the SAME object (its w14:paraId survives) — " +
                     "a fresh Paragraph() has no paraId, so this fails under the F5 mutation " +
                     "(reviewer's report) that always synthesizes a new one instead of reusing " +
                     "the real placeholder — got: \(trailing)")
    }
}
