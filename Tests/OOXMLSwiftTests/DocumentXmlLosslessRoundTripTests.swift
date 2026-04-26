import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Round-trip tests for the `che-word-mcp-document-xml-lossless-roundtrip`
/// Spectra change. Covers Phase 1 (namespace preservation) initially; later
/// phases append bookmark, hyperlink, fldSimple, AlternateContent, and raw-carrier
/// scenarios as they land.
///
/// Specs covered:
/// - `WordDocument preserves <w:document> root element attributes byte-equivalent across no-op round-trip`
final class DocumentXmlLosslessRoundTripTests: XCTestCase {

    // MARK: - Phase 1: Document root namespace preservation

    /// Implements spec scenario "34-namespace document round-trips byte-equivalent root".
    /// Builds a fixture `.docx` whose `<w:document>` root declares the 34 `xmlns:*`
    /// prefixes plus `mc:Ignorable` exactly as enumerated in
    /// `openspec/specs/ooxml-roundtrip-fidelity/spec.md`. The bug ([che-word-mcp#56](https://github.com/PsychQuant/che-word-mcp/issues/56))
    /// is that `DocxWriter.writeDocument` hardcodes only `xmlns:w` + `xmlns:r`, so this
    /// test currently fails on:
    /// 1. compile (after task 1.2 lands the field, this clears),
    /// 2. Reader population (after task 1.3 lands the extraction, this clears),
    /// 3. Writer emission (after task 1.4 lands the rebuild, this clears).
    func testRootNamespacesRoundTripPreservesAll34Declarations() throws {
        let sourceURL = try Self.buildThirtyFourNamespaceFixture()
        defer { try? FileManager.default.removeItem(at: sourceURL) }

        var doc = try DocxReader.read(from: sourceURL)
        defer { doc.close() }

        // Assertion (1): Reader populates every source attribute into the new field.
        for (prefix, uri) in Self.expectedThirtyFourNamespaces {
            let key = "xmlns:\(prefix)"
            XCTAssertEqual(
                doc.documentRootAttributes[key], uri,
                "Reader must populate documentRootAttributes[\(key)] from source"
            )
        }
        XCTAssertEqual(
            doc.documentRootAttributes["mc:Ignorable"], Self.expectedIgnorableValue,
            "Reader must populate documentRootAttributes[mc:Ignorable] from source"
        )

        // Round-trip: force `writeDocument` to regenerate so this test exercises
        // the Writer code path that hardcodes namespaces (the v3.12.0 bug). Without
        // marking the part dirty, overlay mode preserves the byte-identical original
        // and the Writer regression is invisible. The user-facing bug from #56
        // ("open → insert_paragraph → save strips 32/34 namespaces") implicitly
        // marks document.xml dirty via the body mutation; markPartDirty makes the
        // same code path explicit without coupling to body mutation semantics.
        doc.markPartDirty("word/document.xml")
        let savedURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("noop-saved-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: savedURL) }
        try DocxWriter.write(doc, to: savedURL)

        // Assertion (2): Writer emits every source attribute into the saved root tag.
        let savedDocumentXML = try Self.readDocumentXMLString(from: savedURL)
        for (prefix, uri) in Self.expectedThirtyFourNamespaces {
            let needle = #"xmlns:\#(prefix)="\#(uri)""#
            XCTAssertTrue(
                savedDocumentXML.contains(needle),
                "Writer must emit \(needle) in saved word/document.xml root tag"
            )
        }
        XCTAssertTrue(
            savedDocumentXML.contains(#"mc:Ignorable="\#(Self.expectedIgnorableValue)""#),
            "Writer must emit mc:Ignorable in saved word/document.xml root tag"
        )

        // Assertion (3): xmllint --noout parses the saved document.xml cleanly.
        // (Skipped silently when xmllint is unavailable on the host.)
        if let xmllintResult = Self.runXmllintNoOut(on: savedURL) {
            XCTAssertEqual(
                xmllintResult.exitCode, 0,
                "xmllint --noout must report no errors on saved document.xml. stderr: \(xmllintResult.stderr)"
            )
        }
    }

    // MARK: - Phase 2: Bookmark Reader parsing

    /// Implements spec requirement "DocxReader parses bookmark range markers as
    /// Paragraph children" + scenario "Reader populates Paragraph.bookmarks for
    /// source bookmark pair". Builds a paragraph with 3 bookmark pairs at varied
    /// positions, asserts Reader populates `Paragraph.bookmarks` (one per pair)
    /// + `Paragraph.bookmarkMarkers` (start + end with positions), and asserts
    /// Writer emits all 3 pairs after a no-op + dirty-marked round-trip.
    ///
    /// This test currently fails on:
    /// 1. compile (after task 2.2 lands BookmarkRangeMarker, this clears),
    /// 2. compile (after task 2.3 lands Paragraph.bookmarkMarkers, this clears),
    /// 3. Reader assertion (after task 2.4 lands the parser, this clears),
    /// 4. Writer assertion (Phase 4 task 4.5 lands sort-by-position emit).
    /// Phase 2 makes assertions 1–3 green; the Writer assertion lands later
    /// and is verified by `testBookmarkPairsRoundTripPreservesAllThreePairs`
    /// once Phase 4 ships.
    func testBookmarkPairRoundTripsAllAttributes() throws {
        let sourceURL = try Self.buildBookmarkFixture()
        defer { try? FileManager.default.removeItem(at: sourceURL) }

        var doc = try DocxReader.read(from: sourceURL)
        defer { doc.close() }

        // Spec scenario uses a single pair at id=42 name="ref-foo" — assert that
        // exact pair plus the 2 additional pairs the fixture seeds (5 + 7 + 12).
        let bookmarks = doc.getParagraphs().flatMap { $0.bookmarks }
        XCTAssertEqual(bookmarks.count, 3, "Reader must populate Paragraph.bookmarks for all 3 source bookmark pairs")
        let bookmarkIds = Set(bookmarks.map { $0.id })
        XCTAssertEqual(bookmarkIds, Set([42, 7, 12]), "Bookmark ids must round-trip from source")
        let bookmarkNames = Set(bookmarks.map { $0.name })
        XCTAssertEqual(bookmarkNames, Set(["ref-foo", "second-anchor", "third-anchor"]),
                       "Bookmark names must round-trip from source")

        // Per spec: each <w:bookmarkStart> + <w:bookmarkEnd> populates a
        // BookmarkRangeMarker entry (start at one position, end at another).
        let allMarkers = doc.getParagraphs().flatMap { $0.bookmarkMarkers }
        XCTAssertEqual(allMarkers.count, 6, "3 bookmark pairs must produce 6 range markers (3 start + 3 end)")
        let starts = allMarkers.filter { $0.kind == .start }
        let ends = allMarkers.filter { $0.kind == .end }
        XCTAssertEqual(starts.count, 3, "3 start markers expected")
        XCTAssertEqual(ends.count, 3, "3 end markers expected")
        XCTAssertEqual(Set(starts.map { $0.id }), Set([42, 7, 12]))
        XCTAssertEqual(Set(ends.map { $0.id }), Set([42, 7, 12]))

        // Verify the spec scenario exactly: the id-42 pair lives in a paragraph
        // where bookmarkStart precedes the run and bookmarkEnd follows it
        // (positions 0, 1, 2 in source order).
        let para42 = try XCTUnwrap(
            doc.getParagraphs().first(where: { $0.bookmarkMarkers.contains(where: { $0.id == 42 }) }),
            "Must find paragraph carrying id=42 bookmark markers"
        )
        let start42 = try XCTUnwrap(para42.bookmarkMarkers.first(where: { $0.id == 42 && $0.kind == .start }))
        let end42 = try XCTUnwrap(para42.bookmarkMarkers.first(where: { $0.id == 42 && $0.kind == .end }))
        XCTAssertLessThan(start42.position, end42.position,
                          "BookmarkStart must have lower position than BookmarkEnd in source order")
    }

    // MARK: - Phase 3: Wrapper hybrid model — Hyperlink

    /// Implements spec requirement "Structural wrapper elements round-trip
    /// lossless across no-op save" + scenario "External URL hyperlink with
    /// multi-run text round-trips with anchor and runs". Builds a paragraph
    /// containing a `<w:hyperlink r:id="rId7" w:tooltip="external">` whose
    /// inner runs are `"click "` (plain) + `"here"` (bold). Asserts:
    /// 1. Reader populates `Paragraph.hyperlinks` with typed `runs`,
    ///    `relationshipId`, `tooltip`.
    /// 2. The computed `text` property returns the joined run text
    ///    (preserves backward compat for 218 MCP tools reading `hyperlink.text`).
    /// 3. Inner Run formatting (bold) survives the parse.
    ///
    /// Currently fails on:
    /// - compile (Hyperlink lacks `runs` field) — clears after task 3.2
    /// - Reader assertion (no <w:hyperlink> parser) — clears after task 3.3
    func testHyperlinkRunsAndRawAttributesRoundTrip() throws {
        let sourceURL = try Self.buildHyperlinkFixture()
        defer { try? FileManager.default.removeItem(at: sourceURL) }

        var doc = try DocxReader.read(from: sourceURL)
        defer { doc.close() }

        let hyperlinks = doc.getParagraphs().flatMap { $0.hyperlinks }
        XCTAssertEqual(hyperlinks.count, 1, "Reader must populate Paragraph.hyperlinks for the source <w:hyperlink>")
        let link = try XCTUnwrap(hyperlinks.first)

        // Spec scenario: relationshipId == "rId7", anchor == nil, tooltip == "external".
        XCTAssertEqual(link.relationshipId, "rId7", "r:id attribute must round-trip")
        XCTAssertNil(link.anchor, "Hyperlink uses r:id (external), not w:anchor")
        XCTAssertEqual(link.tooltip, "external", "w:tooltip attribute must round-trip")

        // Spec scenario: runs.count == 2; runs[0].text == "click " (no bold);
        // runs[1].text == "here" with properties.bold == true.
        XCTAssertEqual(link.runs.count, 2, "Multi-run hyperlink must populate 2 typed Runs")
        XCTAssertEqual(link.runs[0].text, "click ", "Run 0 text")
        XCTAssertEqual(link.runs[0].properties.bold, false, "Run 0 must not be bold")
        XCTAssertEqual(link.runs[1].text, "here", "Run 1 text")
        XCTAssertEqual(link.runs[1].properties.bold, true, "Run 1 inherits source <w:b/>")

        // Computed `text` property contract: returns concatenated run text.
        XCTAssertEqual(link.text, "click here",
                       "Hyperlink.text computed property must return joined runs.text")
    }

    // MARK: - Phase 3: Wrapper hybrid model — FieldSimple

    /// Implements spec requirement "DocxReader parses fldSimple wrapper as
    /// typed FieldSimple model" + scenario "SEQ Table caption fldSimple parses
    /// with instr and result run". Builds a paragraph
    /// `<w:p><w:r>Table </w:r><w:fldSimple w:instr=" SEQ Table \* ARABIC "><w:r>1</w:r></w:fldSimple><w:r>: caption text</w:r></w:p>`
    /// Asserts:
    /// 1. Reader populates `Paragraph.fieldSimples` with `instr` whitespace
    ///    preserved exactly.
    /// 2. Inner Run "1" parses into `fieldSimples[0].runs`.
    /// 3. Surrounding paragraph runs "Table " and ": caption text" are at
    ///    positions 0 and 2 with the FieldSimple at position 1.
    func testFieldSimpleSEQTableCaptionRoundTrips() throws {
        let sourceURL = try Self.buildFieldSimpleFixture()
        defer { try? FileManager.default.removeItem(at: sourceURL) }

        var doc = try DocxReader.read(from: sourceURL)
        defer { doc.close() }

        let allFieldSimples = doc.getParagraphs().flatMap { $0.fieldSimples }
        XCTAssertEqual(allFieldSimples.count, 1, "Reader must populate Paragraph.fieldSimples for the source <w:fldSimple>")
        let field = try XCTUnwrap(allFieldSimples.first)
        XCTAssertEqual(field.instr, " SEQ Table \\* ARABIC ",
                       "instr whitespace must preserve leading + trailing spaces exactly")
        XCTAssertEqual(field.runs.count, 1, "FieldSimple inner Run count")
        XCTAssertEqual(field.runs[0].text, "1", "FieldSimple inner Run text")

        // v0.19.5+ (#56 R5 P0 #2): childPosition starts at 1 (was 0), so
        // surrounding paragraph runs sit at positions 1 and 3, fieldSimple
        // sits between them at position 2.
        let para = try XCTUnwrap(doc.getParagraphs().first(where: { !$0.fieldSimples.isEmpty }))
        XCTAssertEqual(para.runs.count, 2, "2 surrounding paragraph runs (Table , : caption text)")
        XCTAssertEqual(para.runs[0].text, "Table ")
        XCTAssertEqual(para.runs[1].text, ": caption text")
        XCTAssertEqual(field.position, 2,
                       "FieldSimple position must reflect source order between surrounding runs (R5: positions start at 1)")
        XCTAssertEqual(para.runs[0].position, 1, "First run at source position 1 (R5: positions start at 1)")
        XCTAssertEqual(para.runs[1].position, 3, "Second run at source position 3 (after fieldSimple at position 2)")
    }

    // MARK: - Phase 3: Wrapper hybrid model — AlternateContent

    /// Implements spec requirement "DocxReader parses AlternateContent wrapper
    /// preserving raw XML and extracting fallback runs" + scenario "Math
    /// AlternateContent parses with verbatim raw XML and fallback runs".
    /// Builds a paragraph with an `<mc:AlternateContent>` block containing
    /// a `<mc:Choice>` with a small drawing payload and a `<mc:Fallback>`
    /// carrying text "Pearson (Spearman)". Asserts:
    /// 1. Reader populates `Paragraph.alternateContents` with one entry.
    /// 2. The entry's `rawXML` field round-trips byte-equivalent (preserves
    ///    `<mc:Choice>` content for Word reconciliation).
    /// 3. The entry's `fallbackRuns` extracts the text inside `<mc:Fallback>`
    ///    so MCP tools can edit math fallback text.
    func testAlternateContentMathBlockRoundTrips() throws {
        let sourceURL = try Self.buildAlternateContentFixture()
        defer { try? FileManager.default.removeItem(at: sourceURL) }

        var doc = try DocxReader.read(from: sourceURL)
        defer { doc.close() }

        let allACs = doc.getParagraphs().flatMap { $0.alternateContents }
        XCTAssertEqual(allACs.count, 1, "Reader must populate Paragraph.alternateContents")
        let ac = try XCTUnwrap(allACs.first)

        // rawXML must contain both <mc:Choice> and <mc:Fallback> verbatim.
        XCTAssertTrue(ac.rawXML.contains("<mc:Choice"),
                      "rawXML must preserve <mc:Choice> verbatim. Saw: \(ac.rawXML.prefix(200))")
        XCTAssertTrue(ac.rawXML.contains("<mc:Fallback"),
                      "rawXML must preserve <mc:Fallback> verbatim")
        XCTAssertTrue(ac.rawXML.contains("Pearson (Spearman)"),
                      "rawXML must contain the fallback text content")

        // fallbackRuns extracted from <mc:Fallback> for tool-mediated edits.
        XCTAssertEqual(ac.fallbackRuns.count, 1,
                       "fallbackRuns must extract <mc:Fallback>'s <w:r> children")
        XCTAssertEqual(ac.fallbackRuns[0].text, "Pearson (Spearman)",
                       "fallback Run text must round-trip from <mc:Fallback>")
    }

    // MARK: - Phase 4: Raw-carrier markers (commentRange / perm / proofErr / smartTag / customXml / bidi)

    /// Implements spec scenario "Comment range markers preserved across no-op
    /// round-trip" from `ooxml-document-part-mutations`. Builds a paragraph
    /// with `<w:commentRangeStart w:id="3"/>` + run + `<w:commentRangeEnd w:id="3"/>`
    /// and asserts both markers populate `Paragraph.commentRangeMarkers` with
    /// matching ids and source-order positions (start at 0, end at 2).
    ///
    /// Currently fails on:
    /// - compile (no CommentRangeMarker / no Paragraph.commentRangeMarkers) —
    ///   clears after task 4.2.
    /// - Reader assertion (existing parser pushes `commentRangeStart` w:id only
    ///   to `commentIds` and ignores `commentRangeEnd` entirely) — clears after
    ///   task 4.4.
    func testCommentRangeMarkersRoundTrip() throws {
        let sourceURL = try Self.buildCommentRangeFixture()
        defer { try? FileManager.default.removeItem(at: sourceURL) }

        var doc = try DocxReader.read(from: sourceURL)
        defer { doc.close() }

        let allMarkers = doc.getParagraphs().flatMap { $0.commentRangeMarkers }
        XCTAssertEqual(allMarkers.count, 2, "Reader must populate both commentRangeStart and commentRangeEnd")
        let starts = allMarkers.filter { $0.kind == .start }
        let ends = allMarkers.filter { $0.kind == .end }
        XCTAssertEqual(starts.count, 1, "1 start marker expected")
        XCTAssertEqual(ends.count, 1, "1 end marker expected")
        XCTAssertEqual(starts.first?.id, 3, "start id must round-trip")
        XCTAssertEqual(ends.first?.id, 3, "end id must round-trip")
        XCTAssertLessThan(starts.first?.position ?? Int.max, ends.first?.position ?? Int.min,
                          "start position must precede end position in source order")
    }

    // MARK: - Phase 4: Sort-by-position emit (forces writeDocument regeneration)

    /// Implements spec scenario "Reader assigns sequential positions to
    /// interleaved children" + "Writer is order-stable across mutations".
    /// Builds the exact spec-scenario paragraph
    /// `<w:p><w:r>A</w:r><w:bookmarkStart/><w:r>B</w:r><w:hyperlink><w:r>C</w:r></w:hyperlink><w:bookmarkEnd/><w:r>D</w:r></w:p>`,
    /// forces writeDocument to regenerate, asserts the saved XML emits
    /// children in source order (A, bookmarkStart, B, hyperlink-with-C,
    /// bookmarkEnd, D).
    func testInterleavedChildrenRoundTripPreservesSourceOrder() throws {
        let sourceURL = try Self.buildInterleavedFixture()
        defer { try? FileManager.default.removeItem(at: sourceURL) }

        var doc = try DocxReader.read(from: sourceURL)
        defer { doc.close() }

        // Force writeDocument regeneration so we exercise the new sort-emit path
        // (without dirty mark, overlay mode preserves the byte-identical original).
        doc.markPartDirty("word/document.xml")
        let savedURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("interleaved-saved-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: savedURL) }
        try DocxWriter.write(doc, to: savedURL)

        let savedDocumentXML = try Self.readDocumentXMLString(from: savedURL)

        // Find the test paragraph's <w:p>...</w:p> slice (skip <w:sectPr> which
        // is inside <w:body> after the test paragraph).
        guard let pStart = savedDocumentXML.range(of: "<w:p>") else {
            XCTFail("Saved document.xml missing <w:p>: \(savedDocumentXML)")
            return
        }
        guard let pEnd = savedDocumentXML.range(of: "</w:p>", range: pStart.upperBound..<savedDocumentXML.endIndex) else {
            XCTFail("Saved document.xml missing </w:p>")
            return
        }
        let pSlice = String(savedDocumentXML[pStart.lowerBound..<pEnd.upperBound])

        // Locate each expected element by source-order text marker; assert each
        // marker's index in pSlice is monotonically increasing.
        let needles = [
            ">A<",                                    // <w:r>A</w:r>
            "<w:bookmarkStart w:id=\"100\"",          // <w:bookmarkStart .../>
            ">B<",                                    // <w:r>B</w:r>
            "<w:hyperlink",                           // <w:hyperlink ...>
            ">C<",                                    // <w:r>C</w:r> inside hyperlink
            "</w:hyperlink>",
            "<w:bookmarkEnd w:id=\"100\"",            // <w:bookmarkEnd .../>
            ">D<",                                    // <w:r>D</w:r>
        ]
        var lastIndex = pSlice.startIndex
        for (i, needle) in needles.enumerated() {
            guard let range = pSlice.range(of: needle, range: lastIndex..<pSlice.endIndex) else {
                XCTFail("Saved <w:p> missing needle \(i)='\(needle)' after position \(pSlice.distance(from: pSlice.startIndex, to: lastIndex)). Slice: \(pSlice)")
                return
            }
            lastIndex = range.upperBound
        }
    }

    // MARK: - Phase 5: Builder fixture CI regression test

    /// Implements Phase 5 task 5.2 (`testBuilderFixtureRoundTripsLosslessOn34Namespaces`).
    /// Drives `LosslessRoundTripFixtureBuilder.build()` to synthesize a
    /// fixture exercising every code path Phase 1–4 added, then runs
    /// `open → markPartDirty → save → reload` and asserts:
    /// - bookmark count matches source
    /// - hyperlink count matches source
    /// - fldSimple count matches source
    /// - alternateContent count matches source
    /// - xmllint --noout reports no errors on the saved document.xml
    func testBuilderFixtureRoundTripsLosslessOn34Namespaces() throws {
        let sourceURL = try LosslessRoundTripFixtureBuilder.build()
        defer { try? FileManager.default.removeItem(at: sourceURL) }

        // Read source counts via Reader (acts as ground truth for assertions).
        var sourceDoc = try DocxReader.read(from: sourceURL)
        let sourceBookmarkCount = sourceDoc.getParagraphs().reduce(0) { $0 + $1.bookmarks.count }
        let sourceHyperlinkCount = sourceDoc.getParagraphs().reduce(0) { $0 + $1.hyperlinks.count }
        let sourceFieldSimpleCount = sourceDoc.getParagraphs().reduce(0) { $0 + $1.fieldSimples.count }
        let sourceAlternateContentCount = sourceDoc.getParagraphs().reduce(0) { $0 + $1.alternateContents.count }
        sourceDoc.close()

        // Sanity: builder must populate every wrapper category so the test
        // actually exercises Phase 1–4 paths (otherwise it would be a no-op).
        XCTAssertGreaterThanOrEqual(sourceBookmarkCount, 5, "Builder must seed 5+ bookmarks")
        XCTAssertGreaterThanOrEqual(sourceHyperlinkCount, 3, "Builder must seed 3+ hyperlinks")
        XCTAssertGreaterThanOrEqual(sourceFieldSimpleCount, 2, "Builder must seed 2+ fldSimple")
        XCTAssertGreaterThanOrEqual(sourceAlternateContentCount, 1, "Builder must seed 1+ AlternateContent")

        // Round-trip: load, force writeDocument regeneration, save, reload.
        var doc = try DocxReader.read(from: sourceURL)
        doc.markPartDirty("word/document.xml")
        let savedURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("builder-fixture-saved-\(UUID().uuidString).docx")
        defer {
            doc.close()
            try? FileManager.default.removeItem(at: savedURL)
        }
        try DocxWriter.write(doc, to: savedURL)

        var savedDoc = try DocxReader.read(from: savedURL)
        defer { savedDoc.close() }
        let savedBookmarkCount = savedDoc.getParagraphs().reduce(0) { $0 + $1.bookmarks.count }
        let savedHyperlinkCount = savedDoc.getParagraphs().reduce(0) { $0 + $1.hyperlinks.count }
        let savedFieldSimpleCount = savedDoc.getParagraphs().reduce(0) { $0 + $1.fieldSimples.count }
        let savedAlternateContentCount = savedDoc.getParagraphs().reduce(0) { $0 + $1.alternateContents.count }

        XCTAssertEqual(savedBookmarkCount, sourceBookmarkCount,
                       "bookmark count must round-trip lossless")
        XCTAssertEqual(savedHyperlinkCount, sourceHyperlinkCount,
                       "hyperlink count must round-trip lossless")
        XCTAssertEqual(savedFieldSimpleCount, sourceFieldSimpleCount,
                       "fldSimple count must round-trip lossless")
        XCTAssertEqual(savedAlternateContentCount, sourceAlternateContentCount,
                       "AlternateContent count must round-trip lossless")

        // xmllint --noout: no unbound prefix errors on saved document.xml.
        if let result = Self.runXmllintNoOut(on: savedURL) {
            XCTAssertEqual(result.exitCode, 0,
                           "xmllint --noout on builder fixture must report no errors. stderr: \(result.stderr)")
        }
    }

    // MARK: - v0.19.1 follow-up: pPr double-capture regression guard

    /// Regression test for the v0.19.0 → v0.19.1 follow-up. v0.19.0 silently
    /// captured `<w:pPr>` into `Paragraph.unrecognizedChildren` because the
    /// parseParagraph switch did not have an explicit `case "pPr": break`
    /// branch — pPr fell into `default` even though it was already consumed
    /// by the dedicated `parseParagraphProperties` call above the walker.
    /// Result: `<w:pPr>` got written twice on save (once via the legacy pPr
    /// block, once verbatim from `unrecognizedChildren`), and `unrecognized`
    /// counts grew on every round-trip. This test catches the pattern by
    /// asserting parseParagraph never adds pPr to `unrecognizedChildren`.
    func testParseParagraphSkipsPPrInChildWalker() throws {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("ppr-skip-staging-\(UUID().uuidString)")
        defer { try? FileManager.default.removeItem(at: stagingDir) }
        try FileManager.default.createDirectory(at: stagingDir, withIntermediateDirectories: true)

        func write(_ content: String, to relativePath: String) throws {
            let url = stagingDir.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(),
                withIntermediateDirectories: true
            )
            try content.write(to: url, atomically: true, encoding: .utf8)
        }
        try write(contentTypesMinimalXML, to: "[Content_Types].xml")
        try write(packageRelsMinimalXML, to: "_rels/.rels")
        try write(documentRelsMinimalXML, to: "word/_rels/document.xml.rels")
        let pPrDocXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
        <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>centered</w:t></w:r></w:p>
        <w:sectPr></w:sectPr>
        </w:body>
        </w:document>
        """
        try write(pPrDocXML, to: "word/document.xml")

        let outputURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("ppr-skip-fixture-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outputURL) }
        let archive = try Archive(url: outputURL, accessMode: .create)
        let normalizedStaging = stagingDir.resolvingSymlinksInPath().path
        let basePathLen = normalizedStaging.count + 1
        let enumerator = FileManager.default.enumerator(at: stagingDir, includingPropertiesForKeys: [.isDirectoryKey])!
        for case let fileURL as URL in enumerator {
            let isDir = (try? fileURL.resourceValues(forKeys: [.isDirectoryKey]))?.isDirectory ?? false
            if isDir { continue }
            let normalizedFile = fileURL.resolvingSymlinksInPath().path
            let entryName = String(normalizedFile.dropFirst(basePathLen))
            try archive.addEntry(with: entryName, fileURL: fileURL, compressionMethod: .deflate)
        }

        var doc = try DocxReader.read(from: outputURL)
        defer { doc.close() }
        let paras = doc.getParagraphs()
        XCTAssertGreaterThan(paras.count, 0)
        let pPrCapturedCount = paras.reduce(0) { $0 + $1.unrecognizedChildren.filter { $0.name == "pPr" }.count }
        XCTAssertEqual(pPrCapturedCount, 0,
                       "parseParagraph must NOT add <w:pPr> to unrecognizedChildren — it is consumed by parseParagraphProperties before the child walker. v0.19.0 regression: pPr fell into the default branch and got double-emitted on save.")
    }

    // MARK: - Phase 1: Create-from-scratch minimal namespace set

    /// Implements spec scenario "Create-from-scratch document emits minimal namespace set".
    /// Verifies the `documentRootAttributes` empty-dictionary fallback path: an
    /// initializer-built `WordDocument()` has no source archive, so the Writer must
    /// fall back to emitting exactly `xmlns:w` + `xmlns:r`. Lands in task 1.5.
    func testCreateFromScratchEmitsMinimalNamespaceSet() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "Anchor"))
        let outputURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("scratch-minimal-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outputURL) }
        try DocxWriter.write(doc, to: outputURL)

        let savedDocumentXML = try Self.readDocumentXMLString(from: outputURL)
        let openTagRange = try XCTUnwrap(savedDocumentXML.range(of: "<w:document"))
        let openTagEnd = try XCTUnwrap(savedDocumentXML.range(of: ">", range: openTagRange.upperBound..<savedDocumentXML.endIndex))
        let rootOpenTag = String(savedDocumentXML[openTagRange.lowerBound..<openTagEnd.upperBound])

        XCTAssertTrue(
            rootOpenTag.contains(#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#),
            "Scratch-built document root tag must declare xmlns:w. Saw: \(rootOpenTag)"
        )
        XCTAssertTrue(
            rootOpenTag.contains(#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#),
            "Scratch-built document root tag must declare xmlns:r. Saw: \(rootOpenTag)"
        )
        // No additional xmlns declarations.
        let xmlnsCount = rootOpenTag.components(separatedBy: "xmlns:").count - 1
        XCTAssertEqual(
            xmlnsCount, 2,
            "Scratch-built document must declare exactly 2 xmlns prefixes (w + r). Saw \(xmlnsCount). Tag: \(rootOpenTag)"
        )
    }

    // MARK: - Fixture builder

    /// 34-namespace prefix → URI map matching the spec scenario exactly.
    /// Selected to mirror the NTPU master's thesis encountered in [che-word-mcp#56](https://github.com/PsychQuant/che-word-mcp/issues/56)
    /// (35 listed in the spec's narrative text, deduplicated to 34 unique non-`mc:Ignorable`
    /// `xmlns:*` declarations to match the exact count cited in the requirement).
    static let expectedThirtyFourNamespaces: [(String, String)] = [
        ("w",        "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
        ("r",        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
        ("a",        "http://schemas.openxmlformats.org/drawingml/2006/main"),
        ("m",        "http://schemas.openxmlformats.org/officeDocument/2006/math"),
        ("v",        "urn:schemas-microsoft-com:vml"),
        ("o",        "urn:schemas-microsoft-com:office:office"),
        ("mc",       "http://schemas.openxmlformats.org/markup-compatibility/2006"),
        ("wp",       "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
        ("wpg",      "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"),
        ("wps",      "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"),
        ("w10",      "urn:schemas-microsoft-com:office:word"),
        ("w14",      "http://schemas.microsoft.com/office/word/2010/wordml"),
        ("w15",      "http://schemas.microsoft.com/office/word/2012/wordml"),
        ("w16",      "http://schemas.microsoft.com/office/word/2018/wordml"),
        ("w16cex",   "http://schemas.microsoft.com/office/word/2018/wordml/cex"),
        ("w16cid",   "http://schemas.microsoft.com/office/word/2016/wordml/cid"),
        ("w16du",    "http://schemas.microsoft.com/office/word/2023/wordml/word16du"),
        ("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"),
        ("w16sdtfl", "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock"),
        ("w16se",    "http://schemas.microsoft.com/office/word/2015/wordml/symex"),
        ("wne",      "http://schemas.microsoft.com/office/word/2006/wordml"),
        ("wpc",      "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"),
        ("wpi",      "http://schemas.microsoft.com/office/word/2010/wordprocessingInk"),
        ("cx",       "http://schemas.microsoft.com/office/drawing/2014/chartex"),
        ("cx1",      "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"),
        ("cx2",      "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"),
        ("cx3",      "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex"),
        ("cx4",      "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex"),
        ("cx5",      "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex"),
        ("cx6",      "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex"),
        ("cx7",      "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex"),
        ("cx8",      "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex"),
        ("aink",     "http://schemas.microsoft.com/office/drawing/2016/ink"),
        ("am3d",     "http://schemas.microsoft.com/office/drawing/2017/model3d"),
    ]

    static let expectedIgnorableValue = "w14 w15 w16se w16cid"

    /// Build a minimal in-memory `.docx` whose `<w:document>` root declares
    /// every namespace in `expectedThirtyFourNamespaces` plus `mc:Ignorable`.
    /// Body contains a single trivial paragraph so the Reader / Writer paths
    /// have non-empty content to walk.
    static func buildThirtyFourNamespaceFixture() throws -> URL {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("ns34-staging-\(UUID().uuidString)")
        defer { try? FileManager.default.removeItem(at: stagingDir) }

        try FileManager.default.createDirectory(at: stagingDir, withIntermediateDirectories: true)

        func write(_ content: String, to relativePath: String) throws {
            let url = stagingDir.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(),
                withIntermediateDirectories: true
            )
            try content.write(to: url, atomically: true, encoding: .utf8)
        }

        try write(contentTypesMinimalXML, to: "[Content_Types].xml")
        try write(packageRelsMinimalXML, to: "_rels/.rels")
        try write(documentRelsMinimalXML, to: "word/_rels/document.xml.rels")
        try write(buildDocumentXML(), to: "word/document.xml")

        let outputURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("ns34-fixture-\(UUID().uuidString).docx")
        let archive = try Archive(url: outputURL, accessMode: .create)
        let normalizedStaging = stagingDir.resolvingSymlinksInPath().path
        let basePathLen = normalizedStaging.count + 1
        let enumerator = FileManager.default.enumerator(at: stagingDir, includingPropertiesForKeys: [.isDirectoryKey])!
        for case let fileURL as URL in enumerator {
            let isDir = (try? fileURL.resourceValues(forKeys: [.isDirectoryKey]))?.isDirectory ?? false
            if isDir { continue }
            let normalizedFile = fileURL.resolvingSymlinksInPath().path
            let entryName = String(normalizedFile.dropFirst(basePathLen))
            try archive.addEntry(with: entryName, fileURL: fileURL, compressionMethod: .deflate)
        }
        return outputURL
    }

    /// Build a fixture .docx whose body contains 3 bookmark pairs:
    /// - id=42 name="ref-foo" wrapping a paragraph "X" (matches spec scenario exactly)
    /// - id=7  name="second-anchor" wrapping a paragraph "Y"
    /// - id=12 name="third-anchor" wrapping a paragraph "Z"
    /// Used by `testBookmarkPairRoundTripsAllAttributes` (Phase 2).
    static func buildBookmarkFixture() throws -> URL {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("bm-staging-\(UUID().uuidString)")
        defer { try? FileManager.default.removeItem(at: stagingDir) }

        try FileManager.default.createDirectory(at: stagingDir, withIntermediateDirectories: true)

        func write(_ content: String, to relativePath: String) throws {
            let url = stagingDir.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(),
                withIntermediateDirectories: true
            )
            try content.write(to: url, atomically: true, encoding: .utf8)
        }

        try write(contentTypesMinimalXML, to: "[Content_Types].xml")
        try write(packageRelsMinimalXML, to: "_rels/.rels")
        try write(documentRelsMinimalXML, to: "word/_rels/document.xml.rels")
        try write(buildBookmarkDocumentXML(), to: "word/document.xml")

        let outputURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("bm-fixture-\(UUID().uuidString).docx")
        let archive = try Archive(url: outputURL, accessMode: .create)
        let normalizedStaging = stagingDir.resolvingSymlinksInPath().path
        let basePathLen = normalizedStaging.count + 1
        let enumerator = FileManager.default.enumerator(at: stagingDir, includingPropertiesForKeys: [.isDirectoryKey])!
        for case let fileURL as URL in enumerator {
            let isDir = (try? fileURL.resourceValues(forKeys: [.isDirectoryKey]))?.isDirectory ?? false
            if isDir { continue }
            let normalizedFile = fileURL.resolvingSymlinksInPath().path
            let entryName = String(normalizedFile.dropFirst(basePathLen))
            try archive.addEntry(with: entryName, fileURL: fileURL, compressionMethod: .deflate)
        }
        return outputURL
    }

    /// Document body for bookmark fixture: 3 paragraphs each wrapping one
    /// bookmark pair around a single Run. Mirrors the spec scenario layout
    /// (`<w:bookmarkStart><w:r><w:t>X</w:t></w:r><w:bookmarkEnd>`).
    static func buildBookmarkDocumentXML() -> String {
        return """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
        <w:p><w:bookmarkStart w:id="42" w:name="ref-foo"/><w:r><w:t>X</w:t></w:r><w:bookmarkEnd w:id="42"/></w:p>
        <w:p><w:bookmarkStart w:id="7" w:name="second-anchor"/><w:r><w:t>Y</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>
        <w:p><w:bookmarkStart w:id="12" w:name="third-anchor"/><w:r><w:t>Z</w:t></w:r><w:bookmarkEnd w:id="12"/></w:p>
        <w:sectPr></w:sectPr>
        </w:body>
        </w:document>
        """
    }

    /// Build a fixture .docx with a single multi-run hyperlink matching the
    /// spec scenario "External URL hyperlink with multi-run text round-trips
    /// with anchor and runs". The .rels file needs a matching `rId7` entry
    /// pointing to a hyperlink target, otherwise `RelationshipsCollection`
    /// parsing might silently drop the rel.
    static func buildHyperlinkFixture() throws -> URL {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("hl-staging-\(UUID().uuidString)")
        defer { try? FileManager.default.removeItem(at: stagingDir) }

        try FileManager.default.createDirectory(at: stagingDir, withIntermediateDirectories: true)

        func write(_ content: String, to relativePath: String) throws {
            let url = stagingDir.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(),
                withIntermediateDirectories: true
            )
            try content.write(to: url, atomically: true, encoding: .utf8)
        }

        try write(contentTypesMinimalXML, to: "[Content_Types].xml")
        try write(packageRelsMinimalXML, to: "_rels/.rels")
        try write(hyperlinkDocumentRelsXML, to: "word/_rels/document.xml.rels")
        try write(buildHyperlinkDocumentXML(), to: "word/document.xml")

        let outputURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("hl-fixture-\(UUID().uuidString).docx")
        let archive = try Archive(url: outputURL, accessMode: .create)
        let normalizedStaging = stagingDir.resolvingSymlinksInPath().path
        let basePathLen = normalizedStaging.count + 1
        let enumerator = FileManager.default.enumerator(at: stagingDir, includingPropertiesForKeys: [.isDirectoryKey])!
        for case let fileURL as URL in enumerator {
            let isDir = (try? fileURL.resourceValues(forKeys: [.isDirectoryKey]))?.isDirectory ?? false
            if isDir { continue }
            let normalizedFile = fileURL.resolvingSymlinksInPath().path
            let entryName = String(normalizedFile.dropFirst(basePathLen))
            try archive.addEntry(with: entryName, fileURL: fileURL, compressionMethod: .deflate)
        }
        return outputURL
    }

    static func buildHyperlinkDocumentXML() -> String {
        return """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
        <w:p><w:hyperlink r:id="rId7" w:tooltip="external"><w:r><w:t>click </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>here</w:t></w:r></w:hyperlink></w:p>
        <w:sectPr></w:sectPr>
        </w:body>
        </w:document>
        """
    }

    /// Build a fixture .docx whose body contains the spec-scenario fldSimple:
    /// `<w:p><w:r>Table </w:r><w:fldSimple w:instr=" SEQ Table \* ARABIC "><w:r>1</w:r></w:fldSimple><w:r>: caption text</w:r></w:p>`.
    static func buildFieldSimpleFixture() throws -> URL {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("fs-staging-\(UUID().uuidString)")
        defer { try? FileManager.default.removeItem(at: stagingDir) }

        try FileManager.default.createDirectory(at: stagingDir, withIntermediateDirectories: true)

        func write(_ content: String, to relativePath: String) throws {
            let url = stagingDir.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(),
                withIntermediateDirectories: true
            )
            try content.write(to: url, atomically: true, encoding: .utf8)
        }

        try write(contentTypesMinimalXML, to: "[Content_Types].xml")
        try write(packageRelsMinimalXML, to: "_rels/.rels")
        try write(documentRelsMinimalXML, to: "word/_rels/document.xml.rels")
        try write(buildFieldSimpleDocumentXML(), to: "word/document.xml")

        let outputURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("fs-fixture-\(UUID().uuidString).docx")
        let archive = try Archive(url: outputURL, accessMode: .create)
        let normalizedStaging = stagingDir.resolvingSymlinksInPath().path
        let basePathLen = normalizedStaging.count + 1
        let enumerator = FileManager.default.enumerator(at: stagingDir, includingPropertiesForKeys: [.isDirectoryKey])!
        for case let fileURL as URL in enumerator {
            let isDir = (try? fileURL.resourceValues(forKeys: [.isDirectoryKey]))?.isDirectory ?? false
            if isDir { continue }
            let normalizedFile = fileURL.resolvingSymlinksInPath().path
            let entryName = String(normalizedFile.dropFirst(basePathLen))
            try archive.addEntry(with: entryName, fileURL: fileURL, compressionMethod: .deflate)
        }
        return outputURL
    }

    static func buildFieldSimpleDocumentXML() -> String {
        return """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
        <w:p><w:r><w:t xml:space="preserve">Table </w:t></w:r><w:fldSimple w:instr=" SEQ Table \\* ARABIC "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve">: caption text</w:t></w:r></w:p>
        <w:sectPr></w:sectPr>
        </w:body>
        </w:document>
        """
    }

    /// Build a fixture .docx with a math `<mc:AlternateContent>` block
    /// matching the spec scenario. Includes the `mc:` and `wps:` xmlns
    /// declarations on the root so `<mc:Choice Requires="wps14">` parses.
    static func buildAlternateContentFixture() throws -> URL {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("ac-staging-\(UUID().uuidString)")
        defer { try? FileManager.default.removeItem(at: stagingDir) }

        try FileManager.default.createDirectory(at: stagingDir, withIntermediateDirectories: true)

        func write(_ content: String, to relativePath: String) throws {
            let url = stagingDir.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(),
                withIntermediateDirectories: true
            )
            try content.write(to: url, atomically: true, encoding: .utf8)
        }

        try write(contentTypesMinimalXML, to: "[Content_Types].xml")
        try write(packageRelsMinimalXML, to: "_rels/.rels")
        try write(documentRelsMinimalXML, to: "word/_rels/document.xml.rels")
        try write(buildAlternateContentDocumentXML(), to: "word/document.xml")

        let outputURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("ac-fixture-\(UUID().uuidString).docx")
        let archive = try Archive(url: outputURL, accessMode: .create)
        let normalizedStaging = stagingDir.resolvingSymlinksInPath().path
        let basePathLen = normalizedStaging.count + 1
        let enumerator = FileManager.default.enumerator(at: stagingDir, includingPropertiesForKeys: [.isDirectoryKey])!
        for case let fileURL as URL in enumerator {
            let isDir = (try? fileURL.resourceValues(forKeys: [.isDirectoryKey]))?.isDirectory ?? false
            if isDir { continue }
            let normalizedFile = fileURL.resolvingSymlinksInPath().path
            let entryName = String(normalizedFile.dropFirst(basePathLen))
            try archive.addEntry(with: entryName, fileURL: fileURL, compressionMethod: .deflate)
        }
        return outputURL
    }

    static func buildAlternateContentDocumentXML() -> String {
        return """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
        <w:p><mc:AlternateContent><mc:Choice Requires="wps14"><w:r><w:t>choice-placeholder</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t>Pearson (Spearman)</w:t></w:r></mc:Fallback></mc:AlternateContent></w:p>
        <w:sectPr></w:sectPr>
        </w:body>
        </w:document>
        """
    }

    /// Build a fixture .docx wrapping a paragraph in a comment range:
    /// `<w:p><w:commentRangeStart w:id="3"/><w:r><w:t>commented text</w:t></w:r><w:commentRangeEnd w:id="3"/></w:p>`.
    static func buildCommentRangeFixture() throws -> URL {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("cr-staging-\(UUID().uuidString)")
        defer { try? FileManager.default.removeItem(at: stagingDir) }

        try FileManager.default.createDirectory(at: stagingDir, withIntermediateDirectories: true)

        func write(_ content: String, to relativePath: String) throws {
            let url = stagingDir.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(),
                withIntermediateDirectories: true
            )
            try content.write(to: url, atomically: true, encoding: .utf8)
        }

        try write(contentTypesMinimalXML, to: "[Content_Types].xml")
        try write(packageRelsMinimalXML, to: "_rels/.rels")
        try write(documentRelsMinimalXML, to: "word/_rels/document.xml.rels")
        try write(buildCommentRangeDocumentXML(), to: "word/document.xml")

        let outputURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("cr-fixture-\(UUID().uuidString).docx")
        let archive = try Archive(url: outputURL, accessMode: .create)
        let normalizedStaging = stagingDir.resolvingSymlinksInPath().path
        let basePathLen = normalizedStaging.count + 1
        let enumerator = FileManager.default.enumerator(at: stagingDir, includingPropertiesForKeys: [.isDirectoryKey])!
        for case let fileURL as URL in enumerator {
            let isDir = (try? fileURL.resourceValues(forKeys: [.isDirectoryKey]))?.isDirectory ?? false
            if isDir { continue }
            let normalizedFile = fileURL.resolvingSymlinksInPath().path
            let entryName = String(normalizedFile.dropFirst(basePathLen))
            try archive.addEntry(with: entryName, fileURL: fileURL, compressionMethod: .deflate)
        }
        return outputURL
    }

    static func buildCommentRangeDocumentXML() -> String {
        return """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
        <w:p><w:commentRangeStart w:id="3"/><w:r><w:t>commented text</w:t></w:r><w:commentRangeEnd w:id="3"/></w:p>
        <w:sectPr></w:sectPr>
        </w:body>
        </w:document>
        """
    }

    /// Build a fixture .docx whose body contains the interleaved children
    /// from the spec scenario: A, bookmarkStart, B, hyperlink-with-C,
    /// bookmarkEnd, D. Used to verify Phase 4 sort-by-position emit
    /// preserves source order through `writeDocument` regeneration.
    static func buildInterleavedFixture() throws -> URL {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("intl-staging-\(UUID().uuidString)")
        defer { try? FileManager.default.removeItem(at: stagingDir) }

        try FileManager.default.createDirectory(at: stagingDir, withIntermediateDirectories: true)

        func write(_ content: String, to relativePath: String) throws {
            let url = stagingDir.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(),
                withIntermediateDirectories: true
            )
            try content.write(to: url, atomically: true, encoding: .utf8)
        }

        try write(contentTypesMinimalXML, to: "[Content_Types].xml")
        try write(packageRelsMinimalXML, to: "_rels/.rels")
        try write(documentRelsMinimalXML, to: "word/_rels/document.xml.rels")
        try write(buildInterleavedDocumentXML(), to: "word/document.xml")

        let outputURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("intl-fixture-\(UUID().uuidString).docx")
        let archive = try Archive(url: outputURL, accessMode: .create)
        let normalizedStaging = stagingDir.resolvingSymlinksInPath().path
        let basePathLen = normalizedStaging.count + 1
        let enumerator = FileManager.default.enumerator(at: stagingDir, includingPropertiesForKeys: [.isDirectoryKey])!
        for case let fileURL as URL in enumerator {
            let isDir = (try? fileURL.resourceValues(forKeys: [.isDirectoryKey]))?.isDirectory ?? false
            if isDir { continue }
            let normalizedFile = fileURL.resolvingSymlinksInPath().path
            let entryName = String(normalizedFile.dropFirst(basePathLen))
            try archive.addEntry(with: entryName, fileURL: fileURL, compressionMethod: .deflate)
        }
        return outputURL
    }

    static func buildInterleavedDocumentXML() -> String {
        return """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
        <w:p><w:r><w:t>A</w:t></w:r><w:bookmarkStart w:id="100" w:name="anchor"/><w:r><w:t>B</w:t></w:r><w:hyperlink w:anchor="anchor"><w:r><w:t>C</w:t></w:r></w:hyperlink><w:bookmarkEnd w:id="100"/><w:r><w:t>D</w:t></w:r></w:p>
        <w:sectPr></w:sectPr>
        </w:body>
        </w:document>
        """
    }

    /// Build the document.xml with the 34 xmlns + mc:Ignorable + minimal body.
    static func buildDocumentXML() -> String {
        var lines: [String] = []
        lines.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")
        let nsAttrs = expectedThirtyFourNamespaces
            .map { "xmlns:\($0.0)=\"\($0.1)\"" }
            .joined(separator: " ")
        lines.append("<w:document \(nsAttrs) mc:Ignorable=\"\(expectedIgnorableValue)\">")
        lines.append("<w:body>")
        lines.append("<w:p><w:r><w:t>Anchor</w:t></w:r></w:p>")
        lines.append("<w:sectPr></w:sectPr>")
        lines.append("</w:body>")
        lines.append("</w:document>")
        return lines.joined(separator: "\n")
    }

    // MARK: - Saved document.xml inspection

    /// Read `word/document.xml` from a saved `.docx` (unzip via ZipHelper, return UTF-8 string).
    static func readDocumentXMLString(from docxURL: URL) throws -> String {
        let unzipped = try ZipHelper.unzip(docxURL)
        defer { ZipHelper.cleanup(unzipped) }
        let documentURL = unzipped.appendingPathComponent("word/document.xml")
        return try String(contentsOf: documentURL, encoding: .utf8)
    }

    /// Run `xmllint --noout` on the saved document.xml. Returns nil if xmllint is
    /// unavailable on the host (CI without libxml2 CLI) so the test does not
    /// false-fail. Otherwise returns exit code + stderr.
    static func runXmllintNoOut(on docxURL: URL) -> (exitCode: Int32, stderr: String)? {
        let unzipped: URL
        do {
            unzipped = try ZipHelper.unzip(docxURL)
        } catch {
            return nil
        }
        defer { ZipHelper.cleanup(unzipped) }
        let documentPath = unzipped.appendingPathComponent("word/document.xml").path
        let xmllintPath = "/usr/bin/xmllint"
        guard FileManager.default.isExecutableFile(atPath: xmllintPath) else { return nil }

        let process = Process()
        process.executableURL = URL(fileURLWithPath: xmllintPath)
        process.arguments = ["--noout", documentPath]
        let stderrPipe = Pipe()
        process.standardError = stderrPipe
        process.standardOutput = Pipe()
        do {
            try process.run()
            process.waitUntilExit()
        } catch {
            return nil
        }
        let stderr = String(data: stderrPipe.fileHandleForReading.readDataToEndOfFile(), encoding: .utf8) ?? ""
        return (process.terminationStatus, stderr)
    }
}

// MARK: - Minimal XML boilerplate (no headers/footers/styles needed for Phase 1)

private let contentTypesMinimalXML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""

private let packageRelsMinimalXML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

private let documentRelsMinimalXML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>
"""

/// Hyperlink fixture rels: includes rId7 hyperlink target so the parsed
/// `relationshipId == "rId7"` round-trips meaningfully.
private let hyperlinkDocumentRelsXML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>
</Relationships>
"""
