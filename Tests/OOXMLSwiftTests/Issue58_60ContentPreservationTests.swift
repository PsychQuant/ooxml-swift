import XCTest
@testable import OOXMLSwift

/// Tests for the document-content-preservation Spectra change covering
/// PsychQuant/che-word-mcp#58 (body-level bookmark markers), #59 (whitespace
/// `<w:t>` runs), and #60 (RunProperties field-loss audit).
///
/// Each sub-stack adds tests + a matrix-pin assertion class to the
/// `testDocumentContentEqualityInvariant` test that lands here.
final class Issue58_60ContentPreservationTests: XCTestCase {

    // MARK: - Sub-stack A: #58 BodyChild block-level marker preservation

    /// §1.1 — Body-level `<w:bookmarkStart>` and `<w:bookmarkEnd>` SHALL survive
    /// open → modify → save round-trip. Pre-fix, `parseBodyChildren` switch's
    /// `default: continue` silently drops anything that isn't `<w:p>` / `<w:tbl>`
    /// / `<w:sdt>`, so body-level bookmarks vanish.
    ///
    /// Covers spec requirement: `BodyChild enum SHALL cover EG_BlockLevelElts
    /// members beyond paragraph and table` — Scenarios "Body-level bookmarkStart
    /// preserved through round-trip" and "Body-level bookmarkEnd preserved
    /// through round-trip" (specs/ooxml-paragraph-child-schema-coverage/spec.md).
    func testBodyLevelBookmarkRoundTripPreserved() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p><w:r><w:t>before</w:t></w:r></w:p>
        <w:bookmarkStart w:id="0" w:name="_TocTest"/>
        <w:p><w:r><w:t>middle</w:t></w:r></w:p>
        <w:bookmarkEnd w:id="0"/>
        <w:p><w:r><w:t>after</w:t></w:r></w:p>
        </w:body>
        </w:document>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-§1.1-in-\(UUID().uuidString).docx")
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-§1.1-out-\(UUID().uuidString).docx")
        defer {
            try? FileManager.default.removeItem(at: inURL)
            try? FileManager.default.removeItem(at: outURL)
        }
        try buildMinimalDocx(documentXML: documentXML, to: inURL)

        var doc = try DocxReader.read(from: inURL)
        // Force document.xml re-serialization (mirrors MCP body-mutating save).
        doc.modifiedParts.insert("word/document.xml")
        try DocxWriter.write(doc, to: outURL)

        let outXML = try Self.readDocumentXMLString(from: outURL)
        XCTAssertTrue(
            outXML.contains("<w:bookmarkStart") && outXML.contains("w:name=\"_TocTest\""),
            "body-level <w:bookmarkStart w:name=\"_TocTest\"/> SHALL survive round-trip; output:\n\(outXML)"
        )
        XCTAssertTrue(
            outXML.contains("<w:bookmarkEnd") && outXML.contains("w:id=\"0\""),
            "body-level <w:bookmarkEnd w:id=\"0\"/> SHALL survive round-trip; output:\n\(outXML)"
        )
    }

    /// §1.2 — Unknown body-level elements (e.g., `<w:moveFromRangeStart>`) SHALL
    /// be preserved as raw-XML carriers rather than silently dropped, so future
    /// EG_BlockLevelElts / vendor extensions byte-roundtrip even without typed
    /// parser branches.
    ///
    /// Covers spec requirement: `BodyChild enum SHALL cover EG_BlockLevelElts
    /// members beyond paragraph and table` — Scenario "Unknown body-level element
    /// preserved as raw element" (specs/ooxml-paragraph-child-schema-coverage/spec.md).
    func testBodyLevelUnknownElementPreservedAsRaw() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p><w:r><w:t>x</w:t></w:r></w:p>
        <w:moveFromRangeStart w:id="1" w:name="testMove"/>
        <w:p><w:r><w:t>y</w:t></w:r></w:p>
        </w:body>
        </w:document>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-§1.2-in-\(UUID().uuidString).docx")
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-§1.2-out-\(UUID().uuidString).docx")
        defer {
            try? FileManager.default.removeItem(at: inURL)
            try? FileManager.default.removeItem(at: outURL)
        }
        try buildMinimalDocx(documentXML: documentXML, to: inURL)

        var doc = try DocxReader.read(from: inURL)
        doc.modifiedParts.insert("word/document.xml")
        try DocxWriter.write(doc, to: outURL)

        let outXML = try Self.readDocumentXMLString(from: outURL)
        XCTAssertTrue(
            outXML.contains("moveFromRangeStart"),
            "unknown body-level <w:moveFromRangeStart> SHALL be preserved as raw element; output:\n\(outXML)"
        )
    }

    /// §1.6 — `nextBookmarkId` calibration walker SHALL include body-level
    /// `BookmarkRangeMarker` entries (in addition to paragraph-level
    /// `paragraph.bookmarkMarkers`). Otherwise a future API-built bookmark
    /// could collide with an existing body-level id.
    ///
    /// Covers spec requirement: `nextBookmarkId calibration SHALL include
    /// body-level bookmark markers` — Scenario "nextBookmarkId reflects
    /// body-level bookmarks after read" (specs/ooxml-paragraph-child-schema-coverage/spec.md).
    func testNextBookmarkIdReflectsBodyLevelBookmarksAfterRead() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p><w:bookmarkStart w:id="3" w:name="paraLevel"/><w:bookmarkEnd w:id="3"/><w:r><w:t>x</w:t></w:r></w:p>
        <w:bookmarkStart w:id="7" w:name="bodyLevel"/>
        <w:p><w:r><w:t>y</w:t></w:r></w:p>
        <w:bookmarkEnd w:id="7"/>
        </w:body>
        </w:document>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-§1.6-in-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try buildMinimalDocx(documentXML: documentXML, to: inURL)

        let doc = try DocxReader.read(from: inURL)
        XCTAssertEqual(
            doc.nextBookmarkId, 8,
            "nextBookmarkId SHALL be one greater than the global max bookmark id (7 from body-level + 3 from paragraph-level → 8)"
        )
    }

    // MARK: - Sub-stack A-CONT: container parser entry point + getBookmarks walker

    /// §1.14 — Body-level `<w:bookmarkStart>` / `<w:bookmarkEnd>` inside a
    /// header SHALL survive body-mutating save. Pre-A-CONT-fix
    /// `parseContainerChildBodyChildren` (DocxReader.swift:1291-1322) had only
    /// `case "p"` / `case "tbl"` / `default: continue` — body-level markers in
    /// headers were silently dropped. The dead-code calibration walker added in
    /// sub-stack A (`collectBodyLevelBookmarkIds(header.bodyChildren)`) is the
    /// smoking gun.
    ///
    /// Same fix shape as `parseBodyChildren` (mirror branches into the second
    /// parser entry point).
    func testHeaderBodyLevelBookmarkRoundTripPreserved() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p><w:r><w:t>body</w:t></w:r></w:p>
        </w:body>
        </w:document>
        """
        let headerXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:bookmarkStart w:id="9" w:name="hdrAnchor"/>
        <w:p><w:r><w:t>header text</w:t></w:r></w:p>
        <w:bookmarkEnd w:id="9"/>
        </w:hdr>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-acont-hdr-in-\(UUID().uuidString).docx")
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-acont-hdr-out-\(UUID().uuidString).docx")
        defer {
            try? FileManager.default.removeItem(at: inURL)
            try? FileManager.default.removeItem(at: outURL)
        }
        try buildMinimalDocxWithHeader(documentXML: documentXML, headerXML: headerXML, headerRId: "rId10", to: inURL)

        var doc = try DocxReader.read(from: inURL)
        // Force header re-serialization to exercise the writer path.
        doc.modifiedParts.insert("word/header1.xml")
        try DocxWriter.write(doc, to: outURL)

        let outHeaderXML = try Self.readPartXMLString(from: outURL, partPath: "word/header1.xml")
        XCTAssertTrue(
            outHeaderXML.contains("<w:bookmarkStart") && outHeaderXML.contains("w:name=\"hdrAnchor\""),
            "body-level <w:bookmarkStart w:name=\"hdrAnchor\"/> SHALL survive in header round-trip; output:\n\(outHeaderXML)"
        )
        XCTAssertTrue(
            outHeaderXML.contains("<w:bookmarkEnd") && outHeaderXML.contains("w:id=\"9\""),
            "body-level <w:bookmarkEnd w:id=\"9\"/> SHALL survive in header round-trip; output:\n\(outHeaderXML)"
        )
    }

    /// §1.16 — `Document.getBookmarks()` SHALL surface body-level `.bookmarkMarker`
    /// entries. Pre-A-CONT-fix `getBookmarks()` (Document.swift:2122-2136)
    /// iterated only `case .paragraph` reading `para.bookmarks` — never
    /// `case .bookmarkMarker`. Body-level bookmarks preserved on disk by sub-stack A
    /// were invisible to MCP `list_bookmarks`.
    func testGetBookmarksSurfacesBodyLevelMarkers() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p><w:bookmarkStart w:id="1" w:name="paraLevel"/><w:bookmarkEnd w:id="1"/><w:r><w:t>x</w:t></w:r></w:p>
        <w:bookmarkStart w:id="2" w:name="bodyLevel"/>
        <w:p><w:r><w:t>y</w:t></w:r></w:p>
        <w:bookmarkEnd w:id="2"/>
        </w:body>
        </w:document>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-acont-getbm-in-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try buildMinimalDocx(documentXML: documentXML, to: inURL)

        let doc = try DocxReader.read(from: inURL)
        let bookmarks = doc.getBookmarks()
        let names = bookmarks.map { $0.name }
        XCTAssertTrue(names.contains("paraLevel"),
            "paragraph-level bookmark SHALL be returned; got \(names)")
        XCTAssertTrue(names.contains("bodyLevel"),
            "body-level bookmark SHALL ALSO be returned post-A-CONT; got \(names)")
    }

    // MARK: - Sub-stack A-CONT-2: API-layer container coverage + SDT recursion + matrix-pin fixture

    /// §1.26 — `Document.getBookmarks()` SHALL surface body-level `.bookmarkMarker`
    /// entries from container parts (headers / footers / footnotes / endnotes),
    /// not just `body.children`. A-CONT closed the WRITE-side parser asymmetry but
    /// `getBookmarks()` only iterated `body.children`, so header body-level
    /// bookmarks preserved on disk were invisible to MCP `list_bookmarks`.
    func testGetBookmarksSurfacesContainerBodyLevelMarkers() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p><w:r><w:t>body</w:t></w:r></w:p>
        </w:body>
        </w:document>
        """
        let headerXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:bookmarkStart w:id="100" w:name="hdrBookmark"/>
        <w:p><w:r><w:t>header text</w:t></w:r></w:p>
        <w:bookmarkEnd w:id="100"/>
        </w:hdr>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-acont2-getbm-hdr-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try buildMinimalDocxWithHeader(documentXML: documentXML, headerXML: headerXML, headerRId: "rId10", to: inURL)

        let doc = try DocxReader.read(from: inURL)
        let bookmarks = doc.getBookmarks()
        let names = bookmarks.map { $0.name }
        XCTAssertTrue(names.contains("hdrBookmark"),
            "header body-level bookmark SHALL be surfaced by getBookmarks() (A-CONT-2 P0 #1); got \(names)")
    }

    /// §1.27 — `parseContainerChildBodyChildren` SHALL recursively parse `<w:sdt>`
    /// (block-level Structured Document Tag) into `.contentControl(_, children)`,
    /// not capture as `.rawBlockElement`. A-CONT mirrored 5 of 6 cases from
    /// `parseBodyChildren`; the missing `case "sdt"` meant container-side block-
    /// level SDTs were captured as raw XML — round-trip-preserved but invisible
    /// to typed model walkers (incl. `nextBookmarkId` calibration on nested
    /// bookmark ids).
    func testParseContainerSDTRecursionPreservesNestedBookmark() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p><w:r><w:t>body</w:t></w:r></w:p>
        </w:body>
        </w:document>
        """
        let headerXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:sdt>
        <w:sdtPr><w:id w:val="42"/><w:tag w:val="hdrSDT"/></w:sdtPr>
        <w:sdtContent>
        <w:bookmarkStart w:id="500" w:name="sdtBookmark"/>
        <w:p><w:r><w:t>inside SDT</w:t></w:r></w:p>
        <w:bookmarkEnd w:id="500"/>
        </w:sdtContent>
        </w:sdt>
        </w:hdr>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-acont2-sdt-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try buildMinimalDocxWithHeader(documentXML: documentXML, headerXML: headerXML, headerRId: "rId10", to: inURL)

        let doc = try DocxReader.read(from: inURL)
        guard let header = doc.headers.first else {
            XCTFail("expected at least one header parsed")
            return
        }

        // Assert the SDT was parsed as typed `.contentControl`, NOT `.rawBlockElement`.
        let firstChild = header.bodyChildren.first
        guard let firstChild = firstChild else {
            XCTFail("expected at least one bodyChild in header; got empty")
            return
        }
        if case .rawBlockElement = firstChild {
            XCTFail("header SDT was captured as .rawBlockElement (A-CONT-2 P0 #2 — parseContainerChildBodyChildren missing `case \"sdt\"` recursion)")
            return
        }
        guard case .contentControl(_, let inner) = firstChild else {
            XCTFail("expected `.contentControl` for header SDT; got \(firstChild)")
            return
        }

        // Assert the nested .bookmarkMarker is reachable.
        let hasNestedBookmark = inner.contains { child in
            if case .bookmarkMarker(let marker) = child, marker.id == 500, marker.name == "sdtBookmark" {
                return true
            }
            return false
        }
        XCTAssertTrue(hasNestedBookmark,
            "nested bookmark inside header SDT SHALL be reachable as .bookmarkMarker; got \(inner)")

        // Assert nextBookmarkId calibration picked up the nested id.
        XCTAssertGreaterThan(doc.nextBookmarkId, 500,
            "nextBookmarkId SHALL reflect SDT-nested bookmark id (A-CONT-2 P0 #2 — calibration walker must recurse through .contentControl); got \(doc.nextBookmarkId)")
    }

    /// §1.30 — Synthetic fixture matrix-pin: assert container bookmark count
    /// parity on a fixture that ACTUALLY has body-level bookmarks in headers.
    /// The thesis fixture has 0 such bookmarks across all 12 container parts,
    /// so `assertContainerBookmarkStartParity` against it is regression-blind
    /// (asserts 0=0). This synthetic test ensures the matrix-pin can detect
    /// regressions for real (asserts 2=2 by construction).
    func testMatrixPinCatchesContainerBookmarkRegression() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body><w:p><w:r><w:t>body</w:t></w:r></w:p></w:body>
        </w:document>
        """
        let headerXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:bookmarkStart w:id="201" w:name="hdrBookmark1"/>
        <w:p><w:r><w:t>part 1</w:t></w:r></w:p>
        <w:bookmarkEnd w:id="201"/>
        <w:bookmarkStart w:id="202" w:name="hdrBookmark2"/>
        <w:p><w:r><w:t>part 2</w:t></w:r></w:p>
        <w:bookmarkEnd w:id="202"/>
        </w:hdr>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-acont2-mp-in-\(UUID().uuidString).docx")
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-acont2-mp-out-\(UUID().uuidString).docx")
        defer {
            try? FileManager.default.removeItem(at: inURL)
            try? FileManager.default.removeItem(at: outURL)
        }
        try buildMinimalDocxWithHeader(documentXML: documentXML, headerXML: headerXML, headerRId: "rId10", to: inURL)

        var doc = try DocxReader.read(from: inURL)
        doc.modifiedParts.insert("word/header1.xml")
        try DocxWriter.write(doc, to: outURL)

        // Apply the matrix-pin's container-source assertion to this fixture.
        try Self.assertContainerBookmarkStartParity(srcURL: inURL, outURL: outURL)

        // Sanity: our fixture actually has 2 bookmarkStarts in header1.xml so
        // the assertion above asserts 2=2, not the regression-blind 0=0 it
        // does on the thesis fixture.
        let headerOut = try Self.readPartXMLString(from: outURL, partPath: "word/header1.xml")
        XCTAssertEqual(Self.countBookmarkStartElements(in: headerOut), 2,
            "synthetic fixture sanity: 2 <w:bookmarkStart> SHALL survive in output header; got \(Self.countBookmarkStartElements(in: headerOut))")
    }

    // MARK: - Sub-stack A-CONT-3: deleteBookmark persistence + paragraph-level container coverage + insertBookmark symmetry

    /// §1.39 — `deleteBookmark` for header body-level bookmark SHALL persist to disk.
    /// Pre-A-CONT-3 `Document.swift:2067` does `modifiedParts.insert(headers[i].fileName)`
    /// which inserts BASENAME (`"header1.xml"`); writer overlay-mode dirty-gate at
    /// `DocxWriter.swift:141` checks `dirty.contains("word/\(header.fileName)")` —
    /// looks for FULL PATH (`"word/header1.xml"`). Format mismatch → writer skips
    /// re-emitting → deletion silently lost on disk.
    func testDeleteBookmarkInHeaderPersistsToDisk() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body><w:p><w:r><w:t>body</w:t></w:r></w:p></w:body>
        </w:document>
        """
        let headerXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:bookmarkStart w:id="9" w:name="hdrBookmark"/>
        <w:p><w:r><w:t>x</w:t></w:r></w:p>
        <w:bookmarkEnd w:id="9"/>
        </w:hdr>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-acont3-delhdr-in-\(UUID().uuidString).docx")
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-acont3-delhdr-out-\(UUID().uuidString).docx")
        defer {
            try? FileManager.default.removeItem(at: inURL)
            try? FileManager.default.removeItem(at: outURL)
        }
        try buildMinimalDocxWithHeader(documentXML: documentXML, headerXML: headerXML, headerRId: "rId10", to: inURL)

        var doc = try DocxReader.read(from: inURL)
        try doc.deleteBookmark(name: "hdrBookmark")
        try DocxWriter.write(doc, to: outURL)

        let outHeaderXML = try Self.readPartXMLString(from: outURL, partPath: "word/header1.xml")
        XCTAssertFalse(
            outHeaderXML.contains("hdrBookmark"),
            "header bookmark deletion SHALL persist to disk (A-CONT-3 P0 #1 — modifiedParts path mismatch silently dropped change). Header still contains:\n\(outHeaderXML)"
        )
    }

    /// §1.41 — `getBookmarks()` SHALL surface paragraph-level bookmarks from
    /// container paragraphs (not just body-level container markers).
    /// Pre-A-CONT-3 `collectBodyLevelBookmarkNamesRecursive` skipped `.paragraph`
    /// cases entirely — paragraph-level container bookmarks (the more common
    /// case) were invisible to MCP `list_bookmarks`.
    func testGetBookmarksSurfacesContainerParagraphLevelBookmarks() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body><w:p><w:r><w:t>body</w:t></w:r></w:p></w:body>
        </w:document>
        """
        let headerXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:p><w:bookmarkStart w:id="50" w:name="paraInHdr"/><w:bookmarkEnd w:id="50"/><w:r><w:t>x</w:t></w:r></w:p>
        </w:hdr>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-acont3-paracov-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try buildMinimalDocxWithHeader(documentXML: documentXML, headerXML: headerXML, headerRId: "rId10", to: inURL)

        let doc = try DocxReader.read(from: inURL)
        let names = doc.getBookmarks().map { $0.name }
        XCTAssertTrue(
            names.contains("paraInHdr"),
            "paragraph-level bookmark inside header SHALL be surfaced by getBookmarks() (A-CONT-3 P0 #2); got \(names)"
        )
    }

    /// §1.42 — `insertBookmark` SHALL detect duplicate names across all 5 part
    /// types (body + headers + footers + footnotes + endnotes), not just body.
    /// Pre-A-CONT-3 `insertBookmark` only walked body for duplicate detection;
    /// after A-CONT-2 a TOC anchor in header survived `insertBookmark(name: ...)`
    /// and produced silent name collision.
    func testInsertBookmarkDuplicateNameInContainerThrows() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body><w:p><w:r><w:t>body</w:t></w:r></w:p></w:body>
        </w:document>
        """
        let headerXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:bookmarkStart w:id="3" w:name="dupName"/>
        <w:p><w:r><w:t>x</w:t></w:r></w:p>
        <w:bookmarkEnd w:id="3"/>
        </w:hdr>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-acont3-dup-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try buildMinimalDocxWithHeader(documentXML: documentXML, headerXML: headerXML, headerRId: "rId10", to: inURL)

        var doc = try DocxReader.read(from: inURL)
        // Attempt to insert a bookmark with the same name targeting body para 0.
        XCTAssertThrowsError(
            try doc.insertBookmark(name: "dupName", at: 0),
            "insertBookmark SHALL throw on duplicate name across all 5 part types (A-CONT-3 P0 #3); did not throw"
        )
    }

    // MARK: - Cross-cutting matrix-pin (incremental — sub-stack A initial version)

    /// §1.7 / §2.7 / §3.9 — Cross-cutting content-equality invariant against
    /// the thesis fixture. Asserts that for every preservation class covered
    /// by this Spectra change, the round-tripped `word/document.xml` content
    /// equals the source content.
    ///
    /// **Sub-stack A** (this version): `<w:bookmarkStart>` count parity (#58).
    /// **Sub-stack B** (lands with §2.7): + `<w:t>` total-character parity (#59).
    /// **Sub-stack C** (lands with §3.9): + `<w:rFonts>` / `<w:noProof>` /
    /// `<w:lang>` / `<w:kern>` / `w14:*` count parity (#60).
    ///
    /// The pin asserts CONTENT equality (counts and joined-strings), not BYTE
    /// equality — Word's own canonicalization (e.g., adjacent run consolidation)
    /// is allowed to differ.
    ///
    /// Covers spec requirement: `testDocumentContentEqualityInvariant matrix-pin
    /// SHALL assert content equality across preservation classes` — initial
    /// version covering preservation-class 1 of 3
    /// (specs/ooxml-roundtrip-fidelity/spec.md).
    func testDocumentContentEqualityInvariant() throws {
        let fixturePath = "/Users/che/Developer/macdoc/mcp/che-word-mcp/test-files/thesis-fixture.docx"
        guard FileManager.default.fileExists(atPath: fixturePath) else {
            throw XCTSkip("thesis fixture not present at \(fixturePath); skipping content-equality matrix-pin")
        }
        let srcURL = URL(fileURLWithPath: fixturePath)

        // Read source document.xml directly from the ZIP for ground-truth counts.
        let srcDocXML = try Self.readDocumentXMLString(from: srcURL)
        let srcBookmarkStartCount = Self.countBookmarkStartElements(in: srcDocXML)
        XCTAssertGreaterThan(srcBookmarkStartCount, 0,
            "fixture sanity: source has at least one <w:bookmarkStart>; got \(srcBookmarkStartCount)")

        // Round-trip: read → mark modified → write → re-read document.xml.
        var doc = try DocxReader.read(from: srcURL)
        doc.modifiedParts.insert("word/document.xml")
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("matrix-pin-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(doc, to: outURL)

        let outDocXML = try Self.readDocumentXMLString(from: outURL)
        let outBookmarkStartCount = Self.countBookmarkStartElements(in: outDocXML)

        // Preservation class 1 of 3 (#58): bookmarkStart count parity.
        XCTAssertEqual(
            outBookmarkStartCount, srcBookmarkStartCount,
            "<w:bookmarkStart> count SHALL be preserved across round-trip; src=\(srcBookmarkStartCount), out=\(outBookmarkStartCount)"
        )

        // §1.18 (A-CONT P1): container-source assertion. The thesis fixture has
        // 6 distinct headers and 4 footers; assert bookmarkStart parity for each
        // container part. Pre-A-CONT-fix, body-level bookmarks in headers/
        // footers/footnotes/endnotes were silently dropped on save (parser
        // asymmetry — `parseContainerChildBodyChildren` had only `case "p"` /
        // `case "tbl"` / `default: continue` while `parseBodyChildren` had the
        // typed cases). A-CONT mirrors the fix into the second parser entry
        // point. The assertion below catches future regressions in that fix.
        try Self.assertContainerBookmarkStartParity(srcURL: srcURL, outURL: outURL)

        // Preservation class 2 of 3 (#59): <w:t> total character content parity.
        // §2.23 (B-CONT): the §2.7 placeholder is replaced with a real
        // assertion. Sums the inner-text length of every `<w:t>` element in
        // both source and output `document.xml`, asserts equality. Catches
        // any whitespace-overlay regression — the very class of bug that
        // sub-stack B's BLOCK verify identified (counter desync silently
        // dropping whitespace in tables / mc:Choice / raw-captured wrappers).
        let srcWtCharSum = Self.sumWtElementCharCount(in: srcDocXML)
        let outWtCharSum = Self.sumWtElementCharCount(in: outDocXML)
        XCTAssertEqual(
            outWtCharSum, srcWtCharSum,
            "<w:t> total character content SHALL be preserved across round-trip "
            + "(B-CONT — preservation class 2 of 3 / #59); src=\(srcWtCharSum), out=\(outWtCharSum). "
            + "A mismatch indicates the WhitespaceOverlay counter desynced from parseRun's "
            + "visit order (likely a new prefix-collision class or unhandled raw-capture site)."
        )

        // Preservation class 3 of 3 (#60): <w:rFonts>/<w:noProof>/<w:lang>/<w:kern>/w14
        // round-trip — RUN-LEVEL ONLY scope. Sub-stack C extends RunProperties
        // with rFonts (4-axis), noProof, kern, lang, and rawChildren — for
        // `<w:r><w:rPr>...</w:rPr></w:r>` paths.
        //
        // OUT OF SCOPE for sub-stack C (separate pre-existing bugs surfaced
        // by this matrix-pin and tracked for follow-up):
        //   1. ParagraphProperties has NO markRunProperties field — the
        //      `<w:pPr><w:rPr>...</w:rPr></w:pPr>` (paragraph-mark formatting
        //      that controls the pilcrow-glyph appearance) is silently dropped
        //      at parse time. Accounts for ~50% of <w:lang> loss.
        //   2. Paragraph parser doesn't preserve `w14:paraId`/`w14:textId`
        //      attributes on `<w:p>` (Word's revision-tracking GUIDs).
        //      Accounts for ~95% of the w14:* token loss (2214 of 2359
        //      tokens are these two attributes).
        //
        // To keep the matrix-pin LOAD-BEARING for #60's actual scope while
        // not blocking on these out-of-scope drops, this assertion uses a
        // RATIO floor calibrated to current behavior — any regression below
        // the floor (e.g., my fix breaks instead of preserves) is caught;
        // future improvements to the out-of-scope paths can ratchet floors up.
        //
        // Floors are set conservatively from the post-sub-stack-C measured
        // baseline: rFonts 88%, noProof 92%, lang 50%, kern 84%, w14:* 5%.
        // ANY drop below floor in a future change indicates RunProperties
        // regression and must trip the matrix-pin.
        let preservationClassFloors: [(name: String, floor: Double)] = [
            ("<w:rFonts", 0.85),  // measured: 88% (740 lost in pPr/rPr drop)
            ("<w:noProof", 0.90), // measured: 92%
            ("<w:lang ", 0.45),   // measured: 50% (45 in pPr/rPr drop)
            ("<w:kern ", 0.80),   // measured: 84%
            ("w14:", 0.04)        // measured: 5% — mostly paraId/textId attrs out-of-scope
        ]
        for class3 in preservationClassFloors {
            let srcCount = Self.countSubstring(class3.name, in: srcDocXML)
            let outCount = Self.countSubstring(class3.name, in: outDocXML)
            let ratio = srcCount > 0 ? Double(outCount) / Double(srcCount) : 1.0
            XCTAssertGreaterThanOrEqual(
                ratio, class3.floor,
                "preservation-class-3 (#60): `\(class3.name)` retention SHALL stay >= "
                + "\(class3.floor) across round-trip. src=\(srcCount), out=\(outCount), "
                + "ratio=\(ratio). A drop below floor indicates RunProperties (sub-stack C) "
                + "regression. NOTE: the floor reflects out-of-scope losses already known: "
                + "paragraph-mark rPr (pPr/rPr) drop + w14:paraId/textId paragraph-attribute "
                + "drop. These are separate pre-existing bugs, NOT regressions from sub-stack C."
            )
        }

        // §3.11 (sub-stack C) — thesis fixture round-trip size sanity check.
        // Pre-fix (v0.19.x) document.xml shrunk from 1473896 → 1006805 bytes
        // (32% loss). Sub-stack C reduced this to ~17.75% by recovering rPr
        // typed fields (rFonts 4-axis, noProof, kern, lang) + w14:* rawChildren.
        // Remaining 17.75% loss is driven by known out-of-scope drops:
        // paragraph-mark rPr (`<w:pPr><w:rPr>`) silent drop + `w14:paraId`/
        // `w14:textId` paragraph-attribute drop. Both tracked as follow-up SDD.
        //
        // Floor set to 19% — catches REGRESSIONS that grow loss above the
        // post-sub-stack-C baseline. Future paragraph-mark rPr fix should
        // drop loss to <5% and allow ratcheting this floor down.
        let srcBytes = srcDocXML.utf8.count
        let outBytes = outDocXML.utf8.count
        let sizeLossRatio = Double(srcBytes - outBytes) / Double(srcBytes)
        XCTAssertLessThanOrEqual(
            sizeLossRatio, 0.175,
            "thesis fixture round-trip size SHALL stay within 17.5% of source — "
            + "the post-sub-stack-C-CONT baseline (#60 §3.11 + R2/R5/Codex hotfix). "
            + "src=\(srcBytes) bytes, out=\(outBytes) bytes, loss=\(sizeLossRatio * 100)%. "
            + "Progression: pre-fix v0.19.x 32% loss → sub-stack C v0.20.0 17.75% → "
            + "sub-stack C-CONT v0.20.1 16.66% (recognizedRprChildren trim closed silent "
            + "drop of <w:caps>/<w:smallCaps>/<w:spacing>/<w:position>/<w:shd>/<w:bdr>/<w:em>/etc). "
            + "Remaining ~16.66% is paragraph-mark rPr + w14:paraId/textId drops "
            + "(separate follow-up SDD)."
        )
    }

    /// §3.9 helper: count substring occurrences via simple linear scan.
    /// Same pattern as `countBookmarkStartElements` / `countDelTextElements`.
    static func countSubstring(_ needle: String, in xml: String) -> Int {
        var count = 0
        var searchRange = xml.startIndex..<xml.endIndex
        while let r = xml.range(of: needle, range: searchRange) {
            count += 1
            searchRange = r.upperBound..<xml.endIndex
        }
        return count
    }

    /// §2.23 (B-CONT) helper: sum the inner-text character count across every
    /// `<w:t>` element in the given XML. Used by the matrix-pin to detect
    /// whitespace-overlay regressions — any silent loss of `<w:t>` content
    /// during round-trip shows up as a delta.
    ///
    /// Applies the same tag-name boundary check as `WhitespaceOverlay.scanning`
    /// (rejects prefix collisions like `<w:tab>`, `<w:tbl>`).
    static func sumWtElementCharCount(in xml: String) -> Int {
        var sum = 0
        var searchStart = xml.startIndex
        while let openRange = xml.range(of: "<w:t", range: searchStart..<xml.endIndex) {
            let afterToken = openRange.upperBound
            guard afterToken < xml.endIndex else { break }
            let boundary = xml[afterToken]
            guard boundary == ">" || boundary == " " || boundary == "\t"
                  || boundary == "\n" || boundary == "\r" || boundary == "/" else {
                searchStart = afterToken
                continue
            }
            // Find tag close (handle attributes, self-close).
            guard let tagClose = xml.range(of: ">", range: afterToken..<xml.endIndex) else {
                searchStart = afterToken
                continue
            }
            let tagAttrs = String(xml[afterToken..<tagClose.lowerBound])
            // Self-close `<w:t/>` has no inner content.
            if tagAttrs.hasSuffix("/") {
                searchStart = tagClose.upperBound
                continue
            }
            guard let closeRange = xml.range(of: "</w:t>", range: tagClose.upperBound..<xml.endIndex) else {
                searchStart = tagClose.upperBound
                continue
            }
            let inner = xml[tagClose.upperBound..<closeRange.lowerBound]
            sum += inner.count
            searchStart = closeRange.upperBound
        }
        return sum
    }

    /// Count `<w:bookmarkStart` elements in raw XML via simple substring scan.
    /// Avoids regex compilation overhead and matches both `<w:bookmarkStart`
    /// followed by space/attr or `/>`.
    static func countBookmarkStartElements(in xml: String) -> Int {
        var count = 0
        var searchRange = xml.startIndex..<xml.endIndex
        while let r = xml.range(of: "<w:bookmarkStart", range: searchRange) {
            count += 1
            searchRange = r.upperBound..<xml.endIndex
        }
        return count
    }

    /// §1.18 (A-CONT P1) helper: enumerate all `word/header*.xml`,
    /// `word/footer*.xml`, `word/footnotes.xml`, `word/endnotes.xml` parts in
    /// both source and output `.docx`, count `<w:bookmarkStart>` per part,
    /// assert parity. Catches the parser-asymmetry class of regression.
    static func assertContainerBookmarkStartParity(srcURL: URL, outURL: URL) throws {
        let srcUnzipped = try ZipHelper.unzip(srcURL)
        defer { ZipHelper.cleanup(srcUnzipped) }
        let outUnzipped = try ZipHelper.unzip(outURL)
        defer { ZipHelper.cleanup(outUnzipped) }

        let containerPartPatterns = ["word/header", "word/footer", "word/footnotes.xml", "word/endnotes.xml"]
        let srcWordDir = srcUnzipped.appendingPathComponent("word")
        let fileNames = (try? FileManager.default.contentsOfDirectory(atPath: srcWordDir.path)) ?? []

        for name in fileNames {
            let isContainer = containerPartPatterns.contains { name.hasPrefix($0.replacingOccurrences(of: "word/", with: "")) || "word/\(name)" == $0 }
            guard isContainer, name.hasSuffix(".xml") else { continue }
            let srcXML = (try? String(contentsOf: srcUnzipped.appendingPathComponent("word/\(name)"), encoding: .utf8)) ?? ""
            let outXML = (try? String(contentsOf: outUnzipped.appendingPathComponent("word/\(name)"), encoding: .utf8)) ?? ""
            let srcCount = countBookmarkStartElements(in: srcXML)
            let outCount = countBookmarkStartElements(in: outXML)
            XCTAssertEqual(
                outCount, srcCount,
                "container `word/\(name)` <w:bookmarkStart> count SHALL be preserved across round-trip (#58 A-CONT — parser asymmetry between parseBodyChildren and parseContainerChildBodyChildren); src=\(srcCount), out=\(outCount)"
            )
        }
    }

    // MARK: - Sub-stack B: Whitespace overlay (#59)
    //
    // Methodology lesson from sub-stack A's 4 sub-cycles: matrix-pin needs
    // SYMMETRIC ASSERTIONS BAKED IN FROM DESIGN, not added reactively in
    // response to verify findings. So sub-stack B's tests exercise ALL 6
    // part types upfront — body, header1, footer1, footnotes, endnotes,
    // comments — so the convergence cycle is shorter from the start.

    /// §2.1 — Body `<w:t>` whitespace SHALL survive Reader-side parser limitations.
    /// Foundation `XMLDocument` strips whitespace-only text nodes regardless of
    /// `xml:space="preserve"` AND regardless of `.nodePreserveWhitespace` option
    /// (verified by isolated probe in #59 diagnosis). The fix is a pre-parse
    /// scan over raw XML bytes that captures `<w:t xml:space="preserve">[ws]</w:t>`
    /// content and is consulted by `parseRun` when `t.stringValue.isEmpty`.
    func testWhitespaceOnlyTextRunsRoundTripInBody() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p>
        <w:r><w:t>before</w:t></w:r>
        <w:r><w:t xml:space="preserve">     </w:t></w:r>
        <w:r><w:t>after</w:t></w:r>
        </w:p>
        </w:body>
        </w:document>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue59-body-ws-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try buildMinimalDocx(documentXML: documentXML, to: inURL)

        let doc = try DocxReader.read(from: inURL)
        guard case .paragraph(let para) = doc.body.children.first else {
            XCTFail("expected one paragraph in body")
            return
        }
        XCTAssertEqual(para.runs.count, 3, "expected 3 runs (before / 5-space / after)")
        XCTAssertEqual(
            para.runs[1].text, "     ",
            "5-char whitespace run SHALL survive Reader (sub-stack B P0 — Foundation XMLDocument strips whitespace-only <w:t> stringValue regardless of xml:space=preserve); got \"\(para.runs[1].text)\""
        )
    }

    /// §2.9 — Comprehensive matrix-pin test exercising whitespace preservation
    /// across ALL 6 part types in a single fixture. Pre-implementation: this
    /// test fails on body + header + footer + footnotes + endnotes + comments
    /// (Foundation parser limitation is parser-wide). Post-implementation: all
    /// 6 part types preserve whitespace via WhitespaceOverlay.
    ///
    /// The test design follows the methodology lesson from sub-stack A's
    /// 4 sub-cycles: assert ALL part types from the start, not just one.
    func testWhitespacePreservedAcrossAllSixPartTypes() throws {
        // Build a docx with whitespace-only <w:t> in:
        //   - word/document.xml      (body)
        //   - word/header1.xml       (header)
        //   - word/footer1.xml       (footer)
        //   - word/footnotes.xml     (footnotes)
        //   - word/endnotes.xml      (endnotes)
        //   - word/comments.xml      (comments)
        // Each contains a 5-character whitespace `<w:t xml:space="preserve">     </w:t>`
        // wrapped in non-whitespace runs so the joined-text assertion catches
        // dropped whitespace.
        //
        // For sub-stack B P0: §2.1 covers body in isolation; this test extends
        // coverage to all 6 part types so the convergence cycle catches every
        // symmetric-sibling regression from the START rather than after each
        // verify round.
        //
        // Fixture builder added in §2.9-helpers below.
        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue59-allparts-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try Self.buildAllPartsWhitespaceFixture(to: inURL)

        let doc = try DocxReader.read(from: inURL)

        // Helper: scan a part's bodyChildren array for the 5-char whitespace run.
        func partContainsFiveSpaceRun(in children: [BodyChild], partName: String) -> (found: Bool, message: String) {
            var found = false
            var observedTexts: [String] = []
            for child in children {
                if case .paragraph(let para) = child {
                    for run in para.runs {
                        observedTexts.append(run.text)
                        if run.text == "     " {
                            found = true
                        }
                    }
                }
            }
            return (found, "Part \(partName) — runs observed: \(observedTexts)")
        }

        // Body
        let bodyResult = partContainsFiveSpaceRun(in: doc.body.children, partName: "body")
        XCTAssertTrue(bodyResult.found, "BODY whitespace SHALL survive (sub-stack B). \(bodyResult.message)")

        // Header (use first header)
        if let header = doc.headers.first {
            let r = partContainsFiveSpaceRun(in: header.bodyChildren, partName: "header1")
            XCTAssertTrue(r.found, "HEADER whitespace SHALL survive (sub-stack B). \(r.message)")
        } else {
            XCTFail("expected at least one header parsed")
        }

        // Footer (use first footer)
        if let footer = doc.footers.first {
            let r = partContainsFiveSpaceRun(in: footer.bodyChildren, partName: "footer1")
            XCTAssertTrue(r.found, "FOOTER whitespace SHALL survive (sub-stack B). \(r.message)")
        } else {
            XCTFail("expected at least one footer parsed")
        }

        // Footnotes
        if let footnote = doc.footnotes.footnotes.first {
            let r = partContainsFiveSpaceRun(in: footnote.bodyChildren, partName: "footnotes")
            XCTAssertTrue(r.found, "FOOTNOTES whitespace SHALL survive (sub-stack B). \(r.message)")
        } else {
            XCTFail("expected at least one footnote parsed")
        }

        // Endnotes
        if let endnote = doc.endnotes.endnotes.first {
            let r = partContainsFiveSpaceRun(in: endnote.bodyChildren, partName: "endnotes")
            XCTAssertTrue(r.found, "ENDNOTES whitespace SHALL survive (sub-stack B). \(r.message)")
        } else {
            XCTFail("expected at least one endnote parsed")
        }

        // Comments — comments are in a separate model. Probe via getComments().
        let comments = doc.comments.comments
        if let comment = comments.first {
            // Comment text is a flat string; check it contains 5 spaces.
            XCTAssertTrue(
                comment.text.contains("     "),
                "COMMENTS whitespace SHALL survive (sub-stack B). Comment text: \"\(comment.text)\""
            )
        } else {
            XCTFail("expected at least one comment parsed")
        }
    }

    // MARK: - Sub-stack B-CONT: WhitespaceOverlay counter-desync regressions
    //
    // The sub-stack B 6-AI verify (#59 issuecomment-4323956207) returned BLOCK
    // with 4-reviewer convergence on a P0 counter-desync class. Two root causes:
    //
    // - Root cause A — prefix-match collision in `WhitespaceOverlay.swift:54`:
    //   `xml.range(of: "<w:t", ...)` is a prefix match that also fires on
    //   `<w:tab>`, `<w:tbl>`, `<w:tc>`, `<w:tr>`, `<w:tblPr>`, etc. The DOM
    //   walker `element.elements(forName: "w:t")` is exact-match. Counter
    //   desyncs immediately in any document with tables or tabs.
    //
    // - Root cause B — skipped raw subtrees: `parseAlternateContent` skips
    //   `<mc:Choice>`; `parseInsRevisionWrapper` raw-captures wrappers with
    //   non-run children as `unrecognizedChildren` (see DocxReader §868 path).
    //   Scanner counts `<w:t>` inside skipped/raw-captured ranges; parseRun
    //   never visits them. Counter desyncs per skipped subtree.
    //
    // The 4 tests below exercise each root cause with the exact OOXML pattern
    // from the verify report. Each SHALL fail pre-B-CONT-fix.

    /// §2.16 (B-CONT) — `<w:tab/>` adjacent to whitespace `<w:t>` does not
    /// desync the WhitespaceOverlay counter.
    ///
    /// Pre-fix: scanner's prefix match `<w:t` fires on `<w:tab/>`, increments
    /// `index` to 1; whitespace `<w:t>` then maps to index 2 instead of 1.
    /// parseRun walks DOM (which doesn't see `<w:tab/>` as `<w:t>`) and queries
    /// index 1 → nil → falls back to Foundation's stripped "" → whitespace LOST.
    func testWhitespaceOverlayPrefixMatchTabDoesNotDesyncCounter() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p>
        <w:r><w:tab/></w:r>
        <w:r><w:t xml:space="preserve">     </w:t></w:r>
        <w:r><w:t>after</w:t></w:r>
        </w:p>
        </w:body>
        </w:document>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue59-bcont-tab-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try buildMinimalDocx(documentXML: documentXML, to: inURL)

        let doc = try DocxReader.read(from: inURL)
        guard case .paragraph(let para) = doc.body.children.first else {
            XCTFail("expected one paragraph in body")
            return
        }
        // Find the run that should contain 5 spaces (it's the run between the
        // tab run and the "after" run).
        let whitespaceRun = para.runs.first(where: { $0.text == "     " })
        XCTAssertNotNil(
            whitespaceRun,
            "5-char whitespace run SHALL survive even when preceded by <w:tab/> "
            + "(B-CONT P0 root-cause-A — prefix-match collision in WhitespaceOverlay scanner). "
            + "Observed runs: \(para.runs.map { "\"\($0.text)\"" })"
        )
    }

    /// §2.17 (B-CONT) — `<w:tbl>` between paragraphs does not desync the
    /// WhitespaceOverlay counter or pathologically consume the next `</w:t>`.
    ///
    /// Pre-fix: tables contain `<w:tblPr>`, `<w:tblGrid>`, `<w:tr>`, `<w:trPr>`,
    /// `<w:tc>`, `<w:tcPr>` — all of which prefix-match `<w:t`. Scanner counter
    /// desyncs by 6+ per table. AND for `<w:tbl>` itself, scanner searches
    /// forward for `</w:t>` (not `</w:tbl>`) and consumes the next legitimate
    /// `</w:t>` closer, swallowing real whitespace `<w:t>` elements between.
    /// R5's empirical probe confirmed permanent loss.
    func testWhitespaceOverlayPrefixMatchTableDoesNotDesyncCounter() throws {
        // Empty-cell table forces the bug to be visible: the pathological
        // skip-over (scanner false-matches `<w:tbl>`, searches forward for
        // `</w:t>`) lands on the whitespace `</w:t>` itself (no other `<w:t>`
        // between table-open and the whitespace one), absorbs it. Map ends
        // empty; parseRun queries overlay[N] → nil → whitespace LOST.
        // (A table with non-empty cells like `<w:t>cell</w:t>` masks the bug
        // because parseRun's cell-text visit aligns counters by accident.)
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p><w:r><w:t>before</w:t></w:r></w:p>
        <w:tbl>
        <w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
        <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
        <w:tr>
        <w:trPr><w:trHeight w:val="240"/></w:trPr>
        <w:tc>
        <w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr>
        <w:p/>
        </w:tc>
        </w:tr>
        </w:tbl>
        <w:p><w:r><w:t xml:space="preserve">     </w:t></w:r></w:p>
        </w:body>
        </w:document>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue59-bcont-table-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try buildMinimalDocx(documentXML: documentXML, to: inURL)

        let doc = try DocxReader.read(from: inURL)
        // The post-table paragraph SHALL contain a run with 5 spaces.
        // Body children: [.paragraph(before), .table(_), .paragraph(whitespace)]
        var foundWhitespaceRun = false
        var observedTexts: [String] = []
        for child in doc.body.children {
            if case .paragraph(let para) = child {
                for run in para.runs {
                    observedTexts.append(run.text)
                    if run.text == "     " {
                        foundWhitespaceRun = true
                    }
                }
            }
        }
        XCTAssertTrue(
            foundWhitespaceRun,
            "5-char whitespace run after table SHALL survive (B-CONT P0 root-cause-A — "
            + "<w:tbl>/<w:tblPr>/<w:tr>/<w:trPr>/<w:tc>/<w:tcPr> all prefix-match `<w:t`, "
            + "and pathological skip-over consumes wrong </w:t>). Observed text runs: \(observedTexts)"
        )
    }

    /// §2.18 (B-CONT) — `<mc:AlternateContent>` with `<w:t>` in BOTH Choice
    /// and Fallback branches does not desync counter.
    ///
    /// Pre-fix: `parseAlternateContent` only walks `<mc:Fallback>` runs;
    /// `<mc:Choice>` runs are skipped (raw-stored). But the byte scanner
    /// counts `<w:t>` from BOTH branches → DOM walk visits fewer than scanner
    /// counts → subsequent recoveries shifted by N positions.
    func testWhitespaceOverlayMcAlternateContentDoesNotDesyncCounter() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" \
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" \
        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" \
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" \
        mc:Ignorable="w14">
        <w:body>
        <w:p>
        <w:r>
        <mc:AlternateContent>
        <mc:Choice Requires="w14"><w:r><w:t>choice-text</w:t></w:r></mc:Choice>
        <mc:Fallback><w:r><w:t>fallback-text</w:t></w:r></mc:Fallback>
        </mc:AlternateContent>
        </w:r>
        </w:p>
        <w:p><w:r><w:t xml:space="preserve">     </w:t></w:r></w:p>
        </w:body>
        </w:document>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue59-bcont-mc-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try buildMinimalDocx(documentXML: documentXML, to: inURL)

        let doc = try DocxReader.read(from: inURL)
        var foundWhitespaceRun = false
        var observedTexts: [String] = []
        for child in doc.body.children {
            if case .paragraph(let para) = child {
                for run in para.runs {
                    observedTexts.append(run.text)
                    if run.text == "     " {
                        foundWhitespaceRun = true
                    }
                }
            }
        }
        XCTAssertTrue(
            foundWhitespaceRun,
            "5-char whitespace run after mc:AlternateContent SHALL survive (B-CONT P0 root-cause-B — "
            + "scanner counts <w:t> in both <mc:Choice> AND <mc:Fallback>, parseAlternateContent only "
            + "walks Fallback). Observed text runs: \(observedTexts)"
        )
    }

    /// §2.19 (B-CONT) — `<w:ins>` wrapper with non-run children (forces raw
    /// capture via `hasNonRunChild` path) does not desync counter.
    ///
    /// Pre-fix: when `<w:ins>` has any non-run child (e.g., `<w:bookmarkStart>`),
    /// `parseInsRevisionWrapper` stores the entire wrapper as raw XML in
    /// `unrecognizedChildren` and parseRun is NEVER called for the wrapper's
    /// `<w:t>` elements. But scanner still counts them. Same desync class.
    /// §2.25 (B-CONT P1) — comments containing ONLY whitespace are preserved,
    /// not trimmed away.
    ///
    /// Pre-fix: `parseComments` calls `text.trimmingCharacters(in: .whitespacesAndNewlines)`
    /// at DocxReader.swift:2978, which destroys the recovered overlay text for
    /// any whitespace-only comment. Codex P1 surfaced this — overlay correctly
    /// recovers the whitespace bytes, then trim throws them away.
    ///
    /// Real-world scenario: a Word user inserts a comment containing only
    /// indentation/spacing as a "look-here" annotation. Round-trip silently
    /// changes the comment to empty.
    func testWhitespaceOnlyCommentPreservedNotTrimmed() throws {
        // Build a fixture with a comment whose entire text is "     " (5 spaces)
        let stagingURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue59-bcont-wsonlycomment-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: stagingURL, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: stagingURL) }

        try FileManager.default.createDirectory(at: stagingURL.appendingPathComponent("_rels"), withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: stagingURL.appendingPathComponent("word/_rels"), withIntermediateDirectories: true)

        let contentTypes = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
        </Types>
        """
        try contentTypes.write(to: stagingURL.appendingPathComponent("[Content_Types].xml"), atomically: true, encoding: .utf8)

        let rootRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """
        try rootRels.write(to: stagingURL.appendingPathComponent("_rels/.rels"), atomically: true, encoding: .utf8)

        let docRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
        </Relationships>
        """
        try docRels.write(to: stagingURL.appendingPathComponent("word/_rels/document.xml.rels"), atomically: true, encoding: .utf8)

        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body><w:p><w:r><w:t>body</w:t></w:r></w:p></w:body>
        </w:document>
        """
        try documentXML.write(to: stagingURL.appendingPathComponent("word/document.xml"), atomically: true, encoding: .utf8)

        // Comment with ONLY whitespace text (5 spaces). Pre-fix: parseComments'
        // trim turns this into "".
        let commentsXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:comment w:id="1" w:author="tester" w:date="2026-04-27T00:00:00Z" w:initials="t">
        <w:p><w:r><w:t xml:space="preserve">     </w:t></w:r></w:p>
        </w:comment>
        </w:comments>
        """
        try commentsXML.write(to: stagingURL.appendingPathComponent("word/comments.xml"), atomically: true, encoding: .utf8)

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue59-bcont-wsonlycomment-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try ZipHelper.zip(stagingURL, to: inURL)

        let doc = try DocxReader.read(from: inURL)
        guard let comment = doc.comments.comments.first else {
            XCTFail("expected one comment parsed")
            return
        }
        XCTAssertEqual(
            comment.text, "     ",
            "Whitespace-only comment text SHALL be preserved exactly (B-CONT P1 — "
            + "Codex finding: parseComments' trimmingCharacters(in: .whitespacesAndNewlines) "
            + "destroys recovered overlay text for whitespace-only comments). got \"\(comment.text)\""
        )
    }

    // MARK: - Sub-stack B-CONT-2: parseRun rawElements double-emit + missed raw-capture sites
    //
    // Sub-stack B-CONT 6-AI verify (#59 comment 4324076688) found 6 P0 + 3 P1.
    // CRITICAL: parseRun's recognizedRunChildren = ["rPr", "t", "drawing",
    // "oMath", "oMathPara"] doesn't include "delText". So `<w:delText>` falls
    // through the rawElements path, causing:
    //   1. delTextCounter advances 2x per delText (explicit loop + rawElements)
    //   2. Writer emits delText TWICE on save (once for run.text, once from
    //      rawElements iteration via Run.toXML)
    // v3.13.11 in production silently corrupts every <w:del> round-trip.

    /// §2.33 (B-CONT-2 TIER-0) — `<w:delText>` is emitted EXACTLY ONCE per
    /// source element after round-trip. Pre-fix the rawElements re-emission
    /// path duplicates it.
    func testDelTextEmittedExactlyOncePerSourceElement() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p>
        <w:r><w:t>before</w:t></w:r>
        <w:del w:id="1" w:author="tester" w:date="2026-04-27T00:00:00Z">
        <w:r><w:delText>deleted-text</w:delText></w:r>
        </w:del>
        <w:r><w:t>after</w:t></w:r>
        </w:p>
        </w:body>
        </w:document>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue59-bcont2-deldup-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try buildMinimalDocx(documentXML: documentXML, to: inURL)

        var doc = try DocxReader.read(from: inURL)
        // Force re-serialize so writer emits document.xml.
        doc.modifiedParts.insert("word/document.xml")
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue59-bcont2-deldup-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(doc, to: outURL)

        let outDocXML = try Self.readDocumentXMLString(from: outURL)
        // Count <w:delText> opening tags in source vs output. Source has 1.
        let srcCount = Self.countDelTextElements(in: documentXML)
        let outCount = Self.countDelTextElements(in: outDocXML)
        XCTAssertEqual(srcCount, 1, "fixture sanity: source has exactly 1 <w:delText>")
        XCTAssertEqual(
            outCount, srcCount,
            "<w:delText> count SHALL be preserved across round-trip "
            + "(B-CONT-2 TIER-0 P0 — parseRun's recognizedRunChildren missed 'delText', "
            + "rawElements path causes writer duplicate emission). "
            + "src=\(srcCount), out=\(outCount). Output XML:\n\(outDocXML)"
        )
    }

    /// §2.33-CONTENT (B-CONT-2-CONT — R2 finding) — `<w:delText>` CONTENT
    /// survives round-trip, not just the opening tag.
    ///
    /// The original §2.33 test only counts opening tags, which passed even
    /// after TIER-0 introduced an empty-content regression. R2 verify caught
    /// this: parseRun's `<w:t>` loop never sees `<w:delText>`, so run.text="".
    /// With "delText" in recognizedRunChildren, rawElements is empty too.
    /// Writer gate `!run.text.isEmpty || (run.rawElements?.isEmpty ?? true)`
    /// evaluates `false || true` → emits synthetic `<w:delText></w:delText>`
    /// with empty content. Deleted text silently destroyed.
    func testDelTextContentPreservedThroughRoundTrip() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p>
        <w:del w:id="1" w:author="tester" w:date="2026-04-27T00:00:00Z">
        <w:r><w:delText>deleted-content-marker-xyz</w:delText></w:r>
        </w:del>
        </w:p>
        </w:body>
        </w:document>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue59-bcont2cont-content-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try buildMinimalDocx(documentXML: documentXML, to: inURL)

        var doc = try DocxReader.read(from: inURL)
        doc.modifiedParts.insert("word/document.xml")
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue59-bcont2cont-content-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(doc, to: outURL)

        let outDocXML = try Self.readDocumentXMLString(from: outURL)
        XCTAssertTrue(
            outDocXML.contains("deleted-content-marker-xyz"),
            "<w:delText> CONTENT SHALL survive round-trip — not just the opening tag "
            + "(B-CONT-2-CONT — R2 finding: TIER-0 fix removed delText from rawElements path; "
            + "writer's synthetic-emission gate then emits empty <w:delText></w:delText> when "
            + "run.text is empty AND rawElements is empty/nil). Output XML excerpt:\n"
            + "\(outDocXML.prefix(2000))"
        )
    }

    /// §2.34 (B-CONT-2 TIER-0) — multiple `<w:del>` blocks each containing
    /// whitespace `<w:delText>` round-trip without counter desync.
    func testDeleteTextCounterStaysSyncedAcrossMultipleDels() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p>
        <w:del w:id="1" w:author="tester" w:date="2026-04-27T00:00:00Z">
        <w:r><w:delText xml:space="preserve">     </w:delText></w:r>
        </w:del>
        <w:r><w:t>middle</w:t></w:r>
        <w:del w:id="2" w:author="tester" w:date="2026-04-27T00:00:00Z">
        <w:r><w:delText xml:space="preserve">     </w:delText></w:r>
        </w:del>
        </w:p>
        </w:body>
        </w:document>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue59-bcont2-multidel-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try buildMinimalDocx(documentXML: documentXML, to: inURL)

        let doc = try DocxReader.read(from: inURL)
        guard case .paragraph(let para) = doc.body.children.first else {
            XCTFail("expected one paragraph in body")
            return
        }
        let deletionRevs = para.revisions.filter { $0.type == .deletion }
        XCTAssertEqual(deletionRevs.count, 2, "expected 2 .deletion Revisions")
        for (i, rev) in deletionRevs.enumerated() {
            XCTAssertEqual(
                rev.originalText, "     ",
                "Deletion #\(i+1) originalText SHALL be 5 spaces "
                + "(B-CONT-2 TIER-0 P0 — delTextCounter desync from 2x advance). "
                + "Got: \"\(rev.originalText ?? "<nil>")\""
            )
        }
    }

    /// §2.36 (B-CONT-2 TIER-1) — `parseContainerChildBodyChildren` raw fallback
    /// (DocxReader.swift:1494) doesn't desync the whitespace counter.
    func testWhitespaceOverlayContainerRawFallbackDoesNotDesyncCounter() throws {
        // Build a fixture with header containing an unrecognized body-level
        // child element with `<w:t>` inside, before a paragraph with whitespace.
        let stagingURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue59-bcont2-containerraw-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: stagingURL, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: stagingURL) }

        try FileManager.default.createDirectory(at: stagingURL.appendingPathComponent("_rels"), withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: stagingURL.appendingPathComponent("word/_rels"), withIntermediateDirectories: true)

        let contentTypes = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
        </Types>
        """
        try contentTypes.write(to: stagingURL.appendingPathComponent("[Content_Types].xml"), atomically: true, encoding: .utf8)

        let rootRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """
        try rootRels.write(to: stagingURL.appendingPathComponent("_rels/.rels"), atomically: true, encoding: .utf8)

        let docRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
        </Relationships>
        """
        try docRels.write(to: stagingURL.appendingPathComponent("word/_rels/document.xml.rels"), atomically: true, encoding: .utf8)

        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
        <w:p><w:r><w:t>body</w:t></w:r></w:p>
        <w:sectPr><w:headerReference w:type="default" r:id="rId10"/></w:sectPr>
        </w:body>
        </w:document>
        """
        try documentXML.write(to: stagingURL.appendingPathComponent("word/document.xml"), atomically: true, encoding: .utf8)

        // Header with unrecognized body-level child element containing `<w:t>`,
        // before a paragraph with whitespace `<w:t>`. The unknown element triggers
        // parseContainerChildBodyChildren raw fallback (DocxReader.swift:1494).
        // Without counter advance for the raw subtree, scanner counts the inner
        // `<w:t>` but parseRun never visits it → counter desync → whitespace lost.
        // Use a custom namespace element to ensure unrecognized.
        let header1XML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:vendor="urn:test-vendor">
        <vendor:custom><w:r><w:t>vendor-content</w:t></w:r></vendor:custom>
        <w:p><w:r><w:t xml:space="preserve">     </w:t></w:r></w:p>
        </w:hdr>
        """
        try header1XML.write(to: stagingURL.appendingPathComponent("word/header1.xml"), atomically: true, encoding: .utf8)

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue59-bcont2-containerraw-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try ZipHelper.zip(stagingURL, to: inURL)

        let doc = try DocxReader.read(from: inURL)
        guard let header = doc.headers.first else {
            XCTFail("expected one header parsed")
            return
        }
        var foundWhitespaceRun = false
        var observedTexts: [String] = []
        for child in header.bodyChildren {
            if case .paragraph(let para) = child {
                for run in para.runs {
                    observedTexts.append(run.text)
                    if run.text == "     " {
                        foundWhitespaceRun = true
                    }
                }
            }
        }
        XCTAssertTrue(
            foundWhitespaceRun,
            "Header whitespace `<w:t>` after raw-fallback body-child SHALL survive (B-CONT-2 TIER-1 P0 — "
            + "Codex finding: parseContainerChildBodyChildren raw fallback at DocxReader.swift:1494 "
            + "missed advanceWhitespaceCounter call). Observed: \(observedTexts)"
        )
    }

    /// §2.37 (B-CONT-2 TIER-1) — `parseHyperlink` rawChildren branch
    /// (DocxReader.swift:1644) doesn't desync the whitespace counter.
    func testWhitespaceOverlayHyperlinkRawChildrenDoesNotDesyncCounter() throws {
        // Hyperlink containing nested `<w:fldSimple>` (a non-`<w:r>` child)
        // forces the parseHyperlink rawChildren branch. Without counter advance,
        // post-hyperlink whitespace `<w:t>` is lost.
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
        <w:p>
        <w:hyperlink r:id="rId99">
        <w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>
        </w:hyperlink>
        <w:r><w:t xml:space="preserve">     </w:t></w:r>
        </w:p>
        </w:body>
        </w:document>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue59-bcont2-hyperlink-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try buildMinimalDocx(documentXML: documentXML, to: inURL)

        let doc = try DocxReader.read(from: inURL)
        guard case .paragraph(let para) = doc.body.children.first else {
            XCTFail("expected one paragraph in body")
            return
        }
        let whitespaceRun = para.runs.first(where: { $0.text == "     " })
        XCTAssertNotNil(
            whitespaceRun,
            "Whitespace `<w:t>` after hyperlink-with-rawChildren SHALL survive (B-CONT-2 TIER-1 P0 — "
            + "R2 finding: parseHyperlink rawChildren branch missed advanceWhitespaceCounter). "
            + "Observed: \(para.runs.map { "\"\($0.text)\"" })"
        )
    }

    /// §2.38 (B-CONT-2 TIER-1) — `parseParagraph` smartTag/customXml/dir/bdo
    /// raw-carrier blocks don't desync the whitespace counter.
    func testWhitespaceOverlaySmartTagDoesNotDesyncCounter() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p>
        <w:smartTag w:uri="urn:test" w:element="Test">
        <w:r><w:t>tagged</w:t></w:r>
        </w:smartTag>
        <w:r><w:t xml:space="preserve">     </w:t></w:r>
        </w:p>
        </w:body>
        </w:document>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue59-bcont2-smarttag-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try buildMinimalDocx(documentXML: documentXML, to: inURL)

        let doc = try DocxReader.read(from: inURL)
        guard case .paragraph(let para) = doc.body.children.first else {
            XCTFail("expected one paragraph in body")
            return
        }
        let whitespaceRun = para.runs.first(where: { $0.text == "     " })
        XCTAssertNotNil(
            whitespaceRun,
            "Whitespace `<w:t>` after smartTag SHALL survive (B-CONT-2 TIER-1 P0 — "
            + "R2 finding: parseParagraph smartTag/customXml/dir/bdo cases missed advanceWhitespaceCounter). "
            + "Observed: \(para.runs.map { "\"\($0.text)\"" })"
        )
    }

    /// Helper: count `<w:delText>` opening tags in raw XML.
    static func countDelTextElements(in xml: String) -> Int {
        var count = 0
        var searchStart = xml.startIndex
        while let openRange = xml.range(of: "<w:delText", range: searchStart..<xml.endIndex) {
            let afterToken = openRange.upperBound
            guard afterToken < xml.endIndex else { break }
            let boundary = xml[afterToken]
            if boundary == ">" || boundary == " " || boundary == "\t"
                || boundary == "\n" || boundary == "\r" || boundary == "/" {
                count += 1
            }
            searchStart = afterToken
        }
        return count
    }

    /// §2.24 (B-CONT P1) — `<w:delText xml:space="preserve">[whitespace]</w:delText>`
    /// (tracked-deletion of whitespace) is preserved through round-trip.
    ///
    /// Pre-fix: WhitespaceOverlay only scans `<w:t`, parseRun's delText loop
    /// (DocxReader.swift:970) only reads `delText.stringValue ?? ""` —
    /// Foundation strips whitespace `<w:delText>` stringValue same way as
    /// `<w:t>`. The recovered Revision.originalText loses the deleted
    /// whitespace bytes. R5 finding.
    ///
    /// Real-world scenario: a Word user deletes leading indentation under
    /// track-changes. `accept_revision` would commit the deletion correctly
    /// but `reject_revision` couldn't restore the deleted whitespace.
    func testDeleteTextWhitespaceRoundTrips() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p>
        <w:r><w:t>kept-before</w:t></w:r>
        <w:del w:id="1" w:author="tester" w:date="2026-04-27T00:00:00Z">
        <w:r><w:delText xml:space="preserve">     </w:delText></w:r>
        </w:del>
        <w:r><w:t>kept-after</w:t></w:r>
        </w:p>
        </w:body>
        </w:document>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue59-bcont-deltext-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try buildMinimalDocx(documentXML: documentXML, to: inURL)

        let doc = try DocxReader.read(from: inURL)
        guard case .paragraph(let para) = doc.body.children.first else {
            XCTFail("expected one paragraph in body")
            return
        }
        // The deletion's original text should be the 5-space whitespace.
        let deletionRev = para.revisions.first(where: { $0.type == .deletion })
        XCTAssertNotNil(deletionRev, "expected one .deletion Revision")
        XCTAssertEqual(
            deletionRev?.originalText, "     ",
            "Whitespace `<w:delText xml:space=\"preserve\">` SHALL preserve through "
            + "round-trip (B-CONT P1 — R5 finding: WhitespaceOverlay only scans `<w:t`, "
            + "parseRun delText loop reads stringValue which Foundation strips). "
            + "Got: \"\(deletionRev?.originalText ?? "<nil>")\""
        )
    }

    /// §2.26 (B-CONT P1) — entity-encoded whitespace `&#x09;` (tab),
    /// `&#x20;` (space), `&#x0A;` (newline) is recognized as whitespace by
    /// the overlay and recovered correctly.
    ///
    /// Pre-fix: `WhitespaceOverlay.swift:87` runs `inner.allSatisfy({ $0.isWhitespace })`
    /// over RAW XML bytes — `&`, `#`, `x`, `0`, `9` aren't `Character.isWhitespace`,
    /// so entity-encoded whitespace fails the check and isn't stored in the
    /// map. Foundation later decodes the entities but then strips the
    /// resulting whitespace stringValue → permanent loss. R5 finding.
    func testEntityEncodedWhitespacePreserved() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p>
        <w:r><w:t>before</w:t></w:r>
        <w:r><w:t xml:space="preserve">&#x09;&#x09;</w:t></w:r>
        <w:r><w:t>after</w:t></w:r>
        </w:p>
        </w:body>
        </w:document>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue59-bcont-entity-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try buildMinimalDocx(documentXML: documentXML, to: inURL)

        let doc = try DocxReader.read(from: inURL)
        guard case .paragraph(let para) = doc.body.children.first else {
            XCTFail("expected one paragraph in body")
            return
        }
        // Entity-encoded `&#x09;&#x09;` decodes to two tab characters (\t\t).
        let middleRun = para.runs.first(where: { $0.text == "\t\t" })
        XCTAssertNotNil(
            middleRun,
            "Entity-encoded `&#x09;&#x09;` SHALL decode + preserve as two tabs (B-CONT P1 — "
            + "R5 finding: scanner whitespace check ran on raw bytes including `&`, `#`, `x` "
            + "which aren't Character.isWhitespace; Foundation decodes then strips → loss). "
            + "Observed runs: \(para.runs.map { "\"\($0.text)\"" })"
        )
    }

    func testWhitespaceOverlayInsRevisionWrapperDoesNotDesyncCounter() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p>
        <w:ins w:id="1" w:author="test" w:date="2026-04-27T00:00:00Z">
        <w:bookmarkStart w:id="0" w:name="insAnchor"/>
        <w:r><w:t>inserted</w:t></w:r>
        <w:bookmarkEnd w:id="0"/>
        </w:ins>
        </w:p>
        <w:p><w:r><w:t xml:space="preserve">     </w:t></w:r></w:p>
        </w:body>
        </w:document>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue59-bcont-ins-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try buildMinimalDocx(documentXML: documentXML, to: inURL)

        let doc = try DocxReader.read(from: inURL)
        var foundWhitespaceRun = false
        var observedTexts: [String] = []
        for child in doc.body.children {
            if case .paragraph(let para) = child {
                for run in para.runs {
                    observedTexts.append(run.text)
                    if run.text == "     " {
                        foundWhitespaceRun = true
                    }
                }
            }
        }
        XCTAssertTrue(
            foundWhitespaceRun,
            "5-char whitespace run after raw-captured <w:ins> wrapper SHALL survive (B-CONT P0 "
            + "root-cause-B — wrapper raw-captured because of <w:bookmarkStart> non-run child; "
            + "scanner still counts inner <w:t>, parseRun never visits). Observed text runs: \(observedTexts)"
        )
    }

    // MARK: - Sub-stack C: RunProperties typed + raw rPr children (#60)
    //
    // #60 root cause: RunProperties' `fontName: String?` collapses the
    // 4-axis `<w:rFonts w:ascii=".." w:hAnsi=".." w:eastAsia=".." w:cs="..">`
    // into a single value. parseRunProperties picks one axis (effectively
    // ascii); writer emits all 4 attributes with the same value. Round-trip
    // loses CJK/Complex Script font distinctions (e.g., DFKai-SB for traditional
    // Chinese eastAsia → emitted as `eastAsia="Times"` post-roundtrip).
    //
    // Plus: `<w:noProof/>`, `<w:kern w:val="32"/>`, `<w:lang w:val=".."/>`,
    // and w14:* effects (`<w14:textOutline>`, `<w14:textFill>`, `<w14:glow>`)
    // are silently dropped on read because no typed extraction exists. The
    // `RunProperties.rawChildren: [RawElement]?` field (sub-stack C addition)
    // captures unrecognized rPr children verbatim for byte-equivalent emission.

    /// §3.1 — `<w:rFonts>` 4-axis attributes (ascii, hAnsi, eastAsia, cs)
    /// preserved through round-trip. Pre-fix: only one axis captured into
    /// `fontName`, writer emits all 4 with same value → Chinese eastAsia
    /// font silently replaced with Latin ascii font.
    func testRFontsFourAxisPreservedThroughRoundtrip() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p>
        <w:r>
        <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="DFKai-SB" w:cs="Mangal"/>
        </w:rPr>
        <w:t>mixed-script</w:t>
        </w:r>
        </w:p>
        </w:body>
        </w:document>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue60-rfonts4axis-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try buildMinimalDocx(documentXML: documentXML, to: inURL)

        var doc = try DocxReader.read(from: inURL)
        doc.modifiedParts.insert("word/document.xml")
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue60-rfonts4axis-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(doc, to: outURL)

        let outDocXML = try Self.readDocumentXMLString(from: outURL)
        // Each axis SHALL appear in output with its source value, not collapsed.
        XCTAssertTrue(
            outDocXML.contains("w:ascii=\"Times New Roman\""),
            "rFonts w:ascii SHALL preserve source value (#60). Output:\n\(outDocXML)"
        )
        XCTAssertTrue(
            outDocXML.contains("w:hAnsi=\"Times New Roman\""),
            "rFonts w:hAnsi SHALL preserve source value (#60). Output:\n\(outDocXML)"
        )
        XCTAssertTrue(
            outDocXML.contains("w:eastAsia=\"DFKai-SB\""),
            "rFonts w:eastAsia SHALL preserve source value — pre-fix collapsed to ascii (#60). Output:\n\(outDocXML)"
        )
        XCTAssertTrue(
            outDocXML.contains("w:cs=\"Mangal\""),
            "rFonts w:cs SHALL preserve source value — pre-fix collapsed to ascii (#60). Output:\n\(outDocXML)"
        )
    }

    /// §3.2 — `<w:noProof/>` (suppress spell-check) and `<w:kern w:val="32"/>`
    /// (font kerning threshold) are typed and round-trip preserved. Pre-fix
    /// these were silently dropped on read because parseRunProperties had no
    /// extraction case for them.
    func testNoProofAndKernPreservedThroughRoundtrip() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p>
        <w:r>
        <w:rPr>
        <w:noProof/>
        <w:kern w:val="32"/>
        </w:rPr>
        <w:t>kerned-no-proof</w:t>
        </w:r>
        </w:p>
        </w:body>
        </w:document>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue60-noproof-kern-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try buildMinimalDocx(documentXML: documentXML, to: inURL)

        var doc = try DocxReader.read(from: inURL)
        doc.modifiedParts.insert("word/document.xml")
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue60-noproof-kern-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(doc, to: outURL)

        let outDocXML = try Self.readDocumentXMLString(from: outURL)
        XCTAssertTrue(
            outDocXML.contains("<w:noProof"),
            "<w:noProof/> SHALL survive round-trip (#60). Output:\n\(outDocXML)"
        )
        XCTAssertTrue(
            outDocXML.contains("w:val=\"32\"") && outDocXML.contains("<w:kern"),
            "<w:kern w:val=\"32\"/> SHALL survive round-trip (#60). Output:\n\(outDocXML)"
        )
    }

    /// §3.3 — `<w14:*>` namespace effects (Office 2010 Word DrawingML
    /// extensions like text outline, fill, glow) preserved as raw children
    /// of `<w:rPr>` for byte-equivalent emission. Pre-fix dropped silently.
    func testW14NamespaceEffectsPreservedAsRawChildren() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document \
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" \
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" \
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" \
        mc:Ignorable="w14">
        <w:body>
        <w:p>
        <w:r>
        <w:rPr>
        <w14:textOutline w14:w="9525" w14:cap="rnd" w14:cmpd="sng" w14:algn="ctr">
        <w14:solidFill><w14:srgbClr w14:val="000000"/></w14:solidFill>
        </w14:textOutline>
        </w:rPr>
        <w:t>outlined</w:t>
        </w:r>
        </w:p>
        </w:body>
        </w:document>
        """

        let inURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue60-w14-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: inURL) }
        try buildMinimalDocx(documentXML: documentXML, to: inURL)

        var doc = try DocxReader.read(from: inURL)
        doc.modifiedParts.insert("word/document.xml")
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue60-w14-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(doc, to: outURL)

        let outDocXML = try Self.readDocumentXMLString(from: outURL)
        XCTAssertTrue(
            outDocXML.contains("w14:textOutline"),
            "<w14:textOutline> SHALL survive round-trip as raw rPr child (#60). Output:\n\(outDocXML)"
        )
        XCTAssertTrue(
            outDocXML.contains("w14:srgbClr") && outDocXML.contains("w14:val=\"000000\""),
            "<w14:srgbClr w14:val=\"000000\"/> nested inside textOutline SHALL survive (#60). Output:\n\(outDocXML)"
        )
    }

    // MARK: - Helpers

    /// Build a minimal valid `.docx` with the given `document.xml` content.
    /// Mirrors the `buildMinimalDocx` helper used elsewhere in the test suite
    /// (see `Issue56R4StackTests.buildMinimalDocx` for the original).
    private func buildMinimalDocx(documentXML: String, to url: URL) throws {
        let stagingURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-60-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: stagingURL, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: stagingURL) }

        try FileManager.default.createDirectory(at: stagingURL.appendingPathComponent("_rels"), withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: stagingURL.appendingPathComponent("word/_rels"), withIntermediateDirectories: true)

        let contentTypes = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>
        """
        try contentTypes.write(to: stagingURL.appendingPathComponent("[Content_Types].xml"), atomically: true, encoding: .utf8)

        let rels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """
        try rels.write(to: stagingURL.appendingPathComponent("_rels/.rels"), atomically: true, encoding: .utf8)

        let documentRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        </Relationships>
        """
        try documentRels.write(to: stagingURL.appendingPathComponent("word/_rels/document.xml.rels"), atomically: true, encoding: .utf8)

        try documentXML.write(to: stagingURL.appendingPathComponent("word/document.xml"), atomically: true, encoding: .utf8)

        try ZipHelper.zip(stagingURL, to: url)
    }

    /// Read `word/document.xml` from a saved `.docx` as a UTF-8 string.
    /// Mirrors `DocumentXmlLosslessRoundTripTests.readDocumentXMLString`.
    static func readDocumentXMLString(from docxURL: URL) throws -> String {
        let unzipped = try ZipHelper.unzip(docxURL)
        defer { ZipHelper.cleanup(unzipped) }
        let documentURL = unzipped.appendingPathComponent("word/document.xml")
        return try String(contentsOf: documentURL, encoding: .utf8)
    }

    /// Read an arbitrary part (e.g., `word/header1.xml`) from a saved `.docx`.
    /// Used by A-CONT tests that exercise container parts.
    static func readPartXMLString(from docxURL: URL, partPath: String) throws -> String {
        let unzipped = try ZipHelper.unzip(docxURL)
        defer { ZipHelper.cleanup(unzipped) }
        let url = unzipped.appendingPathComponent(partPath)
        return try String(contentsOf: url, encoding: .utf8)
    }

    /// §2.9 helper: build a comprehensive .docx with whitespace-only `<w:t>` in
    /// ALL 6 part types — body, header1, footer1, footnotes, endnotes, comments.
    /// Used by `testWhitespacePreservedAcrossAllSixPartTypes` to exercise the
    /// full WhitespaceOverlay surface in one fixture so the convergence cycle
    /// catches symmetric-sibling regressions from the START.
    static func buildAllPartsWhitespaceFixture(to url: URL) throws {
        let stagingURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue59-allparts-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: stagingURL, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: stagingURL) }

        try FileManager.default.createDirectory(at: stagingURL.appendingPathComponent("_rels"), withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: stagingURL.appendingPathComponent("word/_rels"), withIntermediateDirectories: true)

        let contentTypes = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
        <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
        <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
        <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
        <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
        </Types>
        """
        try contentTypes.write(to: stagingURL.appendingPathComponent("[Content_Types].xml"), atomically: true, encoding: .utf8)

        let rels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """
        try rels.write(to: stagingURL.appendingPathComponent("_rels/.rels"), atomically: true, encoding: .utf8)

        let documentRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
        <Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
        <Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
        <Relationship Id="rId13" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
        <Relationship Id="rId14" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
        </Relationships>
        """
        try documentRels.write(to: stagingURL.appendingPathComponent("word/_rels/document.xml.rels"), atomically: true, encoding: .utf8)

        // body
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
        <w:p>
        <w:r><w:t>body-before</w:t></w:r>
        <w:r><w:t xml:space="preserve">     </w:t></w:r>
        <w:r><w:t>body-after</w:t></w:r>
        </w:p>
        <w:sectPr><w:headerReference r:id="rId10" w:type="default" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><w:footerReference r:id="rId11" w:type="default" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></w:sectPr>
        </w:body>
        </w:document>
        """
        try documentXML.write(to: stagingURL.appendingPathComponent("word/document.xml"), atomically: true, encoding: .utf8)

        // header1
        let headerXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:p>
        <w:r><w:t>hdr-before</w:t></w:r>
        <w:r><w:t xml:space="preserve">     </w:t></w:r>
        <w:r><w:t>hdr-after</w:t></w:r>
        </w:p>
        </w:hdr>
        """
        try headerXML.write(to: stagingURL.appendingPathComponent("word/header1.xml"), atomically: true, encoding: .utf8)

        // footer1
        let footerXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:p>
        <w:r><w:t>ftr-before</w:t></w:r>
        <w:r><w:t xml:space="preserve">     </w:t></w:r>
        <w:r><w:t>ftr-after</w:t></w:r>
        </w:p>
        </w:ftr>
        """
        try footerXML.write(to: stagingURL.appendingPathComponent("word/footer1.xml"), atomically: true, encoding: .utf8)

        // footnotes (must include separator entries with id=-1, 0)
        let footnotesXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
        <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
        <w:footnote w:id="1">
        <w:p>
        <w:r><w:t>fn-before</w:t></w:r>
        <w:r><w:t xml:space="preserve">     </w:t></w:r>
        <w:r><w:t>fn-after</w:t></w:r>
        </w:p>
        </w:footnote>
        </w:footnotes>
        """
        try footnotesXML.write(to: stagingURL.appendingPathComponent("word/footnotes.xml"), atomically: true, encoding: .utf8)

        // endnotes
        let endnotesXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>
        <w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>
        <w:endnote w:id="1">
        <w:p>
        <w:r><w:t>en-before</w:t></w:r>
        <w:r><w:t xml:space="preserve">     </w:t></w:r>
        <w:r><w:t>en-after</w:t></w:r>
        </w:p>
        </w:endnote>
        </w:endnotes>
        """
        try endnotesXML.write(to: stagingURL.appendingPathComponent("word/endnotes.xml"), atomically: true, encoding: .utf8)

        // comments
        let commentsXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:comment w:id="1" w:author="tester" w:date="2026-04-27T00:00:00Z" w:initials="t">
        <w:p>
        <w:r><w:t>cm-before</w:t></w:r>
        <w:r><w:t xml:space="preserve">     </w:t></w:r>
        <w:r><w:t>cm-after</w:t></w:r>
        </w:p>
        </w:comment>
        </w:comments>
        """
        try commentsXML.write(to: stagingURL.appendingPathComponent("word/comments.xml"), atomically: true, encoding: .utf8)

        try ZipHelper.zip(stagingURL, to: url)
    }

    /// Build a minimal valid `.docx` that includes a header part. Same shape as
    /// `Issue56R4StackTests.buildMinimalDocxWithHeader` (copied here because
    /// that method is private to the other test class).
    private func buildMinimalDocxWithHeader(documentXML: String, headerXML: String, headerRId: String, to url: URL) throws {
        let stagingURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue58-acont-hdr-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: stagingURL, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: stagingURL) }

        try FileManager.default.createDirectory(at: stagingURL.appendingPathComponent("_rels"), withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: stagingURL.appendingPathComponent("word/_rels"), withIntermediateDirectories: true)

        let contentTypes = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
        </Types>
        """
        try contentTypes.write(to: stagingURL.appendingPathComponent("[Content_Types].xml"), atomically: true, encoding: .utf8)

        let rels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """
        try rels.write(to: stagingURL.appendingPathComponent("_rels/.rels"), atomically: true, encoding: .utf8)

        let documentRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="\(headerRId)" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
        </Relationships>
        """
        try documentRels.write(to: stagingURL.appendingPathComponent("word/_rels/document.xml.rels"), atomically: true, encoding: .utf8)

        try documentXML.write(to: stagingURL.appendingPathComponent("word/document.xml"), atomically: true, encoding: .utf8)
        try headerXML.write(to: stagingURL.appendingPathComponent("word/header1.xml"), atomically: true, encoding: .utf8)

        try ZipHelper.zip(stagingURL, to: url)
    }
}
