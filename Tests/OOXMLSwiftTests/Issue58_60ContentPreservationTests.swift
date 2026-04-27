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
        // Lands with §2.7 (sub-stack B). Until then, this assertion is documented
        // but skipped to keep the matrix-pin green for sub-stack A.

        // Preservation class 3 of 3 (#60): <w:rFonts>/<w:noProof>/<w:lang>/<w:kern>/w14 counts.
        // Lands with §3.9 (sub-stack C). Until then, this assertion is documented
        // but skipped to keep the matrix-pin green for sub-stack A.
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
