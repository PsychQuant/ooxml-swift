import XCTest
@testable import OOXMLSwift

/// Test scaffold for `SDTParser` (`packages/ooxml-swift/Sources/OOXMLSwift/IO/SDTParser.swift`).
///
/// Specs covered by this suite (see `openspec/changes/che-word-mcp-content-controls-read-write/specs/ooxml-read-back-parsers/spec.md`):
/// - DocxReader parses w:sdt into structured ContentControl values
/// - SDTParser distinguishes all 12 SDT types
/// - SDTParser handles nested SDTs by preserving tree structure
/// - SDTParser handles block-level SDTs wrapping paragraphs and tables
/// - SDT round-trip preserves byte-level content fidelity
///
/// Fixture generation: see `Fixtures/SDTFixtureBuilder.swift`.
/// Implementation tasks 3.1–3.5 will populate these tests.
final class SDTParserTests: XCTestCase {

    // MARK: - Fixture loading

    /// Loads the SDT fixture into a freshly parsed `WordDocument`.
    /// Returns `nil` if the reader has not yet been extended to parse SDTs
    /// (fixture builder depends only on the writer path, so this helper is
    /// always buildable even before task 3.1 lands).
    private func loadFixtureDocument() throws -> WordDocument {
        let fixtureData = try SDTFixtureBuilder.build()
        let tempURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("sdt-fixture-\(UUID().uuidString).docx")
        try fixtureData.write(to: tempURL)
        defer { try? FileManager.default.removeItem(at: tempURL) }
        return try DocxReader.read(from: tempURL)
    }

    // MARK: - Task 3.1: DocxReader surfaces ContentControls

    /// See: `DocxReader parses w:sdt into structured ContentControl values`.
    func testReaderSurfacesContentControlsAfterTask31() throws {
        let doc = try loadFixtureDocument()
        let allControls = doc.getParagraphs().flatMap { $0.contentControls }
        XCTAssertGreaterThanOrEqual(allControls.count, 10,
            "expected 10+ paragraph-level SDTs from fixture, got \(allControls.count)")

        // Spec scenario: plain-text SDT round-trips with tag/alias/type/content.
        guard let clientName = allControls.first(where: { $0.sdt.tag == "client_name" }) else {
            XCTFail("client_name SDT not surfaced as structured ContentControl")
            return
        }
        XCTAssertEqual(clientName.sdt.tag, "client_name")
        XCTAssertEqual(clientName.sdt.alias, "Client Name")
        XCTAssertEqual(clientName.sdt.type, .plainText)
        XCTAssertTrue(clientName.content.contains("ACME Corp"),
            "plain-text SDT content lost: '\(clientName.content)'")

        // Spec scenario: paragraph runs do not contain SDT XML as rawXML.
        for paragraph in doc.getParagraphs() where !paragraph.contentControls.isEmpty {
            for run in paragraph.runs {
                let raw = run.rawXML ?? ""
                XCTAssertFalse(raw.contains("<w:sdt"),
                    "SDT leaked into Run.rawXML — should be on Paragraph.contentControls")
            }
        }
    }

    // MARK: - Task 3.2: 12-type discrimination

    /// See: `SDTParser distinguishes all 12 SDT types`.
    func testAllElevenFixtureSDTsHaveCorrectTypesAfterTask32() throws {
        let doc = try loadFixtureDocument()
        let byTag: [String: ContentControl] = Dictionary(
            uniqueKeysWithValues: doc.getParagraphs()
                .flatMap { $0.contentControls }
                .compactMap { control in control.sdt.tag.map { ($0, control) } }
        )

        let expectations: [(tag: String, type: SDTType)] = [
            ("intro", .richText),
            ("client_name", .plainText),
            ("logo", .picture),
            ("issue_date", .date),
            ("priority", .dropDownList),
            ("category", .comboBox),
            ("acceptance", .checkbox),
            ("references", .bibliography),
            ("cite_a", .citation),
            ("address_block", .group),
            ("line_items", .repeatingSection),
        ]
        for (tag, expectedType) in expectations {
            guard let control = byTag[tag] else {
                XCTFail("SDT tag='\(tag)' not surfaced from fixture")
                continue
            }
            XCTAssertEqual(control.sdt.type, expectedType,
                "type mismatch for tag='\(tag)'")
        }
    }

    // MARK: - Task 3.3: nested SDT tree

    /// See: `SDTParser handles nested SDTs by preserving tree structure`.
    func testGroupSDTHasNestedPlainTextChildAfterTask33() throws {
        let doc = try loadFixtureDocument()
        let groups = doc.getParagraphs()
            .flatMap { $0.contentControls }
            .filter { $0.sdt.type == .group }
        guard let group = groups.first(where: { $0.sdt.tag == "address_block" }) else {
            XCTFail("address_block group SDT not found")
            return
        }
        XCTAssertEqual(group.children.count, 1,
            "expected 1 nested plainText child inside address_block group")
        guard let city = group.children.first else { return }
        XCTAssertEqual(city.sdt.tag, "city")
        XCTAssertEqual(city.sdt.type, .plainText)
        XCTAssertEqual(city.parentSdtId, group.sdt.id,
            "nested child must reference outer SDT id as parentSdtId")
    }

    // MARK: - Task 3.4: block-level SDT

    /// See: `SDTParser handles block-level SDTs wrapping paragraphs and tables`.
    /// Spec scenario: <w:body><w:sdt>...<w:sdtContent><w:p>A</w:p><w:p>B</w:p></w:sdtContent></w:sdt></w:body>
    func testBlockLevelSDTPreservesChildrenAfterTask34() throws {
        // Build a doc with a block-level SDT wrapping two paragraphs by
        // direct body manipulation (writer round-trip exercises both reader
        // and writer paths).
        var doc = WordDocument()
        let blockSdt = StructuredDocumentTag(
            id: 20001,
            tag: "block_wrapper",
            alias: "Block Wrapper",
            type: .richText
        )
        let paraA = Paragraph(text: "A")
        let paraB = Paragraph(text: "B")
        let control = ContentControl(sdt: blockSdt, content: "")
        doc.body.children = [
            .contentControl(control, children: [.paragraph(paraA), .paragraph(paraB)])
        ]

        let bytes = try DocxWriter.writeData(doc)
        let tempURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("block-sdt-\(UUID().uuidString).docx")
        try bytes.write(to: tempURL)
        defer { try? FileManager.default.removeItem(at: tempURL) }

        let parsed = try DocxReader.read(from: tempURL)
        XCTAssertEqual(parsed.body.children.count, 1, "expected 1 block-level SDT in body")
        guard case .contentControl(let parsedControl, let parsedChildren) = parsed.body.children[0] else {
            XCTFail("expected BodyChild.contentControl, got \(parsed.body.children[0])")
            return
        }
        XCTAssertEqual(parsedControl.sdt.tag, "block_wrapper")
        XCTAssertEqual(parsedControl.sdt.alias, "Block Wrapper")
        XCTAssertEqual(parsedChildren.count, 2)

        var texts: [String] = []
        for child in parsedChildren {
            if case .paragraph(let p) = child { texts.append(p.getText()) }
        }
        XCTAssertEqual(texts, ["A", "B"], "block-level SDT children paragraphs lost")
    }

    // MARK: - Task 3.5: round-trip fidelity

    /// See: `SDT round-trip preserves byte-level content fidelity`.
    /// Reads the fixture, writes it back unchanged, then verifies the
    /// re-read matches the original on id/tag/alias/type for every SDT.
    func testFixtureRoundTripPreservesSDTMetadataAfterTask35() throws {
        let original = try loadFixtureDocument()
        let originalControls = original.getParagraphs().flatMap { $0.contentControls }
        XCTAssertGreaterThanOrEqual(originalControls.count, 10,
            "fixture must surface 10+ SDTs to be a meaningful round-trip test")

        // Write the parsed document back to disk, then re-read.
        let bytes = try DocxWriter.writeData(original)
        let tempURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("sdt-roundtrip-\(UUID().uuidString).docx")
        try bytes.write(to: tempURL)
        defer { try? FileManager.default.removeItem(at: tempURL) }

        let reread = try DocxReader.read(from: tempURL)
        let rereadControls = reread.getParagraphs().flatMap { $0.contentControls }
        XCTAssertEqual(rereadControls.count, originalControls.count,
            "SDT count drift across round-trip: \(originalControls.count) → \(rereadControls.count)")

        // Compare metadata (id/tag/alias/type/lockType) by tag.
        let originalByTag: [String: ContentControl] = Dictionary(
            uniqueKeysWithValues: originalControls.compactMap { c in c.sdt.tag.map { ($0, c) } }
        )
        let rereadByTag: [String: ContentControl] = Dictionary(
            uniqueKeysWithValues: rereadControls.compactMap { c in c.sdt.tag.map { ($0, c) } }
        )
        for tag in originalByTag.keys {
            guard let o = originalByTag[tag], let r = rereadByTag[tag] else {
                XCTFail("tag '\(tag)' lost across round-trip")
                continue
            }
            XCTAssertEqual(o.sdt.id, r.sdt.id, "id drift for tag='\(tag)'")
            XCTAssertEqual(o.sdt.alias, r.sdt.alias, "alias drift for tag='\(tag)'")
            XCTAssertEqual(o.sdt.type, r.sdt.type, "type drift for tag='\(tag)'")
            XCTAssertEqual(o.sdt.lockType, r.sdt.lockType, "lockType drift for tag='\(tag)'")
        }
    }

    // MARK: - Pre-existing sanity: fixture builds without crashing

    /// Confirms the fixture builder produces non-empty .docx bytes.
    /// Independent of SDT parser implementation — verifies task 1.1 only.
    func testFixtureBuilderProducesNonEmptyDocxData() throws {
        let data = try SDTFixtureBuilder.build()
        XCTAssertGreaterThan(data.count, 0, "fixture builder produced empty .docx")
        // .docx files start with PK zip magic bytes.
        XCTAssertEqual(data.prefix(2), Data([0x50, 0x4B]), "fixture is not a valid ZIP archive")
    }

    // MARK: - Helper: ContentControl equality assertion

    /// Compares two parsed ContentControls for metadata equality.
    /// Used by tasks 3.1+ once the reader populates ContentControl structs.
    func assertContentControlMetadata(
        _ actual: ContentControl,
        tag: String? = nil,
        alias: String? = nil,
        type: SDTType? = nil,
        file: StaticString = #filePath,
        line: UInt = #line
    ) {
        if let tag = tag {
            XCTAssertEqual(actual.sdt.tag, tag, "tag mismatch", file: file, line: line)
        }
        if let alias = alias {
            XCTAssertEqual(actual.sdt.alias, alias, "alias mismatch", file: file, line: line)
        }
        if let type = type {
            XCTAssertEqual(actual.sdt.type, type, "type mismatch", file: file, line: line)
        }
    }
}
