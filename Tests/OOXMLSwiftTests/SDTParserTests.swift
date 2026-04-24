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

    // MARK: - Task 3.4 placeholder

    /// See: `SDTParser handles block-level SDTs wrapping paragraphs and tables`.
    /// Enabled by task 3.4.
    func testBlockLevelSDTPreservesChildrenAfterTask34() throws {
        throw XCTSkip("pending task 3.4: block-level SDT")
    }

    // MARK: - Task 3.5 placeholder

    /// See: `SDT round-trip preserves byte-level content fidelity`.
    /// Enabled by task 3.5.
    func testFixtureRoundTripPreservesSDTMetadataAfterTask35() throws {
        throw XCTSkip("pending task 3.5: round-trip fidelity")
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
