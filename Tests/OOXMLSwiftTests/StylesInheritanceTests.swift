import XCTest
@testable import OOXMLSwift

/// Tests for Phase 3 Style inheritance + linking + latentStyles + aliases on
/// WordDocument. Specs covered (see
/// `openspec/changes/che-word-mcp-styles-sections-numbering-foundations/specs/`):
/// - ooxml-document-part-mutations: WordDocument exposes style inheritance traversal
/// - ooxml-document-part-mutations: WordDocument exposes style linking and naming
/// - ooxml-document-part-mutations: WordDocument exposes latentStyles management
///
/// Implementation tasks 3.1-3.4 will populate these tests; until then they
/// XCTSkip so the suite stays green.
final class StylesInheritanceTests: XCTestCase {

    // MARK: - Test fixture helpers

    /// Builds a minimal WordDocument with a chain of styles for inheritance
    /// traversal tests. Each style entry is `(id, name, basedOnId)` where
    /// basedOnId == nil for root styles.
    func makeDocWithStyleChain(_ chain: [(id: String, name: String, basedOn: String?)]) -> WordDocument {
        var doc = WordDocument()
        for entry in chain {
            var style = Style(id: entry.id, name: entry.name, type: .paragraph)
            style.basedOn = entry.basedOn
            doc.styles.append(style)
        }
        return doc
    }

    /// Round-trip a doc through writer + reader so we exercise serialization.
    func roundTrip(_ doc: WordDocument) throws -> WordDocument {
        let bytes = try DocxWriter.writeData(doc)
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("styles-rt-\(UUID().uuidString).docx")
        try bytes.write(to: url)
        defer { try? FileManager.default.removeItem(at: url) }
        return try DocxReader.read(from: url)
    }

    // MARK: - Task 3.1: getStyleInheritanceChain

    /// See: `WordDocument exposes style inheritance traversal`. Spec scenario:
    /// Linear inheritance chain.
    func testInheritanceChainReturnsAncestorsInOrderAfterTask31() throws {
        let doc = makeDocWithStyleChain([
            ("Custom_Normal", "Custom Normal", nil),
            ("Custom_Heading1", "Custom Heading 1", "Custom_Normal"),
            ("Custom_Heading1Bold", "Custom Heading 1 Bold", "Custom_Heading1"),
        ])
        let chain = doc.getStyleInheritanceChain(styleId: "Custom_Heading1Bold")
        XCTAssertEqual(chain.map { $0.id }, ["Custom_Heading1Bold", "Custom_Heading1", "Custom_Normal"])
    }

    /// Spec scenario: Missing style returns empty.
    func testInheritanceChainEmptyForMissingStyleAfterTask31() throws {
        let doc = makeDocWithStyleChain([("X", "X", nil)])
        XCTAssertEqual(doc.getStyleInheritanceChain(styleId: "Nonexistent"), [])
    }

    /// Cycle detection: A→B→A should stop at the revisit, not loop.
    func testInheritanceChainHandlesCycleAfterTask31() throws {
        var doc = WordDocument()
        var a = Style(id: "Cycle_A", name: "A", type: .paragraph)
        a.basedOn = "Cycle_B"
        var b = Style(id: "Cycle_B", name: "B", type: .paragraph)
        b.basedOn = "Cycle_A"
        doc.styles.append(a)
        doc.styles.append(b)
        let chain = doc.getStyleInheritanceChain(styleId: "Cycle_A")
        XCTAssertEqual(chain.map { $0.id }, ["Cycle_A", "Cycle_B"], "should stop at first revisit")
    }

    // MARK: - Task 3.2 + 3.3: linkStyles + addStyleNameAlias

    /// Spec scenario: Link paragraph and character styles.
    func testLinkStylesEmitsBidirectionalLinkAfterTask32() throws {
        var doc = WordDocument()
        doc.styles.append(Style(id: "MyHeading1", name: "MyHeading1", type: .paragraph))
        doc.styles.append(Style(id: "MyHeading1Char", name: "MyHeading1Char", type: .character))
        try doc.linkStyles(paragraphStyleId: "MyHeading1", characterStyleId: "MyHeading1Char")
        XCTAssertEqual(doc.styles.first(where: { $0.id == "MyHeading1" })?.linkedStyleId, "MyHeading1Char")
        XCTAssertEqual(doc.styles.first(where: { $0.id == "MyHeading1Char" })?.linkedStyleId, "MyHeading1")
    }

    func testLinkStylesThrowsOnTypeMismatchAfterTask32() throws {
        var doc = WordDocument()
        doc.styles.append(Style(id: "ParaA", name: "Para A", type: .paragraph))
        doc.styles.append(Style(id: "ParaB", name: "Para B", type: .paragraph))
        XCTAssertThrowsError(try doc.linkStyles(paragraphStyleId: "ParaA", characterStyleId: "ParaB")) { error in
            guard case WordError.typeMismatch = error else { XCTFail("expected typeMismatch"); return }
        }
    }

    /// Spec: addStyleNameAlias replaces same lang.
    func testAddStyleNameAliasReplacesSameLangAfterTask33() throws {
        var doc = WordDocument()
        doc.styles.append(Style(id: "Heading1", name: "Heading 1", type: .paragraph))
        try doc.addStyleNameAlias(styleId: "Heading1", lang: "de-DE", name: "Überschrift 1")
        try doc.addStyleNameAlias(styleId: "Heading1", lang: "de-DE", name: "Überschrift Eins")
        let aliases = doc.styles.first(where: { $0.id == "Heading1" })?.aliases ?? []
        XCTAssertEqual(aliases.count, 1, "same lang should replace, not duplicate")
        XCTAssertEqual(aliases.first?.name, "Überschrift Eins")
    }

    // MARK: - Task 3.4: setLatentStyles

    /// Spec scenario: Set latentStyles persists in styles.xml.
    func testSetLatentStylesPersistsAcrossRoundTripAfterTask34() throws {
        var doc = WordDocument()
        doc.setLatentStyles([
            LatentStyle(name: "Heading 9", uiPriority: 9, semiHidden: true, unhideWhenUsed: false, qFormat: false)
        ])
        let reread = try roundTrip(doc)
        XCTAssertEqual(reread.latentStyles.count, 1)
        XCTAssertEqual(reread.latentStyles.first?.name, "Heading 9")
        XCTAssertEqual(reread.latentStyles.first?.uiPriority, 9)
        XCTAssertTrue(reread.latentStyles.first?.semiHidden ?? false)
    }

    // MARK: - Pre-existing sanity

    /// Sanity check that the test scaffold itself compiles and can build a doc.
    func testFixtureBuilderProducesValidDocument() {
        let baseline = WordDocument().styles.count
        let doc = makeDocWithStyleChain([
            ("CustomA", "Custom A", nil),
            ("CustomB", "Custom B", "CustomA"),
            ("CustomC", "Custom C", "CustomB"),
        ])
        XCTAssertEqual(doc.styles.count, baseline + 3,
            "fixture should append 3 styles on top of WordDocument default styles")
    }
}
