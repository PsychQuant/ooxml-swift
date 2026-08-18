import XCTest
@testable import OOXMLSwift

/// Body-level `<w:sectPr>` round-trip (PsychQuant/ooxml-swift #84).
///
/// **Pre-fix bug**: `DocxReader` never assigned `WordDocument.sectionProperties`
/// — a `grep` for any assignment in the reader returned nothing, while a code
/// comment at the `case "sectPr"` skip claimed it was "parsed separately".
/// `DocxWriter` unconditionally emits `document.sectionProperties.toXML()`, so
/// every save wrote a **default-constructed** section: US Letter page size,
/// generic 1440 margins, no header/footer references.
///
/// The effect is not "an attribute goes missing" but "the section block is
/// replaced with another document's". Measured on a real A4 form after a single
/// `updateCell`:
///
///   pgSz   11906x16838 (A4) -> 12240x15840 (US Letter)
///   pgMar  1276/1077        -> 1440 all round
///   footerReference rId7    -> dropped (footer1.xml stays in the package,
///                              unreferenced, so the footer stops rendering)
///
/// Third instance of the same reader-gap family as #99 (`tcBorders`) and
/// #101 (`tcMar`). Downstream: PsychQuant/macdoc#142,
/// PsychQuant/ntu-claude-plugins#16.
final class SectPrRoundTripTests: XCTestCase {

    private var tempDir: URL!

    override func setUp() {
        super.setUp()
        tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("SectPrRoundTripTests-\(UUID().uuidString)")
        try? FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
    }

    override func tearDown() {
        try? FileManager.default.removeItem(at: tempDir)
        super.tearDown()
    }

    private func roundTrip(_ document: WordDocument) throws -> WordDocument {
        let url = tempDir.appendingPathComponent("t-\(UUID().uuidString).docx")
        try DocxWriter.write(document, to: url)
        return try DocxReader.read(from: url)
    }

    /// A4 must not silently become US Letter.
    func testPageSizeSurvivesRoundTrip() throws {
        var doc = WordDocument()
        doc.sectionProperties.pageSize = PageSize(width: 11906, height: 16838)

        let result = try roundTrip(doc)

        XCTAssertEqual(result.sectionProperties.pageSize.width, 11906, "A4 width lost")
        XCTAssertEqual(result.sectionProperties.pageSize.height, 16838, "A4 height lost")
    }

    func testPageMarginsSurviveRoundTrip() throws {
        var doc = WordDocument()
        var m = doc.sectionProperties.pageMargins
        m.top = 1276; m.right = 1077; m.bottom = 1276; m.left = 1077
        m.header = 720; m.footer = 680
        doc.sectionProperties.pageMargins = m

        let r = try roundTrip(doc).sectionProperties.pageMargins

        XCTAssertEqual(r.top, 1276); XCTAssertEqual(r.right, 1077)
        XCTAssertEqual(r.bottom, 1276); XCTAssertEqual(r.left, 1077)
        XCTAssertEqual(r.footer, 680, "custom footer distance lost")
    }

    /// The footer part stays in the package either way; losing the reference is
    /// what stops it rendering.
    func testFooterReferenceSurvivesRoundTrip() throws {
        var doc = WordDocument()
        doc.sectionProperties.footerReference = "rId7"

        let r = try roundTrip(doc).sectionProperties
        XCTAssertTrue(r.footerReference == "rId7" || r.footerReferences.defaultRef == "rId7",
                      "footerReference dropped")
    }

    func testHeaderReferenceSurvivesRoundTrip() throws {
        var doc = WordDocument()
        doc.sectionProperties.headerReference = "rId5"

        let r = try roundTrip(doc).sectionProperties
        XCTAssertTrue(r.headerReference == "rId5" || r.headerReferences.defaultRef == "rId5",
                      "headerReference dropped")
    }

    func testDocGridSurvivesRoundTrip() throws {
        var doc = WordDocument()
        doc.sectionProperties.docGrid = DocumentGrid(linePitch: 571)

        let r = try roundTrip(doc).sectionProperties
        XCTAssertEqual(r.docGrid?.linePitch, 571, "CJK line pitch reset to the 360 default")
    }

    func testLandscapeOrientationSurvivesRoundTrip() throws {
        var doc = WordDocument()
        doc.sectionProperties.orientation = .landscape
        doc.sectionProperties.pageSize = PageSize(width: 16838, height: 11906)

        let r = try roundTrip(doc).sectionProperties
        XCTAssertEqual(r.orientation, .landscape, "orientation lost")
    }

    func testColumnsSurviveRoundTrip() throws {
        var doc = WordDocument()
        doc.sectionProperties.columns = 3

        XCTAssertEqual(try roundTrip(doc).sectionProperties.columns, 3, "column count lost")
    }

    func testTitlePageDistinctSurvivesRoundTrip() throws {
        var doc = WordDocument()
        doc.sectionProperties.titlePageDistinct = true

        XCTAssertTrue(try roundTrip(doc).sectionProperties.titlePageDistinct, "titlePg lost")
    }

    /// Absence stays absence: a document with no docGrid override must not come
    /// back claiming one the source never had beyond the writer's own default.
    func testDefaultsRoundTripUnchanged() throws {
        let doc = WordDocument()
        let r = try roundTrip(doc).sectionProperties

        XCTAssertEqual(r.pageSize.width, doc.sectionProperties.pageSize.width)
        XCTAssertEqual(r.pageMargins.top, doc.sectionProperties.pageMargins.top)
        XCTAssertNil(r.headerReference)
        XCTAssertNil(r.footerReference)
    }
}

extension SectPrRoundTripTests {
    /// `<w:docGrid w:type="lines">` turns the CJK line grid on. Absent means
    /// "default" (no grid), so dropping it silently changes line layout for
    /// every Chinese/Japanese document.
    func testDocGridTypeSurvivesRoundTrip() throws {
        var doc = WordDocument()
        doc.sectionProperties.docGrid = DocumentGrid(linePitch: 571, type: "lines")

        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("dg-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try DocxWriter.write(doc, to: url)
        let r = try DocxReader.read(from: url).sectionProperties

        XCTAssertEqual(r.docGrid?.type, "lines", "docGrid w:type dropped — CJK line grid silently off")
        XCTAssertEqual(r.docGrid?.linePitch, 571)
    }

    /// A source without `w:type` must not gain one.
    func testDocGridWithoutTypeStaysWithoutType() throws {
        var doc = WordDocument()
        doc.sectionProperties.docGrid = DocumentGrid(linePitch: 400)

        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("dg2-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try DocxWriter.write(doc, to: url)
        let r = try DocxReader.read(from: url).sectionProperties

        XCTAssertNil(r.docGrid?.type, "invented a grid mode the source never had")
        XCTAssertEqual(r.docGrid?.linePitch, 400)
    }
}
