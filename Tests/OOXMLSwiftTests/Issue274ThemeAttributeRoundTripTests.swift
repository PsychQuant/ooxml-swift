import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Regression tests for [PsychQuant/che-word-mcp#274](https://github.com/PsychQuant/che-word-mcp/issues/274).
///
/// Before this fix: `DocxReader.parseRunProperties` read only `w:val` from
/// `<w:color>` and only the five named axes from `<w:rFonts>`. Both elements
/// are in the recognized rPr set, so they were not kept as raw children
/// either. As soon as any edit made the writer regenerate a part from the
/// typed model, every `w:themeColor` / `w:themeTint` / `w:themeShade` and
/// every `w:asciiTheme` / `w:hAnsiTheme` / `w:eastAsiaTheme` / `w:cstheme`
/// in that part disappeared — on untouched runs, paragraph-mark rPr, the
/// prior rPr inside `<w:rPrChange>`, and style rPr alike. An `<w:rFonts>`
/// whose only attribute was a theme attribute vanished entirely.
///
/// After this fix: the theme attributes are typed fields on `RunProperties`
/// / `RFontsProperties`, parsed and re-emitted, and setting an explicit
/// colour or named font clears the theme attribute that would override it
/// (Word gives the theme attribute precedence). A Word-authored
/// `<w:rPrChange>` is still dropped as a whole on typed re-serialization —
/// che-word-mcp#279 — so its prior rPr is not restored by this fix.
final class Issue274ThemeAttributeRoundTripTests: XCTestCase {

    // MARK: - Section A: run-level parse/emit (no docx I/O)

    private func parseRunFragment(_ xml: String) throws -> Run {
        let wrapper = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          \(xml)
        </w:document>
        """
        let doc = try XMLDocument(xmlString: wrapper)
        guard let root = doc.rootElement(),
              let runElement = root.elements(forName: "w:r").first else {
            throw NSError(domain: "test", code: 0, userInfo: [NSLocalizedDescriptionKey: "no <w:r> found"])
        }
        return try DocxReader.parseRun(from: runElement, relationships: RelationshipsCollection())
    }

    func testColorThemeAttributesSurviveParseAndEmit() throws {
        let run = try parseRunFragment(
            "<w:r><w:rPr><w:color w:val=\"000000\" w:themeColor=\"text1\" w:themeTint=\"99\" w:themeShade=\"BF\"/></w:rPr><w:t>a</w:t></w:r>")
        let xml = run.toXML()
        XCTAssertTrue(xml.contains("w:val=\"000000\""), "got: \(xml)")
        XCTAssertTrue(xml.contains("w:themeColor=\"text1\""), "got: \(xml)")
        XCTAssertTrue(xml.contains("w:themeTint=\"99\""), "got: \(xml)")
        XCTAssertTrue(xml.contains("w:themeShade=\"BF\""), "got: \(xml)")
    }

    func testRFontsThemeAttributesSurviveParseAndEmit() throws {
        let run = try parseRunFragment(
            "<w:r><w:rPr><w:rFonts w:ascii=\"Times\" w:asciiTheme=\"minorHAnsi\" w:hAnsiTheme=\"minorHAnsi\" w:eastAsiaTheme=\"minorEastAsia\" w:cstheme=\"minorBidi\"/></w:rPr><w:t>a</w:t></w:r>")
        let xml = run.toXML()
        XCTAssertTrue(xml.contains("w:ascii=\"Times\""), "got: \(xml)")
        XCTAssertTrue(xml.contains("w:asciiTheme=\"minorHAnsi\""), "got: \(xml)")
        XCTAssertTrue(xml.contains("w:hAnsiTheme=\"minorHAnsi\""), "got: \(xml)")
        XCTAssertTrue(xml.contains("w:eastAsiaTheme=\"minorEastAsia\""), "got: \(xml)")
        XCTAssertTrue(xml.contains("w:cstheme=\"minorBidi\""), "got: \(xml)")
    }

    /// The issue's exact shape: an `<w:rFonts>` whose only attribute is a
    /// theme attribute was skipped entirely by the writer.
    func testRFontsCarryingOnlyAThemeAttributeIsNotDropped() throws {
        let run = try parseRunFragment(
            "<w:r><w:rPr><w:rFonts w:eastAsiaTheme=\"minorEastAsia\"/></w:rPr><w:t>a</w:t></w:r>")
        XCTAssertTrue(run.toXML().contains("<w:rFonts w:eastAsiaTheme=\"minorEastAsia\"/>"), "got: \(run.toXML())")
    }

    /// Files without theme attributes must serialize exactly as before.
    func testOutputWithoutThemeAttributesIsUnchanged() throws {
        let run = try parseRunFragment(
            "<w:r><w:rPr><w:rFonts w:ascii=\"A\" w:hAnsi=\"B\" w:eastAsia=\"C\" w:cs=\"D\" w:hint=\"eastAsia\"/><w:color w:val=\"FF0000\"/></w:rPr><w:t>a</w:t></w:r>")
        let xml = run.toXML()
        XCTAssertTrue(xml.contains("<w:rFonts w:ascii=\"A\" w:hAnsi=\"B\" w:eastAsia=\"C\" w:cs=\"D\" w:hint=\"eastAsia\"/>"), "got: \(xml)")
        XCTAssertTrue(xml.contains("<w:color w:val=\"FF0000\"/>"), "got: \(xml)")
    }

    // MARK: - Section A2: merge semantics

    /// Word gives `w:themeColor` precedence over `w:val`. A patch that sets
    /// an explicit colour must therefore drop the stale theme colour, or the
    /// new colour silently has no effect in Word.
    func testMergingAnExplicitColorClearsStaleThemeColor() {
        var base = RunProperties()
        base.color = "000000"
        base.themeColor = "text1"
        base.themeTint = "99"
        base.themeShade = "BF"

        base.merge(with: RunProperties(color: "FF0000"))

        XCTAssertEqual(base.color, "FF0000")
        XCTAssertNil(base.themeColor)
        XCTAssertNil(base.themeTint)
        XCTAssertNil(base.themeShade)
        XCTAssertEqual(base.toXML(), "<w:color w:val=\"FF0000\"/>")
    }

    func testMergingAThemeColorSetsIt() {
        var base = RunProperties()
        base.color = "000000"

        var patch = RunProperties()
        patch.color = "4472C4"
        patch.themeColor = "accent1"
        base.merge(with: patch)

        XCTAssertEqual(base.color, "4472C4")
        XCTAssertEqual(base.themeColor, "accent1")
    }

    /// A patch that does not touch colour leaves the theme colour alone.
    func testMergingAnUnrelatedPropertyKeepsThemeColor() {
        var base = RunProperties()
        base.color = "000000"
        base.themeColor = "text1"

        base.merge(with: RunProperties(bold: true))

        XCTAssertEqual(base.themeColor, "text1")
    }

    // MARK: - Section B: style rPr (styles.xml shares parseRunProperties)

    func testStyleRunPropertiesKeepThemeAttributes() throws {
        let xml = try XMLDocument(xmlString: """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\
            <w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/>\
            <w:rPr><w:rFonts w:eastAsiaTheme="minorEastAsia"/><w:color w:val="000000" w:themeColor="text1"/></w:rPr>\
            </w:style></w:styles>
            """)
        let style = try XCTUnwrap(DocxReader.parseStyles(from: xml).first { $0.id == "Body" })
        let emitted = style.toXML()
        XCTAssertTrue(emitted.contains("w:eastAsiaTheme=\"minorEastAsia\""), "got: \(emitted)")
        XCTAssertTrue(emitted.contains("w:themeColor=\"text1\""), "got: \(emitted)")
    }

    // MARK: - Section C: full-docx — edit one paragraph, the untouched one keeps its theme attributes

    private static let themedRPr =
        "<w:rFonts w:eastAsiaTheme=\"minorEastAsia\"/><w:color w:val=\"000000\" w:themeColor=\"text1\"/>"

    /// Paragraph 0 carries theme attributes in two places: the paragraph
    /// mark rPr and the run rPr. Paragraph 1 ("gamma") is the one the test
    /// edits. (The prior rPr inside `<w:rPrChange>` goes through the same
    /// `parseRunProperties`, but a source `<w:rPrChange>` is itself dropped
    /// on typed re-serialization — che-word-mcp#279.)
    private static let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
        <w:p><w:pPr><w:rPr>\(themedRPr)</w:rPr></w:pPr>\
        <w:r><w:rPr>\(themedRPr)</w:rPr><w:t>alpha</w:t></w:r></w:p>\
        <w:p><w:r><w:t>gamma</w:t></w:r></w:p>\
        <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>\
        </w:body></w:document>
        """

    private func buildFixture(documentXML: String = Issue274ThemeAttributeRoundTripTests.documentXML) throws -> URL {
        let staging = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue274-staging-\(UUID().uuidString)")
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
            .appendingPathComponent("issue274-\(UUID().uuidString).docx")
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
        return docxURL
    }

    private func savedDocumentXML(_ doc: WordDocument) throws -> String {
        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue274-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(doc, to: outURL)
        let archive = try Archive(url: outURL, accessMode: .read)
        let entry = try XCTUnwrap(archive["word/document.xml"])
        var data = Data()
        _ = try archive.extract(entry) { data.append($0) }
        return String(decoding: data, as: UTF8.self)
    }

    /// Returns the `<w:p>…</w:p>` substring that contains `marker`.
    private func paragraph(containing marker: String, in xml: String) throws -> String {
        let pattern = try NSRegularExpression(pattern: "<w:p[ >].*?</w:p>", options: [.dotMatchesLineSeparators])
        let ns = xml as NSString
        for match in pattern.matches(in: xml, range: NSRange(location: 0, length: ns.length)) {
            let p = ns.substring(with: match.range)
            if p.contains(marker) { return p }
        }
        XCTFail("no paragraph containing \(marker)")
        return ""
    }

    /// `appendParagraph` can take the op-log path, whose `RunPayload` knows
    /// three theme fonts but no colour theme and no `w:cstheme`. A run that
    /// carries either must not be squeezed through it and lose them.
    func testAppendedParagraphKeepsEveryThemeAttribute() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        var props = RunProperties()
        props.color = "000000"
        props.themeColor = "text1"
        props.rFonts = RFontsProperties(asciiTheme: "minorHAnsi", hAnsiTheme: "minorHAnsi",
                                        eastAsiaTheme: "minorEastAsia", csTheme: "minorBidi")
        doc.appendParagraph(Paragraph(runs: [Run(text: "delta", properties: props)]))

        let appended = try paragraph(containing: "delta", in: try savedDocumentXML(doc))
        for attribute in ["w:themeColor=\"text1\"", "w:asciiTheme=\"minorHAnsi\"", "w:hAnsiTheme=\"minorHAnsi\"",
                          "w:eastAsiaTheme=\"minorEastAsia\"", "w:cstheme=\"minorBidi\""] {
            XCTAssertTrue(appended.contains(attribute), "\(attribute) missing — got: \(appended)")
        }
    }

    /// Theme fonts alone are representable on the op path; they must be
    /// carried through it rather than dropped.
    func testAppendedParagraphKeepsThemeFontsOnTheOpPath() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        var props = RunProperties()
        props.rFonts = RFontsProperties(asciiTheme: "majorHAnsi", hAnsiTheme: "majorHAnsi",
                                        eastAsiaTheme: "minorEastAsia")
        doc.appendParagraph(Paragraph(runs: [Run(text: "epsilon", properties: props)]))

        let appended = try paragraph(containing: "epsilon", in: try savedDocumentXML(doc))
        for attribute in ["w:asciiTheme=\"majorHAnsi\"", "w:hAnsiTheme=\"majorHAnsi\"",
                          "w:eastAsiaTheme=\"minorEastAsia\""] {
            XCTAssertTrue(appended.contains(attribute), "\(attribute) missing — got: \(appended)")
        }
    }

    func testEditingAnotherParagraphKeepsThemeAttributesOnUntouchedParagraph() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        try doc.updateParagraph(at: 1, text: "GAMMA")
        XCTAssertTrue(doc.modifiedPartsView.contains("word/document.xml"))

        let untouched = try paragraph(containing: "alpha", in: try savedDocumentXML(doc))
        // Two places: paragraph-mark rPr and run rPr.
        XCTAssertEqual(untouched.components(separatedBy: "w:themeColor=\"text1\"").count - 1, 2,
                       "every w:themeColor must survive — got: \(untouched)")
        XCTAssertEqual(untouched.components(separatedBy: "<w:rFonts w:eastAsiaTheme=\"minorEastAsia\"/>").count - 1, 2,
                       "every theme-only <w:rFonts> must survive — got: \(untouched)")
    }

    // MARK: - Section D: the precedence rules live in the model, not only in merge (verify round 1)

    /// che-word-mcp `format_text(as_revision: true)` copies the parsed
    /// properties, assigns `.color`, and hands the result to
    /// `applyRunPropertiesAsRevision`, which replaces the rPr without
    /// merging. A stale theme colour would make the tracked change invisible
    /// in Word, so assigning `color` itself must drop the theme triple.
    func testAssigningColorDirectlyClearsTheThemeTriple() throws {
        let run = try parseRunFragment(
            "<w:r><w:rPr><w:color w:val=\"000000\" w:themeColor=\"text1\" w:themeTint=\"99\" w:themeShade=\"BF\"/></w:rPr><w:t>a</w:t></w:r>")
        var props = run.properties
        props.color = "FF0000"
        XCTAssertNil(props.themeColor)
        XCTAssertNil(props.themeTint)
        XCTAssertNil(props.themeShade)
        XCTAssertEqual(props.toXML(), "<w:color w:val=\"FF0000\"/>")
    }

    func testTrackedColourChangeOnAThemeColouredRunTakesEffect() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }
        doc.enableTrackChanges(author: "Editor")

        var replacement = doc.getParagraphs()[0].runs[0].properties
        replacement.color = "FF0000"
        _ = try doc.applyRunPropertiesAsRevision(atParagraph: 0, atRunIndex: 0, newProperties: replacement)

        let edited = try paragraph(containing: "alpha", in: try savedDocumentXML(doc))
        XCTAssertTrue(edited.contains("<w:color w:val=\"FF0000\"/>"), "new colour must be explicit — got: \(edited)")
        XCTAssertFalse(edited.contains("w:val=\"FF0000\" w:themeColor"), "stale theme colour kept — got: \(edited)")
    }

    /// ECMA-376 §17.3.2.26: a theme font on an axis overrides the named font
    /// on that axis, so naming a font on one axis drops that axis's theme
    /// attribute — and only that one.
    func testAssigningANamedFontOnAnAxisClearsOnlyThatAxisThemeFont() {
        var fonts = RFontsProperties(asciiTheme: "minorHAnsi", hAnsiTheme: "minorHAnsi",
                                     eastAsiaTheme: "minorEastAsia", csTheme: "minorBidi")
        fonts.eastAsia = "PMingLiU"
        XCTAssertNil(fonts.eastAsiaTheme)
        XCTAssertEqual(fonts.asciiTheme, "minorHAnsi")
        XCTAssertEqual(fonts.hAnsiTheme, "minorHAnsi")
        XCTAssertEqual(fonts.csTheme, "minorBidi")

        fonts.ascii = "Times New Roman"; fonts.hAnsi = "Times New Roman"; fonts.cs = "Mangal"
        XCTAssertNil(fonts.asciiTheme)
        XCTAssertNil(fonts.hAnsiTheme)
        XCTAssertNil(fonts.csTheme)
    }

    /// A patch that only tints or shades adjusts the existing theme colour;
    /// it must not drop the theme link.
    func testTintOrShadeOnlyPatchKeepsTheBaseThemeColor() {
        var base = RunProperties()
        base.color = "4472C4"
        base.themeColor = "accent1"

        var patch = RunProperties()
        patch.themeShade = "BF"
        base.merge(with: patch)

        XCTAssertEqual(base.themeColor, "accent1")
        XCTAssertEqual(base.themeShade, "BF")
        XCTAssertEqual(base.color, "4472C4")
    }

    /// A new theme colour replaces the whole triple (stale tint/shade belong
    /// to the old theme colour) and keeps the existing `w:val` as fallback.
    func testThemeColorOnlyPatchReplacesTheTripleAndKeepsTheFallbackValue() {
        var base = RunProperties()
        base.color = "000000"
        base.themeColor = "text1"
        base.themeTint = "99"
        base.themeShade = "BF"

        var patch = RunProperties()
        patch.themeColor = "accent2"
        base.merge(with: patch)

        XCTAssertEqual(base.themeColor, "accent2")
        XCTAssertNil(base.themeTint)
        XCTAssertNil(base.themeShade)
        XCTAssertEqual(base.color, "000000")
    }

    /// CT_Color requires `w:val`. When a theme colour has no explicit value
    /// (a theme-only patch on a run without colour, or a source that omitted
    /// it) the writer emits `w:val="auto"` rather than a schema-invalid element.
    func testThemeColourWithoutAValueIsWrittenWithAuto() throws {
        let run = try parseRunFragment("<w:r><w:rPr><w:color w:themeColor=\"text1\"/></w:rPr><w:t>a</w:t></w:r>")
        XCTAssertNil(run.properties.color)
        XCTAssertEqual(run.properties.themeColor, "text1")
        XCTAssertEqual(run.properties.toXML(), "<w:color w:val=\"auto\" w:themeColor=\"text1\"/>")

        var base = RunProperties()
        var patch = RunProperties()
        patch.themeColor = "accent1"
        base.merge(with: patch)
        XCTAssertEqual(base.toXML(), "<w:color w:val=\"auto\" w:themeColor=\"accent1\"/>")
    }

    /// Each condition that sends an appended run off the op-log path is
    /// checked on its own, so deleting one cannot hide behind the other.
    func testAppendedRunWithOnlyAThemeColourKeepsIt() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        var props = RunProperties()
        props.color = "000000"
        props.themeColor = "text1"
        doc.appendParagraph(Paragraph(runs: [Run(text: "zeta", properties: props)]))

        let appended = try paragraph(containing: "zeta", in: try savedDocumentXML(doc))
        XCTAssertTrue(appended.contains("w:themeColor=\"text1\""), "got: \(appended)")
    }

    func testAppendedRunWithOnlyACsThemeKeepsIt() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        var props = RunProperties()
        props.rFonts = RFontsProperties(csTheme: "minorBidi")
        doc.appendParagraph(Paragraph(runs: [Run(text: "eta", properties: props)]))

        let appended = try paragraph(containing: "eta", in: try savedDocumentXML(doc))
        XCTAssertTrue(appended.contains("w:cstheme=\"minorBidi\""), "got: \(appended)")
    }

    /// The shape #274 reported: the themed run sits inside an existing
    /// tracked insertion, and the edit is `insertParagraph(.afterText)`.
    func testInsertAfterTextKeepsThemeAttributesInsideAnExistingInsertion() throws {
        let xml = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
            <w:p><w:ins w:id="7" w:author="Author A" w:date="2026-10-01T00:00:00Z">\
            <w:r><w:rPr>\(Self.themedRPr)</w:rPr><w:t>theta</w:t></w:r></w:ins></w:p>\
            <w:p><w:r><w:t>gamma</w:t></w:r></w:p>\
            <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>\
            </w:body></w:document>
            """
        let fixture = try buildFixture(documentXML: xml)
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        try doc.insertParagraph(Paragraph(text: "iota"), at: .afterText("gamma", instance: 1))

        let untouched = try paragraph(containing: "theta", in: try savedDocumentXML(doc))
        XCTAssertTrue(untouched.contains("<w:ins w:id=\"7\""), "existing insertion must survive — got: \(untouched)")
        XCTAssertTrue(untouched.contains("w:themeColor=\"text1\""), "got: \(untouched)")
        XCTAssertTrue(untouched.contains("<w:rFonts w:eastAsiaTheme=\"minorEastAsia\"/>"), "got: \(untouched)")
    }

    // MARK: - Section E: verify round 2 — tint / shade belong to a theme colour

    /// A tint-only patch adjusts the existing theme colour (the shape round-1
    /// LOW 14 reported; the shade case is covered above).
    func testTintOnlyPatchKeepsTheBaseThemeColor() {
        var base = RunProperties()
        base.color = "000000"
        base.themeColor = "text1"

        var patch = RunProperties()
        patch.themeTint = "99"
        base.merge(with: patch)

        XCTAssertEqual(base.themeColor, "text1")
        XCTAssertEqual(base.themeTint, "99")
    }

    /// With no theme colour to adjust, a tint- or shade-only patch is a
    /// no-op. It used to write `<w:color w:val="auto" …/>`, overriding the
    /// colour the run inherits from its style.
    func testTintOrShadeOnlyPatchWithoutAThemeColourWritesNoColour() {
        var base = RunProperties(bold: true)
        var patch = RunProperties()
        patch.themeTint = "66"
        patch.themeShade = "BF"
        base.merge(with: patch)

        XCTAssertNil(base.themeTint)
        XCTAssertNil(base.themeShade)
        XCTAssertFalse(base.toXML().contains("<w:color"), "got: \(base.toXML())")
    }

    /// A tint or shade with neither `color` nor `themeColor` has nothing to
    /// apply to; the writer emits no `<w:color>` for it.
    func testOrphanTintWithoutColourOrThemeIsNotWritten() {
        var props = RunProperties()
        props.themeTint = "66"
        XCTAssertEqual(props.toXML(), "")
    }

    /// Assigning `themeColor` directly drops a stale tint / shade — they
    /// belonged to the old theme colour (same rule as `merge`).
    func testAssigningThemeColorDirectlyClearsStaleTintAndShade() throws {
        let run = try parseRunFragment(
            "<w:r><w:rPr><w:color w:val=\"000000\" w:themeColor=\"text1\" w:themeTint=\"99\" w:themeShade=\"BF\"/></w:rPr><w:t>a</w:t></w:r>")
        var props = run.properties
        props.themeColor = "accent2"
        XCTAssertNil(props.themeTint)
        XCTAssertNil(props.themeShade)
        XCTAssertEqual(props.toXML(), "<w:color w:val=\"000000\" w:themeColor=\"accent2\"/>")
    }

    /// An orphan tint / shade next to an explicit value has no theme colour
    /// to modify: written by neither the typed writer nor the op-log path,
    /// so both paths agree (Codex cross-check, round 3).
    func testOrphanTintNextToAValueIsNotWritten() throws {
        let run = try parseRunFragment(
            "<w:r><w:rPr><w:color w:val=\"112233\" w:themeTint=\"80\"/></w:rPr><w:t>a</w:t></w:r>")
        XCTAssertEqual(run.properties.toXML(), "<w:color w:val=\"112233\"/>")
    }
}
