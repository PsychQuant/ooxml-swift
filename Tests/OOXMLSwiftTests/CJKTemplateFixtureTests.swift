// CJKTemplateFixtureTests.swift
// format-alignment-engine Phase A task 1.4 — the committed synthetic CJK
// two-column template (`format-alignment-pipeline` capability, «Template
// fixture policy»; Decision 5). Asserts the generated file parses and carries
// the structural features of the private real templates it stands in for.

import XCTest
@testable import OOXMLSwift

final class CJKTemplateFixtureTests: XCTestCase {

    func testGeneratedTemplateParsesViaDocxReader() throws {
        let url = try CJKTemplateFixtureGenerator.generate()
        defer { try? FileManager.default.removeItem(at: url) }
        // Parses through the real reader without throwing.
        let doc = try DocxReader.read(from: url)
        XCTAssertNil(doc.xmlTreeLoadFailures["word/document.xml"],
                     "document.xml must parse cleanly")
        XCTAssertNotNil(doc.partTree(at: "word/document.xml"))
    }

    func testTemplateHasTwoSectionsWithSecondTwoColumn() throws {
        let url = try CJKTemplateFixtureGenerator.generate()
        defer { try? FileManager.default.removeItem(at: url) }
        let documentXML = try readPartString("word/document.xml", from: url)
        // Two sections: mid-body sectPr (ends section 1) + trailing body sectPr (section 2).
        let sectPrCount = documentXML.components(separatedBy: "<w:sectPr").count - 1
        XCTAssertEqual(sectPrCount, 2, "expected exactly two sections")
        // Second section is two-column.
        XCTAssertTrue(documentXML.contains("<w:cols w:num=\"2\""),
                      "second section must carry w:cols num=\"2\"")
    }

    func testTemplateCarriesEastAsiaFonts() throws {
        let url = try CJKTemplateFixtureGenerator.generate()
        defer { try? FileManager.default.removeItem(at: url) }
        let documentXML = try readPartString("word/document.xml", from: url)
        XCTAssertTrue(documentXML.contains("w:eastAsia="),
                      "runs must declare an eastAsia font")
    }

    func testTemplateHasAtLeastEightStyles() throws {
        let url = try CJKTemplateFixtureGenerator.generate()
        defer { try? FileManager.default.removeItem(at: url) }
        let stylesXML = try readPartString("word/styles.xml", from: url)
        let styleCount = stylesXML.components(separatedBy: "<w:style ").count - 1
        XCTAssertGreaterThanOrEqual(styleCount, 8, "expected at least 8 style definitions")
        XCTAssertTrue(stylesXML.contains("<w:docDefaults"), "styles.xml needs docDefaults")
        XCTAssertTrue(stylesXML.contains("<w:latentStyles"), "styles.xml needs latentStyles")
    }

    func testTemplateHasSettingsSurface() throws {
        let url = try CJKTemplateFixtureGenerator.generate()
        defer { try? FileManager.default.removeItem(at: url) }
        let settingsXML = try readPartString("word/settings.xml", from: url)
        XCTAssertTrue(settingsXML.contains("<w:settings"), "settings.xml must be present")
    }

    // MARK: - Helper

    private func readPartString(_ partPath: String, from url: URL) throws -> String {
        let data = try CorpusFixtureBuilder.readPart(partPath, from: url)
        return String(decoding: data, as: UTF8.self)
    }
}
