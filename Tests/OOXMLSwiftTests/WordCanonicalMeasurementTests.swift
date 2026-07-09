// WordCanonicalMeasurementTests.swift
// word-canonical-forms Phase 1 task 1.2 — gated measurement run. Surveys
// every real .docx under MACDOC_TEMPLATE_DIR: the first form-gap from
// ReverseExtractor plus the full document.xml element/attribute vocabulary
// with unsupported forms flagged. This is the work queue for Phase 2/3
// (Decision 1). Prints byte counts and form names only — never content
// (fixture-privacy per #130 Decision 5). Skips loudly without the gate.
//
//   MACDOC_TEMPLATE_DIR=/path/to/private/docx swift test --filter WordCanonicalMeasurementTests

import XCTest
@testable import OOXMLSwift

final class WordCanonicalMeasurementTests: XCTestCase {

    /// Vocabulary the current extractor accepts (element localNames). A form
    /// outside this set forces the raw channel — the survey flags it.
    private static let supportedElements: Set<String> = [
        "document", "body",
        "p", "pPr", "r", "rPr", "t",
        "pStyle", "numPr", "ilvl", "numId", "spacing", "ind", "jc",
        "rFonts", "b", "i", "color", "sz", "u", "vertAlign",
        "sectPr", "headerReference", "footerReference", "pgSz", "pgMar", "cols",
        "tbl", "tblGrid", "gridCol", "tr", "tc",
    ]

    /// Attribute keys (prefix:localName) the extractor accepts.
    private static let supportedAttrs: Set<String> = [
        "w14:paraId", "w:val",
        "w:before", "w:after", "w:line", "w:lineRule",
        "w:left", "w:right", "w:firstLine", "w:hanging",
        "w:ascii", "w:eastAsia",
        "w:w", "w:h", "w:orient",
        "w:top", "w:bottom", "w:header", "w:footer", "w:gutter",
        "w:num", "w:space", "w:type", "r:id",
        // root namespace declarations the authoring default emits
        "xmlns:w", "xmlns:w14",
    ]

    private func attrKey(_ a: XmlAttribute) -> String {
        (a.prefix.map { "\($0):" } ?? "") + a.localName
    }

    private func walk(_ node: XmlNode, elements: inout Set<String>, attrs: inout Set<String>) {
        if node.kind == .element {
            elements.insert(node.localName)
            for a in node.attributes { attrs.insert(attrKey(a)) }
        }
        for c in node.children { walk(c, elements: &elements, attrs: &attrs) }
    }

    func testSurveyRealTemplatesUnderGate() throws {
        guard let dir = ProcessInfo.processInfo.environment["MACDOC_TEMPLATE_DIR"] else {
            throw XCTSkip("set MACDOC_TEMPLATE_DIR to survey real templates")
        }
        let dirURL = URL(fileURLWithPath: dir)
        let docxURLs = (try? FileManager.default.contentsOfDirectory(
            at: dirURL, includingPropertiesForKeys: nil))?
            .filter { $0.pathExtension == "docx" && !$0.lastPathComponent.hasPrefix("~$") }
            .sorted { $0.lastPathComponent < $1.lastPathComponent } ?? []
        guard !docxURLs.isEmpty else {
            throw XCTSkip("no .docx files under \(dir)")
        }

        for url in docxURLs {
            let name = url.lastPathComponent
            let parts = try RawPartChannel.readAllParts(from: url)
            guard let docData = parts["word/document.xml"] else {
                print("[measure] \(name): NO word/document.xml — skipping")
                continue
            }
            let result = try ReverseExtractor.reverse(parts: parts)
            let coverage = RawPartChannel.partLevelCoverage(parts: parts, dslParts: result.dslParts)

            // Full document.xml vocabulary survey.
            var elements: Set<String> = []
            var attrs: Set<String> = []
            if let tree = try? XmlTreeReader.parse(docData) {
                walk(tree.root, elements: &elements, attrs: &attrs)
            }
            let unsupEl = elements.subtracting(Self.supportedElements).sorted()
            let unsupAttr = attrs.subtracting(Self.supportedAttrs).sorted()

            print("========== [measure] \(name) ==========")
            print("  parts: \(coverage.parts.count) XML, \(coverage.aggregateTotalBytes) bytes; "
                  + "document.xml on DSL channel: \(result.dslParts.contains("word/document.xml"))")
            print("  first form-gap: \(result.formGaps.first.map { "\($0.contentClass) @ \($0.xmlPath.prefix(160))" } ?? "(none)")")
            print("  UNSUPPORTED elements (\(unsupEl.count)): \(unsupEl.joined(separator: ", "))")
            print("  UNSUPPORTED attrs (\(unsupAttr.count)): \(unsupAttr.joined(separator: ", "))")
            print("  ALL elements (\(elements.count)): \(elements.sorted().joined(separator: ", "))")
        }
        // The survey is a measurement probe; the assertion just confirms it ran.
        XCTAssertGreaterThan(docxURLs.count, 0)
    }
}
