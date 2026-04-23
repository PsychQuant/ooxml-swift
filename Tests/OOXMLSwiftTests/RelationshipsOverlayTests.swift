import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Coverage tests for `che-word-mcp#35` document.xml.rels regression fix.
///
/// **Pre-fix behavior** (v0.13.0): Even though `modifiedParts` + overlay-mode
/// skip-when-not-dirty exists, two root causes still caused rels regeneration
/// on no-op round-trip with NTPU-style fixtures:
///
/// - **Root cause A**: `DocxReader.extractImages` falls back to
///   `"rId_\(fileName)"` when its `targetToId` lookup misses, producing
///   `rId_image1.png`-style ids that violate OOXML rId[0-9]+ convention. These
///   forged ids then make `hasNewTypedRelationships` return true (the typed
///   model's image.id doesn't appear in originalRels), forcing rels rewrite.
/// - **Root cause B**: `writeDocumentRelationships` builds rels from typed-model
///   parts list only — drops every rel for parts the typed model doesn't manage
///   (theme / webSettings / customXml / commentsExtensible / commentsIds / people).
///
/// **Post-fix contract**: no-op round-trip preserves `word/_rels/document.xml.rels`
/// byte-equal; even when rels rewrite IS triggered (e.g., addHeader), unknown
/// rels types (theme / people / etc.) survive untouched.
final class RelationshipsOverlayTests: XCTestCase {

    // MARK: - C1: no-op round-trip preserves rels byte-equal

    func testNoOpRoundTripPreservesDocumentRelsByteEqual() throws {
        // The MultiHeaderFooterFixture includes rels for theme (rId4) and
        // people (rId5) — neither is modeled as a typed part in WordDocument,
        // so they're the exact case that v0.13.0 was dropping.
        let fixture = try MultiHeaderFooterFixtureTests.buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }

        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }
        XCTAssertTrue(doc.modifiedPartsView.isEmpty,
                      "Reader must clear modifiedParts on no-op load")

        let dest = FileManager.default.temporaryDirectory
            .appendingPathComponent("rels-overlay-noop-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: dest) }
        try DocxWriter.write(doc, to: dest)

        let srcDir = try unzip(fixture)
        let destDir = try unzip(dest)
        defer {
            try? FileManager.default.removeItem(at: srcDir)
            try? FileManager.default.removeItem(at: destDir)
        }

        let srcRels = try Data(contentsOf: srcDir.appendingPathComponent("word/_rels/document.xml.rels"))
        let destRels = try Data(contentsOf: destDir.appendingPathComponent("word/_rels/document.xml.rels"))
        XCTAssertEqual(srcRels, destRels,
                       "word/_rels/document.xml.rels must be byte-equal after no-op round-trip — "
                       + "v0.13.0 was dropping theme + people rels (root cause B) and forging "
                       + "rId_filename ids (root cause A)")
    }

    // MARK: - C2: addHeader-induced rels rewrite still preserves unknown rels

    func testAddHeaderPreservesUnknownRelsTypes() throws {
        let fixture = try MultiHeaderFooterFixtureTests.buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }

        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        // Trigger a legitimate rels-changing edit
        _ = doc.addHeader(text: "New section header", type: .first)
        XCTAssertTrue(doc.modifiedPartsView.contains("word/_rels/document.xml.rels"),
                      "addHeader must mark rels dirty (v0.13.0 instrumentation)")

        let dest = FileManager.default.temporaryDirectory
            .appendingPathComponent("rels-overlay-addheader-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: dest) }
        try DocxWriter.write(doc, to: dest)

        let destDir = try unzip(dest)
        defer { try? FileManager.default.removeItem(at: destDir) }

        let destRels = try String(
            contentsOf: destDir.appendingPathComponent("word/_rels/document.xml.rels"),
            encoding: .utf8
        )

        // Unknown rels (theme + people) must survive even though rewrite happened.
        XCTAssertTrue(destRels.contains("theme/theme1.xml"),
                      "theme rel (Target=theme/theme1.xml) must survive rels rewrite")
        XCTAssertTrue(destRels.contains("people.xml"),
                      "people rel (Target=people.xml) must survive rels rewrite")
        // The new typed part (the added header) must appear too.
        let firstHeaderFile = doc.headers.first(where: { $0.type == .first })?.fileName ?? "headerFirst.xml"
        XCTAssertTrue(destRels.contains(firstHeaderFile),
                      "Newly-added header (\(firstHeaderFile)) must appear in rewritten rels")
    }

    // MARK: - C3: rels never contain non-numeric rId pattern

    func testRelsNeverProducesNonNumericIds() throws {
        let fixture = try MultiHeaderFooterFixtureTests.buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }

        // Both code paths should be safe: no-op round-trip AND mutation-triggered rewrite
        for shouldEdit in [false, true] {
            var doc = try DocxReader.read(from: fixture)
            defer { doc.close() }
            if shouldEdit {
                _ = doc.addHeader(text: "Trigger rels rewrite", type: .first)
            }

            let dest = FileManager.default.temporaryDirectory
                .appendingPathComponent("rels-overlay-id-\(UUID().uuidString).docx")
            defer { try? FileManager.default.removeItem(at: dest) }
            try DocxWriter.write(doc, to: dest)

            let destDir = try unzip(dest)
            defer { try? FileManager.default.removeItem(at: destDir) }
            let destRels = try String(
                contentsOf: destDir.appendingPathComponent("word/_rels/document.xml.rels"),
                encoding: .utf8
            )

            // No id like rId_image1.png / rId_xyz / etc.
            // OOXML spec: rId followed only by digits.
            let nonNumericPattern = #"Id="rId_[^"]+""#
            let regex = try NSRegularExpression(pattern: nonNumericPattern)
            let nsRange = NSRange(destRels.startIndex..., in: destRels)
            let matches = regex.matches(in: destRels, range: nsRange)
            XCTAssertEqual(matches.count, 0,
                           "rels must not contain non-numeric rId_xxx pattern "
                           + "(shouldEdit=\(shouldEdit)); found: "
                           + matches.map { (destRels as NSString).substring(with: $0.range) }.joined(separator: ", "))
        }
    }

    // MARK: - Helpers

    private func unzip(_ docx: URL) throws -> URL {
        let dir = FileManager.default.temporaryDirectory
            .appendingPathComponent("rels-overlay-verify-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        try FileManager.default.unzipItem(at: docx, to: dir)
        return dir
    }
}
