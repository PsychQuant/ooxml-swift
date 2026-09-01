import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// PsychQuant/macdoc#175 — append-mode `insertImage` on a disk-loaded document
/// silently loses the `<w:drawing>`.
///
/// `appendParagraph` takes an op-emission fast path whenever the document came
/// from disk (`xmlTrees["word/document.xml"] != nil`). That path projects the
/// paragraph into `ParagraphPayload` + `RunPayload` — and `RunPayload` has no
/// drawing field, so a drawing-bearing run projects to an empty text run. The
/// relationship and media parts are still written, leaving the saved package
/// with rels = media = N+1 but body drawings = N (row 6 of the issue's
/// reproduction matrix). The typed view retains the drawing, which is why any
/// later typed-dirty operation resurrects the backlog (rows 4–5).
///
/// Same branch lineage as #96 (introduced the fast path) and #104 (fixed its
/// sibling resync-overwrite defect).
final class Issue175AppendImageTests: XCTestCase {

    // MARK: - Fixtures

    /// 1×1 PNG (same bytes as ActorIsolationStressTests / RelationshipIdAllocatorMutationTests).
    private let png1x1: [UInt8] = [
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D,
        0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00,
        0x0D, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x62, 0x00, 0x01, 0x00, 0x00,
        0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49,
        0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ]

    private func writeTempPNG() throws -> URL {
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue175-\(UUID().uuidString).png")
        try Data(png1x1).write(to: url)
        return url
    }

    /// Minimal three-paragraph docx on disk — mirrors the issue's `base.docx`
    /// (a macdoc-convert product with drawing count 0). Loading it through
    /// `DocxReader.read(from:wireTreeBackedViews:true)` populates `xmlTrees`,
    /// which is the precondition for the buggy fast path.
    private func threeParagraphDocx() throws -> URL {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" \
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" \
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body>\
        <w:p w14:paraId="11111111" w14:textId="11111111"><w:r><w:t>One</w:t></w:r></w:p>\
        <w:p w14:paraId="22222222" w14:textId="22222222"><w:r><w:t>Two</w:t></w:r></w:p>\
        <w:p w14:paraId="33333333" w14:textId="33333333"><w:r><w:t>Three</w:t></w:r></w:p>\
        <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>\
        </w:body></w:document>
        """
        let contentTypes = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\
        <Default Extension="xml" ContentType="application/xml"/>\
        <Override PartName="/word/document.xml" \
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>
        """
        let rootRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\
        <Relationship Id="rId1" \
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" \
        Target="word/document.xml"/></Relationships>
        """
        let docRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
        """
        let parts: [String: Data] = [
            "[Content_Types].xml": Data(contentTypes.utf8),
            "_rels/.rels": Data(rootRels.utf8),
            "word/document.xml": Data(documentXML.utf8),
            "word/_rels/document.xml.rels": Data(docRels.utf8),
        ]
        let archive = try Archive(accessMode: .create)
        for name in parts.keys.sorted() {
            let data = parts[name]!
            try archive.addEntry(with: name, type: .file,
                                 uncompressedSize: Int64(data.count),
                                 compressionMethod: .deflate) { position, size in
                let start = data.startIndex.advanced(by: Int(position))
                return data.subdata(in: start..<start.advanced(by: size))
            }
        }
        guard let bytes = archive.data else {
            XCTFail("could not obtain in-memory archive bytes")
            return URL(fileURLWithPath: "/dev/null")
        }
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue175-\(UUID().uuidString).docx")
        try bytes.write(to: url)
        return url
    }

    // MARK: - Saved-package inspection

    private struct PackageCounts {
        var bodyDrawings: Int
        var imageRels: Int
        var mediaEntries: Int
    }

    private func counts(of packageData: Data) throws -> PackageCounts {
        let archive = try Archive(data: packageData, accessMode: .read)
        func entryText(_ path: String) throws -> String {
            guard let entry = archive[path] else { return "" }
            var data = Data()
            _ = try archive.extract(entry) { data.append($0) }
            return String(decoding: data, as: UTF8.self)
        }
        let documentXML = try entryText("word/document.xml")
        let relsXML = try entryText("word/_rels/document.xml.rels")
        let media = archive.filter { $0.path.hasPrefix("word/media/") }.count
        return PackageCounts(
            bodyDrawings: documentXML.components(separatedBy: "<w:drawing").count - 1,
            imageRels: relsXML.components(separatedBy: "relationships/image").count - 1,
            mediaEntries: media)
    }

    // MARK: - The defect (issue reproduction matrix)

    /// Row 6 — the core case: append one image, save immediately. The saved
    /// package must contain the drawing, and the three counts must agree.
    func testAppendedImageSurvivesImmediateSave() throws {
        let docURL = try threeParagraphDocx()
        let pngURL = try writeTempPNG()
        defer {
            try? FileManager.default.removeItem(at: docURL)
            try? FileManager.default.removeItem(at: pngURL)
        }

        var doc = try DocxReader.read(from: docURL, wireTreeBackedViews: true)
        _ = try doc.insertImage(path: pngURL.path, widthPx: 100, heightPx: 100, at: nil)

        let saved = try DocxWriter.writeData(doc)
        let c = try counts(of: saved)
        XCTAssertEqual(c.bodyDrawings, 1,
                       "the appended image's <w:drawing> SHALL be in the saved body")
        XCTAssertEqual(c.imageRels, 1, "one image relationship expected")
        XCTAssertEqual(c.mediaEntries, 1, "one media entry expected")
    }

    /// Row 1 — two appended images, then save: both must survive.
    /// (Two distinct source files — reusing one path exercises media-name
    /// collision, which is a different concern from #175.)
    func testTwoAppendedImagesSurviveSave() throws {
        let docURL = try threeParagraphDocx()
        let pngA = try writeTempPNG()
        let pngB = try writeTempPNG()
        defer {
            try? FileManager.default.removeItem(at: docURL)
            try? FileManager.default.removeItem(at: pngA)
            try? FileManager.default.removeItem(at: pngB)
        }

        var doc = try DocxReader.read(from: docURL, wireTreeBackedViews: true)
        _ = try doc.insertImage(path: pngA.path, widthPx: 100, heightPx: 100, at: nil)
        _ = try doc.insertImage(path: pngB.path, widthPx: 100, heightPx: 100, at: nil)

        let c = try counts(of: try DocxWriter.writeData(doc))
        XCTAssertEqual(c.bodyDrawings, 2, "both appended drawings SHALL survive the save")
        XCTAssertEqual(c.imageRels, 2)
        XCTAssertEqual(c.mediaEntries, 2)
    }

    /// Row 2 — append image, then append a text paragraph, then save: the text
    /// lands either way; the image must too.
    func testAppendedImageSurvivesWhenFollowedByTextAppend() throws {
        let docURL = try threeParagraphDocx()
        let pngURL = try writeTempPNG()
        defer {
            try? FileManager.default.removeItem(at: docURL)
            try? FileManager.default.removeItem(at: pngURL)
        }

        var doc = try DocxReader.read(from: docURL, wireTreeBackedViews: true)
        _ = try doc.insertImage(path: pngURL.path, widthPx: 100, heightPx: 100, at: nil)
        doc.appendParagraph(Paragraph(text: "trailing text"))

        let saved = try DocxWriter.writeData(doc)
        let c = try counts(of: saved)
        XCTAssertEqual(c.bodyDrawings, 1,
                       "a later text append SHALL NOT be required for the image to materialize — and must not mask its loss")
        let archive = try Archive(data: saved, accessMode: .read)
        var docXML = Data()
        if let entry = archive["word/document.xml"] {
            _ = try archive.extract(entry) { docXML.append($0) }
        }
        XCTAssertTrue(String(decoding: docXML, as: UTF8.self).contains("trailing text"),
                      "the text append still lands")
    }

    // MARK: - No-regress guards

    /// Plain-text appends on a tree-backed document keep using the op-emission
    /// fast path — the fix must narrow, not remove, it.
    func testPlainTextAppendStillEmitsOps() throws {
        let docURL = try threeParagraphDocx()
        defer { try? FileManager.default.removeItem(at: docURL) }

        var doc = try DocxReader.read(from: docURL, wireTreeBackedViews: true)
        let opsBefore = doc.operationLog.entries.count
        doc.appendParagraph(Paragraph(text: "plain"))
        XCTAssertGreaterThan(doc.operationLog.entries.count, opsBefore,
                             "text-only appendParagraph SHALL keep emitting ops (fast path preserved)")
    }

    /// The anchored insert path (`InsertLocation`) was never affected; pin its
    /// behavior so the fix cannot regress it.
    func testAnchoredInsertStillWritesDrawing() throws {
        let docURL = try threeParagraphDocx()
        let pngURL = try writeTempPNG()
        defer {
            try? FileManager.default.removeItem(at: docURL)
            try? FileManager.default.removeItem(at: pngURL)
        }

        var doc = try DocxReader.read(from: docURL, wireTreeBackedViews: true)
        _ = try doc.insertImage(path: pngURL.path, widthPx: 100, heightPx: 100,
                                at: .afterText("Two", instance: 1))

        let c = try counts(of: try DocxWriter.writeData(doc))
        XCTAssertEqual(c.bodyDrawings, 1)
        XCTAssertEqual(c.imageRels, 1)
        XCTAssertEqual(c.mediaEntries, 1)
    }

    // MARK: - PackageInspector (defense-in-depth API)

    /// A healthy save reports consistent — no orphan image relationships.
    func testConsistencyReportOnHealthySave() throws {
        let docURL = try threeParagraphDocx()
        let pngURL = try writeTempPNG()
        defer {
            try? FileManager.default.removeItem(at: docURL)
            try? FileManager.default.removeItem(at: pngURL)
        }

        var doc = try DocxReader.read(from: docURL, wireTreeBackedViews: true)
        _ = try doc.insertImage(path: pngURL.path, widthPx: 100, heightPx: 100, at: nil)

        let report = try PackageInspector.imageConsistencyReport(of: try DocxWriter.writeData(doc))
        XCTAssertTrue(report.isConsistent, "orphans: \(report.orphanImageRelationshipIds)")
        XCTAssertEqual(report.bodyDrawingCount, 1)
        XCTAssertEqual(report.imageRelationshipCount, 1)
        XCTAssertEqual(report.mediaEntryCount, 1)
    }

    /// A package with an image relationship + media entry but no referencing
    /// drawing — the #175 signature — is flagged as inconsistent.
    func testConsistencyReportFlagsOrphanRelationship() throws {
        let docURL = try threeParagraphDocx()
        defer { try? FileManager.default.removeItem(at: docURL) }
        var packageData = try Data(contentsOf: docURL)

        // Rewrite the (empty) document rels to declare an image relationship
        // nothing references, and add the media entry — simulating the
        // pre-fix save output.
        let archive = try Archive(data: packageData, accessMode: .update)
        let relsXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\
        <Relationship Id="rId9" \
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" \
        Target="media/image1.png"/></Relationships>
        """
        if let old = archive["word/_rels/document.xml.rels"] { try archive.remove(old) }
        let relsData = Data(relsXML.utf8)
        try archive.addEntry(with: "word/_rels/document.xml.rels", type: .file,
                             uncompressedSize: Int64(relsData.count),
                             compressionMethod: .deflate) { position, size in
            let start = relsData.startIndex.advanced(by: Int(position))
            return relsData.subdata(in: start..<start.advanced(by: size))
        }
        let mediaData = Data(png1x1)
        try archive.addEntry(with: "word/media/image1.png", type: .file,
                             uncompressedSize: Int64(mediaData.count),
                             compressionMethod: .deflate) { position, size in
            let start = mediaData.startIndex.advanced(by: Int(position))
            return mediaData.subdata(in: start..<start.advanced(by: size))
        }
        packageData = archive.data ?? packageData

        let report = try PackageInspector.imageConsistencyReport(of: packageData)
        XCTAssertFalse(report.isConsistent)
        XCTAssertEqual(report.orphanImageRelationshipIds, ["rId9"])
        XCTAssertEqual(report.bodyDrawingCount, 0)
        XCTAssertEqual(report.imageRelationshipCount, 1)
        XCTAssertEqual(report.mediaEntryCount, 1)
    }
}
