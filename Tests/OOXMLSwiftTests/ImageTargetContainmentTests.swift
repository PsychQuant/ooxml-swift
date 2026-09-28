import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Security regression tests for the `extractImages` path-traversal
/// vulnerability reported against `DocxReader.extractImages`
/// (`Sources/OOXMLSwift/IO/DocxReader.swift`, che-word-mcp `export_all_images`
/// path-traversal report). Branch `idd/ox-image-target-containment`.
///
/// `word/_rels/document.xml.rels` image relationship `Target` attributes used
/// to be joined onto `wordDir`/`tempDir` with `URL.appendingPathComponent`
/// and read with `Data(contentsOf:)` — no containment check at all. A
/// `Target` of `../../../<path>` walks straight out of the extraction
/// directory and onto the real filesystem (the OS resolves `..` during the
/// actual file open, `appendingPathComponent` never collapses it lexically).
/// A second candidate that "sanitized" the target by deleting the literal
/// substring `"../"` was itself bypassable via `"..././"`-style inputs
/// (deleting the middle occurrence makes the untouched first and last
/// characters adjacent into a NEW `"../"` a single left-to-right,
/// non-overlapping scan never re-examines).
///
/// Every test below builds its own malicious/legitimate `.docx` by hand
/// (ZIPFoundation, mirroring `Issue7XMLHardeningTests.buildHardeningFixture`)
/// and places any "outside the package" sentinel file inside a temp
/// directory the test itself creates — never a real system path.
final class ImageTargetContainmentTests: XCTestCase {

    private let imageType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

    // MARK: - Fixture builder

    private struct ImageRel {
        let id: String
        let target: String
        let targetMode: String?
        init(id: String = "rId1", target: String, targetMode: String? = nil) {
            self.id = id
            self.target = target
            self.targetMode = targetMode
        }
    }

    /// Builds a minimal-but-valid `.docx` whose `word/_rels/document.xml.rels`
    /// declares exactly the given image relationships, plus any extra parts
    /// (package-relative path -> bytes) the test wants physically present in
    /// the package (e.g. a legitimate `word/media/image1.png`, or a
    /// package-root `media/image1.png`).
    private func buildImagePackage(
        imageRelationships: [ImageRel],
        extraParts: [String: Data] = [:]
    ) throws -> URL {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("img-containment-fixture-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: stagingDir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: stagingDir) }

        func write(_ data: Data, to rel: String) throws {
            let url = stagingDir.appendingPathComponent(rel)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(), withIntermediateDirectories: true)
            try data.write(to: url)
        }
        func writeText(_ s: String, to rel: String) throws { try write(Data(s.utf8), to: rel) }

        try writeText(#"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="png" ContentType="image/png"/>
          <Default Extension="txt" ContentType="text/plain"/>
          <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>
        """#, to: "[Content_Types].xml")

        try writeText(#"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdMain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """#, to: "_rels/.rels")

        var relsXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">

        """
        for rel in imageRelationships {
            let modeAttr = rel.targetMode.map { " TargetMode=\"\($0)\"" } ?? ""
            relsXML += "  <Relationship Id=\"\(rel.id)\" Type=\"\(imageType)\" Target=\"\(rel.target)\"\(modeAttr)/>\n"
        }
        relsXML += "</Relationships>"
        try writeText(relsXML, to: "word/_rels/document.xml.rels")

        try writeText(#"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body>
        </w:document>
        """#, to: "word/document.xml")

        for (path, data) in extraParts {
            try write(data, to: path)
        }

        let outputURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("img-containment-\(UUID().uuidString).docx")
        let archive = try Archive(url: outputURL, accessMode: .create)
        let normalizedStaging = stagingDir.resolvingSymlinksInPath().path
        let basePathLen = normalizedStaging.count + 1
        let enumerator = FileManager.default.enumerator(at: stagingDir, includingPropertiesForKeys: [.isDirectoryKey])!
        for case let fileURL as URL in enumerator {
            let isDir = (try? fileURL.resourceValues(forKeys: [.isDirectoryKey]))?.isDirectory ?? false
            if isDir { continue }
            let normalizedFile = fileURL.resolvingSymlinksInPath().path
            let entryName = String(normalizedFile.dropFirst(basePathLen))
            try archive.addEntry(with: entryName, fileURL: fileURL, compressionMethod: .deflate)
        }
        return outputURL
    }

    /// Writes a uniquely-named sentinel file directly under the system temp
    /// directory (`FileManager.default.temporaryDirectory` — the same root
    /// `ZipHelper.unzip` extracts every package under, always exactly
    /// `<tempDir>/\(ZipHelper.readerNamespace)/<uuid>/`) and returns its
    /// filename plus content. Three `..` segments from `word/` (the base
    /// directory image Targets resolve against) climb: `word` -> the per-read
    /// uuid directory -> the `\(ZipHelper.readerNamespace)` namespace
    /// directory -> the system temp directory itself — exactly where this
    /// sentinel lives. This is "the temp directory's own parent", per the
    /// task's own suggested shape, never a real system path.
    private func writeSentinelOutsideExtractionRoot() throws -> (fileName: String, content: String, cleanup: () -> Void) {
        let name = "ox-sec-sentinel-\(UUID().uuidString).txt"
        let content = "OX-SEC-SENTINEL-\(UUID().uuidString)-DO-NOT-LEAK"
        let url = FileManager.default.temporaryDirectory.appendingPathComponent(name)
        try Data(content.utf8).write(to: url)
        return (name, content, { try? FileManager.default.removeItem(at: url) })
    }

    private func assertNoLeak(_ document: WordDocument, sentinel: String, file: StaticString = #filePath, line: UInt = #line) {
        for image in document.images {
            XCTAssertFalse(
                image.data.contains(Data(sentinel.utf8)),
                "sentinel content leaked into document.images via relationship '\(image.id)'",
                file: file, line: line
            )
        }
    }

    // MARK: - RED cases: direct escape (candidates 1 & 3, pre-fix)

    func testExtractImagesRefusesDirectParentTraversalEscapingThePackage() throws {
        let sentinel = try writeSentinelOutsideExtractionRoot()
        defer { sentinel.cleanup() }

        let url = try buildImagePackage(imageRelationships: [
            ImageRel(target: "../../../\(sentinel.fileName)")
        ])
        defer { try? FileManager.default.removeItem(at: url) }

        var document = try DocxReader.read(from: url)
        defer { document.close() }

        assertNoLeak(document, sentinel: sentinel.content)
        XCTAssertTrue(document.images.isEmpty, "the escaping relationship must not materialize as an image at all")
        XCTAssertNotNil(document.imageTargetContainmentFailures["rId1"],
                        "the rejection must be recorded in the loud-diagnostic channel")
    }

    // MARK: - RED case: the naive "../" substring-strip bypass (candidates 2 & 4, pre-fix)

    func testExtractImagesRefusesTheSubstringStripBypassPattern() throws {
        // "..././".replacingOccurrences(of: "../", with: "") == "../" — the
        // untouched first two characters and the untouched trailing
        // characters become adjacent into a NEW "../" the naive single-pass
        // strip never re-examines. Three repeats reproduce "../../../" after
        // one strip pass, reaching the same sentinel as the direct-escape
        // test above via the (now-removed) stripped candidate.
        let sentinel = try writeSentinelOutsideExtractionRoot()
        defer { sentinel.cleanup() }

        let bypassTarget = "..././..././..././\(sentinel.fileName)"
        XCTAssertEqual(bypassTarget.replacingOccurrences(of: "../", with: ""), "../../../\(sentinel.fileName)",
                       "fixture sanity check: confirms this Target is the intended strip-bypass shape")

        let url = try buildImagePackage(imageRelationships: [
            ImageRel(target: bypassTarget)
        ])
        defer { try? FileManager.default.removeItem(at: url) }

        var document = try DocxReader.read(from: url)
        defer { document.close() }

        assertNoLeak(document, sentinel: sentinel.content)
        XCTAssertTrue(document.images.isEmpty)
        XCTAssertNotNil(document.imageTargetContainmentFailures["rId1"])
    }

    // MARK: - Completeness: other traversal encodings (safe already, must stay safe)

    func testExtractImagesRefusesPercentEncodedTraversal() throws {
        let sentinel = try writeSentinelOutsideExtractionRoot()
        defer { sentinel.cleanup() }

        let url = try buildImagePackage(imageRelationships: [
            ImageRel(target: "%2e%2e%2f%2e%2e%2f%2e%2e%2f\(sentinel.fileName)")
        ])
        defer { try? FileManager.default.removeItem(at: url) }

        var document = try DocxReader.read(from: url)
        defer { document.close() }

        assertNoLeak(document, sentinel: sentinel.content)
        XCTAssertTrue(document.images.isEmpty)
    }

    func testExtractImagesRefusesBackslashSeparatedTraversal() throws {
        let sentinel = try writeSentinelOutsideExtractionRoot()
        defer { sentinel.cleanup() }

        let url = try buildImagePackage(imageRelationships: [
            ImageRel(target: "..\\..\\..\\\(sentinel.fileName)")
        ])
        defer { try? FileManager.default.removeItem(at: url) }

        var document = try DocxReader.read(from: url)
        defer { document.close() }

        assertNoLeak(document, sentinel: sentinel.content)
        XCTAssertTrue(document.images.isEmpty)
    }

    func testExtractImagesRefusesTargetModeExternal() throws {
        let url = try buildImagePackage(imageRelationships: [
            ImageRel(target: "http://evil.example/x", targetMode: "External")
        ])
        defer { try? FileManager.default.removeItem(at: url) }

        var document = try DocxReader.read(from: url)
        defer { document.close() }

        XCTAssertTrue(document.images.isEmpty, "TargetMode=\"External\" must never be read as a package-internal part")
        XCTAssertNotNil(document.imageTargetContainmentFailures["rId1"])
    }

    func testExtractImagesRefusesATargetThatResolvesToADirectory() throws {
        let url = try buildImagePackage(
            imageRelationships: [ImageRel(target: "media")],
            extraParts: ["word/media/dummy.png": Data("not-the-target".utf8)]
        )
        defer { try? FileManager.default.removeItem(at: url) }

        var document = try DocxReader.read(from: url)
        defer { document.close() }

        XCTAssertTrue(document.images.isEmpty, "a Target resolving to a directory must be refused, not read as a part")
        XCTAssertNotNil(document.imageTargetContainmentFailures["rId1"])
    }

    // MARK: - Leading `/`: package-root-relative, never filesystem-root-relative

    func testExtractImagesResolvesLeadingSlashTargetAgainstThePackageRootNotFilesystemRoot() throws {
        let rootImageData = Data("PACKAGE-ROOT-IMAGE-BYTES".utf8)
        let url = try buildImagePackage(
            imageRelationships: [ImageRel(target: "/media/rootimg.png")],
            extraParts: ["media/rootimg.png": rootImageData]   // package root, sibling of word/ — NOT word/media/
        )
        defer { try? FileManager.default.removeItem(at: url) }

        var document = try DocxReader.read(from: url)
        defer { document.close() }

        XCTAssertEqual(document.images.count, 1)
        XCTAssertEqual(document.images.first?.data, rootImageData,
                       "a leading-`/` Target resolves against the OPC package root, not the real filesystem root")
        XCTAssertTrue(document.imageTargetContainmentFailures.isEmpty)
    }

    // MARK: - Legitimate shapes must keep working

    func testExtractImagesStillReadsALegitimateMediaRelativeTarget() throws {
        let pngData = Data("legit-png-bytes".utf8)
        let url = try buildImagePackage(
            imageRelationships: [ImageRel(target: "media/image1.png")],
            extraParts: ["word/media/image1.png": pngData]
        )
        defer { try? FileManager.default.removeItem(at: url) }

        var document = try DocxReader.read(from: url)
        defer { document.close() }

        XCTAssertEqual(document.images.count, 1)
        XCTAssertEqual(document.images.first?.data, pngData)
        XCTAssertEqual(document.images.first?.fileName, "image1.png")
        XCTAssertTrue(document.imageTargetContainmentFailures.isEmpty)
    }

    func testExtractImagesStillReadsALegitimateParentRelativeTarget() throws {
        // `word/_rels/document.xml.rels` resolves relative to `word/` (the
        // directory `word/document.xml` itself lives in, per OPC). A Target
        // of `../media/image1.png` therefore legitimately names the
        // PACKAGE-ROOT part `media/image1.png` (a sibling of `word/`), not
        // `word/media/image1.png` — some real-world producers lay out media
        // this way. This must keep resolving correctly, not merely "not
        // crash".
        let pngData = Data("legit-sibling-png-bytes".utf8)
        let url = try buildImagePackage(
            imageRelationships: [ImageRel(target: "../media/image1.png")],
            extraParts: ["media/image1.png": pngData]
        )
        defer { try? FileManager.default.removeItem(at: url) }

        var document = try DocxReader.read(from: url)
        defer { document.close() }

        XCTAssertEqual(document.images.count, 1)
        XCTAssertEqual(document.images.first?.data, pngData)
        XCTAssertTrue(document.imageTargetContainmentFailures.isEmpty)
    }

    // MARK: - Defense-in-depth: the shared primitive itself, against a real symlink

    /// `ZipHelper.unzip` already refuses to extract any zip entry of type
    /// `.symlink` (see `ZipHelper.swift`), so a real `.docx` can never smuggle
    /// a symlink into `tempDir` through `DocxReader.read` — this scenario is
    /// tested directly against `resolveContainedOOXMLTarget` (the shared
    /// containment primitive) with an on-disk tree the test builds by hand,
    /// as defense-in-depth for any caller that prepares a "package root"
    /// directory a different way.
    func testResolveContainedOOXMLTargetRefusesASymlinkEscapingThePackageRoot() throws {
        let root = FileManager.default.temporaryDirectory.appendingPathComponent("symlink-root-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: root.appendingPathComponent("word/media"), withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: root) }

        let outsideFile = FileManager.default.temporaryDirectory.appendingPathComponent("symlink-outside-\(UUID().uuidString).txt")
        try Data("OUTSIDE-VIA-SYMLINK".utf8).write(to: outsideFile)
        defer { try? FileManager.default.removeItem(at: outsideFile) }

        let symlinkPath = root.appendingPathComponent("word/media/evil.png")
        try FileManager.default.createSymbolicLink(at: symlinkPath, withDestinationURL: outsideFile)

        let resolved = resolveContainedOOXMLTarget(
            "media/evil.png", targetMode: nil, relativeToPartDirectory: "/word/", packageRoot: root
        )
        XCTAssertNil(resolved, "a symlink inside the package pointing outside it must be refused")
    }

    func testResolveContainedOOXMLTargetAcceptsARegularFileStrictlyInsideTheRoot() throws {
        let root = FileManager.default.temporaryDirectory.appendingPathComponent("symlink-root-ok-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: root.appendingPathComponent("word/media"), withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: root) }
        try Data("ok".utf8).write(to: root.appendingPathComponent("word/media/image1.png"))

        let resolved = resolveContainedOOXMLTarget(
            "media/image1.png", targetMode: nil, relativeToPartDirectory: "/word/", packageRoot: root
        )
        XCTAssertNotNil(resolved)
        XCTAssertEqual(try? Data(contentsOf: XCTUnwrap(resolved)), Data("ok".utf8))
    }

    // MARK: - End-to-end: a rejected image must not resurface through a subsequent save

    func testMaliciousTargetDoesNotResurfaceInAWrittenBackPackage() throws {
        let sentinel = try writeSentinelOutsideExtractionRoot()
        defer { sentinel.cleanup() }

        let legitData = Data("legit-roundtrip-png".utf8)
        let url = try buildImagePackage(
            imageRelationships: [
                ImageRel(id: "rIdEvil", target: "../../../\(sentinel.fileName)"),
                ImageRel(id: "rIdOk", target: "media/image1.png"),
            ],
            extraParts: ["word/media/image1.png": legitData]
        )
        defer { try? FileManager.default.removeItem(at: url) }

        var document = try DocxReader.read(from: url)
        defer { document.close() }
        assertNoLeak(document, sentinel: sentinel.content)
        XCTAssertEqual(document.images.count, 1, "only the legitimate image relationship survives extraction")

        let outURL = FileManager.default.temporaryDirectory.appendingPathComponent("img-containment-roundtrip-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(document, to: outURL)

        let writtenParts = try RawPartChannel.readAllParts(from: outURL)
        for (part, data) in writtenParts {
            XCTAssertFalse(data.contains(Data(sentinel.content.utf8)),
                           "sentinel content must not be embedded anywhere in the written-back package (part \(part))")
        }
        // The legitimate image must still be there — the fix must not have
        // collaterally dropped a perfectly safe image.
        XCTAssertTrue(writtenParts.values.contains(where: { $0 == legitData }))
    }
}
