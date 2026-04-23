import XCTest
@testable import OOXMLSwift

/// Round-trip fidelity regression suite for the preserve-by-default architecture
/// added in v0.12.0 (`che-word-mcp-ooxml-roundtrip-fidelity` Spectra change).
///
/// Specs covered:
/// - `WordDocument retains source archive tempDir for round-trip preservation`
/// - `WordDocument.close() releases the archive tempDir`
/// - `DocxWriter overlay mode preserves unknown OOXML parts byte-for-byte`
/// - `ContentTypesOverlay merges typed parts with preserved Override entries`
/// - `RelationshipIdAllocator generates collision-free rIds across preserved and typed relationships`
/// - `Round-trip fidelity regression test asserts byte-equality on minimal-multipart fixture`
final class RoundTripFidelityTests: XCTestCase {

    // MARK: - WordDocument.archiveTempDir + close()

    func testInitializerBuiltDocumentHasNoArchiveTempDir() {
        let doc = WordDocument()
        XCTAssertNil(doc.archiveTempDir)
    }

    func testCloseOnInitializerBuiltDocumentIsNoOp() {
        var doc = WordDocument()
        doc.close()   // must not throw
        XCTAssertNil(doc.archiveTempDir)
    }

    func testCloseIsIdempotent() {
        var doc = WordDocument()
        doc.close()
        doc.close()   // second call must not throw
        XCTAssertNil(doc.archiveTempDir)
    }

    // MARK: - DocxReader retains tempDir

    /// Builds a minimal in-memory `.docx` fixture by reusing DocxWriter (scratch mode),
    /// then reads it back to exercise tempDir retention. Once Task 1.6 lands the binary
    /// fixture, this helper can be replaced with that more-complete fixture.
    private func makeMinimalDocxFixture() throws -> URL {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "Round-trip fixture body"))
        let tempURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("rt-fidelity-fixture-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: tempURL)
        return tempURL
    }

    func testReaderLoadedDocumentCarriesArchiveTempDir() throws {
        let fixture = try makeMinimalDocxFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }

        var doc = try DocxReader.read(from: fixture)
        defer { doc.close() }

        XCTAssertNotNil(doc.archiveTempDir)
        if let tempDir = doc.archiveTempDir {
            var isDir: ObjCBool = false
            XCTAssertTrue(
                FileManager.default.fileExists(atPath: tempDir.path, isDirectory: &isDir),
                "archiveTempDir directory must still exist after read()"
            )
            XCTAssertTrue(isDir.boolValue, "archiveTempDir must be a directory")
            XCTAssertTrue(
                FileManager.default.fileExists(
                    atPath: tempDir.appendingPathComponent("word/document.xml").path
                ),
                "archiveTempDir must contain word/document.xml"
            )
        }
    }

    // MARK: - RelationshipIdAllocator

    func testAllocatorAvoidsCollisionWithPreservedRelId() {
        let originalRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="..." Target="styles.xml"/>
          <Relationship Id="rId7" Type="..." Target="header1.xml"/>
        </Relationships>
        """
        let allocator = RelationshipIdAllocator(originalRelsXML: originalRels)
        XCTAssertEqual(allocator.allocate(), "rId8")
        XCTAssertEqual(allocator.allocate(), "rId9")
        XCTAssertEqual(allocator.allocate(), "rId10")
    }

    func testAllocatorReserveMarksIdAsTaken() {
        let allocator = RelationshipIdAllocator(originalRelsXML: "")
        allocator.reserve("rId12")
        let next = allocator.allocate()
        XCTAssertNotEqual(next, "rId12")
        XCTAssertEqual(next, "rId13")
    }

    func testAllocatorSkipsNonNumericRIdSuffix() {
        let originalRels = """
        <Relationships>
          <Relationship Id="rIdAbc" Type="..." Target="x.xml"/>
          <Relationship Id="rId5" Type="..." Target="y.xml"/>
        </Relationships>
        """
        let allocator = RelationshipIdAllocator(originalRelsXML: originalRels)
        // Non-numeric "rIdAbc" is ignored; max numeric observed is 5
        XCTAssertEqual(allocator.allocate(), "rId6")
    }

    func testAllocatorMergesAdditionalReservedIds() {
        let allocator = RelationshipIdAllocator(
            originalRelsXML: "",
            additionalReservedIds: ["rId3", "rId7", "rIdSkip"]
        )
        XCTAssertEqual(allocator.allocate(), "rId8")
    }

    func testAllocatorOnEmptyInputStartsAtRId1() {
        let allocator = RelationshipIdAllocator(originalRelsXML: "")
        XCTAssertEqual(allocator.allocate(), "rId1")
        XCTAssertEqual(allocator.allocate(), "rId2")
    }

    // MARK: - ContentTypesOverlay

    private static let documentXMLContentType =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
    private static let themeContentType =
        "application/vnd.openxmlformats-officedocument.theme+xml"
    private static let imagePNGContentType = "image/png"

    func testOverlayPreservesThemeOverrideAndReplacesDocumentOverride() {
        let originalCT = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/word/document.xml" ContentType="\(Self.documentXMLContentType)"/>
          <Override PartName="/word/theme/theme1.xml" ContentType="\(Self.themeContentType)"/>
        </Types>
        """
        let overlay = ContentTypesOverlay(originalContentTypesXML: originalCT)
        let merged = overlay.merge(typedParts: [
            PartDescriptor(partName: "/word/document.xml", contentType: Self.documentXMLContentType)
        ])

        // Theme Override preserved.
        XCTAssertTrue(merged.contains("PartName=\"/word/theme/theme1.xml\""))
        XCTAssertTrue(merged.contains("ContentType=\"\(Self.themeContentType)\""))
        // Document Override present exactly once (no duplication after replace).
        let docOccurrences = merged.components(separatedBy: "PartName=\"/word/document.xml\"").count - 1
        XCTAssertEqual(docOccurrences, 1, "document.xml Override must appear exactly once")
        // Default extensions preserved.
        XCTAssertTrue(merged.contains("Extension=\"rels\""))
        XCTAssertTrue(merged.contains("Extension=\"xml\""))
    }

    func testOverlayAddsNewImageOverride() {
        let originalCT = """
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Override PartName="/word/document.xml" ContentType="\(Self.documentXMLContentType)"/>
        </Types>
        """
        let overlay = ContentTypesOverlay(originalContentTypesXML: originalCT)
        let merged = overlay.merge(typedParts: [
            PartDescriptor(partName: "/word/document.xml", contentType: Self.documentXMLContentType),
            PartDescriptor(partName: "/word/media/imageNew.png", contentType: Self.imagePNGContentType)
        ])

        XCTAssertTrue(merged.contains("PartName=\"/word/media/imageNew.png\""))
        XCTAssertTrue(merged.contains("ContentType=\"image/png\""))
    }

    func testOverlayDropsDeletedTypedPart() {
        // Simulates `delete_footnote` removing /word/footnotes.xml from the typed model.
        let originalCT = """
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Override PartName="/word/document.xml" ContentType="\(Self.documentXMLContentType)"/>
          <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
        </Types>
        """
        let overlay = ContentTypesOverlay(originalContentTypesXML: originalCT)
        let merged = overlay.merge(
            typedParts: [
                PartDescriptor(partName: "/word/document.xml", contentType: Self.documentXMLContentType)
                // footnotes.xml absent → deleted by typed model
            ],
            typedManagedPatterns: ["/word/document.xml", "/word/footnotes.xml"]
        )

        XCTAssertFalse(
            merged.contains("PartName=\"/word/footnotes.xml\""),
            "deleted typed part must be dropped from merged Content_Types"
        )
    }

    func testCloseReleasesReaderLoadedTempDir() throws {
        let fixture = try makeMinimalDocxFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }

        var doc = try DocxReader.read(from: fixture)
        guard let tempDir = doc.archiveTempDir else {
            XCTFail("archiveTempDir was nil immediately after read()")
            return
        }
        XCTAssertTrue(FileManager.default.fileExists(atPath: tempDir.path))

        doc.close()

        XCTAssertNil(doc.archiveTempDir, "archiveTempDir must be nil after close()")
        XCTAssertFalse(
            FileManager.default.fileExists(atPath: tempDir.path),
            "tempDir directory must be removed from disk after close()"
        )
    }

    // MARK: - DocxWriter overlay mode

    /// Build a fixture with an extra "unknown" part (theme1.xml) injected by
    /// editing the scratch-built ZIP. After read+write round-trip in overlay
    /// mode, theme1.xml MUST survive byte-for-byte.
    func testOverlayModePreservesUnknownPart() throws {
        // 1. Build a basic fixture
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "Body content"))
        let baseFixture = FileManager.default.temporaryDirectory
            .appendingPathComponent("overlay-base-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: baseFixture)
        defer { try? FileManager.default.removeItem(at: baseFixture) }

        // 2. Inject an "unknown" theme1.xml into the ZIP via re-zip
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("overlay-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: stagingDir, withIntermediateDirectories: true)
        defer { ZipHelper.cleanup(stagingDir) }
        try FileManager.default.unzipItem(at: baseFixture, to: stagingDir)

        let themeBytes = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"OverlayTest\"></a:theme>"
        let themeDir = stagingDir.appendingPathComponent("word/theme")
        try FileManager.default.createDirectory(at: themeDir, withIntermediateDirectories: true)
        try themeBytes.write(
            to: themeDir.appendingPathComponent("theme1.xml"),
            atomically: true,
            encoding: .utf8
        )

        // Add Override to Content_Types so Word would consider theme1 valid
        let ctURL = stagingDir.appendingPathComponent("[Content_Types].xml")
        var ctContent = try String(contentsOf: ctURL, encoding: .utf8)
        ctContent = ctContent.replacingOccurrences(
            of: "</Types>",
            with: #"<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>"#
        )
        try ctContent.write(to: ctURL, atomically: true, encoding: .utf8)

        // Re-zip to make the source fixture
        let sourceWithTheme = FileManager.default.temporaryDirectory
            .appendingPathComponent("overlay-source-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: sourceWithTheme) }
        try ZipHelper.zip(stagingDir, to: sourceWithTheme)

        // 3. Read + write without modification (overlay mode triggered by archiveTempDir)
        var loaded = try DocxReader.read(from: sourceWithTheme)
        defer { loaded.close() }
        XCTAssertNotNil(loaded.archiveTempDir, "round-trip requires archiveTempDir for overlay")

        let dest = FileManager.default.temporaryDirectory
            .appendingPathComponent("overlay-dest-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: dest) }
        try DocxWriter.write(loaded, to: dest)

        // 4. Verify theme1.xml survives byte-for-byte
        let verifyDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("overlay-verify-\(UUID().uuidString)")
        defer { ZipHelper.cleanup(verifyDir) }
        try FileManager.default.unzipItem(at: dest, to: verifyDir)

        let preservedThemeURL = verifyDir.appendingPathComponent("word/theme/theme1.xml")
        XCTAssertTrue(
            FileManager.default.fileExists(atPath: preservedThemeURL.path),
            "theme1.xml must survive overlay round-trip"
        )
        let preservedBytes = try String(contentsOf: preservedThemeURL, encoding: .utf8)
        XCTAssertEqual(preservedBytes, themeBytes, "theme1.xml content must be byte-identical")

        // 5. Verify Content_Types still has the theme Override
        let preservedCT = try String(
            contentsOf: verifyDir.appendingPathComponent("[Content_Types].xml"),
            encoding: .utf8
        )
        XCTAssertTrue(
            preservedCT.contains("PartName=\"/word/theme/theme1.xml\""),
            "Content_Types Override for theme1.xml must be preserved by overlay"
        )
    }

    func testOverlayRoundTripPreservesZipEntryList() throws {
        // Build a fixture with multiple unknown parts (theme1.xml + custom file in customXml/)
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "Entry-list test"))
        let baseFixture = FileManager.default.temporaryDirectory
            .appendingPathComponent("entry-list-base-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: baseFixture)
        defer { try? FileManager.default.removeItem(at: baseFixture) }

        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("entry-list-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: stagingDir, withIntermediateDirectories: true)
        defer { ZipHelper.cleanup(stagingDir) }
        try FileManager.default.unzipItem(at: baseFixture, to: stagingDir)

        // Inject 2 unknown parts: word/theme/theme1.xml + customXml/item1.xml
        let themeDir = stagingDir.appendingPathComponent("word/theme")
        try FileManager.default.createDirectory(at: themeDir, withIntermediateDirectories: true)
        try "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>"
            .write(to: themeDir.appendingPathComponent("theme1.xml"),
                   atomically: true, encoding: .utf8)
        let customDir = stagingDir.appendingPathComponent("customXml")
        try FileManager.default.createDirectory(at: customDir, withIntermediateDirectories: true)
        try "<custom><value>preserve me</value></custom>"
            .write(to: customDir.appendingPathComponent("item1.xml"),
                   atomically: true, encoding: .utf8)

        let sourceFixture = FileManager.default.temporaryDirectory
            .appendingPathComponent("entry-list-src-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: sourceFixture) }
        try ZipHelper.zip(stagingDir, to: sourceFixture)

        let sourceEntries = try Self.zipEntryList(of: sourceFixture)

        // Read + write without modification
        var loaded = try DocxReader.read(from: sourceFixture)
        defer { loaded.close() }
        let dest = FileManager.default.temporaryDirectory
            .appendingPathComponent("entry-list-dest-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: dest) }
        try DocxWriter.write(loaded, to: dest)

        let destEntries = try Self.zipEntryList(of: dest)

        // Both unknown parts MUST appear in dest
        XCTAssertTrue(destEntries.contains("word/theme/theme1.xml"),
                      "theme1.xml must survive round-trip; got entries: \(destEntries.sorted())")
        XCTAssertTrue(destEntries.contains("customXml/item1.xml"),
                      "customXml/item1.xml must survive round-trip; got entries: \(destEntries.sorted())")
        // dest entry list MUST be a superset of source (no entries lost)
        let lostEntries = sourceEntries.subtracting(destEntries)
        XCTAssertTrue(lostEntries.isEmpty,
                      "Round-trip lost entries: \(lostEntries)")
    }

    func testOverlayRoundTripPreservesContentTypesOverrideSet() throws {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "Override-set test"))
        let baseFixture = FileManager.default.temporaryDirectory
            .appendingPathComponent("ct-base-\(UUID().uuidString).docx")
        try DocxWriter.write(doc, to: baseFixture)
        defer { try? FileManager.default.removeItem(at: baseFixture) }

        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("ct-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: stagingDir, withIntermediateDirectories: true)
        defer { ZipHelper.cleanup(stagingDir) }
        try FileManager.default.unzipItem(at: baseFixture, to: stagingDir)

        // Inject theme1.xml + corresponding Content_Types Override
        let themeDir = stagingDir.appendingPathComponent("word/theme")
        try FileManager.default.createDirectory(at: themeDir, withIntermediateDirectories: true)
        try "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>"
            .write(to: themeDir.appendingPathComponent("theme1.xml"),
                   atomically: true, encoding: .utf8)
        let ctURL = stagingDir.appendingPathComponent("[Content_Types].xml")
        var ctContent = try String(contentsOf: ctURL, encoding: .utf8)
        ctContent = ctContent.replacingOccurrences(
            of: "</Types>",
            with: #"<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>"#
        )
        try ctContent.write(to: ctURL, atomically: true, encoding: .utf8)

        let sourceFixture = FileManager.default.temporaryDirectory
            .appendingPathComponent("ct-src-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: sourceFixture) }
        try ZipHelper.zip(stagingDir, to: sourceFixture)

        let sourceCT = try Self.extractEntry(from: sourceFixture, path: "[Content_Types].xml")
        let sourcePartNames = Self.overridePartNames(in: sourceCT)

        // Read + write without modification
        var loaded = try DocxReader.read(from: sourceFixture)
        defer { loaded.close() }
        let dest = FileManager.default.temporaryDirectory
            .appendingPathComponent("ct-dest-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: dest) }
        try DocxWriter.write(loaded, to: dest)

        let destCT = try Self.extractEntry(from: dest, path: "[Content_Types].xml")
        let destPartNames = Self.overridePartNames(in: destCT)

        // Source PartNames MUST be a subset of dest PartNames (no Overrides lost)
        let lostOverrides = sourcePartNames.subtracting(destPartNames)
        XCTAssertTrue(lostOverrides.isEmpty,
                      "Content_Types Overrides lost in round-trip: \(lostOverrides)")
        // theme1 specifically
        XCTAssertTrue(destPartNames.contains("/word/theme/theme1.xml"))
    }

    // MARK: - Helpers for ZIP introspection

    private static func zipEntryList(of url: URL) throws -> Set<String> {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("zip-list-\(UUID().uuidString)")
        defer { ZipHelper.cleanup(stagingDir) }
        try FileManager.default.createDirectory(at: stagingDir, withIntermediateDirectories: true)
        try FileManager.default.unzipItem(at: url, to: stagingDir)

        // Normalize through realpath to handle macOS /var → /private/var symlinks.
        let resolvedRoot = stagingDir.resolvingSymlinksInPath().path
        var entries: Set<String> = []
        if let enumerator = FileManager.default.enumerator(at: stagingDir,
                                                            includingPropertiesForKeys: [.isRegularFileKey]) {
            for case let fileURL as URL in enumerator {
                let isFile = (try? fileURL.resourceValues(forKeys: [.isRegularFileKey]).isRegularFile) ?? false
                guard isFile else { continue }
                let resolvedPath = fileURL.resolvingSymlinksInPath().path
                if resolvedPath.hasPrefix(resolvedRoot + "/") {
                    let rel = String(resolvedPath.dropFirst(resolvedRoot.count + 1))
                    entries.insert(rel)
                }
            }
        }
        return entries
    }

    private static func extractEntry(from zipURL: URL, path: String) throws -> String {
        let stagingDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("extract-\(UUID().uuidString)")
        defer { ZipHelper.cleanup(stagingDir) }
        try FileManager.default.createDirectory(at: stagingDir, withIntermediateDirectories: true)
        try FileManager.default.unzipItem(at: zipURL, to: stagingDir)
        return try String(contentsOf: stagingDir.appendingPathComponent(path), encoding: .utf8)
    }

    private static func overridePartNames(in ctXML: String) -> Set<String> {
        var names: Set<String> = []
        let pattern = #"<Override\s[^>]*PartName="([^"]+)""#
        guard let regex = try? NSRegularExpression(pattern: pattern) else { return names }
        let nsString = ctXML as NSString
        let matches = regex.matches(in: ctXML, range: NSRange(location: 0, length: nsString.length))
        for match in matches where match.numberOfRanges >= 2 {
            let range = match.range(at: 1)
            guard range.location != NSNotFound else { continue }
            names.insert(nsString.substring(with: range))
        }
        return names
    }

    func testScratchModeUnchangedForInitializerBuiltDocument() throws {
        // Initializer-built doc has archiveTempDir == nil, so write() falls back to scratch mode.
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "Scratch path"))
        XCTAssertNil(doc.archiveTempDir)

        let dest = FileManager.default.temporaryDirectory
            .appendingPathComponent("scratch-mode-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: dest) }
        try DocxWriter.write(doc, to: dest)

        // Re-read should succeed and have body content
        var rt = try DocxReader.read(from: dest)
        defer { rt.close() }
        XCTAssertEqual(rt.body.children.count, 1)
    }
}
