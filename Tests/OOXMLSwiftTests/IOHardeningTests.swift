import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Regression coverage for the untrusted-`.docx` IO-hardening pass:
/// PsychQuant/ooxml-swift #93, #107, #110, #143, #145, #146, #147, #150.
///
/// Each `test...` method below documents, in its own name and comment, which
/// issue it pins. #130 / #157 (size and streaming limits) have their own
/// file, `Issue130SizeLimitsTests.swift`.
final class IOHardeningTests: XCTestCase {

    private let imageType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    private let pkgNS = "http://schemas.openxmlformats.org/package/2006/relationships"

    private func cleanupNamespace(_ ns: String) {
        ZipHelper.removeTreeForcibly(FileManager.default.temporaryDirectory.appendingPathComponent(ns))
    }

    // MARK: - #93 · writeAuthoringPackage containment

    /// `apply(operations:)` already refuses an unsafe `carryPart` path when a
    /// script op-log populates `carriedParts`. But `carriedParts` is
    /// `public internal(set)`, and `SyncOrchestrator` writes it DIRECTLY from
    /// parts read off a Word-written package — a route that never calls
    /// `apply()`. This simulates that bypass and checks the write boundary
    /// itself refuses the path, regardless of how it got into the dictionary.
    func testWriteAuthoringPackageRejectsAnUnsafePathThatReachedCarriedPartsDirectly() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        doc.carriedParts["../../../etc/evil.xml"] = Data("<x/>".utf8)
        let out = FileManager.default.temporaryDirectory.appendingPathComponent("i93-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: out) }
        XCTAssertThrowsError(try doc.writeAuthoringPackage(to: out)) { error in
            XCTAssertTrue(String(describing: error).contains("safe relative"), String(describing: error))
        }
        XCTAssertFalse(FileManager.default.fileExists(atPath: out.path), "nothing is produced on refusal")
    }

    /// A percent-encoded traversal sequence reaching `carriedParts` the same
    /// (bypassed) way — `isSafeRelativeOOXMLPath` rejects this class too, not
    /// only literal `..` segments.
    func testWriteAuthoringPackageRejectsPercentEncodedTraversalInCarriedParts() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        doc.carriedParts["word/media/%2e%2e%2fescaped.png"] = Data("x".utf8)
        let out = FileManager.default.temporaryDirectory.appendingPathComponent("i93-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: out) }
        XCTAssertThrowsError(try doc.writeAuthoringPackage(to: out)) { error in
            XCTAssertTrue(String(describing: error).contains("safe relative"), String(describing: error))
        }
    }

    /// A legitimate nested carried part is not disturbed by the new gate.
    func testWriteAuthoringPackageStillAcceptsALegitimateNestedCarriedPart() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        doc.carriedParts["word/media/image1.png"] = Data("fake-png-bytes".utf8)
        let out = FileManager.default.temporaryDirectory.appendingPathComponent("i93-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: out) }
        try doc.writeAuthoringPackage(to: out)
        XCTAssertTrue(FileManager.default.fileExists(atPath: out.path))
    }

    // MARK: - #107 · RawPartChannel.readAllParts has no size/count bound

    private func bombFile(entryBytes: Int, entries: Int = 1, name: String = "i107") throws -> URL {
        let archive = try Archive(accessMode: .create)
        let payload = Data(repeating: 0x41, count: entryBytes)
        try archive.addEntry(with: "word/document.xml", type: .file, uncompressedSize: Int64(payload.count),
                             compressionMethod: .deflate,
                             provider: { pos, size in payload.subdata(in: Int(pos)..<Int(pos) + size) })
        for i in 1..<max(entries, 1) {
            try archive.addEntry(with: "word/media/blob\(i).bin", type: .file, uncompressedSize: Int64(payload.count),
                                 compressionMethod: .deflate,
                                 provider: { pos, size in payload.subdata(in: Int(pos)..<Int(pos) + size) })
        }
        let file = FileManager.default.temporaryDirectory.appendingPathComponent("\(name)-\(UUID().uuidString).docx")
        try (archive.data ?? Data()).write(to: file)
        return file
    }

    func testReadAllPartsRefusesAnEntryOverTheDeclaredSizeLimit() throws {
        let file = try bombFile(entryBytes: 4 << 20)
        defer { try? FileManager.default.removeItem(at: file) }
        let l = ZipHelper.Limits(maximumEntryBytes: 1 << 20)
        XCTAssertThrowsError(try RawPartChannel.readAllParts(from: file, limits: l)) { error in
            let m = String(describing: error)
            XCTAssertTrue(m.contains("over the"), m)
            XCTAssertTrue(m.contains("document.xml"), m)
        }
    }

    func testReadAllPartsRefusesMoreEntriesThanTheCountLimit() throws {
        let file = try bombFile(entryBytes: 16, entries: 8)
        defer { try? FileManager.default.removeItem(at: file) }
        let l = ZipHelper.Limits(maximumEntryCount: 5)
        XCTAssertThrowsError(try RawPartChannel.readAllParts(from: file, limits: l)) { error in
            XCTAssertTrue(String(describing: error).contains("entries"), String(describing: error))
        }
    }

    func testReadAllPartsRefusesTheRunningTotalEvenWhenEveryEntryFits() throws {
        let file = try bombFile(entryBytes: 2 << 20, entries: 4)     // 4 x 2 MB = 8 MB total, each entry legal
        defer { try? FileManager.default.removeItem(at: file) }
        let l = ZipHelper.Limits(maximumEntryBytes: 4 << 20, maximumTotalBytes: 6 << 20)
        XCTAssertThrowsError(try RawPartChannel.readAllParts(from: file, limits: l)) { error in
            XCTAssertTrue(String(describing: error).contains("in total"), String(describing: error))
        }
    }

    /// Build a package, then rewrite its central directory to declare
    /// `declared` for the one entry while its real (compressed) bytes keep
    /// decompressing to `actualBytes`. ZIPFoundation's own `addEntry` always
    /// writes exactly as many bytes as it is told to declare, so this is the
    /// only way to produce an archive that genuinely understates an entry —
    /// the same technique `Issue130SizeLimitsTests` uses for the same reason.
    private func packageDeclaring(_ declared: UInt64, actualBytes: Int) throws -> Data {
        let archive = try Archive(accessMode: .create)
        let payload = Data(repeating: 0x42, count: actualBytes)
        try archive.addEntry(with: "word/document.xml", type: .file, uncompressedSize: Int64(payload.count),
                             compressionMethod: .deflate,
                             provider: { pos, size in payload.subdata(in: Int(pos)..<Int(pos) + size) })
        var raw = [UInt8](archive.data ?? Data())
        guard let i = raw.firstRange(of: Array("PK\u{01}\u{02}".utf8)) else {
            XCTFail("no central directory"); return Data()
        }
        let base = i.lowerBound   // uncompressed-size field is 4 bytes at +24
        for k in 0..<4 { raw[base + 24 + k] = UInt8((declared >> (8 * UInt64(k))) & 0xFF) }
        return Data(raw)
    }

    /// Checked per DECOMPRESSED chunk, not after the whole entry is buffered:
    /// an archive whose central directory UNDERSTATES an entry's size is
    /// refused within one chunk of the limit, not after the whole thing is
    /// already sitting in memory.
    func testReadAllPartsStreamingCheckCatchesAnEntryThatDecompressesLargerThanItDeclared() throws {
        let data = try packageDeclaring(1024, actualBytes: 4 << 20)   // declares 1 KB, actually 4 MB
        let file = FileManager.default.temporaryDirectory.appendingPathComponent("i107-lie-\(UUID().uuidString).docx")
        try data.write(to: file)
        defer { try? FileManager.default.removeItem(at: file) }
        let l = ZipHelper.Limits(maximumEntryBytes: 1 << 20)   // over the honest 1 KB declaration, under the real 4 MB
        XCTAssertThrowsError(try RawPartChannel.readAllParts(from: file, limits: l)) { error in
            XCTAssertTrue(String(describing: error).contains("decompresses"), String(describing: error))
        }
    }

    func testReadAllPartsStillReadsALegitimateSmallPackage() throws {
        let file = try bombFile(entryBytes: 64)
        defer { try? FileManager.default.removeItem(at: file) }
        let parts = try RawPartChannel.readAllParts(from: file)
        XCTAssertEqual(parts["word/document.xml"]?.count, 64)
    }

    /// R2 (PsychQuant/ooxml-swift#130 LOW-1, review sweep): the `fromDirectory:`
    /// overload (reads an already-extracted directory, not a zip) had no size
    /// bound at all — `public`, so reachable by an external caller pointing it
    /// at a directory it extracted itself, outside `ZipHelper`'s own bounded
    /// namespaces.
    func testReadAllPartsFromDirectoryRefusesAFileOverTheLimit() throws {
        let root = FileManager.default.temporaryDirectory.appendingPathComponent("i130-dir-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: root.appendingPathComponent("word"), withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: root) }
        try Data(repeating: 0x41, count: 4 << 20).write(to: root.appendingPathComponent("word/document.xml"))
        let l = ZipHelper.Limits(maximumEntryBytes: 1 << 20, maximumPartBytes: 1 << 20)
        XCTAssertThrowsError(try RawPartChannel.readAllParts(fromDirectory: root, limits: l)) { error in
            XCTAssertTrue(String(describing: error).contains("limit for a single part"), String(describing: error))
        }
    }

    func testReadAllPartsFromDirectoryStillReadsALegitimateSmallDirectory() throws {
        let root = FileManager.default.temporaryDirectory.appendingPathComponent("i130-dir-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: root.appendingPathComponent("word"), withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: root) }
        try Data(repeating: 0x41, count: 64).write(to: root.appendingPathComponent("word/document.xml"))
        let parts = try RawPartChannel.readAllParts(fromDirectory: root)
        XCTAssertEqual(parts["word/document.xml"]?.count, 64)
    }

    // MARK: - #110 · ZipHelper.zip must not delete a directory destination

    func testZipRefusesAnExistingDirectoryAsItsDestinationInsteadOfDeletingIt() throws {
        let srcDir = FileManager.default.temporaryDirectory.appendingPathComponent("i110-src-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: srcDir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: srcDir) }
        try Data("hello".utf8).write(to: srcDir.appendingPathComponent("a.txt"))

        let destDir = FileManager.default.temporaryDirectory.appendingPathComponent("i110-dest-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: destDir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: destDir) }
        let canary = destDir.appendingPathComponent("must-survive.txt")
        try Data("precious".utf8).write(to: canary)

        XCTAssertThrowsError(try ZipHelper.zip(srcDir, to: destDir)) { error in
            XCTAssertTrue(String(describing: error).contains("directory"), String(describing: error))
        }
        // The regression this issue named: the tree must still be there.
        XCTAssertTrue(FileManager.default.fileExists(atPath: canary.path), "the destination directory must not be deleted")
    }

    func testZipStillReplacesAnExistingFileDestination() throws {
        let srcDir = FileManager.default.temporaryDirectory.appendingPathComponent("i110-src-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: srcDir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: srcDir) }
        try Data("hello".utf8).write(to: srcDir.appendingPathComponent("a.txt"))

        let dest = FileManager.default.temporaryDirectory.appendingPathComponent("i110-dest-\(UUID().uuidString).zip")
        try Data("stale".utf8).write(to: dest)
        defer { try? FileManager.default.removeItem(at: dest) }

        try ZipHelper.zip(srcDir, to: dest)
        XCTAssertTrue(FileManager.default.fileExists(atPath: dest.path))
        XCTAssertNotEqual(try Data(contentsOf: dest), Data("stale".utf8))
    }

    // MARK: - #143 · strict-namespace image relationship vs RelationshipType.image

    func testStrictNamespaceImageRelationshipIsNotDeclaredAsAnImage() {
        // The reader's `RelationshipType.image` matches only the transitional
        // URL exactly. The inspector's predicate must not be looser.
        let strictImageType = "http://purl.oclc.org/ooxml/officeDocument/relationships/image"
        let relsXML = #"<Relationships xmlns="\#(pkgNS)"><Relationship Id="rId4" Type="\#(strictImageType)" Target="media/image1.png"/></Relationships>"#
        XCTAssertEqual(PackageInspector.imageRelationshipIds(inRels: relsXML), [],
                      "the reader never treats a strict-namespace Type as an image, so the inspector must not declare one as an image either")
    }

    func testTransitionalNamespaceImageRelationshipIsStillDeclared() {
        let relsXML = #"<Relationships xmlns="\#(pkgNS)"><Relationship Id="rId4" Type="\#(imageType)" Target="media/image1.png"/></Relationships>"#
        XCTAssertEqual(PackageInspector.imageRelationshipIds(inRels: relsXML), ["rId4"])
    }

    // MARK: - #145 · ZipHelper.unzip's path policy now reuses isSafeRelativeOOXMLPath

    /// A gap the pre-existing hand-rolled check (splitting only on `/`, only
    /// looking for a literal `..` segment) did not close: a `..` segment
    /// reachable only by splitting on backslash too.
    func testUnzipRefusesBackslashSeparatedTraversalThatTheOldCheckMissed() throws {
        let archive = try Archive(accessMode: .create)
        let body = Data("x".utf8)
        try archive.addEntry(with: #"word\..\..\evil.xml"#, type: .file, uncompressedSize: Int64(body.count),
                             provider: { pos, size in body.subdata(in: Int(pos)..<Int(pos) + size) })
        let ns = "i145-bs-\(UUID().uuidString)"
        defer { cleanupNamespace(ns) }
        XCTAssertThrowsError(try ZipHelper.unzip(data: try XCTUnwrap(archive.data), namespace: ns)) { error in
            XCTAssertTrue(String(describing: error).contains("safe relative OPC path"), String(describing: error))
        }
    }

    /// Percent-encoded traversal: the archive-level check never decoded `%`
    /// sequences at all, so this passed the old checks outright.
    func testUnzipRefusesPercentEncodedTraversalThatTheOldCheckMissed() throws {
        let archive = try Archive(accessMode: .create)
        let body = Data("x".utf8)
        try archive.addEntry(with: "word/media/%2e%2e%2fescaped.png", type: .file, uncompressedSize: Int64(body.count),
                             provider: { pos, size in body.subdata(in: Int(pos)..<Int(pos) + size) })
        let ns = "i145-pct-\(UUID().uuidString)"
        defer { cleanupNamespace(ns) }
        XCTAssertThrowsError(try ZipHelper.unzip(data: try XCTUnwrap(archive.data), namespace: ns)) { error in
            XCTAssertTrue(String(describing: error).contains("safe relative OPC path"), String(describing: error))
        }
    }

    /// A normal, legitimate entry path is unaffected by the added check.
    func testUnzipStillAcceptsOrdinaryEntryPaths() throws {
        let archive = try Archive(accessMode: .create)
        let body = Data("<?xml version=\"1.0\"?><w:document xmlns:w=\"x\"><w:body/></w:document>".utf8)
        try archive.addEntry(with: "word/document.xml", type: .file, uncompressedSize: Int64(body.count),
                             provider: { pos, size in body.subdata(in: Int(pos)..<Int(pos) + size) })
        try archive.addEntry(with: "word/media/image1.png", type: .file, uncompressedSize: Int64(3),
                             provider: { _, _ in Data("png".utf8) })
        let ns = "i145-ok-\(UUID().uuidString)"
        defer { cleanupNamespace(ns) }
        let out = try ZipHelper.unzip(data: try XCTUnwrap(archive.data), namespace: ns)
        ZipHelper.cleanup(out)
    }

    // MARK: - #146 · reader extraction failures never leak a raw POSIX/Cocoa error

    /// The M1_dir_0500 shape from the issue: a directory entry stored with a
    /// mode lacking owner-write, listed before its children. ZIPFoundation
    /// creates the directory, chmods it to the declared (unwritable) mode,
    /// then EACCESes writing the child inside it. Before fix round 7 this
    /// reached the caller as a raw `NSPOSIXErrorDomain` error; it must now be
    /// a named `WordError` with no temporary path in it.
    func testDirectoryEntryWithNoOwnerWriteDoesNotLeakARawPOSIXError() throws {
        let archive = try Archive(accessMode: .create)
        try archive.addEntry(with: "word/", type: .directory, uncompressedSize: Int64(0), permissions: 0o500,
                             provider: { _, _ in Data() })
        let body = Data("<?xml version=\"1.0\"?><w:document/>".utf8)
        try archive.addEntry(with: "word/document.xml", type: .file, uncompressedSize: Int64(body.count),
                             provider: { pos, size in body.subdata(in: Int(pos)..<Int(pos) + size) })
        let ns = "i146-dir500-\(UUID().uuidString)"
        defer { cleanupNamespace(ns) }
        XCTAssertThrowsError(try ZipHelper.unzip(data: try XCTUnwrap(archive.data), namespace: ns)) { error in
            XCTAssertTrue(error is WordError, "a raw Cocoa/POSIX error must not escape: \(error)")
            let m = String(describing: error)
            XCTAssertFalse(m.contains("/var/"), "no temporary path leaks into the message: \(m)")
            XCTAssertFalse(m.contains(ns), "the namespace name itself must not leak either: \(m)")
        }
    }

    /// The DocxReader-level half of #146: a rels part with a mismatched
    /// quote (`Id='rId9"`) makes `XmlTreeReader` throw `XmlTreeReaderError`,
    /// which none of `DocxReader`'s parse call sites catch — so it reaches
    /// `DocxReader.read`'s caller exactly as thrown, with only a byte offset
    /// and no other explanation. `errorDescription` now gives it one.
    func testMismatchedQuoteRelsProducesAnExplainedErrorNotABareByteOffset() throws {
        let relsXML = #"<Relationships xmlns="\#(pkgNS)"><Relationship Id='rId9" Type="\#(imageType)" Target="media/image1.png"/></Relationships>"#
        XCTAssertThrowsError(try XmlTreeReader.parse(Data(relsXML.utf8))) { error in
            guard let treeError = error as? XmlTreeReaderError else {
                return XCTFail("expected XmlTreeReaderError, got \(type(of: error))")
            }
            let description = treeError.localizedDescription
            // Not just the bare `caseName(at: N)` rendering `String(describing:)`
            // would give — an actual sentence explaining what went wrong.
            XCTAssertGreaterThan(description.count, 20, description)
            XCTAssertTrue(description.contains(" "), "a real sentence, not a case label: \(description)")
        }
    }

    func testXmlTreeReaderErrorDescriptionIsNonNilForEveryCase() {
        let cases: [XmlTreeReaderError] = [
            .emptyInput, .unexpectedEOF(at: 1), .unexpectedCharacter("x", at: 1),
            .malformedAttribute(at: 1), .malformedTag(at: 1),
            .mismatchedClosingTag(expected: "a", found: "b", at: 1),
            .invalidEntityReference(reference: "x", at: 1),
            .invalidNumericReference(reference: "1", at: 1),
            .unexpectedTrailingContent(at: 1), .nestingTooDeep(limit: 10, at: 1),
        ]
        for c in cases {
            XCTAssertNotNil(c.errorDescription, "\(c)")
            XCTAssertFalse(c.errorDescription!.isEmpty, "\(c)")
        }
    }

    // MARK: - #150 R2 · XmlTreeReaderError.errorDescription must not re-embed
    // attacker-controlled text unbounded (review finding HIGH-1)

    /// R2 finding: `decodeXmlEntities` captures "from `&` to the end of the
    /// text node" into `invalidEntityReference`'s `reference` when no `;`
    /// follows — unbounded, limited only by the part's own size (up to
    /// `maximumPartBytes`, 512 MB by default). Round-1's `errorDescription`
    /// interpolated it verbatim; the review measured a 2,000,045-character
    /// message from this exact shape.
    func testInvalidEntityReferenceErrorDescriptionIsBoundedEvenWithA2MillionCharacterReference() throws {
        let junk = String(repeating: "A", count: 2_000_000)
        let xml = "<w:t xmlns:w=\"ns\">&" + junk + "</w:t>"
        XCTAssertThrowsError(try XmlTreeReader.parse(Data(xml.utf8))) { error in
            guard case .invalidEntityReference = error as? XmlTreeReaderError else {
                return XCTFail("expected .invalidEntityReference, got \(error)")
            }
            let description = (error as! XmlTreeReaderError).errorDescription ?? ""
            XCTAssertLessThan(description.count, 600,
                              "must not reach anywhere near the 2,000,000-character input: \(description.prefix(100))…")
        }
    }

    /// R2 finding: a numeric reference's digits run to the next `;`, also
    /// unbounded — `&#` followed by two million digits and no valid scalar.
    func testInvalidNumericReferenceErrorDescriptionIsBoundedEvenWithA2MillionDigitReference() throws {
        let junk = String(repeating: "9", count: 2_000_000)
        let xml = "<w:t xmlns:w=\"ns\">&#" + junk + ";</w:t>"
        XCTAssertThrowsError(try XmlTreeReader.parse(Data(xml.utf8))) { error in
            guard case .invalidNumericReference = error as? XmlTreeReaderError else {
                return XCTFail("expected .invalidNumericReference, got \(error)")
            }
            let description = (error as! XmlTreeReaderError).errorDescription ?? ""
            XCTAssertLessThan(description.count, 600,
                              "must not reach anywhere near the 2,000,000-digit input: \(description.prefix(100))…")
        }
    }

    /// R2 finding: `parseName()` reads NameChar bytes with no length cap, so a
    /// closing tag can carry a multi-million-character name that does not
    /// match the element it closes.
    func testMismatchedClosingTagErrorDescriptionIsBoundedEvenWithA2MillionCharacterTagName() throws {
        let junk = String(repeating: "b", count: 2_000_000)
        let xml = "<a xmlns=\"ns\"></" + junk + ">"
        XCTAssertThrowsError(try XmlTreeReader.parse(Data(xml.utf8))) { error in
            guard case .mismatchedClosingTag = error as? XmlTreeReaderError else {
                return XCTFail("expected .mismatchedClosingTag, got \(error)")
            }
            let description = (error as! XmlTreeReaderError).errorDescription ?? ""
            XCTAssertLessThan(description.count, 600,
                              "must not reach anywhere near the 2,000,000-character tag name: \(description.prefix(100))…")
        }
    }

    /// R2 finding: `<\n>` throws `unexpectedCharacter('\n', ...)`, and
    /// interpolating a raw `Character` drops a literal U+000A into the
    /// message — the same "a newline forges a second line" shape #150
    /// closed for `RawChannelSlotSurgery`'s parser-error text, reopened here.
    func testUnexpectedCharacterErrorDescriptionEscapesARealNewline() throws {
        let xml = "<\n>"
        XCTAssertThrowsError(try XmlTreeReader.parse(Data(xml.utf8))) { error in
            guard case .unexpectedCharacter = error as? XmlTreeReaderError else {
                return XCTFail("expected .unexpectedCharacter, got \(error)")
            }
            let description = (error as! XmlTreeReaderError).errorDescription ?? ""
            XCTAssertFalse(description.contains("\n"),
                           "a literal newline must not reach the message: \(description.debugDescription)")
        }
    }

    // MARK: - #147 · directory entry permission bits (documented, not fixed here)

    /// PsychQuant/ooxml-swift#147: a readable package (content intact, only
    /// a directory entry's permission bits are unusual) is currently refused
    /// rather than opened. Confirmed still reproducible — fail-closed, no
    /// crash, no residue — and deliberately NOT changed in this pass: doing
    /// so means extracting file entries ourselves instead of `fm.unzipItem`
    /// (the only way to stop ZIPFoundation applying an archive-declared mode
    /// to a directory before its children are written), which re-opens the
    /// exact extraction path #137-#139's many hardening rounds were careful
    /// to leave untouched. Left to a dedicated follow-up. This test pins
    /// CURRENT behavior so a future change to it is a deliberate decision,
    /// not an accidental regression.
    func testDirectoryEntryPermissionBitsCurrentlyRefuseAnOtherwiseReadablePackage() throws {
        let archive = try Archive(accessMode: .create)
        try archive.addEntry(with: "word/", type: .directory, uncompressedSize: Int64(0), permissions: 0o500,
                             provider: { _, _ in Data() })
        let body = Data("<?xml version=\"1.0\"?><w:document/>".utf8)
        try archive.addEntry(with: "word/document.xml", type: .file, uncompressedSize: Int64(body.count),
                             provider: { pos, size in body.subdata(in: Int(pos)..<Int(pos) + size) })
        let ns = "i147-\(UUID().uuidString)"
        defer { cleanupNamespace(ns) }
        // Documents today's behavior (refused); this is NOT the fix the
        // issue asks for (a readable package like this one should ideally
        // open). See the doc comment above for why that is out of scope here.
        XCTAssertThrowsError(try ZipHelper.unzip(data: try XCTUnwrap(archive.data), namespace: ns))
    }

    // MARK: - #150 · a parser's own error text must not reach a message unbounded

    /// Reproduces the shape from the issue: an attribute value packed with
    /// thousands of invalid numeric character references, which makes
    /// `XMLParser`'s own `parserError` description large — the exact
    /// `String(describing: parser.parserError)` expression
    /// `RawChannelSlotSurgery`'s post-surgery well-formedness check used to
    /// embed verbatim. Routed through `ZipHelper.displayName` now, the same
    /// escaping/length cap already applied to attacker-controlled entry names
    /// and relationship ids.
    func testXMLParserOwnErrorTextIsCappedAndEscapedRatherThanEmbeddedVerbatim() {
        let storm = String(repeating: "&#1;", count: 9000)
        let xml = "<r a=\"\(storm)\"/>"
        let parser = XMLParser(data: Data(xml.utf8))
        XCTAssertFalse(parser.parse(), "the storm must actually fail to parse, or this test proves nothing")
        let raw = String(describing: parser.parserError as Any)
        let capped = ZipHelper.displayName(raw)
        // displayName's own contract: at most 120 scalars of the source shown,
        // rendered text capped at 480 characters plus a truncation marker.
        XCTAssertLessThanOrEqual(capped.count, 520, "capped, whatever the raw text's size (\(raw.count) chars): \(capped.prefix(200))…")
        XCTAssertFalse(capped.contains("\n"), "newlines are escaped (\\n), not passed through as real line breaks")
    }

    // MARK: - #150 R2 · XMLHardeningError.errorDescription (review sweep finding)

    /// `splitAttributes`'s "read until `=`" loop, and the production
    /// `parseContainerRootAttributes` path via `Foundation.XMLParser`'s own
    /// attribute dictionary, both cap nothing about an attribute NAME's
    /// length before `validateAttrName` rejects it — the NameChar regex only
    /// constrains character class, not length. A name that starts with a
    /// digit (illegal) followed by two million legal NameChars reaches
    /// `invalidAttributeName` unbounded.
    func testInvalidAttributeNameErrorDescriptionIsBoundedEvenWithA2MillionCharacterName() {
        let junk = "9" + String(repeating: "A", count: 2_000_000)
        let error = XMLHardeningError.invalidAttributeName(name: junk, context: "split-attributes")
        let description = error.errorDescription ?? ""
        XCTAssertLessThan(description.count, 600,
                          "must not reach anywhere near the 2,000,001-character name: \(description.prefix(100))…")
    }

    func testAttributeValueTooLargeErrorDescriptionNameIsBoundedEvenWhenUnbounded() {
        let junk = String(repeating: "a", count: 2_000_000)
        let error = XMLHardeningError.attributeValueTooLarge(name: junk, byteSize: 999_999, cap: 65_536)
        let description = error.errorDescription ?? ""
        XCTAssertLessThan(description.count, 600,
                          "the NAME must not reach anywhere near 2,000,000 characters: \(description.prefix(100))…")
    }

    /// Integration: `document.xmlTreeLoadFailures` is a `public` property.
    /// `String(describing: error)` — what this used to store — bypasses
    /// `LocalizedError.errorDescription` entirely (Swift's reflection-based
    /// description of an enum case renders its associated values raw), so a
    /// non-essential part (anything the typed model doesn't already consume
    /// — `word/fontTable.xml` here) that fails to parse with a 2,000,000-
    /// character entity reference used to leak the whole thing into a public
    /// dictionary even after `errorDescription` itself was fixed.
    func testXmlTreeLoadFailuresDoesNotLeakUnboundedContentForANonEssentialPart() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        let junk = String(repeating: "A", count: 2_000_000)
        let malformed = "<a xmlns=\"ns\">&" + junk + "</a>"
        doc.carriedParts["word/fontTable.xml"] = Data(malformed.utf8)
        let out = FileManager.default.temporaryDirectory.appendingPathComponent("i150-r2-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: out) }
        try doc.writeAuthoringPackage(to: out)

        // #130's compression-ratio guard is not what this test is about: the
        // 2,000,000-character payload is a run of one byte, which compresses
        // far past the default 500x limit. Lifted within this test's own
        // scope, matching the other fixtures in this file that build
        // deliberately repetitive content.
        let ratioExempt = ZipHelper.Limits(maximumCompressionRatio: .infinity)
        var read = try DocxReader.read(from: out, limits: ratioExempt)
        defer { read.close() }
        let failure = try XCTUnwrap(read.xmlTreeLoadFailures["word/fontTable.xml"],
                                    "the malformed part must be recorded as a load failure")
        XCTAssertLessThan(failure.count, 600,
                          "must not reach anywhere near the 2,000,000-character input: \(failure.prefix(100))…")
    }
}
