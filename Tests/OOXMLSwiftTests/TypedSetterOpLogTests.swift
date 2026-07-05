import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// word-aligned-state-sync Phase 2 task 3.15 — wire typed mutations through
/// the operation log ("Decision 4: Typed APIs as views, not as the model",
/// end-to-end; `ooxml-operation-log` Scenario "Swift-originated operation").
///
/// The document-scoped setter is the Decision-4 surface: Swift value
/// semantics prevent a free-standing `paragraph.text =` from reaching the
/// document-owned `OperationLog`, so the typed write path is
/// `document.setParagraphText(id:_:)` — same ownership shape the EditAlgebra
/// `apply(_ edit:)` surface established (WordDocument owns the log).
///
/// End-to-end contract verified here:
///   caller → op appended (source .swift) → reducer materializes tree →
///   typed view reads the new value → replay reproduces the state →
///   save persists the mutation.
final class TypedSetterOpLogTests: XCTestCase {

    // MARK: - Fixture

    private func buildFixture() throws -> URL {
        let staging = FileManager.default.temporaryDirectory
            .appendingPathComponent("typed-setter-staging-\(UUID().uuidString)")
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
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="0AB7C123"><w:r><w:t>original first</w:t></w:r></w:p><w:p w14:paraId="0DEF4567"><w:r><w:t>original second</w:t></w:r></w:p></w:body></w:document>
            """, to: "word/document.xml")

        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("typed-setter-\(UUID().uuidString).docx")
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

    private func firstParagraphID(_ doc: WordDocument) throws -> ElementID {
        for child in doc.body.children {
            if case .paragraph(let p) = child {
                guard let id = p.elementID else { break }
                return id
            }
        }
        throw XCTSkip("fixture paragraph must expose an ElementID")
    }

    // MARK: - Decision 4 end-to-end

    func testSetParagraphTextAppendsSwiftSourcedOp() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture, wireTreeBackedViews: true)
        let pid = try firstParagraphID(doc)
        let countBefore = doc.operationLog.entries.count

        try doc.setParagraphText(id: pid, "Hello")

        XCTAssertEqual(doc.operationLog.entries.count, countBefore + 1,
                       "typed setter must append exactly one op to the log")
        guard let entry = doc.operationLog.entries.last else {
            return XCTFail("missing appended entry")
        }
        XCTAssertEqual(entry.source, .swift,
                       "spec scenario 'Swift-originated operation': source must be swift")
        guard case .setText(let target, let text) = entry.op else {
            return XCTFail("appended op must be setText, got \(entry.op)")
        }
        XCTAssertEqual(target, pid, "op must reference the paragraph by ElementID, not position")
        XCTAssertEqual(text, "Hello")
    }

    func testSetParagraphTextMaterializesTreeAndTypedView() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture, wireTreeBackedViews: true)
        let pid = try firstParagraphID(doc)

        try doc.setParagraphText(id: pid, "Hello")

        // Tree materialized.
        guard let docRoot = doc.xmlTrees["word/document.xml"]?.root else {
            return XCTFail("document tree missing")
        }
        func allText(_ node: XmlNode) -> String {
            if node.kind == .text { return node.textContent }
            return node.children.map(allText).joined()
        }
        XCTAssertTrue(allText(docRoot).contains("Hello"),
                      "reducer must materialize the new text into the tree")
        XCTAssertFalse(allText(docRoot).contains("original first"),
                       "old text of the targeted paragraph must be replaced in the tree")

        // Typed view reads the new value ("paragraph.text now reads Hello
        // from the tree" — Decision 4).
        guard case .paragraph(let p) = doc.body.children.first else {
            return XCTFail("first body child must remain a paragraph")
        }
        XCTAssertEqual(p.text, "Hello",
                       "typed view must read the mutated value through the tree")

        // The untargeted paragraph is untouched.
        XCTAssertTrue(allText(docRoot).contains("original second"),
                      "non-targeted paragraph must be untouched")
    }

    func testReplayReproducesMaterializedState() throws {
        // Design goal: `state(t) = replay(ops[0..t])` — replaying the
        // persisted log against the original base tree reproduces the
        // mutated state (normalized fingerprint equality).
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }

        let baseDoc = try DocxReader.read(from: fixture, wireTreeBackedViews: true)
        guard let baseTree = baseDoc.xmlTrees["word/document.xml"] else {
            return XCTFail("base tree missing")
        }

        var doc = baseDoc
        let pid = try firstParagraphID(doc)
        try doc.setParagraphText(id: pid, "Hello")

        let replayed = try OperationReducer.materialize(
            log: doc.operationLog, base: baseTree)
        XCTAssertEqual(
            replayed.root.normalizedFingerprint(),
            doc.xmlTrees["word/document.xml"]?.root.normalizedFingerprint(),
            "replaying the log on the base tree must reproduce the live state")
    }

    func testSetParagraphTextPersistsThroughSave() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture, wireTreeBackedViews: true)
        let pid = try firstParagraphID(doc)
        try doc.setParagraphText(id: pid, "Hello")

        let outURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("typed-setter-out-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: outURL) }
        try DocxWriter.write(doc, to: outURL)

        let reread = try DocxReader.read(from: outURL)
        guard case .paragraph(let p) = reread.body.children.first else {
            return XCTFail("first body child must be a paragraph after re-read")
        }
        XCTAssertEqual(p.text, "Hello", "the mutation must survive save + re-read")
    }

    func testSetParagraphTextUnknownIDThrows() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        var doc = try DocxReader.read(from: fixture, wireTreeBackedViews: true)

        XCTAssertThrowsError(
            try doc.setParagraphText(id: ElementID(rawString: "w14:paraId=DEADBEEF"), "x"),
            "a setText targeting a nonexistent ElementID must throw, not silently no-op")
    }

    // MARK: - Paragraph.elementID convenience

    func testParagraphElementIDDerivesFromParaId() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }
        let doc = try DocxReader.read(from: fixture, wireTreeBackedViews: true)

        guard case .paragraph(let p) = doc.body.children.first else {
            return XCTFail("first body child must be a paragraph")
        }
        guard let eid = p.elementID else {
            return XCTFail("tree-backed paragraph must expose an ElementID")
        }
        XCTAssertTrue(eid.raw.contains("0AB7C123"),
                      "ElementID must derive from the existing w14:paraId")
    }
}
