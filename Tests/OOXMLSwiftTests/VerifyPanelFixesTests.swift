import XCTest
@testable import OOXMLSwift

/// 7.x verify panel — still-open findings batch (codex + v32-correctness):
/// 1. setText paragraph branch drops bookmark/commentRange/ins-del siblings
/// 2. saveWithSidecars has no backup/rollback (torn-write window)
/// 3. treeFreshParts stale-shadow: later typed mutation is overwritten by
///    the stale tree at write time
/// 4. op-refreshed generic parts (customXml/theme/…) never reach the package
/// 5. importFromDisk persists only the log — reopening replays the diff
final class VerifyPanelFixesTests: XCTestCase {

    private let ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    // MARK: 1 — setText keeps non-run siblings

    func testSetTextKeepsBookmarkAndCommentRangeSiblings() throws {
        let p = XmlNode.element(prefix: "w", localName: "p", namespaceURI: ns)
        p.setAttribute(prefix: "w14", localName: "paraId", value: "P1")
        let pPr = XmlNode.element(prefix: "w", localName: "pPr", namespaceURI: ns)
        let bmStart = XmlNode.element(prefix: "w", localName: "bookmarkStart", namespaceURI: ns)
        let r = XmlNode.element(prefix: "w", localName: "r", namespaceURI: ns,
            children: [XmlNode.element(prefix: "w", localName: "t", namespaceURI: ns,
                                       children: [XmlNode.text("old")])])
        let bmEnd = XmlNode.element(prefix: "w", localName: "bookmarkEnd", namespaceURI: ns)
        p.children = [pPr, bmStart, r, bmEnd]
        let body = XmlNode.element(prefix: "w", localName: "body", namespaceURI: ns, children: [p])
        let doc = XmlNode.element(prefix: "w", localName: "document", namespaceURI: ns, children: [body])

        var log = OperationLog()
        log.append(.setText(target: ElementID(rawString: "w14:paraId=P1"), text: "new"), source: .swift)
        let out = try OperationReducer.materialize(log: log, base: XmlTree.synthesized(root: doc))

        let outP = out.root.children[0].children[0]
        let names = outP.children.filter { $0.kind == .element }.map(\.localName)
        XCTAssertTrue(names.contains("bookmarkStart") && names.contains("bookmarkEnd"),
                      "setText must not drop bookmark markers; got \(names)")
        XCTAssertTrue(names.contains("pPr"))
        XCTAssertEqual(names.filter { $0 == "r" }.count, 1, "runs replaced by one fresh run")
    }

    // MARK: 2 — saveWithSidecars rollback

    func testSaveWithSidecarsRollsBackDocxWhenSidecarWriteFails() throws {
        let dir = FileManager.default.temporaryDirectory
            .appendingPathComponent("sws-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: dir) }
        let url = dir.appendingPathComponent("doc.docx")

        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: [.appendParagraph(in: nil, paragraph: ParagraphPayload(
            text: "x", styleId: nil, paraId: "p1"))], source: .swift)

        // Force the SECOND write (oplog sidecar) to fail: pre-create its
        // path as a directory.
        try FileManager.default.createDirectory(
            at: SidecarStore.oplogURL(for: url), withIntermediateDirectories: true)

        XCTAssertThrowsError(try doc.saveWithSidecars(to: url))
        XCTAssertFalse(FileManager.default.fileExists(atPath: url.path),
                       "docx must roll back when a sidecar write fails (torn-write window)")
    }

    // MARK: 3 — typed mutation invalidates tree freshness

    func testTypedMutationAfterOpApplyInvalidatesTreeFreshness() throws {
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: [.appendParagraph(in: nil, paragraph: ParagraphPayload(
            text: "from op", styleId: nil, paraId: "p1"))], source: .swift)
        XCTAssertTrue(doc.treeFreshParts.contains("word/document.xml"))

        // Legacy direct-typed mutation marks the part dirty (all typed
        // paths route through markTypedDirty) — freshness must drop, or
        // the stale tree would overwrite this change at write time.
        doc.markTypedDirty("word/document.xml")
        XCTAssertFalse(doc.treeFreshParts.contains("word/document.xml"),
                       "a typed-dirty mark must invalidate tree freshness (stale-shadow P1)")
    }

    // MARK: 4 — generic treeFresh parts reach the package

    func testOpRefreshedGenericPartReachesThePackage() throws {
        let dir = FileManager.default.temporaryDirectory
            .appendingPathComponent("generic-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: dir) }
        let url = dir.appendingPathComponent("doc.docx")

        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: [.appendParagraph(in: nil, paragraph: ParagraphPayload(
            text: "x", styleId: nil, paraId: "p1"))], source: .swift)
        try doc.writeAuthoringPackage(to: url)

        var reread = try DocxReader.read(from: url)
        defer { reread.close() }
        // Mutate a generic (non-typed-writer) part through the op path.
        let target = reread.xmlTrees["word/document.xml"]!.root.children[0].children[0]
        _ = target // paragraph exists; now hit a generic part: fabricate customXml via insertNode
        // Simpler: mutate document.xml via op, then verify a treeFresh part
        // that has NO typed-writer branch (docProps/app.xml is typed; use
        // a synthetic custom part) is emitted. Insert the custom part tree
        // through the public authoring apply on a custom-part-addressed op:
        reread.modifiedParts.insert("word/webSettings.xml") // typed writer has no branch for this
        // give it a tree + freshness through the internal test hook
        let root = XmlNode.element(prefix: "w", localName: "webSettings", namespaceURI: ns)
        reread.xmlTrees["word/webSettings.xml"] = XmlTree.synthesized(root: root)
        reread.treeFreshParts.insert("word/webSettings.xml")

        let out = dir.appendingPathComponent("out.docx")
        try DocxWriter.write(reread, to: out)

        var final = try DocxReader.read(from: out)
        defer { final.close() }
        XCTAssertNotNil(final.xmlTrees["word/webSettings.xml"],
                        "an op-refreshed generic part must reach the package (P1: silent drop)")
    }

    // MARK: 5 — import persists the snapshot too

    func testImportFromDiskPersistsSnapshot() throws {
        let dir = FileManager.default.temporaryDirectory
            .appendingPathComponent("imp-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: dir) }
        let url = dir.appendingPathComponent("doc.docx")

        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: [.appendParagraph(in: nil, paragraph: ParagraphPayload(
            text: "base", styleId: nil, paraId: "p1"))], source: .swift)
        try doc.saveWithSidecars(to: url)
        let snapBefore = try SidecarStore.loadSnapshot(alongside: url)

        var orchestrator = try SyncOrchestrator.bootstrapFromDocx(url: url)
        // Simulate a Word-side edit: rewrite the docx with different text.
        var editor = try WordDocument.openWithSidecars(from: url)
        try editor.apply(operations: [.setText(
            target: ElementID(rawString: "w14:paraId=p1"), text: "edited outside")], source: .word)
        try editor.writeAuthoringPackage(to: url)

        _ = try orchestrator.importFromDisk()

        let snapAfter = try SidecarStore.loadSnapshot(alongside: url)
        XCTAssertNotEqual(snapAfter?.docxSHA256, snapBefore?.docxSHA256,
                          "importFromDisk must persist a fresh snapshot — a stale snapshot replays the same Word diff on reopen")
    }
}
