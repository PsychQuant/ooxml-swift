import XCTest
@testable import OOXMLSwift

final class OperationReducerDeclaredSurfaceTests: XCTestCase {
    private struct Fixture {
        let tree: XmlTree
        let bodyID: ElementID
        let firstParagraphID: ElementID
        let secondParagraphID: ElementID
        let firstRunID: ElementID
    }

    private func makeFixture() -> Fixture {
        let bodyUUID = UUID()
        let firstParagraphUUID = UUID()
        let secondParagraphUUID = UUID()
        let firstRunUUID = UUID()

        let firstRun = XmlNode.element(
            prefix: "w", localName: "r", children: [
                XmlNode.element(prefix: "w", localName: "t", children: [.text("first")])
            ])
        firstRun.libraryUUID = firstRunUUID
        let first = XmlNode.element(prefix: "w", localName: "p", children: [firstRun])
        first.libraryUUID = firstParagraphUUID

        let second = XmlNode.element(
            prefix: "w", localName: "p", children: [
                XmlNode.element(prefix: "w", localName: "r", children: [
                    XmlNode.element(prefix: "w", localName: "t", children: [.text("second")])
                ])
            ])
        second.libraryUUID = secondParagraphUUID

        let body = XmlNode.element(prefix: "w", localName: "body", children: [first, second])
        body.libraryUUID = bodyUUID
        let root = XmlNode.element(prefix: "w", localName: "document", children: [body])
        return Fixture(
            tree: .synthesized(root: root),
            bodyID: .init(libraryUUID: bodyUUID),
            firstParagraphID: .init(libraryUUID: firstParagraphUUID),
            secondParagraphID: .init(libraryUUID: secondParagraphUUID),
            firstRunID: .init(libraryUUID: firstRunUUID)
        )
    }

    private func descendants(named localName: String, in node: XmlNode) -> [XmlNode] {
        var result: [XmlNode] = []
        if node.kind == .element, node.localName == localName { result.append(node) }
        for child in node.children { result += descendants(named: localName, in: child) }
        return result
    }

    private func joinedText(_ node: XmlNode) -> String {
        descendants(named: "t", in: node)
            .flatMap(\.children)
            .filter { $0.kind == .text }
            .map(\.textContent)
            .joined()
    }

    func testTableOperationsReplayInsertEditAndRemove() throws {
        let fixture = makeFixture()
        let tableOpID = UUID()
        var insertion = OperationLog()
        insertion.append(
            .insertTable(
                at: fixture.firstParagraphID,
                table: TablePayload(rows: 2, columns: 2, cells: [["A", "B"], ["C", "D"]])
            ),
            source: .swift,
            opID: tableOpID
        )
        insertion.append(
            .setCellText(
                table: .init(libraryUUID: tableOpID), row: 1, column: 0, text: "changed"
            ),
            source: .swift
        )

        let inserted = try OperationReducer.materialize(log: insertion, base: fixture.tree)
        let tables = descendants(named: "tbl", in: inserted.root)
        XCTAssertEqual(tables.count, 1)
        XCTAssertEqual(tables[0].libraryUUID, tableOpID)
        XCTAssertEqual(descendants(named: "tc", in: tables[0]).map(joinedText), ["A", "B", "changed", "D"])

        var removal = insertion
        removal.append(.removeTable(id: .init(libraryUUID: tableOpID)), source: .swift)
        let removed = try OperationReducer.materialize(log: removal, base: fixture.tree)
        XCTAssertTrue(descendants(named: "tbl", in: removed.root).isEmpty)
    }

    func testInsertRunAndEveryDeclaredRunFormatFieldReplay() throws {
        let fixture = makeFixture()
        let insertedRunID = UUID()
        var log = OperationLog()
        log.append(
            .insertRun(
                in: fixture.firstParagraphID,
                position: 0,
                run: RunPayload(text: "formatted")
            ),
            source: .swift,
            opID: insertedRunID
        )
        log.append(
            .setRunFormat(
                target: .init(libraryUUID: insertedRunID),
                format: RunFormatPayload(
                    bold: true,
                    italic: true,
                    underline: true,
                    fontSizeHalfPoints: 28,
                    color: "C00000"
                )
            ),
            source: .swift
        )

        let result = try OperationReducer.materialize(log: log, base: fixture.tree)
        let run = try XCTUnwrap(OperationReducer.findNode(
            elementID: .init(libraryUUID: insertedRunID), in: result))
        XCTAssertEqual(joinedText(run), "formatted")
        let rPr = try XCTUnwrap(run.children.first { $0.localName == "rPr" })
        XCTAssertEqual(rPr.children.map(\.localName), ["b", "i", "color", "sz", "u"])
        XCTAssertEqual(rPr.children.first { $0.localName == "color" }?
            .attributeValue(prefix: "w", localName: "val"), "C00000")
        XCTAssertEqual(rPr.children.first { $0.localName == "sz" }?
            .attributeValue(prefix: "w", localName: "val"), "28")
    }

    func testBookmarkAndCommentOperationsEmitCompleteAnchorMarkers() throws {
        let fixture = makeFixture()
        var log = OperationLog()
        log.append(
            .insertBookmark(at: fixture.firstParagraphID, bookmarkId: 7, name: "target"),
            source: .swift
        )
        log.append(
            .insertComment(
                anchor: fixture.secondParagraphID,
                commentId: 9,
                text: "review note",
                author: "Reviewer"
            ),
            source: .swift
        )

        let result = try OperationReducer.materialize(log: log, base: fixture.tree)
        let first = try XCTUnwrap(OperationReducer.findNode(
            elementID: fixture.firstParagraphID, in: result))
        XCTAssertEqual(descendants(named: "bookmarkStart", in: first).first?
            .attributeValue(prefix: "w", localName: "name"), "target")
        XCTAssertEqual(descendants(named: "bookmarkStart", in: first).first?
            .attributeValue(prefix: "w", localName: "id"), "7")
        XCTAssertEqual(descendants(named: "bookmarkEnd", in: first).first?
            .attributeValue(prefix: "w", localName: "id"), "7")

        let second = try XCTUnwrap(OperationReducer.findNode(
            elementID: fixture.secondParagraphID, in: result))
        XCTAssertEqual(descendants(named: "commentRangeStart", in: second).first?
            .attributeValue(prefix: "w", localName: "id"), "9")
        XCTAssertEqual(descendants(named: "commentRangeEnd", in: second).first?
            .attributeValue(prefix: "w", localName: "id"), "9")
        XCTAssertEqual(descendants(named: "commentReference", in: second).first?
            .attributeValue(prefix: "w", localName: "id"), "9")
    }

    func testTreeFallbackOperationsReplayInOrder() throws {
        let fixture = makeFixture()
        let insertedID = UUID()
        var log = OperationLog()
        log.append(
            .insertNode(
                parent: fixture.bodyID,
                position: 1,
                nodeXML: #"<w:p><w:r><w:t>inserted</w:t></w:r></w:p>"#
            ),
            source: .swift,
            opID: insertedID
        )
        log.append(
            .updateAttribute(
                target: .init(libraryUUID: insertedID),
                prefix: "w14",
                localName: "paraId",
                value: "ABCDEF12"
            ),
            source: .swift
        )
        log.append(
            .moveNode(
                source: fixture.secondParagraphID,
                destinationParent: fixture.bodyID,
                destinationIndex: 0
            ),
            source: .swift
        )

        let moved = try OperationReducer.materialize(log: log, base: fixture.tree)
        let body = try XCTUnwrap(OperationReducer.findNode(elementID: fixture.bodyID, in: moved))
        XCTAssertEqual(body.children.filter { $0.localName == "p" }.map(joinedText), ["second", "first", "inserted"])
        let inserted = try XCTUnwrap(descendants(named: "p", in: moved.root).first {
            joinedText($0) == "inserted"
        })
        XCTAssertEqual(inserted.attributeValue(prefix: "w14", localName: "paraId"), "ABCDEF12")

        var remove = log
        remove.append(
            .removeNode(target: .init(rawString: "w14:paraId=ABCDEF12")),
            source: .swift
        )
        let removed = try OperationReducer.materialize(log: remove, base: fixture.tree)
        XCTAssertFalse(descendants(named: "p", in: removed.root).contains { joinedText($0) == "inserted" })
    }

    func testRedoEntryReappliesPreviouslyUndoneOperation() throws {
        let fixture = makeFixture()
        let textOpID = UUID()
        var log = OperationLog()
        log.append(
            .setText(target: fixture.firstParagraphID, text: "changed"),
            source: .swift,
            opID: textOpID
        )
        log.append(.undo(targetOpID: textOpID), source: .swift)
        log.append(.redo(targetOpID: textOpID), source: .swift)

        let result = try OperationReducer.materialize(log: log, base: fixture.tree)
        let paragraph = try XCTUnwrap(OperationReducer.findNode(
            elementID: fixture.firstParagraphID, in: result))
        XCTAssertEqual(joinedText(paragraph), "changed")
    }

    func testUndoEntryRestoresStateBeforeTargetOperation() throws {
        let fixture = makeFixture()
        let firstChangeID = UUID()
        let secondChangeID = UUID()
        var log = OperationLog()
        log.append(
            .setText(target: fixture.firstParagraphID, text: "first change"),
            source: .swift,
            opID: firstChangeID
        )
        log.append(
            .setText(target: fixture.firstParagraphID, text: "second change"),
            source: .swift,
            opID: secondChangeID
        )
        log.append(.undo(targetOpID: secondChangeID), source: .swift)

        let result = try OperationReducer.materialize(log: log, base: fixture.tree)
        let paragraph = try XCTUnwrap(OperationReducer.findNode(
            elementID: fixture.firstParagraphID, in: result))
        XCTAssertEqual(joinedText(paragraph), "first change")
    }

    func testUndoAndRedoOfInterveningOperationReplayLaterHistory() throws {
        let fixture = makeFixture()
        let a = UUID(), b = UUID(), c = UUID()
        var undoLog = OperationLog()
        undoLog.append(.setText(target: fixture.firstParagraphID, text: "A"),
                       source: .swift, opID: a)
        undoLog.append(.setText(target: fixture.firstParagraphID, text: "B"),
                       source: .swift, opID: b)
        undoLog.append(.setText(target: fixture.firstParagraphID, text: "C"),
                       source: .swift, opID: c)
        undoLog.append(.undo(targetOpID: b), source: .swift)

        let undone = try OperationReducer.materialize(log: undoLog, base: fixture.tree)
        let undoneParagraph = try XCTUnwrap(OperationReducer.findNode(
            elementID: fixture.firstParagraphID, in: undone))
        XCTAssertEqual(joinedText(undoneParagraph), "C",
                       "later operations must replay as if B never existed")

        var redoLog = undoLog
        redoLog.append(.redo(targetOpID: b), source: .swift)
        let redone = try OperationReducer.materialize(log: redoLog, base: fixture.tree)
        let redoneParagraph = try XCTUnwrap(OperationReducer.findNode(
            elementID: fixture.firstParagraphID, in: redone))
        XCTAssertEqual(joinedText(redoneParagraph), "C",
                       "redo restores B in source order, before later C")
    }

    func testPublicApplyInterpretsUndoRedoAgainstPersistentEntryIDs() throws {
        var document = WordDocument.emptyAuthoringDocument()
        try document.apply(operations: [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "base", paraId: "CONTROL1")),
            .setText(target: .init(rawString: "w14:paraId=CONTROL1"), text: "A"),
            .setText(target: .init(rawString: "w14:paraId=CONTROL1"), text: "B"),
        ])
        let targetID = document.operationLog.entries[2].opID

        try document.apply(operations: [.undo(targetOpID: targetID)])
        var paragraph = try XCTUnwrap(document.xmlTrees["word/document.xml"]?
            .root.children.first { $0.localName == "body" }?
            .children.first { $0.localName == "p" })
        XCTAssertEqual(joinedText(paragraph), "A")

        try document.apply(operations: [.redo(targetOpID: targetID)])
        paragraph = try XCTUnwrap(document.xmlTrees["word/document.xml"]?
            .root.children.first { $0.localName == "body" }?
            .children.first { $0.localName == "p" })
        XCTAssertEqual(joinedText(paragraph), "B")
    }

    func testPublicApplyUndoAcrossCallsRestoresTheDocumentBaseline() throws {
        var document = WordDocument.emptyAuthoringDocument()
        try document.apply(operations: [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "base", paraId: "CONTROL2")),
        ])
        try document.apply(operations: [
            .setText(target: .init(rawString: "w14:paraId=CONTROL2"), text: "changed"),
        ])
        let changeID = document.operationLog.entries.last!.opID

        try document.apply(operations: [.undo(targetOpID: changeID)])
        let paragraph = try XCTUnwrap(document.xmlTrees["word/document.xml"]?
            .root.children.first { $0.localName == "body" }?
            .children.first { $0.localName == "p" })
        XCTAssertEqual(joinedText(paragraph), "base")
    }

    func testPublicApplyCanUndoStructuralBatchAsOneUnit() throws {
        var document = WordDocument.emptyAuthoringDocument()
        try document.apply(operations: [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "existing", paraId: "EXISTING")),
        ])
        let batchID = UUID()
        var log = OperationLog()
        log.append(.batchBegin(label: "structural"), source: .swift, opID: batchID)
        log.append(.appendParagraph(in: nil, paragraph: ParagraphPayload(
            text: "new", paraId: "NEWPARA")), source: .swift)
        log.append(.batchEnd, source: .swift)
        log.append(.undo(targetOpID: batchID), source: .swift)

        try document.apply(log: log)
        let body = try XCTUnwrap(document.xmlTrees["word/document.xml"]?
            .root.children.first { $0.localName == "body" })
        XCTAssertEqual(body.children.filter { $0.localName == "p" }.map(joinedText),
                       ["existing"])
    }

    func testDanglingUndoFailsLoudlyInReducerAndPublicApply() throws {
        let missing = UUID()
        let fixture = makeFixture()
        var log = OperationLog()
        log.append(.undo(targetOpID: missing), source: .swift)
        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: fixture.tree)) {
            XCTAssertEqual($0 as? ReducerError, .cannotUndo(targetOpID: missing))
        }

        var document = WordDocument.emptyAuthoringDocument()
        XCTAssertThrowsError(try document.apply(operations: [.undo(targetOpID: missing)]))
        XCTAssertTrue(document.operationLog.entries.isEmpty,
                      "a rejected control operation must not be committed")
    }

    func testTypedMutationBecomesTheBaselineForTheNextOperationSuffix() throws {
        var document = WordDocument.emptyAuthoringDocument()
        try document.apply(operations: [
            .appendParagraph(
                in: nil,
                paragraph: ParagraphPayload(text: "changed", paraId: "FIRST")),
        ])
        document.insertParagraph(Paragraph(text: "typed-only"), at: 1)

        try document.apply(operations: [
            .appendParagraph(
                in: nil,
                paragraph: ParagraphPayload(text: "next-op", paraId: "THIRD")),
        ])

        let body = try XCTUnwrap(document.xmlTrees["word/document.xml"]?
            .root.children.first { $0.localName == "body" })
        XCTAssertEqual(
            body.children.filter { $0.localName == "p" }.map(joinedText),
            ["changed", "typed-only", "next-op"],
            "starting a new operation suffix must not resurrect the stale pre-typed tree")
    }

    func testTypedToOperationBridgeUsesAnIsolatedOverlayStagingCopy() throws {
        let directory = FileManager.default.temporaryDirectory
            .appendingPathComponent("typed-op-overlay-\(UUID().uuidString)")
        let url = directory.appendingPathComponent("source.docx")
        try FileManager.default.createDirectory(
            at: directory, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: directory) }
        var seed = WordDocument()
        seed.appendParagraph(Paragraph(text: "seed"))
        try DocxWriter.write(seed, to: url)

        var document = try DocxReader.read(from: url)
        defer { document.close() }
        let archiveDocumentURL = try XCTUnwrap(document.archiveTempDir)
            .appendingPathComponent("word/document.xml")
        let archiveBefore = try Data(contentsOf: archiveDocumentURL)
        document.insertParagraph(Paragraph(text: "typed-only"), at: 1)
        try document.apply(operations: [
            .appendParagraph(
                in: nil,
                paragraph: ParagraphPayload(text: "next-op", paraId: "OVERLAY3")),
        ])

        let body = try XCTUnwrap(document.xmlTrees["word/document.xml"]?
            .root.children.first { $0.localName == "body" })
        XCTAssertEqual(
            body.children.filter { $0.localName == "p" }.map(joinedText),
            ["seed", "typed-only", "next-op"])
        XCTAssertEqual(
            try Data(contentsOf: archiveDocumentURL), archiveBefore,
            "tree refresh must not mutate the source archive staging directory")
    }

    func testUndoBatchRevertsEveryMemberAsOneLogicalUnit() throws {
        let fixture = makeFixture()
        let batchID = UUID()
        var log = OperationLog()
        log.append(.batchBegin(label: "bulk"), source: .swift, opID: batchID)
        log.append(.setText(target: fixture.firstParagraphID, text: "bulk first"),
                   source: .swift)
        log.append(.setText(target: fixture.secondParagraphID, text: "bulk second"),
                   source: .swift)
        log.append(.batchEnd, source: .swift)
        log.append(.undo(targetOpID: batchID), source: .swift)

        let result = try OperationReducer.materialize(log: log, base: fixture.tree)
        let first = try XCTUnwrap(OperationReducer.findNode(
            elementID: fixture.firstParagraphID, in: result))
        let second = try XCTUnwrap(OperationReducer.findNode(
            elementID: fixture.secondParagraphID, in: result))
        XCTAssertEqual(joinedText(first), "first")
        XCTAssertEqual(joinedText(second), "second")
    }

    func testInsertCommentAuthoringWritesDefinitionAndRelationship() throws {
        var document = WordDocument.emptyAuthoringDocument()
        let paragraphID = "COMMENT01"
        try document.apply(operations: [
            .appendParagraph(
                in: nil,
                paragraph: ParagraphPayload(text: "annotated", paraId: paragraphID)
            ),
            .insertComment(
                anchor: .init(rawString: "w14:paraId=\(paragraphID)"),
                commentId: 4,
                text: "review <this>",
                author: "A & B"
            ),
        ])

        let output = FileManager.default.temporaryDirectory
            .appendingPathComponent("reducer-comment-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: output) }
        try document.writeAuthoringPackage(to: output)

        let comments = String(decoding: try CorpusFixtureBuilder.readPart(
            "word/comments.xml", from: output), as: UTF8.self)
        XCTAssertTrue(comments.contains("review &lt;this&gt;"))
        XCTAssertTrue(comments.contains("A &amp; B"))
        XCTAssertTrue(comments.contains("w:id=\"4\""))

        let relationships = String(decoding: try CorpusFixtureBuilder.readPart(
            "word/_rels/document.xml.rels", from: output), as: UTF8.self)
        XCTAssertTrue(relationships.contains("relationships/comments"))
        XCTAssertTrue(relationships.contains("Target=\"comments.xml\""))

        let contentTypes = String(decoding: try CorpusFixtureBuilder.readPart(
            "[Content_Types].xml", from: output), as: UTF8.self)
        XCTAssertTrue(contentTypes.contains("/word/comments.xml"))
    }

    func testInsertCommentIntoExistingPackageAddsContentTypeAndRelationship() throws {
        let directory = FileManager.default.temporaryDirectory
            .appendingPathComponent("reducer-comment-overlay-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: directory, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: directory) }

        let seedURL = directory.appendingPathComponent("seed.docx")
        var seed = WordDocument.emptyAuthoringDocument()
        try seed.apply(operations: [
            .defineStyle(payload: StylePayload(styleId: "ExistingStyle")),
            .appendParagraph(
                in: nil,
                paragraph: ParagraphPayload(text: "annotated", paraId: "COMMENT02")
            )
        ])
        try seed.writeAuthoringPackage(to: seedURL)

        var document = try DocxReader.read(from: seedURL)
        defer { document.close() }
        try document.apply(operations: [
            .insertComment(
                anchor: .init(rawString: "w14:paraId=COMMENT02"),
                commentId: 5,
                text: "overlay note",
                author: "Reviewer"
            )
        ])

        let output = directory.appendingPathComponent("output.docx")
        try DocxWriter.write(document, to: output)

        let contentTypes = String(decoding: try CorpusFixtureBuilder.readPart(
            "[Content_Types].xml", from: output), as: UTF8.self)
        XCTAssertTrue(contentTypes.contains("/word/comments.xml"))

        let relationships = String(decoding: try CorpusFixtureBuilder.readPart(
            "word/_rels/document.xml.rels", from: output), as: UTF8.self)
        XCTAssertTrue(relationships.contains("relationships/comments"))
        XCTAssertTrue(relationships.contains("relationships/styles"),
                      "adding comments must preserve existing relationships")
        XCTAssertTrue(relationships.contains(
            #"xmlns="http://schemas.openxmlformats.org/package/2006/relationships""#))
        _ = try CorpusFixtureBuilder.readPart("word/comments.xml", from: output)
    }

    func testTypedBridgeMetadataCannotShadowLaterCommentMetadata() throws {
        let directory = FileManager.default.temporaryDirectory
            .appendingPathComponent("typed-comment-metadata-\(UUID().uuidString)")
        try FileManager.default.createDirectory(
            at: directory, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: directory) }

        let seedURL = directory.appendingPathComponent("seed.docx")
        var seed = WordDocument.emptyAuthoringDocument()
        try seed.apply(operations: [
            .appendParagraph(
                in: nil,
                paragraph: ParagraphPayload(text: "annotated", paraId: "COMMENT03")),
        ])
        try seed.writeAuthoringPackage(to: seedURL)

        var document = try DocxReader.read(from: seedURL)
        defer { document.close() }
        _ = document.addHeader(text: "typed header")
        try document.apply(operations: [
            .insertComment(
                anchor: .init(rawString: "w14:paraId=COMMENT03"),
                commentId: 6,
                text: "after typed bridge",
                author: "Reviewer"),
        ])

        let output = directory.appendingPathComponent("output.docx")
        try DocxWriter.write(document, to: output)
        let contentTypes = String(decoding: try CorpusFixtureBuilder.readPart(
            "[Content_Types].xml", from: output), as: UTF8.self)
        XCTAssertTrue(contentTypes.contains("/word/header1.xml"))
        XCTAssertTrue(
            contentTypes.contains("/word/comments.xml"),
            "metadata touched by a later operation must not be overwritten by a stale fresh tree")
    }

    func testAddRelationshipIntoExistingPackageUsesFreshRelationshipTree() throws {
        let directory = FileManager.default.temporaryDirectory
            .appendingPathComponent("reducer-rels-overlay-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: directory, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: directory) }

        let seedURL = directory.appendingPathComponent("seed.docx")
        var seed = WordDocument.emptyAuthoringDocument()
        try seed.apply(operations: [
            .defineStyle(payload: StylePayload(styleId: "ExistingStyle")),
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "link", paraId: "REL00001"))
        ])
        try seed.writeAuthoringPackage(to: seedURL)

        var document = try DocxReader.read(from: seedURL)
        defer { document.close() }
        try document.apply(operations: [
            .addRelationship(
                part: "word/_rels/document.xml.rels",
                id: "rIdExternal",
                type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                target: "https://example.com",
                targetMode: "External")
        ])

        let output = directory.appendingPathComponent("output.docx")
        try DocxWriter.write(document, to: output)
        let relationships = String(decoding: try CorpusFixtureBuilder.readPart(
            "word/_rels/document.xml.rels", from: output), as: UTF8.self)
        XCTAssertTrue(relationships.contains("Id=\"rIdExternal\""))
        XCTAssertTrue(relationships.contains("Target=\"https://example.com\""))
        XCTAssertTrue(relationships.contains("relationships/styles"),
                      "adding a relationship must not replace the existing set")
        XCTAssertTrue(relationships.contains(
            #"xmlns="http://schemas.openxmlformats.org/package/2006/relationships""#))
    }

    func testDefineStyleIntoExistingPackageAddsMetadataAndPart() throws {
        let directory = FileManager.default.temporaryDirectory
            .appendingPathComponent("reducer-style-overlay-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: directory, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: directory) }

        let seedURL = directory.appendingPathComponent("seed.docx")
        var seed = WordDocument.emptyAuthoringDocument()
        try seed.apply(operations: [
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "styled later", paraId: "STYLE001")),
        ])
        try seed.writeAuthoringPackage(to: seedURL)

        var document = try DocxReader.read(from: seedURL)
        defer { document.close() }
        try document.apply(operations: [
            .defineStyle(payload: StylePayload(styleId: "NewStyle")),
        ])

        let output = directory.appendingPathComponent("output.docx")
        try DocxWriter.write(document, to: output)
        let styles = String(decoding: try CorpusFixtureBuilder.readPart(
            "word/styles.xml", from: output), as: UTF8.self)
        XCTAssertTrue(styles.contains("NewStyle"))

        let relationships = String(decoding: try CorpusFixtureBuilder.readPart(
            "word/_rels/document.xml.rels", from: output), as: UTF8.self)
        XCTAssertTrue(relationships.contains("relationships/styles"))

        let contentTypes = String(decoding: try CorpusFixtureBuilder.readPart(
            "[Content_Types].xml", from: output), as: UTF8.self)
        XCTAssertTrue(contentTypes.contains("/word/styles.xml"))
    }
}
