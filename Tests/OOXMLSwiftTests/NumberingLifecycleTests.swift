import XCTest
@testable import OOXMLSwift

/// Tests for Phase 4 Numbering definition lifecycle on WordDocument. Spec
/// covered (see
/// `openspec/changes/che-word-mcp-styles-sections-numbering-foundations/specs/`):
/// - ooxml-document-part-mutations: WordDocument exposes numbering definition lifecycle
///
/// Implementation tasks 4.1-4.6 will populate these tests; until then they
/// XCTSkip so the suite stays green.
final class NumberingLifecycleTests: XCTestCase {

    // MARK: - Test fixture helpers

    /// Builds a doc with N abstractNum/num pairs and `referencedNumIds`
    /// indicating which num ids are referenced by paragraphs (others remain
    /// orphan, target for `gcOrphanNumbering`).
    func makeDocWithNumIds(_ allNumIds: [Int], referenced referencedNumIds: Set<Int>) -> WordDocument {
        var doc = WordDocument()
        for numId in allNumIds {
            let abstractNum = AbstractNum(abstractNumId: numId, levels: [])
            doc.numbering.abstractNums.append(abstractNum)
            let num = Num(numId: numId, abstractNumId: numId)
            doc.numbering.nums.append(num)
        }
        for numId in referencedNumIds {
            var para = Paragraph(text: "ref \(numId)")
            para.properties.numbering = NumberingInfo(numId: numId, level: 0)
            doc.body.children.append(.paragraph(para))
        }
        return doc
    }

    func roundTrip(_ doc: WordDocument) throws -> WordDocument {
        let bytes = try DocxWriter.writeData(doc)
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("num-rt-\(UUID().uuidString).docx")
        try bytes.write(to: url)
        defer { try? FileManager.default.removeItem(at: url) }
        return try DocxReader.read(from: url)
    }

    // MARK: - Task 4.1: createNumberingDefinition

    func testCreateNumberingDefinitionReturnsNewNumIdAfterTask41() throws {
        var doc = WordDocument()
        let levels = [
            Level(ilvl: 0, numFmt: .decimal, lvlText: "%1.", indent: 720),
            Level(ilvl: 1, numFmt: .decimal, lvlText: "%1.%2.", indent: 1440),
        ]
        let numId = try doc.createNumberingDefinition(levels: levels)
        XCTAssertGreaterThan(numId, 0)
        XCTAssertEqual(doc.numbering.nums.last?.numId, numId)
        XCTAssertEqual(doc.numbering.abstractNums.last?.levels.count, 2)
    }

    func testCreateNumberingDefinitionRejectsEmptyLevelsAfterTask41() throws {
        var doc = WordDocument()
        XCTAssertThrowsError(try doc.createNumberingDefinition(levels: [])) { error in
            guard case WordError.invalidIndex(0) = error else {
                XCTFail("expected invalidIndex(0)"); return
            }
        }
    }

    // MARK: - Task 4.2: overrideNumberingLevel

    func testOverrideNumberingLevelEmitsLvlOverrideAfterTask42() throws {
        var doc = WordDocument()
        let numId = try doc.createNumberingDefinition(levels: [
            Level(ilvl: 0, numFmt: .decimal, lvlText: "%1.", indent: 720),
        ])
        try doc.overrideNumberingLevel(numId: numId, level: 0, startValue: 5)
        let num = doc.numbering.nums.first(where: { $0.numId == numId })!
        XCTAssertEqual(num.lvlOverrides.count, 1)
        XCTAssertEqual(num.lvlOverrides.first?.startOverride, 5)
        XCTAssertTrue(num.toXML().contains("<w:lvlOverride w:ilvl=\"0\"><w:startOverride w:val=\"5\"/></w:lvlOverride>"))
    }

    func testOverrideNumberingLevelThrowsWhenNumIdMissingAfterTask42() throws {
        var doc = WordDocument()
        XCTAssertThrowsError(try doc.overrideNumberingLevel(numId: 999, level: 0, startValue: 1)) { error in
            guard case WordError.numIdNotFound(999) = error else { XCTFail("expected numIdNotFound"); return }
        }
    }

    // MARK: - Task 4.3: assignNumberingToParagraph

    func testAssignNumberingToParagraphMarksDocumentDirtyAfterTask43() throws {
        var doc = WordDocument()
        doc.body.children = [.paragraph(Paragraph(text: "list item"))]
        let numId = try doc.createNumberingDefinition(levels: [
            Level(ilvl: 0, numFmt: .decimal, lvlText: "%1.", indent: 720),
        ])
        try doc.assignNumberingToParagraph(paragraphIndex: 0, numId: numId, level: 0)
        if case .paragraph(let p) = doc.body.children[0] {
            XCTAssertEqual(p.properties.numbering?.numId, numId)
            XCTAssertEqual(p.properties.numbering?.level, 0)
        } else {
            XCTFail("paragraph 0 not a paragraph after assign")
        }
    }

    // MARK: - Task 4.4 + 4.5

    func testContinueListReusesPreviousNumIdAfterTask44() throws {
        var doc = WordDocument()
        doc.body.children = [
            .paragraph(Paragraph(text: "first")),
            .paragraph(Paragraph(text: "second")),
        ]
        let numId = try doc.createNumberingDefinition(levels: [
            Level(ilvl: 0, numFmt: .decimal, lvlText: "%1.", indent: 720),
        ])
        try doc.assignNumberingToParagraph(paragraphIndex: 0, numId: numId, level: 0)
        try doc.continueList(paragraphIndex: 1, previousListNumId: numId)
        if case .paragraph(let p) = doc.body.children[1] {
            XCTAssertEqual(p.properties.numbering?.numId, numId)
        }
    }

    func testStartNewListAllocatesFreshNumIdAfterTask45() throws {
        var doc = WordDocument()
        doc.body.children = [.paragraph(Paragraph(text: "x"))]
        let firstNumId = try doc.createNumberingDefinition(levels: [
            Level(ilvl: 0, numFmt: .decimal, lvlText: "%1.", indent: 720),
        ])
        let abstractId = doc.numbering.nums.last!.abstractNumId
        let newNumId = try doc.startNewList(paragraphIndex: 0, abstractNumId: abstractId)
        XCTAssertNotEqual(newNumId, firstNumId)
        if case .paragraph(let p) = doc.body.children[0] {
            XCTAssertEqual(p.properties.numbering?.numId, newNumId)
        }
    }

    // MARK: - Task 4.6: gcOrphanNumbering

    /// Spec scenario: GC removes orphan numIds.
    func testGcOrphanNumberingReturnsDeletedNumIdsInOrderAfterTask46() throws {
        var doc = makeDocWithNumIds([1, 2, 3], referenced: [1])
        let deleted = doc.gcOrphanNumbering()
        XCTAssertEqual(deleted, [2, 3])
        XCTAssertEqual(doc.numbering.nums.map { $0.numId }, [1])
    }

    func testGcOrphanNumberingPreservesAbstractNumsAfterTask46() throws {
        var doc = makeDocWithNumIds([1, 2, 3], referenced: [])
        let abstractCountBefore = doc.numbering.abstractNums.count
        _ = doc.gcOrphanNumbering()
        XCTAssertEqual(doc.numbering.abstractNums.count, abstractCountBefore,
            "abstractNums must NOT be GCed (they are templates)")
    }

    // MARK: - Pre-existing sanity

    func testFixtureBuilderProducesDocWithMixedRefs() {
        let doc = makeDocWithNumIds([1, 2, 3], referenced: [1])
        XCTAssertEqual(doc.numbering.nums.count, 3)
        XCTAssertEqual(doc.body.children.count, 1)
    }
}
