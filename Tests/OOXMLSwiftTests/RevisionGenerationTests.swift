import XCTest
@testable import OOXMLSwift

/// Unit tests for `WordDocument` revision-generating mutations and the
/// supporting `Revision.toXML` writer. Spec coverage:
///   openspec/changes/che-word-mcp-track-changes-programmatic-generation/specs/
///     ooxml-document-part-mutations/spec.md
///
/// Scaffold landed by task 1.1. Body-level tests fill in across tasks
/// 2.x (Revision.toXML verification) and 3.x (WordDocument generators).
/// Until each task lands, the scenario tests XCTSkip so the suite stays green.
final class RevisionGenerationTests: XCTestCase {

    // MARK: - Builders

    /// Build a `WordDocument` with `n` paragraphs (each containing one Run with
    /// the matching text from `texts`), optionally enabling track changes with
    /// the given author at construction time.
    func makeDoc(texts: [String], trackChangesAuthor: String? = nil) -> WordDocument {
        var doc = WordDocument()
        for text in texts {
            let run = Run(text: text)
            let paragraph = Paragraph(runs: [run])
            doc.body.children.append(.paragraph(paragraph))
        }
        if let author = trackChangesAuthor {
            doc.enableTrackChanges(author: author)
        }
        return doc
    }

    /// Convenience: returns the `Paragraph` at body index `i` (skips non-paragraph
    /// children — fails the test if no paragraph at that count).
    func paragraph(_ doc: WordDocument, at i: Int) -> Paragraph {
        let paragraphs = doc.body.children.compactMap { child -> Paragraph? in
            if case .paragraph(let p) = child { return p }
            return nil
        }
        XCTAssertLessThan(i, paragraphs.count, "paragraph index \(i) out of range")
        return paragraphs[i]
    }

    // MARK: - Writer XML Assertion Helpers

    /// Assert the XML string contains a `<w:ins ...>` opening tag with the
    /// given id and author attributes. Date is checked for ISO 8601 prefix
    /// only (year-month-day) since the exact timestamp is not deterministic.
    func assertContainsInsTag(_ xml: String, id: Int, author: String,
                              file: StaticString = #file, line: UInt = #line) {
        let pattern = "<w:ins w:id=\"\(id)\" w:author=\"\(author)\""
        XCTAssertTrue(xml.contains(pattern),
                      "expected <w:ins> tag with id=\(id) author=\(author); xml=\(xml)",
                      file: file, line: line)
    }

    func assertContainsDelTag(_ xml: String, id: Int, author: String,
                              file: StaticString = #file, line: UInt = #line) {
        let pattern = "<w:del w:id=\"\(id)\" w:author=\"\(author)\""
        XCTAssertTrue(xml.contains(pattern),
                      "expected <w:del> tag with id=\(id) author=\(author); xml=\(xml)",
                      file: file, line: line)
    }

    func assertContainsMoveFromTag(_ xml: String, id: Int, author: String,
                                   file: StaticString = #file, line: UInt = #line) {
        let pattern = "<w:moveFrom w:id=\"\(id)\" w:author=\"\(author)\""
        XCTAssertTrue(xml.contains(pattern),
                      "expected <w:moveFrom> tag with id=\(id) author=\(author); xml=\(xml)",
                      file: file, line: line)
    }

    func assertContainsMoveToTag(_ xml: String, id: Int, author: String,
                                 file: StaticString = #file, line: UInt = #line) {
        let pattern = "<w:moveTo w:id=\"\(id)\" w:author=\"\(author)\""
        XCTAssertTrue(xml.contains(pattern),
                      "expected <w:moveTo> tag with id=\(id) author=\(author); xml=\(xml)",
                      file: file, line: line)
    }

    func assertContainsRPrChangeTag(_ xml: String, id: Int, author: String,
                                    file: StaticString = #file, line: UInt = #line) {
        let pattern = "<w:rPrChange w:id=\"\(id)\" w:author=\"\(author)\""
        XCTAssertTrue(xml.contains(pattern),
                      "expected <w:rPrChange> tag with id=\(id) author=\(author); xml=\(xml)",
                      file: file, line: line)
    }

    func assertContainsPPrChangeTag(_ xml: String, id: Int, author: String,
                                    file: StaticString = #file, line: UInt = #line) {
        let pattern = "<w:pPrChange w:id=\"\(id)\" w:author=\"\(author)\""
        XCTAssertTrue(xml.contains(pattern),
                      "expected <w:pPrChange> tag with id=\(id) author=\(author); xml=\(xml)",
                      file: file, line: line)
    }

    // MARK: - Round-Trip Helpers

    /// Save document to a temporary .docx, re-read it via DocxReader, and
    /// return the parsed `WordDocument`. Test ensures temp file is cleaned up.
    func roundTrip(_ doc: WordDocument) throws -> WordDocument {
        let tempURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("revision-rt-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: tempURL) }
        try DocxWriter.write(doc, to: tempURL)
        return try DocxReader.read(from: tempURL)
    }

    // MARK: - Smoke Tests (always pass; verify scaffold compiles)

    func testScaffoldCompiles() {
        let doc = makeDoc(texts: ["Hello", "World"], trackChangesAuthor: "Tester")
        XCTAssertTrue(doc.isTrackChangesEnabled())
        XCTAssertEqual(doc.body.children.count, 2)
    }

    // MARK: - Task 2.1: Revision.toXML multi-run wrapping

    /// When a Paragraph has 3 runs all linked to the same insertion-type Revision
    /// (via `Run.revisionId`), `Paragraph.toXML()` SHALL emit exactly one
    /// `<w:ins>...</w:ins>` pair containing all 3 `<w:r>` siblings.
    func testRevisionToXMLWrapsMultipleRunsInSingleInsTag() throws {
        var paragraph = Paragraph()
        let revision = Revision(id: 7, type: .insertion, author: "Alice")
        paragraph.revisions = [revision]

        var run1 = Run(text: "Hello ")
        run1.revisionId = 7
        var run2 = Run(text: "brave ")
        run2.revisionId = 7
        var run3 = Run(text: "world")
        run3.revisionId = 7
        paragraph.runs = [run1, run2, run3]

        let xml = paragraph.toXML()

        let openingCount = xml.components(separatedBy: "<w:ins ").count - 1
        let closingCount = xml.components(separatedBy: "</w:ins>").count - 1
        XCTAssertEqual(openingCount, 1,
            "expected exactly one <w:ins> opening tag, got \(openingCount); xml=\(xml)")
        XCTAssertEqual(closingCount, 1,
            "expected exactly one </w:ins> closing tag, got \(closingCount); xml=\(xml)")

        let runCount = xml.components(separatedBy: "<w:r>").count - 1
        XCTAssertEqual(runCount, 3,
            "expected 3 <w:r> children inside the ins wrapper, got \(runCount); xml=\(xml)")

        XCTAssertTrue(xml.contains("<w:ins w:id=\"7\" w:author=\"Alice\""),
            "expected <w:ins> with id=7 author=Alice; xml=\(xml)")
    }

    // MARK: - Task 2.2: <w:t> → <w:delText> substitution

    /// When a Run is linked to a deletion-type Revision, `Paragraph.toXML()`
    /// SHALL emit `<w:delText>` instead of `<w:t>` inside the `<w:del>` wrapper.
    func testRevisionToXMLSubstitutesDelTextOnDeletion() throws {
        var paragraph = Paragraph()
        let revision = Revision(id: 3, type: .deletion, author: "Bob")
        paragraph.revisions = [revision]

        var run = Run(text: "World")
        run.revisionId = 3
        paragraph.runs = [run]

        let xml = paragraph.toXML()

        XCTAssertTrue(xml.contains("<w:del w:id=\"3\" w:author=\"Bob\""),
            "expected <w:del> wrapper with id=3 author=Bob; xml=\(xml)")
        XCTAssertTrue(xml.contains("<w:delText xml:space=\"preserve\">World</w:delText>"),
            "expected <w:delText> wrapping the deleted text; xml=\(xml)")
        XCTAssertFalse(xml.contains("<w:t xml:space=\"preserve\">World</w:t>"),
            "expected <w:t> NOT to appear for deleted text; xml=\(xml)")
        XCTAssertTrue(xml.contains("</w:del>"),
            "expected </w:del> closing tag; xml=\(xml)")
    }

    // MARK: - Task 3.1: WordDocument.allocateRevisionId

    func testAllocateRevisionIdReturns1OnEmptyDocument() throws {
        let doc = WordDocument()
        XCTAssertEqual(doc.allocateRevisionId(), 1)
    }

    func testAllocateRevisionIdReturnsMaxPlus1WithExistingRevisions() throws {
        var doc = WordDocument()
        doc.revisions.revisions.append(Revision(id: 1, type: .insertion, author: "A"))
        doc.revisions.revisions.append(Revision(id: 7, type: .deletion, author: "B"))
        doc.revisions.revisions.append(Revision(id: 3, type: .insertion, author: "C"))
        XCTAssertEqual(doc.allocateRevisionId(), 8)
    }

    func testAllocateRevisionIdIsIdempotentAcrossCallsWithoutAppend() throws {
        var doc = WordDocument()
        doc.revisions.revisions.append(Revision(id: 5, type: .insertion, author: "A"))
        XCTAssertEqual(doc.allocateRevisionId(), 6)
        XCTAssertEqual(doc.allocateRevisionId(), 6,
            "consecutive calls without appending should return the same id")
    }

    // MARK: - Task 3.3: WordDocument.insertTextAsRevision

    func testInsertTextAsRevisionAppendsRevisionAndMarksDirty() throws {
        var doc = makeDoc(texts: ["Hello"], trackChangesAuthor: "Reviewer A")
        let id = try doc.insertTextAsRevision(text: " World", atParagraph: 0,
                                              position: 5, author: nil, date: nil)
        XCTAssertEqual(id, 1)

        let p = paragraph(doc, at: 0)
        XCTAssertEqual(p.runs.count, 2)
        XCTAssertEqual(p.runs.map(\.text).joined(), "Hello World")
        XCTAssertEqual(p.runs[1].text, " World")
        XCTAssertEqual(p.runs[1].revisionId, 1)

        XCTAssertEqual(p.revisions.count, 1)
        XCTAssertEqual(p.revisions[0].author, "Reviewer A")
        XCTAssertEqual(p.revisions[0].type, .insertion)
        XCTAssertEqual(p.revisions[0].id, 1)

        XCTAssertTrue(doc.modifiedParts.contains("word/document.xml"))
    }

    func testInsertTextAsRevisionAtPositionZeroPrepends() throws {
        var doc = makeDoc(texts: ["World"], trackChangesAuthor: "A")
        _ = try doc.insertTextAsRevision(text: "Hello ", atParagraph: 0,
                                         position: 0, author: nil, date: nil)
        let p = paragraph(doc, at: 0)
        XCTAssertEqual(p.runs.count, 2)
        XCTAssertEqual(p.runs[0].text, "Hello ")
        XCTAssertEqual(p.runs[0].revisionId, 1)
        XCTAssertEqual(p.runs[1].text, "World")
        XCTAssertNil(p.runs[1].revisionId)
    }

    func testInsertTextAsRevisionInMiddleSplitsRun() throws {
        // Single run "HelloWorld" — position 5 splits into ["Hello", "<new>", "World"]
        var doc = makeDoc(texts: ["HelloWorld"], trackChangesAuthor: "A")
        _ = try doc.insertTextAsRevision(text: " brave ", atParagraph: 0,
                                         position: 5, author: nil, date: nil)
        let p = paragraph(doc, at: 0)
        XCTAssertEqual(p.runs.count, 3)
        XCTAssertEqual(p.runs[0].text, "Hello")
        XCTAssertNil(p.runs[0].revisionId)
        XCTAssertEqual(p.runs[1].text, " brave ")
        XCTAssertEqual(p.runs[1].revisionId, 1)
        XCTAssertEqual(p.runs[2].text, "World")
        XCTAssertNil(p.runs[2].revisionId)
    }

    func testInsertTextAsRevisionExplicitAuthorOverridesSettings() throws {
        var doc = makeDoc(texts: ["X"], trackChangesAuthor: "Default")
        _ = try doc.insertTextAsRevision(text: "Y", atParagraph: 0, position: 1,
                                         author: "Override", date: nil)
        XCTAssertEqual(paragraph(doc, at: 0).revisions[0].author, "Override")
    }

    func testInsertTextAsRevisionResolvesAuthorFallback() throws {
        // No explicit author + no track-changes settings author → "Unknown"
        var doc = makeDoc(texts: ["X"])
        doc.enableTrackChanges()  // default author = "Unknown"
        _ = try doc.insertTextAsRevision(text: "Y", atParagraph: 0, position: 1,
                                         author: nil, date: nil)
        XCTAssertEqual(paragraph(doc, at: 0).revisions[0].author, "Unknown")
    }

    func testInsertTextAsRevisionThrowsWhenTrackChangesOff() throws {
        var doc = makeDoc(texts: ["Hello"])
        XCTAssertThrowsError(try doc.insertTextAsRevision(
            text: " World", atParagraph: 0, position: 5, author: nil, date: nil
        )) { error in
            guard let we = error as? WordError, case .trackChangesNotEnabled = we else {
                XCTFail("expected WordError.trackChangesNotEnabled, got \(error)")
                return
            }
        }
        // Document state must be unchanged.
        XCTAssertEqual(paragraph(doc, at: 0).runs.count, 1)
        XCTAssertEqual(paragraph(doc, at: 0).runs[0].text, "Hello")
    }

    func testInsertTextAsRevisionThrowsOnOutOfBoundsParagraphIndex() throws {
        var doc = makeDoc(texts: ["Hello"], trackChangesAuthor: "A")
        XCTAssertThrowsError(try doc.insertTextAsRevision(
            text: " World", atParagraph: 99, position: 0, author: nil, date: nil
        )) { error in
            guard let we = error as? WordError, case .invalidIndex = we else {
                XCTFail("expected WordError.invalidIndex, got \(error)")
                return
            }
        }
    }

    func testInsertTextAsRevisionThrowsOnOutOfBoundsPosition() throws {
        var doc = makeDoc(texts: ["Hello"], trackChangesAuthor: "A")
        XCTAssertThrowsError(try doc.insertTextAsRevision(
            text: "X", atParagraph: 0, position: 999, author: nil, date: nil
        )) { error in
            guard let we = error as? WordError, case .invalidIndex = we else {
                XCTFail("expected WordError.invalidIndex, got \(error)")
                return
            }
        }
    }

    // MARK: - Task 3.4: WordDocument.deleteTextAsRevision

    func testDeleteTextAsRevisionMarksRunsAndAppendsRevision() throws {
        var doc = makeDoc(texts: ["Hello World"], trackChangesAuthor: "Bob")
        let id = try doc.deleteTextAsRevision(atParagraph: 0, start: 6, end: 11,
                                              author: nil, date: nil)
        XCTAssertEqual(id, 1)

        let p = paragraph(doc, at: 0)
        // Original "Hello World" splits into ["Hello ", "World"]; "World" gets revisionId.
        XCTAssertEqual(p.runs.count, 2)
        XCTAssertEqual(p.runs[0].text, "Hello ")
        XCTAssertNil(p.runs[0].revisionId)
        XCTAssertEqual(p.runs[1].text, "World")
        XCTAssertEqual(p.runs[1].revisionId, 1)

        XCTAssertEqual(p.revisions.count, 1)
        XCTAssertEqual(p.revisions[0].type, .deletion)
        XCTAssertEqual(p.revisions[0].author, "Bob")
        XCTAssertEqual(p.revisions[0].originalText, "World")

        XCTAssertTrue(doc.modifiedParts.contains("word/document.xml"))
    }

    func testDeleteTextAsRevisionEmitsDelWithDelText() throws {
        var doc = makeDoc(texts: ["Hello World"], trackChangesAuthor: "Bob")
        _ = try doc.deleteTextAsRevision(atParagraph: 0, start: 6, end: 11,
                                         author: nil, date: nil)
        let xml = paragraph(doc, at: 0).toXML()

        XCTAssertTrue(xml.contains("<w:del w:id=\"1\" w:author=\"Bob\""),
                      "expected <w:del> wrapper; xml=\(xml)")
        XCTAssertTrue(xml.contains("<w:delText xml:space=\"preserve\">World</w:delText>"),
                      "expected <w:delText> for deleted run; xml=\(xml)")
        XCTAssertFalse(xml.contains("<w:t xml:space=\"preserve\">World</w:t>"),
                       "expected <w:t> NOT used for deleted run; xml=\(xml)")
        XCTAssertTrue(xml.contains("</w:del>"))
    }

    func testDeleteTextAsRevisionSplitsRunWhenRangeStartsMidRun() throws {
        // "HelloWorld" (single run); delete chars 5..10 → ["Hello", "World"(deleted)]
        var doc = makeDoc(texts: ["HelloWorld"], trackChangesAuthor: "A")
        _ = try doc.deleteTextAsRevision(atParagraph: 0, start: 5, end: 10,
                                         author: nil, date: nil)
        let p = paragraph(doc, at: 0)
        XCTAssertEqual(p.runs.count, 2)
        XCTAssertEqual(p.runs[0].text, "Hello")
        XCTAssertNil(p.runs[0].revisionId)
        XCTAssertEqual(p.runs[1].text, "World")
        XCTAssertEqual(p.runs[1].revisionId, 1)
    }

    func testDeleteTextAsRevisionSplitsBothEndsWhenRangeIsInterior() throws {
        // "HelloWorld" (single run); delete chars 2..7 → ["He", "lloWo"(deleted), "rld"]
        var doc = makeDoc(texts: ["HelloWorld"], trackChangesAuthor: "A")
        _ = try doc.deleteTextAsRevision(atParagraph: 0, start: 2, end: 7,
                                         author: nil, date: nil)
        let p = paragraph(doc, at: 0)
        XCTAssertEqual(p.runs.count, 3)
        XCTAssertEqual(p.runs[0].text, "He")
        XCTAssertNil(p.runs[0].revisionId)
        XCTAssertEqual(p.runs[1].text, "lloWo")
        XCTAssertEqual(p.runs[1].revisionId, 1)
        XCTAssertEqual(p.runs[2].text, "rld")
        XCTAssertNil(p.runs[2].revisionId)
        XCTAssertEqual(p.revisions[0].originalText, "lloWo")
    }

    func testDeleteTextAsRevisionThrowsWhenTrackChangesOff() throws {
        var doc = makeDoc(texts: ["Hello"])
        XCTAssertThrowsError(try doc.deleteTextAsRevision(
            atParagraph: 0, start: 0, end: 5, author: nil, date: nil
        )) { error in
            guard let we = error as? WordError, case .trackChangesNotEnabled = we else {
                XCTFail("expected WordError.trackChangesNotEnabled, got \(error)")
                return
            }
        }
    }

    func testDeleteTextAsRevisionThrowsOnOutOfBoundsRange() throws {
        var doc = makeDoc(texts: ["Hello"], trackChangesAuthor: "A")
        XCTAssertThrowsError(try doc.deleteTextAsRevision(
            atParagraph: 0, start: 0, end: 999, author: nil, date: nil
        )) { error in
            guard let we = error as? WordError, case .invalidIndex = we else {
                XCTFail("expected WordError.invalidIndex, got \(error)")
                return
            }
        }
    }

    // MARK: - Task 3.5: WordDocument.moveTextAsRevision

    func testMoveTextAsRevisionAllocatesAdjacentIds() throws {
        var doc = makeDoc(texts: ["Hello World", "Greetings"], trackChangesAuthor: "A")
        let result = try doc.moveTextAsRevision(
            fromParagraph: 0, fromStart: 6, fromEnd: 11,
            toParagraph: 1, toPosition: 0,
            author: nil, date: nil
        )
        XCTAssertEqual(result.fromId, 1)
        XCTAssertEqual(result.toId, 2)

        let p0 = paragraph(doc, at: 0)
        XCTAssertEqual(p0.revisions.count, 1)
        XCTAssertEqual(p0.revisions[0].type, .moveFrom)
        XCTAssertEqual(p0.revisions[0].id, 1)

        let p1 = paragraph(doc, at: 1)
        XCTAssertEqual(p1.revisions.count, 1)
        XCTAssertEqual(p1.revisions[0].type, .moveTo)
        XCTAssertEqual(p1.revisions[0].id, 2)
    }

    func testMoveTextAsRevisionMarksSourceRunsAndInsertsAtDestination() throws {
        var doc = makeDoc(texts: ["Hello World", "Greetings"], trackChangesAuthor: "A")
        _ = try doc.moveTextAsRevision(
            fromParagraph: 0, fromStart: 6, fromEnd: 11,
            toParagraph: 1, toPosition: 0,
            author: nil, date: nil
        )

        let p0 = paragraph(doc, at: 0)
        // Source paragraph keeps the moved text but marked as moveFrom.
        XCTAssertEqual(p0.runs.map(\.text).joined(), "Hello World")
        let movedRun = p0.runs.first { $0.revisionId == 1 }
        XCTAssertNotNil(movedRun)
        XCTAssertEqual(movedRun?.text, "World")

        let p1 = paragraph(doc, at: 1)
        // Destination paragraph gains the moved text at position 0.
        XCTAssertEqual(p1.runs.first?.text, "World")
        XCTAssertEqual(p1.runs.first?.revisionId, 2)
    }

    func testMoveTextAsRevisionThrowsWhenTrackChangesOff() throws {
        var doc = makeDoc(texts: ["Hello", "World"])
        XCTAssertThrowsError(try doc.moveTextAsRevision(
            fromParagraph: 0, fromStart: 0, fromEnd: 5,
            toParagraph: 1, toPosition: 0,
            author: nil, date: nil
        )) { error in
            guard let we = error as? WordError, case .trackChangesNotEnabled = we else {
                XCTFail("expected trackChangesNotEnabled, got \(error)")
                return
            }
        }
    }

    func testMoveTextAsRevisionRejectsSameParagraphMove() throws {
        // Single-paragraph move is out of scope for v0.18.0 — the run-shift
        // arithmetic when from and to overlap in the same paragraph is fragile;
        // reject early so callers don't get silent corruption.
        var doc = makeDoc(texts: ["Hello World"], trackChangesAuthor: "A")
        XCTAssertThrowsError(try doc.moveTextAsRevision(
            fromParagraph: 0, fromStart: 6, fromEnd: 11,
            toParagraph: 0, toPosition: 0,
            author: nil, date: nil
        )) { error in
            guard let we = error as? WordError, case .invalidParameter = we else {
                XCTFail("expected invalidParameter, got \(error)")
                return
            }
        }
    }

    // MARK: - Task 3.6: WordDocument.applyRunPropertiesAsRevision

    func testApplyRunPropertiesAsRevisionCapturesPreviousFormat() throws {
        var doc = makeDoc(texts: ["important"], trackChangesAuthor: "Alice")
        // Initial properties are default (not bold). Apply bold as revision.
        var newProps = RunProperties()
        newProps.bold = true
        let id = try doc.applyRunPropertiesAsRevision(
            atParagraph: 0, atRunIndex: 0, newProperties: newProps,
            author: nil, date: nil
        )
        XCTAssertEqual(id, 1)

        let p = paragraph(doc, at: 0)
        XCTAssertTrue(p.runs[0].properties.bold)
        XCTAssertEqual(p.runs[0].formatChangeRevisionId, 1)

        XCTAssertEqual(p.revisions.count, 1)
        XCTAssertEqual(p.revisions[0].type, .formatChange)
        XCTAssertEqual(p.revisions[0].author, "Alice")
        XCTAssertNotNil(p.revisions[0].previousFormat)
        XCTAssertFalse(p.revisions[0].previousFormat?.bold ?? true,
                       "previousFormat must capture the pre-mutation (non-bold) state")

        XCTAssertTrue(doc.modifiedParts.contains("word/document.xml"))
    }

    func testApplyRunPropertiesAsRevisionEmitsRPrChange() throws {
        var doc = makeDoc(texts: ["important"], trackChangesAuthor: "Alice")
        var newProps = RunProperties()
        newProps.bold = true
        _ = try doc.applyRunPropertiesAsRevision(
            atParagraph: 0, atRunIndex: 0, newProperties: newProps,
            author: nil, date: nil
        )
        let xml = paragraph(doc, at: 0).toXML()

        XCTAssertTrue(xml.contains("<w:rPrChange w:id=\"1\" w:author=\"Alice\""),
                      "expected <w:rPrChange>; xml=\(xml)")
        XCTAssertTrue(xml.contains("</w:rPrChange>"),
                      "expected </w:rPrChange>; xml=\(xml)")
        // Current run is bold; previous (inside rPrChange) must NOT have <w:b/>.
        // Crude check: count <w:b/> occurrences — should be exactly 1 (the current).
        let boldCount = xml.components(separatedBy: "<w:b/>").count - 1
        XCTAssertEqual(boldCount, 1, "expected exactly one <w:b/>; xml=\(xml)")
    }

    func testApplyRunPropertiesAsRevisionThrowsWhenTrackChangesOff() throws {
        var doc = makeDoc(texts: ["X"])
        XCTAssertThrowsError(try doc.applyRunPropertiesAsRevision(
            atParagraph: 0, atRunIndex: 0, newProperties: RunProperties(),
            author: nil, date: nil
        )) { error in
            guard let we = error as? WordError, case .trackChangesNotEnabled = we else {
                XCTFail("expected trackChangesNotEnabled, got \(error)")
                return
            }
        }
    }

    func testApplyRunPropertiesAsRevisionThrowsOnInvalidRunIndex() throws {
        var doc = makeDoc(texts: ["X"], trackChangesAuthor: "A")
        XCTAssertThrowsError(try doc.applyRunPropertiesAsRevision(
            atParagraph: 0, atRunIndex: 99, newProperties: RunProperties(),
            author: nil, date: nil
        )) { error in
            guard let we = error as? WordError, case .invalidIndex = we else {
                XCTFail("expected invalidIndex, got \(error)")
                return
            }
        }
    }

    // MARK: - Task 3.7: WordDocument.applyParagraphPropertiesAsRevision

    func testApplyParagraphPropertiesAsRevisionCapturesPreviousFormat() throws {
        var doc = makeDoc(texts: ["Hello"], trackChangesAuthor: "Alice")
        var newProps = ParagraphProperties()
        newProps.alignment = .center
        let id = try doc.applyParagraphPropertiesAsRevision(
            atParagraph: 0, newProperties: newProps, author: nil, date: nil
        )
        XCTAssertEqual(id, 1)

        let p = paragraph(doc, at: 0)
        XCTAssertEqual(p.properties.alignment, .center)
        XCTAssertEqual(p.paragraphFormatChangeRevisionId, 1)
        XCTAssertNotNil(p.previousProperties)
        XCTAssertNil(p.previousProperties?.alignment,
                     "previousProperties must capture the pre-mutation (no-alignment) state")

        XCTAssertEqual(p.revisions.count, 1)
        XCTAssertEqual(p.revisions[0].type, .paragraphChange)

        XCTAssertTrue(doc.modifiedParts.contains("word/document.xml"))
    }

    func testApplyParagraphPropertiesAsRevisionEmitsPPrChange() throws {
        var doc = makeDoc(texts: ["Hello"], trackChangesAuthor: "Alice")
        var newProps = ParagraphProperties()
        newProps.alignment = .center
        _ = try doc.applyParagraphPropertiesAsRevision(
            atParagraph: 0, newProperties: newProps, author: nil, date: nil
        )
        let xml = paragraph(doc, at: 0).toXML()

        XCTAssertTrue(xml.contains("<w:pPrChange w:id=\"1\" w:author=\"Alice\""),
                      "expected <w:pPrChange>; xml=\(xml)")
        XCTAssertTrue(xml.contains("</w:pPrChange>"))
        // Current alignment is center; previous (inside pPrChange) lacks <w:jc>.
        XCTAssertTrue(xml.contains("<w:jc w:val=\"center\""),
                      "expected current center alignment; xml=\(xml)")
    }

    func testApplyParagraphPropertiesAsRevisionThrowsWhenTrackChangesOff() throws {
        var doc = makeDoc(texts: ["X"])
        XCTAssertThrowsError(try doc.applyParagraphPropertiesAsRevision(
            atParagraph: 0, newProperties: ParagraphProperties(), author: nil, date: nil
        )) { error in
            guard let we = error as? WordError, case .trackChangesNotEnabled = we else {
                XCTFail("expected trackChangesNotEnabled, got \(error)")
                return
            }
        }
    }
}
