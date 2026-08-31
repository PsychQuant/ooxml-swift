import XCTest
@testable import OOXMLSwift

/// Cross-document OMath splice tests.
///
/// Spec: openspec/changes/cross-document-omath-splice/specs/omath-splice/spec.md
/// Issue: PsychQuant/ooxml-swift#57
final class OMathSpliceTests: XCTestCase {

    // MARK: - Fixture builders

    /// Parse an inline `<w:p>` XML into a `Paragraph` via `DocxReader.parseParagraph`
    /// (same pattern as Issue99FlattenReplaceOMMLBilateralTests).
    private func parseParagraph(xml: String) throws -> Paragraph {
        let data = xml.data(using: .utf8)!
        let doc = try XMLDocument(data: data)
        guard let root = doc.rootElement() else {
            throw NSError(domain: "OMathSpliceTests", code: 1)
        }
        let document = WordDocument()
        return try DocxReader.parseParagraph(
            from: root,
            relationships: RelationshipsCollection(),
            styles: document.styles,
            numbering: document.numbering
        )
    }

    /// Wrap a paragraph as a single-paragraph WordDocument body for splice target.
    private func makeDocument(with paragraph: Paragraph) -> WordDocument {
        var doc = WordDocument()
        doc.body.children = [.paragraph(paragraph)]
        return doc
    }

    /// Wrap multiple paragraphs as body.
    private func makeDocument(with paragraphs: [Paragraph]) -> WordDocument {
        var doc = WordDocument()
        doc.body.children = paragraphs.map { .paragraph($0) }
        return doc
    }

    // MARK: - XML constants

    private static let mNS = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\""
    private static let mmlNS = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mml=\"http://schemas.openxmlformats.org/officeDocument/2006/math\""

    // MARK: - Source paragraph fixtures

    /// Source paragraph with one inline OMath in a Run: `所得出的參數進行 <m:oMath>t</m:oMath> 檢定:`
    private static let sourceInlineRunOMath = """
    <w:p \(mNS)>
      <w:r><w:t>所得出的參數進行 </w:t></w:r>
      <w:r><m:oMath><m:r><m:t>t</m:t></m:r></m:oMath></w:r>
      <w:r><w:t> 檢定：</w:t></w:r>
    </w:p>
    """

    /// Source paragraph with one direct-child OMath (Pandoc display math style):
    /// `<w:p><w:r>before</w:r><m:oMath>α</m:oMath><w:r>after</w:r></w:p>`
    private static let sourceDirectChildOMath = """
    <w:p \(mNS)>
      <w:r><w:t>before</w:t></w:r>
      <m:oMath><m:r><m:t>α</m:t></m:r></m:oMath>
      <w:r><w:t>after</w:t></w:r>
    </w:p>
    """

    /// Source paragraph with three inline OMath blocks at different positions for batch tests.
    private static let sourceMultipleOMath = """
    <w:p \(mNS)>
      <w:r><w:t>進行 </w:t></w:r>
      <w:r><m:oMath><m:r><m:t>t</m:t></m:r></m:oMath></w:r>
      <w:r><w:t> 檢定，係數 </w:t></w:r>
      <w:r><m:oMath><m:r><m:t>α</m:t></m:r></m:oMath></w:r>
      <w:r><w:t> 與 </w:t></w:r>
      <w:r><m:oMath><m:r><m:t>β</m:t></m:r></m:oMath></w:r>
      <w:r><w:t>。</w:t></w:r>
    </w:p>
    """

    /// Source paragraph with no OMath.
    private static let sourcePureText = """
    <w:p \(mNS)>
      <w:r><w:t>進行統計分析，計算各項參數。</w:t></w:r>
    </w:p>
    """

    /// Source paragraph using mml: prefix instead of m:.
    private static let sourceMMLPrefixOMath = """
    <w:p \(mmlNS)>
      <w:r><w:t>變數 </w:t></w:r>
      <w:r><mml:oMath><mml:r><mml:t>x</mml:t></mml:r></mml:oMath></w:r>
      <w:r><w:t> 是輸入。</w:t></w:r>
    </w:p>
    """

    // MARK: - Target paragraph fixtures

    /// Target paragraph that has corresponding prose anchors but no OMath
    /// (this is the rescue use case — text was preserved but inline math lost).
    private static let targetWithMatchingAnchors = """
    <w:p \(mNS)>
      <w:r><w:t>所得出的參數進行  檢定：</w:t></w:r>
    </w:p>
    """

    private static let targetEmpty = """
    <w:p \(mNS)>
      <w:r><w:t></w:t></w:r>
    </w:p>
    """

    // MARK: - Tests

    /// Test 6.2: Inline OMath spliced from source Run.rawXML to target paragraph end.
    /// Covers: Inline OMath spliced from source Run.rawXML to target paragraph end scenario.
    func testInlineRunRawXMLSpliceAtEnd() throws {
        let sourcePara = try parseParagraph(xml: Self.sourceInlineRunOMath)
        let targetPara = try parseParagraph(xml: Self.targetEmpty)
        var target = makeDocument(with: targetPara)

        let count = try target.spliceOMath(
            from: sourcePara,
            toBodyParagraphIndex: 0,
            position: .atEnd,
            omathIndex: 0
        )

        XCTAssertEqual(count, 1)
        guard case .paragraph(let resultPara) = target.body.children[0] else {
            XCTFail("Expected paragraph"); return
        }

        // Verify a Run with rawXML containing OMath was added.
        let omathRuns = resultPara.runs.filter {
            ($0.rawXML ?? "").contains("oMath")
        }
        XCTAssertEqual(omathRuns.count, 1, "Expected exactly one OMath run in target")

        // Verify the rawXML byte-equals (or substring-matches) the source OMath block.
        let spliced = omathRuns[0].rawXML ?? ""
        XCTAssertTrue(
            spliced.contains("<m:t>t</m:t>"),
            "Spliced OMath should preserve source OMath content; got: \(spliced)"
        )
    }

    /// Test 6.3: Direct-child OMath spliced preserving carrier.
    /// Covers: Direct-child OMath spliced preserving carrier scenario.
    func testDirectChildOMathSplicePreservesCarrier() throws {
        let sourcePara = try parseParagraph(xml: Self.sourceDirectChildOMath)
        let targetPara = try parseParagraph(xml: Self.targetEmpty)
        var target = makeDocument(with: targetPara)

        try target.spliceOMath(
            from: sourcePara,
            toBodyParagraphIndex: 0,
            position: .atEnd,
            omathIndex: 0
        )

        guard case .paragraph(let resultPara) = target.body.children[0] else {
            XCTFail("Expected paragraph"); return
        }

        // Direct-child OMath should appear in unrecognizedChildren, NOT in a Run.
        let directChildOMath = resultPara.unrecognizedChildren.filter {
            $0.name == "oMath" || $0.name == "oMathPara"
        }
        XCTAssertGreaterThanOrEqual(directChildOMath.count, 1,
            "Direct-child OMath should be added to unrecognizedChildren")

        // Verify no OMath was wrapped into a Run (carrier preservation).
        let runWithOMath = resultPara.runs.first { ($0.rawXML ?? "").contains("oMath") }
        XCTAssertNil(runWithOMath,
            "Direct-child source OMath should NOT be wrapped into a Run on target")
    }

    /// Test 6.4: Source paragraph has no OMath → throws .sourceHasNoOMath.
    func testSourceHasNoOMathThrows() throws {
        let sourcePara = try parseParagraph(xml: Self.sourcePureText)
        let targetPara = try parseParagraph(xml: Self.targetEmpty)
        var target = makeDocument(with: targetPara)

        XCTAssertThrowsError(
            try target.spliceOMath(from: sourcePara, toBodyParagraphIndex: 0, position: .atEnd, omathIndex: 0)
        ) { error in
            XCTAssertEqual(error as? OMathSpliceError, .sourceHasNoOMath)
        }
    }

    /// Test 6.5: omathIndex out of range → throws .omathIndexOutOfRange.
    func testOMathIndexOutOfRangeThrows() throws {
        let sourcePara = try parseParagraph(xml: Self.sourceInlineRunOMath)
        let targetPara = try parseParagraph(xml: Self.targetEmpty)
        var target = makeDocument(with: targetPara)

        XCTAssertThrowsError(
            try target.spliceOMath(from: sourcePara, toBodyParagraphIndex: 0, position: .atEnd, omathIndex: 5)
        ) { error in
            if case let .omathIndexOutOfRange(requested, available) = error as? OMathSpliceError {
                XCTAssertEqual(requested, 5)
                XCTAssertEqual(available, 1)
            } else {
                XCTFail("Expected .omathIndexOutOfRange, got: \(error)")
            }
        }
    }

    /// Test 6.6: Target paragraph index out of range → throws .targetParagraphOutOfRange.
    func testTargetParagraphOutOfRangeThrows() throws {
        let sourcePara = try parseParagraph(xml: Self.sourceInlineRunOMath)
        let targetPara = try parseParagraph(xml: Self.targetEmpty)
        var target = makeDocument(with: targetPara)

        XCTAssertThrowsError(
            try target.spliceOMath(from: sourcePara, toBodyParagraphIndex: 9999, position: .atEnd, omathIndex: 0)
        ) { error in
            if case let .targetParagraphOutOfRange(idx) = error as? OMathSpliceError {
                XCTAssertEqual(idx, 9999)
            } else {
                XCTFail("Expected .targetParagraphOutOfRange, got: \(error)")
            }
        }
    }

    /// Test 6.7: Mid-paragraph splice with anchor → run split into prefix/OMath/suffix.
    func testMidParagraphSpliceWithRunSplit() throws {
        let sourcePara = try parseParagraph(xml: Self.sourceInlineRunOMath)
        let targetPara = try parseParagraph(xml: Self.targetWithMatchingAnchors)
        var target = makeDocument(with: targetPara)

        try target.spliceOMath(
            from: sourcePara,
            toBodyParagraphIndex: 0,
            position: .afterText("所得出的參數進行 ", instance: 1),
            omathIndex: 0
        )

        guard case .paragraph(let resultPara) = target.body.children[0] else {
            XCTFail("Expected paragraph"); return
        }

        // After mid-paragraph splice, target paragraph should have:
        // [prefix run "所得出的參數進行 ", omath run, suffix run " 檢定："]
        // (At minimum: the prefix text appears before the OMath run, suffix after.)
        let runs = resultPara.runs
        let omathIdx = runs.firstIndex { ($0.rawXML ?? "").contains("oMath") }
        XCTAssertNotNil(omathIdx, "Expected an OMath run in result")
        guard let oi = omathIdx else { return }

        // Check that runs preceding the OMath run contain the prefix text.
        let prefixText = runs[0..<oi].map { $0.text }.joined()
        XCTAssertTrue(prefixText.contains("所得出的參數進行 "),
            "Expected prefix text before OMath; got prefixText='\(prefixText)'")

        // Check that runs following the OMath run contain the suffix text.
        let suffixText = runs[(oi + 1)...].map { $0.text }.joined()
        XCTAssertTrue(suffixText.contains("檢定"),
            "Expected suffix '檢定' after OMath; got suffixText='\(suffixText)'")
    }

    /// Test 6.8: Anchor not found → throws .anchorNotFound.
    func testAnchorNotFoundThrows() throws {
        let sourcePara = try parseParagraph(xml: Self.sourceInlineRunOMath)
        let targetPara = try parseParagraph(xml: Self.targetEmpty)
        var target = makeDocument(with: targetPara)

        XCTAssertThrowsError(
            try target.spliceOMath(
                from: sourcePara,
                toBodyParagraphIndex: 0,
                position: .afterText("nonexistent text", instance: 1),
                omathIndex: 0
            )
        ) { error in
            if case let .anchorNotFound(text, _) = error as? OMathSpliceError {
                XCTAssertEqual(text, "nonexistent text")
            } else {
                XCTFail("Expected .anchorNotFound, got: \(error)")
            }
        }
    }

    // MARK: - rPr propagation tests (6.9)

    /// .full mode copies all rPr fields verbatim.
    func testRpRModeFullCopiesVerbatim() throws {
        let sourceXML = """
        <w:p \(Self.mNS)>
          <w:r>
            <w:rPr><w:rFonts w:ascii="Cambria Math"/><w:sz w:val="24"/></w:rPr>
            <m:oMath><m:r><m:t>α</m:t></m:r></m:oMath>
          </w:r>
        </w:p>
        """
        let sourcePara = try parseParagraph(xml: sourceXML)
        let targetPara = try parseParagraph(xml: Self.targetEmpty)
        var target = makeDocument(with: targetPara)

        try target.spliceOMath(
            from: sourcePara,
            toBodyParagraphIndex: 0,
            position: .atEnd,
            omathIndex: 0,
            rPrMode: .full
        )

        guard case .paragraph(let resultPara) = target.body.children[0] else { XCTFail(); return }
        let omathRun = resultPara.runs.first { ($0.rawXML ?? "").contains("oMath") }
        XCTAssertNotNil(omathRun)
        XCTAssertEqual(omathRun?.properties.fontName, "Cambria Math",
            "fontName should propagate via rFonts.ascii in .full mode")
        XCTAssertEqual(omathRun?.properties.fontSize, 24,
            "fontSize 24 (12pt) should propagate in .full mode")
    }

    /// .discard mode resets to default rPr.
    func testRpRModeDiscardResetsToDefault() throws {
        let sourceXML = """
        <w:p \(Self.mNS)>
          <w:r>
            <w:rPr><w:rFonts w:ascii="Cambria Math"/><w:sz w:val="24"/></w:rPr>
            <m:oMath><m:r><m:t>α</m:t></m:r></m:oMath>
          </w:r>
        </w:p>
        """
        let sourcePara = try parseParagraph(xml: sourceXML)
        let targetPara = try parseParagraph(xml: Self.targetEmpty)
        var target = makeDocument(with: targetPara)

        try target.spliceOMath(
            from: sourcePara,
            toBodyParagraphIndex: 0,
            position: .atEnd,
            omathIndex: 0,
            rPrMode: .discard
        )

        guard case .paragraph(let resultPara) = target.body.children[0] else { XCTFail(); return }
        let omathRun = resultPara.runs.first { ($0.rawXML ?? "").contains("oMath") }
        XCTAssertNotNil(omathRun)
        XCTAssertNil(omathRun?.properties.fontName, ".discard should clear fontName")
        XCTAssertNil(omathRun?.properties.fontSize, ".discard should clear fontSize")
    }

    func testRpRModeOMathOnlyPreservesLanguageAndDropsStyle() throws {
        let sourceXML = """
        <w:p \(Self.mNS)>
          <w:r>
            <w:rPr>
              <w:rStyle w:val="SourceOnlyStyle"/>
              <w:rFonts w:ascii="Cambria Math"/>
              <w:sz w:val="24"/>
              <w:lang w:val="en-US" w:eastAsia="zh-TW"/>
            </w:rPr>
            <m:oMath><m:r><m:t>α</m:t></m:r></m:oMath>
          </w:r>
        </w:p>
        """
        let source = try parseParagraph(xml: sourceXML)
        var target = makeDocument(with: Paragraph(runs: [Run(text: "target")]))
        try target.spliceOMath(
            from: source,
            toBodyParagraphIndex: 0,
            position: .atEnd,
            rPrMode: .omathOnly
        )

        guard case .paragraph(let result) = target.body.children[0],
              let run = result.runs.first(where: { $0.rawXML?.contains("oMath") == true }) else {
            return XCTFail("Missing OMath Run")
        }
        XCTAssertEqual(run.properties.lang, LanguageProperties(val: "en-US", eastAsia: "zh-TW"))
        XCTAssertNil(run.properties.rStyle)
        XCTAssertEqual(run.properties.rFonts, RFontsProperties(ascii: "Cambria Math"))
        XCTAssertEqual(run.properties.fontName, "Cambria Math")
        XCTAssertEqual(run.properties.fontSize, 24)
        XCTAssertFalse(run.properties.strikethrough)
        XCTAssertNil(run.properties.color)
    }

    // MARK: - Namespace policy tests (6.10)

    /// .lenient (default) accepts prefix mismatch when URI is the same.
    func testNamespaceLenientAcceptsPrefixMismatch() throws {
        let sourcePara = try parseParagraph(xml: Self.sourceMMLPrefixOMath)  // mml: prefix
        let targetPara = try parseParagraph(xml: Self.targetEmpty)            // m: default
        var target = makeDocument(with: targetPara)

        // Should not throw — both use the standard OMML URI.
        XCTAssertNoThrow(
            try target.spliceOMath(
                from: sourcePara,
                toBodyParagraphIndex: 0,
                position: .atEnd,
                omathIndex: 0,
                namespacePolicy: .lenient
            )
        )
    }

    /// .strict rejects prefix mismatch even when URI is the same.
    func testNamespaceStrictRejectsPrefixMismatch() throws {
        let sourcePara = try parseParagraph(xml: Self.sourceMMLPrefixOMath)  // mml: prefix

        // Target needs to have an existing m: OMath so prefix can be detected.
        let targetWithOMath = """
        <w:p \(Self.mNS)>
          <w:r><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:r>
        </w:p>
        """
        let targetPara = try parseParagraph(xml: targetWithOMath)
        var target = makeDocument(with: targetPara)

        XCTAssertThrowsError(
            try target.spliceOMath(
                from: sourcePara,
                toBodyParagraphIndex: 0,
                position: .atEnd,
                omathIndex: 0,
                namespacePolicy: .strict
            )
        ) { error in
            guard case .namespaceMismatch = error as? OMathSpliceError else {
                XCTFail("Expected .namespaceMismatch, got: \(error)")
                return
            }
        }
    }

    // MARK: - Batch API tests (6.11)

    /// All OMath blocks spliced in source order via spliceParagraphOMath.
    /// Covers: All OMath blocks spliced in source order scenario.
    func testParagraphBatchSpliceAllOMath() throws {
        let sourcePara = try parseParagraph(xml: Self.sourceMultipleOMath)
        // Target with all 3 anchors but no OMath (the typical rescue scenario).
        let targetXML = """
        <w:p \(Self.mNS)>
          <w:r><w:t>進行  檢定，係數  與 。</w:t></w:r>
        </w:p>
        """
        let targetPara = try parseParagraph(xml: targetXML)
        var target = makeDocument(with: targetPara)

        let spliced = try target.spliceParagraphOMath(
            from: sourcePara,
            toBodyParagraphIndex: 0
        )

        XCTAssertEqual(spliced, 3, "Expected all 3 OMath blocks to be spliced")

        guard case .paragraph(let resultPara) = target.body.children[0] else { XCTFail(); return }
        let omathRuns = resultPara.runs.filter { ($0.rawXML ?? "").contains("oMath") }
        XCTAssertEqual(omathRuns.count, 3, "Expected 3 OMath runs in result")

        let xml = resultPara.toXML()
        let orderedNeedles = ["進行", "<m:t>t</m:t>", "檢定，係數", "<m:t>α</m:t>", "與", "<m:t>β</m:t>", "。"]
        let positions = try orderedNeedles.map { needle in
            try XCTUnwrap(xml.range(of: needle)?.lowerBound, "Missing \(needle): \(xml)")
        }
        for pair in zip(positions, positions.dropFirst()) {
            XCTAssertLessThan(pair.0, pair.1, "Batch equation ordering is wrong: \(xml)")
        }
    }

    // MARK: - Round-trip lossless guarantee (6.13)

    /// After splice + DocxWriter.write + DocxReader.read, the OMath XML in target
    /// preserves the source OMath content (substring match — ECMA-376 attribute
    /// reordering may shift exact byte sequence, but visible content is preserved).
    func testRoundTripPreservesOMathContent() throws {
        let sourcePara = try parseParagraph(xml: Self.sourceInlineRunOMath)
        let targetPara = try parseParagraph(xml: Self.targetEmpty)
        var target = makeDocument(with: targetPara)

        try target.spliceOMath(
            from: sourcePara,
            toBodyParagraphIndex: 0,
            position: .atEnd,
            omathIndex: 0
        )

        // Write to a temp file then reload.
        let tempURL = URL(fileURLWithPath: NSTemporaryDirectory())
            .appendingPathComponent("OMathSpliceTests-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: tempURL) }

        try DocxWriter.write(target, to: tempURL)
        let reloaded = try DocxReader.read(from: tempURL)

        // After round-trip, inline OMath that was stored as Run.rawXML on the write side
        // is re-parsed by DocxReader. The exact carrier on the read side depends on whether
        // DocxWriter emitted the OMath inside <w:r> or as direct child — current Run.toXML
        // behavior emits Run.rawXML verbatim (without <w:r> wrapper), so the OMath ends
        // up as direct-child in the round-tripped paragraph's unrecognizedChildren.
        // The round-trip lossless guarantee is at the **content** level: the OMath glyph
        // is preserved regardless of which carrier holds it.
        guard case .paragraph(let reloadedPara) = reloaded.body.children[0] else { XCTFail(); return }

        let runOMathContent = reloadedPara.runs
            .compactMap { $0.rawXML }
            .filter { $0.contains("oMath") }
            .joined()
        let directChildOMathContent = reloadedPara.unrecognizedChildren
            .filter { $0.name == "oMath" || $0.name == "oMathPara" }
            .map { $0.rawXML }
            .joined()
        let allOMathContent = runOMathContent + directChildOMathContent

        XCTAssertTrue(
            allOMathContent.contains("<m:t>t</m:t>") || allOMathContent.contains("<mml:t>t</mml:t>"),
            "Round-tripped OMath should contain original glyph regardless of carrier; got: \(allOMathContent)"
        )
    }

    // MARK: - Boundary and batch regressions (#122, #124, #125)

    func testAtEndSerializesAfterAPIBuiltRun() throws {
        let source = try parseParagraph(xml: Self.sourceInlineRunOMath)
        var target = makeDocument(with: Paragraph(runs: [Run(text: "target")]))

        try target.spliceOMath(from: source, toBodyParagraphIndex: 0, position: .atEnd)
        guard case .paragraph(let paragraph) = target.body.children[0] else { XCTFail(); return }
        let xml = paragraph.toXML()
        let textIndex = try XCTUnwrap(xml.range(of: "target")?.lowerBound)
        let mathIndex = try XCTUnwrap(xml.range(of: "<m:oMath")?.lowerBound)
        XCTAssertLessThan(textIndex, mathIndex, "atEnd must follow position-zero API-built text: \(xml)")
    }

    func testAtEndFollowsMixedPositiveAndPostContentCarriers() throws {
        let source = try parseParagraph(xml: Self.sourceInlineRunOMath)
        let mixedXML = """
        <w:p \(Self.mNS) xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:r><w:t>A</w:t></w:r>
          <m:oMath><m:r><m:t>β</m:t></m:r></m:oMath>
          <w:hyperlink r:id="rId9"><w:r><w:t>H</w:t></w:r></w:hyperlink>
        </w:p>
        """
        var targetParagraph = try parseParagraph(xml: mixedXML)
        targetParagraph.runs.append(Run(text: "Z"))
        var target = makeDocument(with: targetParagraph)

        try target.spliceOMath(from: source, toBodyParagraphIndex: 0, position: .atEnd)
        guard case .paragraph(let paragraph) = target.body.children[0] else { XCTFail(); return }
        let xml = paragraph.toXML()
        let orderedNeedles = [">A<", "<m:t>β</m:t>", ">H<", ">Z<", "<m:t>t</m:t>"]
        let positions = try orderedNeedles.map { needle in
            try XCTUnwrap(xml.range(of: needle)?.lowerBound, "Missing \(needle): \(xml)")
        }
        for pair in zip(positions, positions.dropFirst()) {
            XCTAssertLessThan(pair.0, pair.1, "Mixed carrier order changed: \(xml)")
        }
    }

    func testAtStartPrecedesMixedCarriers() throws {
        let source = try parseParagraph(xml: Self.sourceInlineRunOMath)
        var targetParagraph = try parseParagraph(xml: Self.sourceDirectChildOMath)
        targetParagraph.runs.append(Run(text: "tail"))
        var target = makeDocument(with: targetParagraph)

        try target.spliceOMath(from: source, toBodyParagraphIndex: 0, position: .atStart)
        guard case .paragraph(let paragraph) = target.body.children[0] else { XCTFail(); return }
        let xml = paragraph.toXML()
        let inserted = try XCTUnwrap(xml.range(of: "<m:t>t</m:t>")?.lowerBound)
        let firstExisting = try XCTUnwrap(xml.range(of: "before")?.lowerBound)
        XCTAssertLessThan(inserted, firstExisting, "atStart must precede every existing carrier: \(xml)")
    }

    func testDirectChildAtEndFollowsAPIBuiltRun() throws {
        let source = try parseParagraph(xml: Self.sourceDirectChildOMath)
        var target = makeDocument(with: Paragraph(runs: [Run(text: "target")]))

        try target.spliceOMath(from: source, toBodyParagraphIndex: 0, position: .atEnd)
        guard case .paragraph(let paragraph) = target.body.children[0] else { XCTFail(); return }
        let xml = paragraph.toXML()
        let textIndex = try XCTUnwrap(xml.range(of: "target")?.lowerBound)
        let mathIndex = try XCTUnwrap(xml.range(of: "<m:oMath")?.lowerBound)
        XCTAssertLessThan(textIndex, mathIndex, "Direct-child atEnd must follow API-built text: \(xml)")
    }

    func testDirectChildAtStartPrecedesPositiveAndAPIBuiltRuns() throws {
        let source = try parseParagraph(xml: Self.sourceDirectChildOMath)
        var targetParagraph = try parseParagraph(xml: "<w:p \(Self.mNS)><w:r><w:t>positive</w:t></w:r></w:p>")
        targetParagraph.runs.append(Run(text: "post"))
        var target = makeDocument(with: targetParagraph)

        try target.spliceOMath(from: source, toBodyParagraphIndex: 0, position: .atStart)
        guard case .paragraph(let paragraph) = target.body.children[0] else { XCTFail(); return }
        let xml = paragraph.toXML()
        let math = try XCTUnwrap(xml.range(of: "<m:oMath")?.lowerBound)
        let positive = try XCTUnwrap(xml.range(of: "positive")?.lowerBound)
        let post = try XCTUnwrap(xml.range(of: "post")?.lowerBound)
        XCTAssertLessThan(math, positive, "Direct-child atStart must precede positive content: \(xml)")
        XCTAssertLessThan(math, post, "Direct-child atStart must precede post-content: \(xml)")
    }

    func testDirectChildOMathParaPreservesNameThroughRoundTrip() throws {
        let sourceXML = """
        <w:p \(Self.mNS)>
          <m:oMathPara><m:oMath><m:r><m:t>γ</m:t></m:r></m:oMath></m:oMathPara>
        </w:p>
        """
        let source = try parseParagraph(xml: sourceXML)
        var target = makeDocument(with: Paragraph(runs: [Run(text: "target")]))
        try target.spliceOMath(from: source, toBodyParagraphIndex: 0, position: .atEnd)

        guard case .paragraph(let paragraph) = target.body.children[0] else { XCTFail(); return }
        let child = try XCTUnwrap(paragraph.unrecognizedChildren.last)
        XCTAssertEqual(child.name, "oMathPara")
        XCTAssertTrue(child.rawXML.hasPrefix("<m:oMathPara"))

        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("OMathSpliceTests-oMathPara-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try DocxWriter.write(target, to: url)
        let reloaded = try DocxReader.read(from: url)
        guard case .paragraph(let reloadedParagraph) = reloaded.body.children[0] else { XCTFail(); return }
        XCTAssertEqual(reloadedParagraph.unrecognizedChildren.last?.name, "oMathPara")
        let reextracted = OMathExtractor.extract(from: reloadedParagraph)
        XCTAssertEqual(reextracted.count, 1)
        XCTAssertEqual(reextracted[0].directChildName, "oMathPara")
        XCTAssertTrue(reextracted[0].xml.contains("<m:t>γ</m:t>"))

        var secondTarget = makeDocument(with: Paragraph(runs: [Run(text: "second")]))
        try secondTarget.spliceOMath(from: reloadedParagraph, toBodyParagraphIndex: 0, position: .atEnd)
        guard case .paragraph(let resplicedParagraph) = secondTarget.body.children[0] else { XCTFail(); return }
        XCTAssertEqual(resplicedParagraph.unrecognizedChildren.last?.name, "oMathPara")
        XCTAssertTrue(resplicedParagraph.unrecognizedChildren.last?.rawXML.contains("<m:t>γ</m:t>") == true)
    }

    func testBatchAnchorMatchesAfterNamespaceSelfContainment() throws {
        var mathRun = Run(text: "")
        mathRun.rawXML = "<mml:oMath><mml:r><mml:t>δ</mml:t></mml:r></mml:oMath>"
        let source = Paragraph(runs: [Run(text: "prefix "), mathRun, Run(text: " suffix")])
        var target = makeDocument(with: Paragraph(runs: [Run(text: "prefix  suffix")]))

        XCTAssertEqual(try target.spliceParagraphOMath(from: source, toBodyParagraphIndex: 0), 1)
        guard case .paragraph(let paragraph) = target.body.children[0] else { XCTFail(); return }
        let xml = paragraph.toXML()
        let prefix = try XCTUnwrap(xml.range(of: ">prefix<")?.lowerBound)
        let math = try XCTUnwrap(xml.range(of: "<mml:oMath")?.lowerBound)
        let suffix = try XCTUnwrap(xml.range(of: ">  suffix<")?.lowerBound)
        XCTAssertLessThan(prefix, math, "Batch math must follow its prefix: \(xml)")
        XCTAssertLessThan(math, suffix, "Batch math must precede its suffix: \(xml)")
    }

    func testBatchLeadingOMathMapsToAtStart() throws {
        var mathRun = Run(text: "")
        mathRun.rawXML = "<m:oMath><m:r><m:t>λ</m:t></m:r></m:oMath>"
        let source = Paragraph(runs: [mathRun, Run(text: " suffix")])
        var target = makeDocument(with: Paragraph(runs: [Run(text: " suffix")]))

        XCTAssertEqual(try target.spliceParagraphOMath(from: source, toBodyParagraphIndex: 0), 1)
        guard case .paragraph(let paragraph) = target.body.children[0] else { XCTFail(); return }
        let xml = paragraph.toXML()
        let math = try XCTUnwrap(xml.range(of: "<m:oMath")?.lowerBound)
        let suffix = try XCTUnwrap(xml.range(of: " suffix")?.lowerBound)
        XCTAssertLessThan(math, suffix, "Leading source math must map to target start: \(xml)")
    }

    func testBatchConsecutiveLeadingOMathPreservesSourceOrder() throws {
        var alpha = Run(text: "")
        alpha.rawXML = "<m:oMath><m:r><m:t>α</m:t></m:r></m:oMath>"
        var beta = Run(text: "")
        beta.rawXML = "<m:oMath><m:r><m:t>β</m:t></m:r></m:oMath>"
        let source = Paragraph(runs: [alpha, beta, Run(text: " suffix")])
        var target = makeDocument(with: Paragraph(runs: [Run(text: " suffix")]))

        XCTAssertEqual(try target.spliceParagraphOMath(from: source, toBodyParagraphIndex: 0), 2)
        guard case .paragraph(let paragraph) = target.body.children[0] else { XCTFail(); return }
        let xml = paragraph.toXML()
        let alphaIndex = try XCTUnwrap(xml.range(of: "<m:t>α</m:t>")?.lowerBound)
        let betaIndex = try XCTUnwrap(xml.range(of: "<m:t>β</m:t>")?.lowerBound)
        let suffixIndex = try XCTUnwrap(xml.range(of: " suffix")?.lowerBound)
        XCTAssertLessThan(alphaIndex, betaIndex, "Leading equations reversed: \(xml)")
        XCTAssertLessThan(betaIndex, suffixIndex, "Leading equations must precede suffix: \(xml)")
    }

    func testBatchConsecutiveOMathWithSharedNonemptyAnchorPreservesSourceOrder() throws {
        var alpha = Run(text: "")
        alpha.rawXML = "<m:oMath><m:r><m:t>α</m:t></m:r></m:oMath>"
        var beta = Run(text: "")
        beta.rawXML = "<m:oMath><m:r><m:t>β</m:t></m:r></m:oMath>"
        let source = Paragraph(runs: [Run(text: "prefix "), alpha, beta, Run(text: " suffix")])
        var target = makeDocument(with: Paragraph(runs: [Run(text: "prefix  suffix")]))

        XCTAssertEqual(try target.spliceParagraphOMath(from: source, toBodyParagraphIndex: 0), 2)
        guard case .paragraph(let paragraph) = target.body.children[0] else { XCTFail(); return }
        let xml = paragraph.toXML()
        let prefixIndex = try XCTUnwrap(xml.range(of: ">prefix<")?.lowerBound)
        let alphaIndex = try XCTUnwrap(xml.range(of: "<m:t>α</m:t>")?.lowerBound)
        let betaIndex = try XCTUnwrap(xml.range(of: "<m:t>β</m:t>")?.lowerBound)
        let suffixIndex = try XCTUnwrap(xml.range(of: ">  suffix<")?.lowerBound)
        XCTAssertLessThan(prefixIndex, alphaIndex, "Equations must follow their shared anchor: \(xml)")
        XCTAssertLessThan(alphaIndex, betaIndex, "Shared-anchor equations reversed: \(xml)")
        XCTAssertLessThan(betaIndex, suffixIndex, "Equations must precede suffix: \(xml)")
    }

    func testBatchUnmatchedDirectChildThrowsInsteadOfAppending() throws {
        var source = Paragraph()
        source.unrecognizedChildren = [
            UnrecognizedChild(
                name: "oMath",
                rawXML: "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r/></m:oMath>",
                position: nil
            )
        ]
        var target = makeDocument(with: Paragraph(runs: [Run(text: "target")]))

        XCTAssertThrowsError(
            try target.spliceParagraphOMath(from: source, toBodyParagraphIndex: 0)
        ) { error in
            guard case .contextAnchorNotFound(let index, _) = error as? OMathSpliceError else {
                return XCTFail("Expected contextAnchorNotFound, got \(error)")
            }
            XCTAssertEqual(index, 0)
        }
    }

    func testAtEndFollowsLegacyPostContentCarriers() throws {
        let source = try parseParagraph(xml: Self.sourceInlineRunOMath)
        var paragraph = Paragraph(runs: [Run(text: "body")])
        paragraph.footnoteIds = [7]
        paragraph.endnoteIds = [9]
        paragraph.bookmarks = [Bookmark(id: 11, name: "legacyBookmark")]
        var target = makeDocument(with: paragraph)

        try target.spliceOMath(from: source, toBodyParagraphIndex: 0, position: .atEnd)
        guard case .paragraph(let result) = target.body.children[0] else { XCTFail(); return }
        let xml = result.toXML()
        let orderedNeedles = ["body", "w:footnoteReference", "w:endnoteReference", "w:bookmarkEnd", "<m:oMath"]
        let positions = try orderedNeedles.map { needle in
            try XCTUnwrap(xml.range(of: needle)?.lowerBound, "Missing \(needle): \(xml)")
        }
        for pair in zip(positions, positions.dropFirst()) {
            XCTAssertLessThan(pair.0, pair.1, "atEnd did not follow every legacy carrier: \(xml)")
        }
    }

    func testAtStartPrecedesLegacyPreContentCarriers() throws {
        let source = try parseParagraph(xml: Self.sourceInlineRunOMath)
        var paragraph = Paragraph(runs: [Run(text: "body")])
        paragraph.hasPageBreak = true
        paragraph.bookmarks = [Bookmark(id: 11, name: "legacyBookmark")]
        paragraph.commentIds = [13]
        var target = makeDocument(with: paragraph)

        try target.spliceOMath(from: source, toBodyParagraphIndex: 0, position: .atStart)
        guard case .paragraph(let result) = target.body.children[0] else { XCTFail(); return }
        let xml = result.toXML()
        let math = try XCTUnwrap(xml.range(of: "<m:oMath")?.lowerBound)
        for needle in ["w:br w:type=\"page\"", "w:bookmarkStart", "w:commentRangeStart", "body"] {
            let existing = try XCTUnwrap(xml.range(of: needle)?.lowerBound, "Missing \(needle): \(xml)")
            XCTAssertLessThan(math, existing, "atStart did not precede \(needle): \(xml)")
        }
    }

    func testBoundaryInsertionCanonicalizesIntMaxPositionWithoutTrapping() throws {
        let source = try parseParagraph(xml: Self.sourceInlineRunOMath)

        var endRun = Run(text: "end-body")
        endRun.position = Int.max
        var endTarget = makeDocument(with: Paragraph(runs: [endRun]))
        try endTarget.spliceOMath(from: source, toBodyParagraphIndex: 0, position: .atEnd)
        guard case .paragraph(let endResult) = endTarget.body.children[0] else { XCTFail(); return }
        let endXML = endResult.toXML()
        XCTAssertLessThan(
            try XCTUnwrap(endXML.range(of: "end-body")?.lowerBound),
            try XCTUnwrap(endXML.range(of: "<m:oMath")?.lowerBound)
        )

        var startRun = Run(text: "start-body")
        startRun.position = Int.max
        var startTarget = makeDocument(with: Paragraph(runs: [startRun]))
        try startTarget.spliceOMath(from: source, toBodyParagraphIndex: 0, position: .atStart)
        guard case .paragraph(let startResult) = startTarget.body.children[0] else { XCTFail(); return }
        let startXML = startResult.toXML()
        XCTAssertLessThan(
            try XCTUnwrap(startXML.range(of: "<m:oMath")?.lowerBound),
            try XCTUnwrap(startXML.range(of: "start-body")?.lowerBound)
        )
    }

    func testDirectChildTextAnchorSplicesAtResolvedBoundary() throws {
        let source = try parseParagraph(xml: Self.sourceDirectChildOMath)
        let targetParagraph = try parseParagraph(xml: "<w:p \(Self.mNS)><w:r><w:t>before after</w:t></w:r></w:p>")
        var target = makeDocument(with: targetParagraph)

        try target.spliceOMath(
            from: source,
            toBodyParagraphIndex: 0,
            position: .afterText("before")
        )

        guard case .paragraph(let result) = target.body.children[0] else { XCTFail(); return }
        let xml = result.toXML()
        let before = try XCTUnwrap(xml.range(of: ">before<")?.lowerBound)
        let math = try XCTUnwrap(xml.range(of: "<m:oMath")?.lowerBound)
        let after = try XCTUnwrap(xml.range(of: "> after<")?.lowerBound)
        XCTAssertLessThan(before, math, "Direct child must follow resolved anchor: \(xml)")
        XCTAssertLessThan(math, after, "Direct child must precede suffix: \(xml)")
    }

    func testDirectChildMissingTextAnchorThrows() throws {
        let source = try parseParagraph(xml: Self.sourceDirectChildOMath)
        var target = makeDocument(with: Paragraph(runs: [Run(text: "target")]))

        XCTAssertThrowsError(
            try target.spliceOMath(
                from: source,
                toBodyParagraphIndex: 0,
                position: .afterText("missing")
            )
        ) { error in
            guard case .anchorNotFound("missing", 1) = error as? OMathSpliceError else {
                return XCTFail("Expected anchorNotFound, got \(error)")
            }
        }
    }

    func testDirectChildBeforeTextSplicesAtResolvedBoundary() throws {
        let source = try parseParagraph(xml: Self.sourceDirectChildOMath)
        var target = makeDocument(with: Paragraph(runs: [Run(text: "before after")]))

        try target.spliceOMath(
            from: source,
            toBodyParagraphIndex: 0,
            position: .beforeText("after")
        )

        guard case .paragraph(let result) = target.body.children[0] else { XCTFail(); return }
        let xml = result.toXML()
        let before = try XCTUnwrap(xml.range(of: ">before <")?.lowerBound)
        let math = try XCTUnwrap(xml.range(of: "<m:oMath")?.lowerBound)
        let after = try XCTUnwrap(xml.range(of: ">after<")?.lowerBound)
        XCTAssertLessThan(before, math, "Direct child must follow prefix: \(xml)")
        XCTAssertLessThan(math, after, "Direct child must precede beforeText anchor: \(xml)")
    }

    func testBatchDirectChildUsesDerivedTextAnchor() throws {
        let source = try parseParagraph(xml: Self.sourceDirectChildOMath)
        var target = makeDocument(with: Paragraph(runs: [Run(text: "before after")]))

        XCTAssertEqual(try target.spliceParagraphOMath(from: source, toBodyParagraphIndex: 0), 1)
        guard case .paragraph(let result) = target.body.children[0] else { XCTFail(); return }
        let xml = result.toXML()
        let before = try XCTUnwrap(xml.range(of: ">before<")?.lowerBound)
        let math = try XCTUnwrap(xml.range(of: "<m:oMath")?.lowerBound)
        let after = try XCTUnwrap(xml.range(of: "> after<")?.lowerBound)
        XCTAssertLessThan(before, math, "Batch direct child must follow derived anchor: \(xml)")
        XCTAssertLessThan(math, after, "Batch direct child must precede suffix: \(xml)")
    }

    func testBatchDirectChildMissingTargetAnchorFailsWithoutMutation() throws {
        let source = try parseParagraph(xml: Self.sourceDirectChildOMath)
        var target = makeDocument(with: Paragraph(runs: [Run(text: "target")]))

        XCTAssertThrowsError(
            try target.spliceParagraphOMath(from: source, toBodyParagraphIndex: 0)
        ) { error in
            guard case .contextAnchorNotFound(0, "before") = error as? OMathSpliceError else {
                return XCTFail("Expected contextAnchorNotFound, got \(error)")
            }
        }
        guard case .paragraph(let result) = target.body.children[0] else { XCTFail(); return }
        XCTAssertFalse(result.toXML().contains("<m:oMath"))
    }

    func testBatchRepeatedTextUsesSourceOccurrenceInstance() throws {
        var alpha = Run(text: "")
        alpha.rawXML = "<m:oMath><m:r><m:t>α</m:t></m:r></m:oMath>"
        var beta = Run(text: "")
        beta.rawXML = "<m:oMath><m:r><m:t>β</m:t></m:r></m:oMath>"
        let source = Paragraph(runs: [
            Run(text: "1234567890"), alpha,
            Run(text: " separator 1234567890"), beta,
            Run(text: " suffix"),
        ])
        var target = makeDocument(with: Paragraph(runs: [Run(text: "1234567890 separator 1234567890 suffix")]))

        XCTAssertEqual(try target.spliceParagraphOMath(from: source, toBodyParagraphIndex: 0), 2)
        guard case .paragraph(let result) = target.body.children[0] else { XCTFail(); return }
        let xml = result.toXML()
        let firstPrefix = try XCTUnwrap(xml.range(of: "1234567890")?.lowerBound)
        let alphaIndex = try XCTUnwrap(xml.range(of: "<m:t>α</m:t>")?.lowerBound)
        let secondPrefix = try XCTUnwrap(xml.range(of: "1234567890", range: alphaIndex..<xml.endIndex)?.lowerBound)
        let betaIndex = try XCTUnwrap(xml.range(of: "<m:t>β</m:t>")?.lowerBound)
        XCTAssertLessThan(firstPrefix, alphaIndex, "First equation must use first occurrence: \(xml)")
        XCTAssertLessThan(alphaIndex, secondPrefix, "First equation must not move to second occurrence: \(xml)")
        XCTAssertLessThan(secondPrefix, betaIndex, "Second equation must use second occurrence: \(xml)")
    }

    func testBatchFailurePreservesOnlyEarlierSourceGroups() throws {
        var alpha = Run(text: "")
        alpha.rawXML = "<m:oMath><m:r><m:t>α</m:t></m:r></m:oMath>"
        var beta = Run(text: "")
        beta.rawXML = "<m:oMath><m:r><m:t>β</m:t></m:r></m:oMath>"
        let source = Paragraph(runs: [
            Run(text: "good"), alpha,
            Run(text: " missing-anchor"), beta,
        ])
        var target = makeDocument(with: Paragraph(runs: [Run(text: "good target")]))

        XCTAssertThrowsError(
            try target.spliceParagraphOMath(from: source, toBodyParagraphIndex: 0)
        ) { error in
            guard case .contextAnchorNotFound(let index, _) = error as? OMathSpliceError else {
                return XCTFail("Expected contextAnchorNotFound, got \(error)")
            }
            XCTAssertEqual(index, 1)
        }

        guard case .paragraph(let result) = target.body.children[0] else { XCTFail(); return }
        let xml = result.toXML()
        XCTAssertTrue(xml.contains("<m:t>α</m:t>"), "Earlier source group must remain: \(xml)")
        XCTAssertFalse(xml.contains("<m:t>β</m:t>"), "Later failing group must not mutate target: \(xml)")
    }

    func testMathScriptInsensitiveAnchorOptionIsApplied() throws {
        let source = try parseParagraph(xml: Self.sourceInlineRunOMath)
        var target = makeDocument(with: Paragraph(runs: [Run(text: "H₀ result")]))

        try target.spliceOMath(
            from: source,
            toBodyParagraphIndex: 0,
            position: .afterText(
                "H0",
                options: AnchorLookupOptions(mathScriptInsensitive: true)
            )
        )

        guard case .paragraph(let result) = target.body.children[0] else { XCTFail(); return }
        let xml = result.toXML()
        XCTAssertLessThan(
            try XCTUnwrap(xml.range(of: "H₀")?.lowerBound),
            try XCTUnwrap(xml.range(of: "<m:oMath")?.lowerBound)
        )
    }

    func testSingleQuotedNamespaceMismatchIsRejected() throws {
        var sourceRun = Run(text: "")
        sourceRun.rawXML = "<m:oMath xmlns:m='https://example.invalid/vendor-math'><m:r/></m:oMath>"
        let source = Paragraph(runs: [sourceRun])
        var target = makeDocument(with: Paragraph(runs: [Run(text: "target")]))

        XCTAssertEqual(OMathNamespace.extractURI(from: sourceRun.rawXML!), "https://example.invalid/vendor-math")
        XCTAssertThrowsError(
            try target.spliceOMath(from: source, toBodyParagraphIndex: 0, position: .atEnd)
        ) { error in
            guard case .namespaceMismatch = error as? OMathSpliceError else {
                return XCTFail("Expected namespaceMismatch, got \(error)")
            }
        }
    }

    func testDefaultNamespaceOMathIsMadeSelfContained() {
        let normalized = OMathExtractor.ensureXmlnsDeclared(in: "<oMath><r/></oMath>")
        XCTAssertTrue(
            normalized.contains("xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/math\""),
            "Default-namespace OMath fragment must be self-contained: \(normalized)"
        )
    }

    func testDescendantNamespaceDeclarationDoesNotBindOMathRoot() {
        let normalized = OMathExtractor.ensureXmlnsDeclared(
            in: "<m:oMath><m:r xmlns:m='https://example.invalid/descendant'/></m:oMath>"
        )
        XCTAssertTrue(
            normalized.hasPrefix("<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\""),
            "Root must receive its own binding: \(normalized)"
        )
        XCTAssertEqual(
            OMathNamespace.extractURI(from: normalized),
            "http://schemas.openxmlformats.org/officeDocument/2006/math"
        )
    }

    func testBoundarySplicePreservesTypedLegacyStateAndNoteDeletion() throws {
        let source = try parseParagraph(xml: Self.sourceInlineRunOMath)
        var target = makeDocument(with: Paragraph(runs: [Run(text: "body")]))
        guard case .paragraph(var paragraph) = target.body.children[0] else { XCTFail(); return }
        paragraph.hasPageBreak = true
        target.body.children[0] = .paragraph(paragraph)
        let footnoteId = try target.insertFootnote(text: "foot", paragraphIndex: 0)
        let endnoteId = try target.insertEndnote(text: "end", paragraphIndex: 0)

        try target.spliceOMath(from: source, toBodyParagraphIndex: 0, position: .atEnd)
        guard case .paragraph(let spliced) = target.body.children[0] else { XCTFail(); return }
        XCTAssertTrue(spliced.hasPageBreak, "Boundary splice must not rewrite typed page-break state")
        XCTAssertEqual(spliced.footnoteIds, [footnoteId])
        XCTAssertEqual(spliced.endnoteIds, [endnoteId])

        try target.deleteFootnote(footnoteId: footnoteId)
        try target.deleteEndnote(endnoteId: endnoteId)
        guard case .paragraph(let deleted) = target.body.children[0] else { XCTFail(); return }
        let xml = deleted.toXML()
        XCTAssertFalse(xml.contains("w:footnoteReference"), "Deleted footnote must not leave a raw dangling reference: \(xml)")
        XCTAssertFalse(xml.contains("w:endnoteReference"), "Deleted endnote must not leave a raw dangling reference: \(xml)")
        XCTAssertTrue(deleted.hasPageBreak)
    }

    func testRootQNameTokenizerHandlesDefaultVendorFakeAndDottedNamespaces() throws {
        let standardURI = "http://schemas.openxmlformats.org/officeDocument/2006/math"
        let explicitDefault = "<oMath xmlns:m='urn:unused' xmlns='\(standardURI)'><r/></oMath>"
        XCTAssertNil(OMathNamespace.extractPrefix(from: explicitDefault))
        XCTAssertEqual(OMathNamespace.extractURI(from: explicitDefault), standardURI)
        XCTAssertEqual(OMathExtractor.ensureXmlnsDeclared(in: explicitDefault), explicitDefault)

        let vendorDefault = "<oMath xmlns='https://example.invalid/vendor'><r/></oMath>"
        XCTAssertEqual(OMathNamespace.extractURI(from: vendorDefault), "https://example.invalid/vendor")
        var vendorRun = Run(text: "")
        vendorRun.rawXML = vendorDefault
        var target = makeDocument(with: Paragraph(runs: [Run(text: "target")]))
        XCTAssertThrowsError(
            try target.spliceOMath(
                from: Paragraph(runs: [vendorRun]),
                toBodyParagraphIndex: 0,
                position: .atEnd
            )
        ) { error in
            guard case .namespaceMismatch = error as? OMathSpliceError else {
                return XCTFail("Expected namespaceMismatch, got \(error)")
            }
        }

        let fakeAttribute = "<m:oMath data-note=\"xmlns:m='https://example.invalid/fake'\"><m:r/></m:oMath>"
        let normalizedFake = OMathExtractor.ensureXmlnsDeclared(in: fakeAttribute)
        XCTAssertEqual(OMathNamespace.extractPrefix(from: normalizedFake), "m")
        XCTAssertEqual(OMathNamespace.extractURI(from: normalizedFake), standardURI)
        XCTAssertTrue(normalizedFake.contains("data-note=\"xmlns:m='https://example.invalid/fake'\""))

        let dotted = "  <m.math:oMath><m.math:r/></m.math:oMath>"
        let normalizedDotted = OMathExtractor.ensureXmlnsDeclared(in: dotted)
        XCTAssertEqual(OMathNamespace.extractPrefix(from: normalizedDotted), "m.math")
        XCTAssertEqual(OMathNamespace.extractURI(from: normalizedDotted), standardURI)
        XCTAssertTrue(normalizedDotted.contains("xmlns:m.math=\"\(standardURI)\""))
    }

    func testMathScriptInsensitiveAnchorPreservesCrossRunCombiningAndNonBMPSuffix() throws {
        let source = try parseParagraph(xml: Self.sourceInlineRunOMath)
        let combining = "\u{0304}"
        var target = makeDocument(with: Paragraph(runs: [
            Run(text: "X"),
            Run(text: combining),
            Run(text: "😀tail"),
        ]))

        try target.spliceOMath(
            from: source,
            toBodyParagraphIndex: 0,
            position: .afterText(
                "X",
                options: AnchorLookupOptions(mathScriptInsensitive: true)
            )
        )

        guard case .paragraph(let result) = target.body.children[0] else { XCTFail(); return }
        let mathIndex = try XCTUnwrap(result.runs.firstIndex { $0.rawXML?.contains("oMath") == true })
        XCTAssertEqual(result.runs[..<mathIndex].map(\.text).joined(), "X\(combining)")
        XCTAssertEqual(result.runs[(mathIndex + 1)...].map(\.text).joined(), "😀tail")
    }

    func testAllParagraphCarrierMatrixRetainsOrderAtBothBoundaries() throws {
        let allCarrierXML = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
          <w:r><w:t>RUN_SENTINEL</w:t></w:r>
          <w:hyperlink r:id="rId1"><w:r><w:t>HYPER_SENTINEL</w:t></w:r></w:hyperlink>
          <w:fldSimple w:instr="FIELD_SENTINEL"><w:r><w:t>FIELD_TEXT</w:t></w:r></w:fldSimple>
          <mc:AlternateContent><mc:Fallback><w:r><w:t>ALT_SENTINEL</w:t></w:r></mc:Fallback></mc:AlternateContent>
          <w:bookmarkStart w:id="1" w:name="positionedBookmark"/>
          <w:bookmarkEnd w:id="1"/>
          <w:commentRangeStart w:id="2"/>
          <w:commentRangeEnd w:id="2"/>
          <w:permStart w:id="perm-sentinel" w:edGrp="everyone"/>
          <w:permEnd w:id="perm-sentinel"/>
          <w:proofErr w:type="spellStart"/>
          <w:smartTag w:uri="urn:smart-sentinel" w:element="SMART_SENTINEL"><w:r><w:t>SMART_TEXT</w:t></w:r></w:smartTag>
          <w:customXml w:element="CUSTOM_SENTINEL"><w:r><w:t>CUSTOM_TEXT</w:t></w:r></w:customXml>
          <w:dir w:val="ltr"><w:r><w:t>BIDI_SENTINEL</w:t></w:r></w:dir>
          <w:unknownCarrier w:sentinel="UNKNOWN_SENTINEL"/>
          <w:sdt><w:sdtPr><w:id w:val="77"/><w:tag w:val="SDT_SENTINEL"/></w:sdtPr><w:sdtContent><w:r><w:t>SDT_TEXT</w:t></w:r></w:sdtContent></w:sdt>
        </w:p>
        """

        func makeTarget() throws -> WordDocument {
            var paragraph = try parseParagraph(xml: allCarrierXML)
            paragraph.hasPageBreak = true
            paragraph.bookmarks.append(Bookmark(id: 91, name: "legacyBookmark"))
            paragraph.commentIds = [88]
            paragraph.footnoteIds = [89]
            paragraph.endnoteIds = [90]
            var zeroRun = Run(text: "ZERO_SENTINEL")
            zeroRun.position = 0
            paragraph.runs.append(zeroRun)
            return makeDocument(with: paragraph)
        }

        let existingTokens = [
            "w:br w:type=\"page\"",
            "w:bookmarkStart w:id=\"91\"",
            "w:commentRangeStart w:id=\"88\"",
            "RUN_SENTINEL", "HYPER_SENTINEL", "FIELD_SENTINEL", "ALT_SENTINEL",
            "w:bookmarkStart w:id=\"1\"", "w:bookmarkEnd w:id=\"1\"",
            "w:commentRangeStart w:id=\"2\"", "w:commentRangeEnd w:id=\"2\"",
            "w:permStart", "w:permEnd", "w:proofErr",
            "SMART_SENTINEL", "CUSTOM_SENTINEL", "BIDI_SENTINEL",
            "UNKNOWN_SENTINEL", "SDT_SENTINEL", "ZERO_SENTINEL",
            "w:commentRangeEnd w:id=\"88\"",
            "w:footnoteReference w:id=\"89\"",
            "w:endnoteReference w:id=\"90\"",
            "w:bookmarkEnd w:id=\"91\"",
        ]

        let sources = [
            try parseParagraph(xml: Self.sourceInlineRunOMath),
            try parseParagraph(xml: Self.sourceDirectChildOMath),
        ]
        for source in sources {
            for position in [OMathSplicePosition.atStart, .atEnd] {
                var target = try makeTarget()
                try target.spliceOMath(from: source, toBodyParagraphIndex: 0, position: position)
                guard case .paragraph(let result) = target.body.children[0] else { XCTFail(); return }
                let xml = result.toXML()
                let tokenPositions = try existingTokens.map { token in
                    try XCTUnwrap(xml.range(of: token)?.lowerBound, "Missing \(token): \(xml)")
                }
                for pair in zip(tokenPositions, tokenPositions.dropFirst()) {
                    XCTAssertLessThan(pair.0, pair.1, "Existing carrier order changed: \(xml)")
                }
                let math = try XCTUnwrap(xml.range(of: "<m:oMath")?.lowerBound)
                if position == .atStart {
                    XCTAssertLessThan(math, tokenPositions[0], "atStart must precede full carrier matrix: \(xml)")
                } else {
                    XCTAssertLessThan(tokenPositions.last!, math, "atEnd must follow full carrier matrix: \(xml)")
                }
            }
        }
    }

    func testOMathReloadContractUsesSemanticXMLFingerprint() throws {
        let uri = "http://schemas.openxmlformats.org/officeDocument/2006/math"
        let lexicalA = "<m:oMath xmlns:m=\"\(uri)\" data-a=\"1\" data-b=\"2\"><m:r><m:ctrl m:val=\"on\"/><m:t>&#x3B1;</m:t></m:r></m:oMath>"
        let lexicalB = "<math:oMath data-b='2' xmlns:math='\(uri)' data-a='1'><math:r><math:ctrl math:val='on'/><math:t>α</math:t></math:r></math:oMath>"
        let different = "<math:oMath data-b='2' xmlns:math='\(uri)' data-a='1'><math:r><math:ctrl math:val='off'/><math:t>α</math:t></math:r></math:oMath>"

        XCTAssertNotEqual(lexicalA, lexicalB, "Fixture must exercise lexical differences")
        XCTAssertEqual(
            try OMathSemanticXML.canonicalRepresentation(of: lexicalA),
            try OMathSemanticXML.canonicalRepresentation(of: lexicalB),
            "Entity spelling, attribute order/quotes, namespace placement, and element/attribute prefixes are semantically equivalent"
        )
        XCTAssertNotEqual(
            try OMathSemanticXML.canonicalRepresentation(of: lexicalA),
            try OMathSemanticXML.canonicalRepresentation(of: different),
            "Different semantic attribute values must remain different"
        )

        let elementDifference = lexicalB.replacingOccurrences(of: "math:ctrl", with: "math:other")
        let childOrderDifference = "<math:oMath data-b='2' xmlns:math='\(uri)' data-a='1'><math:r><math:t>α</math:t><math:ctrl math:val='on'/></math:r></math:oMath>"
        let textDifference = lexicalB.replacingOccurrences(of: ">α<", with: ">β<")
        for candidate in [elementDifference, childOrderDifference, textDifference] {
            XCTAssertFalse(try OMathSemanticXML.isEquivalent(lexicalA, candidate))
        }

        let adjacentText = "<m:oMath xmlns:m='\(uri)'><m:r><m:t>α<![CDATA[β]]>γ</m:t></m:r></m:oMath>"
        let mergedText = "<m:oMath xmlns:m='\(uri)'><m:r><m:t>αβγ</m:t></m:r></m:oMath>"
        XCTAssertTrue(try OMathSemanticXML.isEquivalent(adjacentText, mergedText))

        XCTAssertFalse(try OMathSemanticXML.isEquivalent("<xml:oMath/>", "<oMath/>"))
        XCTAssertFalse(try OMathSemanticXML.isEquivalent(
            "<m:oMath xmlns:m='\(uri)' xmlns:p='unbound:p' p:flag='1'/>",
            "<m:oMath xmlns:m='\(uri)' p:flag='1'/>"
        ))

        let undeclaredDefault = "<oMath xmlns='\(uri)'><r><t xmlns=''>plain</t></r></oMath>"
        let neverDefaulted = "<m:oMath xmlns:m='\(uri)'><m:r><t>plain</t></m:r></m:oMath>"
        let mathNamespacedText = "<m:oMath xmlns:m='\(uri)'><m:r><m:t>plain</m:t></m:r></m:oMath>"
        XCTAssertTrue(
            try OMathSemanticXML.isEquivalent(undeclaredDefault, neverDefaulted),
            "xmlns='' must remove the inherited default namespace"
        )
        XCTAssertFalse(try OMathSemanticXML.isEquivalent(undeclaredDefault, mathNamespacedText))

        let literalTab = "<m:oMath xmlns:m='\(uri)' note='a\tb'/>"
        let normalizedSpace = "<m:oMath xmlns:m='\(uri)' note='a b'/>"
        let referencedTab = "<m:oMath xmlns:m='\(uri)' note='a&#x9;b'/>"
        XCTAssertTrue(
            try OMathSemanticXML.isEquivalent(literalTab, normalizedSpace),
            "XML processors normalize literal attribute whitespace to spaces"
        )
        XCTAssertFalse(
            try OMathSemanticXML.isEquivalent(literalTab, referencedTab),
            "A character-referenced tab remains a tab and must not collide with normalized literal whitespace"
        )
    }

    func testOMathSemanticContractSurvivesActualSpliceWriteReload() throws {
        let uri = "http://schemas.openxmlformats.org/officeDocument/2006/math"
        let sourceXML = "<m:oMath xmlns:m='\(uri)' data-b='2' data-a='1'><m:r><m:ctrl m:val='on'/><m:t>&#x3B1;</m:t></m:r></m:oMath>"
        var sourceRun = Run(text: "")
        sourceRun.rawXML = sourceXML
        let source = Paragraph(runs: [sourceRun])
        let expected = try XCTUnwrap(OMathExtractor.extract(from: source).first?.xml)
        var target = makeDocument(with: Paragraph(runs: [Run(text: "target")]))
        try target.spliceOMath(from: source, toBodyParagraphIndex: 0, position: .atEnd)

        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("OMathSpliceTests-semantic-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try DocxWriter.write(target, to: url)
        let reloaded = try DocxReader.read(from: url)
        guard case .paragraph(let paragraph) = reloaded.body.children[0] else { XCTFail(); return }
        let actual = try XCTUnwrap(OMathExtractor.extract(from: paragraph).first?.xml)
        XCTAssertTrue(try OMathSemanticXML.isEquivalent(expected, actual))
    }

    func testExtractorUsesAbsoluteBoundarySerializationOrder() throws {
        let alphaSource = try parseParagraph(xml: Self.sourceDirectChildOMath)
        let betaXML = "<w:p \(Self.mNS)><m:oMath><m:r><m:t>β</m:t></m:r></m:oMath></w:p>"
        let betaSource = try parseParagraph(xml: betaXML)
        var target = makeDocument(with: Paragraph(runs: [Run(text: "body")]))

        try target.spliceOMath(from: alphaSource, toBodyParagraphIndex: 0, position: .atStart)
        try target.spliceOMath(from: betaSource, toBodyParagraphIndex: 0, position: .atStart)
        guard case .paragraph(let paragraph) = target.body.children[0] else { XCTFail(); return }

        let extracted = OMathExtractor.extract(from: paragraph)
        XCTAssertEqual(extracted.count, 2)
        XCTAssertTrue(extracted[0].xml.contains("<m:t>β</m:t>"), "Extractor must follow boundary XML order")
        XCTAssertTrue(extracted[1].xml.contains("<m:t>α</m:t>"), "Extractor must follow boundary XML order")

        let xml = paragraph.toXML()
        XCTAssertLessThan(
            try XCTUnwrap(xml.range(of: "<m:t>β</m:t>")?.lowerBound),
            try XCTUnwrap(xml.range(of: "<m:t>α</m:t>")?.lowerBound)
        )
    }

    func testBoundaryDirectChildCanBeReusedAsBatchSourceBeforeReload() throws {
        let source = try parseParagraph(xml: Self.sourceDirectChildOMath)
        var intermediate = makeDocument(with: Paragraph(runs: [Run(text: " suffix")]))
        try intermediate.spliceOMath(from: source, toBodyParagraphIndex: 0, position: .atStart)
        guard case .paragraph(let boundarySource) = intermediate.body.children[0] else { XCTFail(); return }

        var target = makeDocument(with: Paragraph(runs: [Run(text: " suffix")]))
        XCTAssertEqual(
            try target.spliceParagraphOMath(from: boundarySource, toBodyParagraphIndex: 0),
            1
        )
        guard case .paragraph(let result) = target.body.children[0] else { XCTFail(); return }
        let xml = result.toXML()
        XCTAssertLessThan(
            try XCTUnwrap(xml.range(of: "<m:oMath")?.lowerBound),
            try XCTUnwrap(xml.range(of: " suffix")?.lowerBound)
        )
    }

    func testBoundaryEndDirectChildrenRemainAtEndAndInSourceOrderDuringBatchBeforeReload() throws {
        let alphaSource = try parseParagraph(xml: Self.sourceDirectChildOMath)
        let betaSource = try parseParagraph(
            xml: "<w:p \(Self.mNS)><m:oMath><m:r><m:t>β</m:t></m:r></m:oMath></w:p>"
        )
        var intermediate = makeDocument(with: Paragraph())
        try intermediate.spliceOMath(from: alphaSource, toBodyParagraphIndex: 0, position: .atEnd)
        try intermediate.spliceOMath(from: betaSource, toBodyParagraphIndex: 0, position: .atEnd)
        guard case .paragraph(let boundarySource) = intermediate.body.children[0] else {
            return XCTFail("Missing boundary source")
        }

        var target = makeDocument(with: Paragraph(runs: [Run(text: "BODY_SENTINEL")]))
        XCTAssertEqual(
            try target.spliceParagraphOMath(from: boundarySource, toBodyParagraphIndex: 0),
            2
        )
        guard case .paragraph(let result) = target.body.children[0] else { XCTFail(); return }
        let xml = result.toXML()
        let body = try XCTUnwrap(xml.range(of: "BODY_SENTINEL")?.lowerBound)
        let alpha = try XCTUnwrap(xml.range(of: "<m:t>α</m:t>")?.lowerBound)
        let beta = try XCTUnwrap(xml.range(of: "<m:t>β</m:t>")?.lowerBound)
        XCTAssertLessThan(body, alpha, "Absolute end-lane OMath must remain after prose")
        XCTAssertLessThan(alpha, beta, "Batch must retain source order within the absolute end lane")
    }

    func testNegativePositionCarrierSurvivesAbsoluteBoundarySplice() throws {
        let source = try parseParagraph(xml: Self.sourceInlineRunOMath)
        var negative = Run(text: "NEGATIVE_SENTINEL")
        negative.position = -7
        var target = makeDocument(with: Paragraph(runs: [negative]))

        try target.spliceOMath(from: source, toBodyParagraphIndex: 0, position: .atEnd)
        guard case .paragraph(let result) = target.body.children[0] else { XCTFail(); return }
        let xml = result.toXML()
        XCTAssertTrue(xml.contains("NEGATIVE_SENTINEL"))
        XCTAssertLessThan(
            try XCTUnwrap(xml.range(of: "NEGATIVE_SENTINEL")?.lowerBound),
            try XCTUnwrap(xml.range(of: "<m:oMath")?.lowerBound)
        )
        XCTAssertEqual(OMathExtractor.extract(from: result).count, 1)
    }

    func testNamespacePolicyDecodesEntitiesFailsClosedAndHandlesDefaultStrict() throws {
        let uri = "http://schemas.openxmlformats.org/officeDocument/2006/math"
        let entityURI = "http://schemas.openxmlformats.org/officeDocument/2006/&#x6D;ath"
        XCTAssertEqual(OMathNamespace.extractURI(from: "<m:oMath xmlns:m='\(entityURI)'/>"), uri)

        var malformedRun = Run(text: "")
        malformedRun.rawXML = "<m:oMath xmlns:m='\(uri)><m:r/></m:oMath>"
        var malformedTarget = makeDocument(with: Paragraph(runs: [Run(text: "target")]))
        XCTAssertThrowsError(
            try malformedTarget.spliceOMath(
                from: Paragraph(runs: [malformedRun]),
                toBodyParagraphIndex: 0,
                position: .atEnd
            )
        ) { error in
            guard case .malformedOMathXML = error as? OMathSpliceError else {
                return XCTFail("Expected malformedOMathXML, got \(error)")
            }
        }
        guard case .paragraph(let unchanged) = malformedTarget.body.children[0] else { XCTFail(); return }
        XCTAssertEqual(unchanged.text, "target")
        XCTAssertFalse(unchanged.toXML().contains("oMath"))

        var defaultSourceRun = Run(text: "")
        defaultSourceRun.rawXML = "<oMath xmlns='\(uri)'><r><t>δ</t></r></oMath>"
        var defaultTarget = Paragraph(runs: [Run(text: "target")])
        defaultTarget.unrecognizedChildren = [
            UnrecognizedChild(
                name: "oMath",
                rawXML: "<oMath xmlns='\(uri)'><r><t>existing</t></r></oMath>",
                position: 2
            )
        ]
        var defaultDocument = makeDocument(with: defaultTarget)
        XCTAssertNoThrow(try defaultDocument.spliceOMath(
            from: Paragraph(runs: [defaultSourceRun]),
            toBodyParagraphIndex: 0,
            position: .atEnd,
            namespacePolicy: .strict
        ))

        var mixedTarget = Paragraph()
        var firstRun = Run(text: "")
        firstRun.rawXML = "<m:oMath xmlns:m='\(uri)'><m:r/></m:oMath>"
        firstRun.position = 1
        mixedTarget.runs = [firstRun]
        mixedTarget.unrecognizedChildren = [
            UnrecognizedChild(
                name: "oMath",
                rawXML: "<mml:oMath xmlns:mml='\(uri)'><mml:r/></mml:oMath>",
                position: 2
            )
        ]
        var mixedDocument = makeDocument(with: mixedTarget)
        var mSource = Run(text: "")
        mSource.rawXML = "<m:oMath xmlns:m='\(uri)'><m:r/></m:oMath>"
        XCTAssertNoThrow(try mixedDocument.spliceOMath(
            from: Paragraph(runs: [mSource]),
            toBodyParagraphIndex: 0,
            position: .atEnd,
            namespacePolicy: .strict
        ))

        var directFirst = Paragraph()
        directFirst.unrecognizedChildren = [
            UnrecognizedChild(
                name: "oMath",
                rawXML: "<mml:oMath xmlns:mml='\(uri)'><mml:r/></mml:oMath>",
                position: 1
            )
        ]
        var laterRun = Run(text: "")
        laterRun.rawXML = "<m:oMath xmlns:m='\(uri)'><m:r/></m:oMath>"
        laterRun.position = 2
        directFirst.runs = [laterRun]
        var directFirstDocument = makeDocument(with: directFirst)
        XCTAssertThrowsError(try directFirstDocument.spliceOMath(
            from: Paragraph(runs: [mSource]),
            toBodyParagraphIndex: 0,
            position: .atEnd,
            namespacePolicy: .strict
        )) { error in
            guard case .namespaceMismatch = error as? OMathSpliceError else {
                return XCTFail("Expected first direct-child prefix mismatch, got \(error)")
            }
        }
    }

    func testNamespaceValidityGateRejectsInvalidXMLWithoutMutation() throws {
        let uri = "http://schemas.openxmlformats.org/officeDocument/2006/math"
        let invalidFragments = [
            "<m:oMath xmlns:m='\(uri)' xmlns:m='\(uri)'/>",
            "<m:oMath xmlns:m='\(uri)'><m:r p:flag='1'/></m:oMath>",
            "<:oMath xmlns='\(uri)'/>",
            "<m:oMath xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/&#X6D;ath'/>",
            "<m:oMath xmlns:m='\(uri)'><m:t>&#0;</m:t></m:oMath>",
            "<?xml version='1.0'?><m:oMath xmlns:m='\(uri)'/>",
            "<!DOCTYPE m:oMath><m:oMath xmlns:m='\(uri)'/>",
            "<!--document framing--><m:oMath xmlns:m='\(uri)'/>",
            "framing text<m:oMath xmlns:m='\(uri)'/>",
            "\u{FEFF}<m:oMath xmlns:m='\(uri)'/>",
            "<m:oMath xmlns:m='\(uri)'/>\u{FEFF}",
            "\u{00A0}<m:oMath xmlns:m='\(uri)'/>",
            "<![CDATA[]]><m:oMath xmlns:m='\(uri)'/>",
            "<m:oMath xmlns:m='\(uri)'/><![CDATA[]]>",
            "<m:oMath xmlns:m='\(uri)'/><m:oMath xmlns:m='\(uri)'/>",
            "<wrapper><m:oMath xmlns:m='\(uri)'/></wrapper>",
        ]

        for fragment in invalidFragments {
            var run = Run(text: "")
            run.rawXML = fragment
            var target = makeDocument(with: Paragraph(runs: [Run(text: "UNCHANGED")]))
            XCTAssertThrowsError(
                try target.spliceOMath(
                    from: Paragraph(runs: [run]),
                    toBodyParagraphIndex: 0,
                    position: .atEnd
                ),
                "Expected malformed rejection for \(fragment)"
            ) { error in
                guard case .malformedOMathXML = error as? OMathSpliceError else {
                    return XCTFail("Expected malformedOMathXML, got \(error) for \(fragment)")
                }
            }
            guard case .paragraph(let unchanged) = target.body.children[0] else { XCTFail(); return }
            XCTAssertEqual(unchanged.text, "UNCHANGED")
            XCTAssertFalse(unchanged.toXML().contains("oMath"))
        }
    }

    func testNamespaceValidityGateAllowsDoctypeTextInsidePayloadCommentAndCDATA() throws {
        let uri = "http://schemas.openxmlformats.org/officeDocument/2006/math"
        var run = Run(text: "")
        run.rawXML = "<m:oMath xmlns:m='\(uri)'><!-- literal <!DOCTYPE text --><m:r><m:t><![CDATA[<!DOCTYPE text]]></m:t></m:r></m:oMath>"
        var target = makeDocument(with: Paragraph(runs: [Run(text: "UNCHANGED")]))

        XCTAssertNoThrow(try target.spliceOMath(
            from: Paragraph(runs: [run]),
            toBodyParagraphIndex: 0,
            position: .atEnd
        ))
        guard case .paragraph(let result) = target.body.children[0] else { XCTFail(); return }
        XCTAssertTrue(result.toXML().contains("<!DOCTYPE text"))
    }

    func testNamespaceValidityGateAllowsOnlyXMLSOutsidePayloadRoot() throws {
        let uri = "http://schemas.openxmlformats.org/officeDocument/2006/math"
        let framings = [
            (" \t\n", "\r "),
            ("\r\n", ""),
            ("", "\r\n"),
            ("\r\n", "\r\n"),
            (" \t\r\n", "\r\n "),
        ]

        for (leading, trailing) in framings {
            var run = Run(text: "")
            run.rawXML = "\(leading)<m:oMath xmlns:m='\(uri)'/>\(trailing)"
            var target = makeDocument(with: Paragraph(runs: [Run(text: "UNCHANGED")]))

            XCTAssertNoThrow(try target.spliceOMath(
                from: Paragraph(runs: [run]),
                toBodyParagraphIndex: 0,
                position: .atEnd
            ), "Expected XML S framing to remain legal: \(leading.debugDescription) / \(trailing.debugDescription)")
            guard case .paragraph(let result) = target.body.children[0] else { XCTFail(); return }
            XCTAssertTrue(result.toXML().contains("<m:oMath"))
        }
    }

    // MARK: - No regression (6.14)

    /// Pre-existing OMath in target paragraph must be preserved during splice.
    func testNoRegressionOnExistingOMathInTarget() throws {
        let sourcePara = try parseParagraph(xml: Self.sourceInlineRunOMath)  // contains 't'
        // Target already has α and β.
        let targetWithExistingOMath = """
        <w:p \(Self.mNS)>
          <w:r><w:t>變數 </w:t></w:r>
          <w:r><m:oMath><m:r><m:t>α</m:t></m:r></m:oMath></w:r>
          <w:r><w:t> 與 </w:t></w:r>
          <w:r><m:oMath><m:r><m:t>β</m:t></m:r></m:oMath></w:r>
          <w:r><w:t>。</w:t></w:r>
        </w:p>
        """
        let targetPara = try parseParagraph(xml: targetWithExistingOMath)
        var target = makeDocument(with: targetPara)

        try target.spliceOMath(
            from: sourcePara,
            toBodyParagraphIndex: 0,
            position: .atEnd,
            omathIndex: 0
        )

        guard case .paragraph(let resultPara) = target.body.children[0] else { XCTFail(); return }
        let omathRuns = resultPara.runs.filter { ($0.rawXML ?? "").contains("oMath") }

        // Should have α + β (original) + t (spliced) = 3 OMath runs total.
        XCTAssertEqual(omathRuns.count, 3,
            "Expected 3 OMath runs (α, β preserved + t spliced); got \(omathRuns.count)")

        // Verify each glyph is present.
        let allOMathContent = omathRuns.compactMap { $0.rawXML }.joined()
        XCTAssertTrue(allOMathContent.contains("<m:t>α</m:t>"), "Expected α preserved")
        XCTAssertTrue(allOMathContent.contains("<m:t>β</m:t>"), "Expected β preserved")
        XCTAssertTrue(allOMathContent.contains("<m:t>t</m:t>"), "Expected t spliced")
    }
}
