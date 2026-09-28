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

    /// PsychQuant/ooxml-swift#117: a spliced-into Run carries its OMath payload
    /// as a verbatim `Run.rawElements` child (so `Run.toXML()` still emits
    /// `<w:rPr>` before it), not the legacy `Run.rawXML` full-run-replacement
    /// channel that a document loaded straight off disk uses. Both are valid
    /// "this Run is OMath" signals; tests should not care which one a given
    /// Run happens to use, so every OMath-detection assertion in this file
    /// goes through this helper instead of inspecting `.rawXML` directly.
    private func omathPayload(of run: Run) -> String? {
        if let raw = run.rawXML, raw.contains("oMath") { return raw }
        if let elements = run.rawElements {
            for element in elements where element.name == "oMath" || element.name == "oMathPara" {
                return element.xml
            }
        }
        return nil
    }

    private func isOMathRun(_ run: Run) -> Bool {
        omathPayload(of: run) != nil
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

    /// Source paragraph with one direct-child `oMathPara` (display-math wrapper,
    /// distinct from bare `oMath`). Used by PsychQuant/ooxml-swift#124.
    private static let sourceDirectChildOMathPara = """
    <w:p \(mNS)>
      <w:r><w:t>before</w:t></w:r>
      <m:oMathPara><m:oMath><m:r><m:t>γ</m:t></m:r></m:oMath></m:oMathPara>
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

    /// Test 6.2: Inline OMath spliced from source into target paragraph end.
    /// Covers: Inline OMath spliced from source Run to target paragraph end scenario.
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

        // Verify a Run carrying an OMath payload was added (PsychQuant/ooxml-swift#117:
        // the carrier is `rawElements`, not `rawXML` — see `isOMathRun`).
        let omathRuns = resultPara.runs.filter { isOMathRun($0) }
        XCTAssertEqual(omathRuns.count, 1, "Expected exactly one OMath run in target")

        // Verify the payload preserves the source OMath content.
        let spliced = omathPayload(of: omathRuns[0]) ?? ""
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
        let runWithOMath = resultPara.runs.first { isOMathRun($0) }
        XCTAssertNil(runWithOMath,
            "Direct-child source OMath should NOT be wrapped into a Run on target")
    }

    /// PsychQuant/ooxml-swift#124: a direct-child `oMathPara` must round-trip as
    /// `oMathPara`, not be hardcoded to `oMath` — typed metadata (`name`) and the
    /// verbatim payload (`rawXML`, whose root is literally `<m:oMathPara>`) must agree.
    func testDirectChildOMathParaPreservesElementName() throws {
        let sourcePara = try parseParagraph(xml: Self.sourceDirectChildOMathPara)
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

        let spliced = resultPara.unrecognizedChildren.filter { $0.rawXML.contains("γ") }
        XCTAssertEqual(spliced.count, 1, "Expected exactly one spliced oMathPara child")
        XCTAssertEqual(spliced.first?.name, "oMathPara",
            "Typed name must agree with the payload root <m:oMathPara>, not hardcode 'oMath'")
        XCTAssertTrue(spliced.first?.rawXML.contains("<m:oMathPara") ?? false)
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
        let omathIdx = runs.firstIndex { isOMathRun($0) }
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
        let omathRun = resultPara.runs.first { isOMathRun($0) }
        XCTAssertNotNil(omathRun)
        XCTAssertEqual(omathRun?.properties.fontName, "Cambria Math",
            "fontName should propagate via rFonts.ascii in .full mode")
        XCTAssertEqual(omathRun?.properties.fontSize, 24,
            "fontSize 24 (12pt) should propagate in .full mode")
    }

    /// PsychQuant/ooxml-swift#117: `.full` must not just set `Run.properties`
    /// in memory — the copied rPr must actually be emitted by `Run.toXML()`,
    /// *before* the OMath body, inside the same `<w:r>` element. Pre-fix,
    /// `Run.toXML()` returned the OMath carrier immediately and never reached
    /// the `<w:rPr>` emission code at all.
    func testRpRModeFullEmitsCopiedRPrInSerializedXML() throws {
        let sourceXML = """
        <w:p \(Self.mNS)>
          <w:r>
            <w:rPr><w:rFonts w:ascii="Cambria Math"/><w:sz w:val="24"/><w:b/></w:rPr>
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
        guard let omathRun = resultPara.runs.first(where: { isOMathRun($0) }) else {
            XCTFail("Expected an OMath run"); return
        }

        let runXML = omathRun.toXML()
        guard let rPrRange = runXML.range(of: "<w:rPr>"),
              let omathRange = runXML.range(of: "<m:oMath") else {
            XCTFail("Expected both <w:rPr> and <m:oMath> in serialized run XML; got: \(runXML)")
            return
        }
        XCTAssertTrue(rPrRange.lowerBound < omathRange.lowerBound,
            "<w:rPr> must be serialized before <m:oMath> within the same <w:r>; got: \(runXML)")
        XCTAssertTrue(runXML.contains("Cambria Math"), "Expected rFonts to survive serialization; got: \(runXML)")
        XCTAssertTrue(runXML.contains("w:val=\"24\""), "Expected sz to survive serialization; got: \(runXML)")
        XCTAssertTrue(runXML.contains("<w:b/>"), "Expected explicit bold to survive serialization; got: \(runXML)")
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
        let omathRun = resultPara.runs.first { isOMathRun($0) }
        XCTAssertNotNil(omathRun)
        XCTAssertNil(omathRun?.properties.fontName, ".discard should clear fontName")
        XCTAssertNil(omathRun?.properties.fontSize, ".discard should clear fontSize")

        // PsychQuant/ooxml-swift#117: also assert the *serialized* XML — the
        // run must still exist (with the OMath body) but with no copied rPr
        // fields leaking through.
        let runXML = omathRun?.toXML() ?? ""
        XCTAssertTrue(runXML.contains("<m:oMath"), "Expected OMath body still serialized; got: \(runXML)")
        XCTAssertFalse(runXML.contains("Cambria Math"), ".discard must not leak source rFonts into serialized XML")
        XCTAssertFalse(runXML.contains("w:val=\"24\""), ".discard must not leak source sz into serialized XML")
    }

    /// PsychQuant/ooxml-swift#117: `.omathOnly` emits only the documented
    /// whitelist (rFonts / sz / szCs / lang / specified bold / specified
    /// italic) and excludes non-whitelisted fields (rStyle / color), while
    /// still preserving an explicit-off bold as `<w:b w:val="0"/>` rather than
    /// silently dropping it (PsychQuant/ooxml-swift#115 tri-state contract).
    func testRpRModeOmathOnlyEmitsWhitelistOnly() throws {
        let sourceXML = """
        <w:p \(Self.mNS)>
          <w:r>
            <w:rPr>
              <w:rStyle w:val="MathChar"/>
              <w:rFonts w:ascii="Cambria Math"/>
              <w:b w:val="0"/>
              <w:color w:val="FF0000"/>
              <w:sz w:val="24"/>
            </w:rPr>
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
            rPrMode: .omathOnly
        )

        guard case .paragraph(let resultPara) = target.body.children[0] else { XCTFail(); return }
        guard let omathRun = resultPara.runs.first(where: { isOMathRun($0) }) else {
            XCTFail("Expected an OMath run"); return
        }

        XCTAssertEqual(omathRun.properties.fontName, "Cambria Math")
        XCTAssertEqual(omathRun.properties.fontSize, 24)
        XCTAssertEqual(omathRun.properties.specifiedBold, false,
            "Explicit-off bold must survive .omathOnly as false, not absent")
        XCTAssertNil(omathRun.properties.rStyle, "rStyle is not in the .omathOnly whitelist")
        XCTAssertNil(omathRun.properties.color, "color is not in the .omathOnly whitelist")

        let runXML = omathRun.toXML()
        XCTAssertTrue(runXML.contains("Cambria Math"))
        XCTAssertTrue(runXML.contains("w:val=\"24\""))
        XCTAssertTrue(runXML.contains("<w:b w:val=\"0\"/>"),
            "Explicit-off bold must serialize as <w:b w:val=\"0\"/>, not be omitted; got: \(runXML)")
        XCTAssertFalse(runXML.contains("MathChar"), "rStyle must not leak through .omathOnly; got: \(runXML)")
        XCTAssertFalse(runXML.contains("FF0000"), "color must not leak through .omathOnly; got: \(runXML)")
    }

    /// PsychQuant/ooxml-swift#117: rPr correctness must survive a *later*
    /// dirty resave — i.e. some unrelated mutation forces the whole paragraph
    /// to regenerate via `Paragraph.toXML()` (not just the freshly-spliced
    /// `Run.toXML()` in isolation) before the document is saved.
    func testRpRModeFullSurvivesLaterDirtyResave() throws {
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

        // Unrelated mutation: append a plain text Run to the same paragraph,
        // forcing a full paragraph regeneration on next `toXML()`/save.
        guard case .paragraph(var resultPara) = target.body.children[0] else { XCTFail(); return }
        resultPara.runs.append(Run(text: "後續文字"))
        target.body.children[0] = .paragraph(resultPara)

        let tempURL = URL(fileURLWithPath: NSTemporaryDirectory())
            .appendingPathComponent("OMathSpliceTests-dirty-resave-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: tempURL) }

        try DocxWriter.write(target, to: tempURL)
        let reloaded = try DocxReader.read(from: tempURL)

        guard case .paragraph(let reloadedPara) = reloaded.body.children[0] else { XCTFail(); return }
        guard let reloadedOMathRun = reloadedPara.runs.first(where: { isOMathRun($0) }) else {
            XCTFail("Expected OMath run to survive dirty resave + reload"); return
        }
        XCTAssertEqual(reloadedOMathRun.properties.fontName, "Cambria Math",
            "rPr must survive a later dirty resave, not just the immediate in-memory splice")
        XCTAssertEqual(reloadedOMathRun.properties.fontSize, 24)
        XCTAssertTrue(reloadedPara.runs.contains { $0.text == "後續文字" })
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
        let omathRuns = resultPara.runs.filter { isOMathRun($0) }
        XCTAssertEqual(omathRuns.count, 3, "Expected 3 OMath runs in result")
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

        // PsychQuant/ooxml-swift#117 fix: the spliced Run now serializes as a
        // proper `<w:r>...<m:oMath>...</m:oMath></w:r>` (rPr wrapper preserved,
        // OMath carried as a Run child via `rawElements` rather than replacing
        // the whole run). Re-parsed by DocxReader, this is indistinguishable
        // from a normal inline OMath run loaded from any other document, so it
        // round-trips as a `Run` (via the legacy `rawXML` convention on the read
        // side) — NOT as a direct child of `<w:p>`. Pre-fix, `Run.toXML()`
        // returned the bare `<m:oMath>` fragment with no `<w:r>` wrapper at
        // all, so re-parsing misclassified it as a direct child.
        guard case .paragraph(let reloadedPara) = reloaded.body.children[0] else { XCTFail(); return }

        let runOMathContent = reloadedPara.runs
            .compactMap { $0.rawXML }
            .filter { $0.contains("oMath") }
            .joined()
        let directChildOMathContent = reloadedPara.unrecognizedChildren
            .filter { $0.name == "oMath" || $0.name == "oMathPara" }
            .map { $0.rawXML }
            .joined()

        XCTAssertTrue(
            runOMathContent.contains("<m:t>t</m:t>") || runOMathContent.contains("<mml:t>t</mml:t>"),
            "Round-tripped OMath should be carried by a Run; got runOMathContent: \(runOMathContent)"
        )
        XCTAssertTrue(directChildOMathContent.isEmpty,
            "OMath must round-trip as a Run child, not leak into direct-child unrecognizedChildren; got: \(directChildOMathContent)")
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
        let omathRuns = resultPara.runs.filter { isOMathRun($0) }

        // Should have α + β (original) + t (spliced) = 3 OMath runs total.
        XCTAssertEqual(omathRuns.count, 3,
            "Expected 3 OMath runs (α, β preserved + t spliced); got \(omathRuns.count)")

        // Verify each glyph is present.
        let allOMathContent = omathRuns.compactMap { omathPayload(of: $0) }.joined()
        XCTAssertTrue(allOMathContent.contains("<m:t>α</m:t>"), "Expected α preserved")
        XCTAssertTrue(allOMathContent.contains("<m:t>β</m:t>"), "Expected β preserved")
        XCTAssertTrue(allOMathContent.contains("<m:t>t</m:t>"), "Expected t spliced")
    }
}
