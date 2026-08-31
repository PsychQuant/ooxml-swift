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

    private func omathXML(in run: Run) -> String? {
        if let rawXML = run.rawXML, rawXML.contains("oMath") {
            return rawXML
        }
        return run.rawElements?.first(where: {
            $0.name == "oMath" || $0.name == "oMathPara"
        })?.xml
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

        // Verify a Run carrier containing OMath was added.
        let omathRuns = resultPara.runs.filter { omathXML(in: $0) != nil }
        XCTAssertEqual(omathRuns.count, 1, "Expected exactly one OMath run in target")

        // Verify the rawXML byte-equals (or substring-matches) the source OMath block.
        let spliced = omathXML(in: omathRuns[0]) ?? ""
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
        let runWithOMath = resultPara.runs.first { omathXML(in: $0) != nil }
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
        let omathIdx = runs.firstIndex { omathXML(in: $0) != nil }
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
        let omathRun = resultPara.runs.first { omathXML(in: $0) != nil }
        XCTAssertNotNil(omathRun)
        XCTAssertEqual(omathRun?.properties.fontName, "Cambria Math",
            "fontName should propagate via rFonts.ascii in .full mode")
        XCTAssertEqual(omathRun?.properties.fontSize, 24,
            "fontSize 24 (12pt) should propagate in .full mode")
        XCTAssertNotNil(omathRun?.rawXML, "inline OMath remains visible through the compatibility rawXML model")
        let xml = resultPara.toXML()
        XCTAssertTrue(xml.contains("</w:rPr><m:oMath"), "Got: \(xml)")
        XCTAssertTrue(xml.contains(#"<w:rFonts w:ascii="Cambria Math""#), "Got: \(xml)")
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
        let omathRun = resultPara.runs.first { omathXML(in: $0) != nil }
        XCTAssertNotNil(omathRun)
        XCTAssertNil(omathRun?.properties.fontName, ".discard should clear fontName")
        XCTAssertNil(omathRun?.properties.fontSize, ".discard should clear fontSize")
        let xml = resultPara.toXML()
        XCTAssertTrue(xml.contains("<w:r><m:oMath"), "Got: \(xml)")
        XCTAssertFalse(xml.contains("<w:rPr>"), "discard must not emit copied rPr: \(xml)")
    }

    func testRpRModeOMathOnlyEmitsWhitelistInsideRunCarrier() throws {
        let sourceXML = """
        <w:p \(Self.mNS)>
          <w:r>
            <w:rPr><w:rStyle w:val="MathStyle"/><w:rFonts w:ascii="Cambria Math" w:eastAsia="PMingLiU"/><w:b/><w:i/><w:color w:val="FF0000"/><w:sz w:val="24"/><w:lang w:val="en-US" w:eastAsia="zh-TW"/></w:rPr>
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
        let xml = resultPara.toXML()
        XCTAssertTrue(xml.contains(#"<w:rFonts w:ascii="Cambria Math""#), "Got: \(xml)")
        XCTAssertTrue(xml.contains("<w:b/>"), "Got: \(xml)")
        XCTAssertTrue(xml.contains("<w:i/>"), "Got: \(xml)")
        XCTAssertTrue(xml.contains(#"<w:sz w:val="24"/>"#), "Got: \(xml)")
        XCTAssertTrue(xml.contains(#"<w:szCs w:val="24"/>"#), "Got: \(xml)")
        XCTAssertTrue(xml.contains(#"<w:lang w:val="en-US" w:eastAsia="zh-TW"/>"#), "Got: \(xml)")
        XCTAssertTrue(xml.contains("</w:rPr><m:oMath"), "Got: \(xml)")
        XCTAssertFalse(xml.contains("<w:rStyle"), "Got: \(xml)")
        XCTAssertFalse(xml.contains("<w:color"), "Got: \(xml)")
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
        let omathRuns = resultPara.runs.filter { omathXML(in: $0) != nil }
        XCTAssertEqual(omathRuns.count, 3, "Expected 3 OMath runs in result")
    }

    // MARK: - Round-trip lossless guarantee (6.13)

    /// Save/reload and a second dirty save must both preserve the inline Run
    /// carrier, exact self-contained OMath bytes, and each rPr mode's semantics.
    func testRoundTripPreservesOMathCarrierBytesAndRpRModesAcrossDirtyResave() throws {
        let sourceXML = """
        <w:p \(Self.mNS)>
          <w:r>
            <w:rPr><w:rStyle w:val="MathStyle"/><w:rFonts w:ascii="Cambria Math" w:eastAsia="PMingLiU"/><w:b/><w:i/><w:color w:val="FF0000"/><w:sz w:val="24"/><w:lang w:val="en-US" w:eastAsia="zh-TW"/></w:rPr>
            <m:oMath><m:r><m:t>α</m:t></m:r></m:oMath>
          </w:r>
        </w:p>
        """
        let sourcePara = try parseParagraph(xml: sourceXML)
        let expectedOMath = try XCTUnwrap(OMathExtractor.extract(from: sourcePara).first?.xml)

        let cases: [(name: String, mode: OMathSpliceRpRMode)] = [
            ("full", .full),
            ("omathOnly", .omathOnly),
            ("discard", .discard),
        ]

        for entry in cases {
            let firstURL = URL(fileURLWithPath: NSTemporaryDirectory())
                .appendingPathComponent("OMathSpliceTests-\(entry.name)-first-\(UUID().uuidString).docx")
            let secondURL = URL(fileURLWithPath: NSTemporaryDirectory())
                .appendingPathComponent("OMathSpliceTests-\(entry.name)-second-\(UUID().uuidString).docx")
            defer {
                try? FileManager.default.removeItem(at: firstURL)
                try? FileManager.default.removeItem(at: secondURL)
            }

            var target = makeDocument(with: try parseParagraph(xml: Self.targetEmpty))
            try target.spliceOMath(
                from: sourcePara,
                toBodyParagraphIndex: 0,
                position: .atEnd,
                omathIndex: 0,
                rPrMode: entry.mode
            )
            try DocxWriter.write(target, to: firstURL)
            var firstReload = try DocxReader.read(from: firstURL)

            try assertReloadedOMath(
                in: firstReload,
                expectedXML: expectedOMath,
                mode: entry.mode,
                stage: "\(entry.name) first reload"
            )

            // Force typed document.xml serialization on the second save. A
            // rawXML OMath short-circuit used to drop both rPr and <w:r> here.
            firstReload.markPartDirty("word/document.xml")
            try DocxWriter.write(firstReload, to: secondURL)
            let secondReload = try DocxReader.read(from: secondURL)
            try assertReloadedOMath(
                in: secondReload,
                expectedXML: expectedOMath,
                mode: entry.mode,
                stage: "\(entry.name) dirty resave"
            )
        }
    }

    private func assertReloadedOMath(
        in document: WordDocument,
        expectedXML: String,
        mode: OMathSpliceRpRMode,
        stage: String
    ) throws {
        guard case .paragraph(let paragraph) = document.body.children[0] else {
            XCTFail("\(stage): expected paragraph")
            return
        }
        let run = try XCTUnwrap(
            paragraph.runs.first(where: { omathXML(in: $0) != nil }),
            "\(stage): inline OMath must remain in a Run carrier"
        )
        XCTAssertTrue(
            paragraph.unrecognizedChildren.filter { $0.name == "oMath" || $0.name == "oMathPara" }.isEmpty,
            "\(stage): inline OMath must not become a direct paragraph child"
        )
        XCTAssertEqual(omathXML(in: run), expectedXML, "\(stage): OMath bytes changed")

        switch mode {
        case .full:
            XCTAssertEqual(run.properties.rStyle, "MathStyle", "\(stage): full rStyle")
            XCTAssertEqual(run.properties.color, "FF0000", "\(stage): full color")
            XCTAssertEqual(run.properties.rFonts?.eastAsia, "PMingLiU", "\(stage): full rFonts")
            XCTAssertEqual(run.properties.fontSize, 24, "\(stage): full sz/szCs")
            XCTAssertEqual(run.properties.lang?.eastAsia, "zh-TW", "\(stage): full lang")
            XCTAssertTrue(run.properties.bold, "\(stage): full bold")
            XCTAssertTrue(run.properties.italic, "\(stage): full italic")
        case .omathOnly:
            XCTAssertNil(run.properties.rStyle, "\(stage): omathOnly strips rStyle")
            XCTAssertNil(run.properties.color, "\(stage): omathOnly strips color")
            XCTAssertEqual(run.properties.rFonts?.eastAsia, "PMingLiU", "\(stage): omathOnly rFonts")
            XCTAssertEqual(run.properties.fontSize, 24, "\(stage): omathOnly sz/szCs")
            XCTAssertEqual(run.properties.lang?.eastAsia, "zh-TW", "\(stage): omathOnly lang")
            XCTAssertTrue(run.properties.bold, "\(stage): omathOnly bold")
            XCTAssertTrue(run.properties.italic, "\(stage): omathOnly italic")
        case .discard:
            XCTAssertEqual(run.properties, RunProperties(), "\(stage): discard must keep default rPr")
        }
    }

    func testOMathOnlyExplicitOffSurvivesCarrierAndDirtyResave() throws {
        let sourceXML = """
        <w:p \(Self.mNS)>
          <w:r>
            <w:rPr><w:rFonts w:ascii="Cambria Math"/><w:b w:val="0"/><w:i w:val="off"/><w:lang w:val="en-US"/></w:rPr>
            <m:oMath><m:r><m:t>α</m:t></m:r></m:oMath>
          </w:r>
        </w:p>
        """
        let sourcePara = try parseParagraph(xml: sourceXML)
        let expectedOMath = try XCTUnwrap(OMathExtractor.extract(from: sourcePara).first?.xml)
        let firstURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("OMathSpliceTests-explicit-off-first-\(UUID().uuidString).docx")
        let secondURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("OMathSpliceTests-explicit-off-second-\(UUID().uuidString).docx")
        defer {
            try? FileManager.default.removeItem(at: firstURL)
            try? FileManager.default.removeItem(at: secondURL)
        }

        func assertExplicitOff(_ document: WordDocument, stage: String) throws {
            guard case .paragraph(let paragraph) = document.body.children[0] else {
                XCTFail("\(stage): expected paragraph")
                return
            }
            let run = try XCTUnwrap(
                paragraph.runs.first(where: { omathXML(in: $0) != nil }),
                "\(stage): expected inline OMath Run"
            )
            XCTAssertEqual(run.properties.specifiedBold, false, "\(stage): explicit bold off")
            XCTAssertEqual(run.properties.specifiedItalic, false, "\(stage): explicit italic off")
            XCTAssertEqual(run.properties.lang?.val, "en-US", "\(stage): lang")
            XCTAssertEqual(omathXML(in: run), expectedOMath, "\(stage): OMath bytes")
            let xml = paragraph.toXML()
            XCTAssertTrue(xml.contains(#"<w:b w:val="0"/>"#), "\(stage): \(xml)")
            XCTAssertTrue(xml.contains(#"<w:i w:val="0"/>"#), "\(stage): \(xml)")
            XCTAssertTrue(xml.contains("</w:rPr><m:oMath"), "\(stage): \(xml)")
        }

        var target = makeDocument(with: try parseParagraph(xml: Self.targetEmpty))
        try target.spliceOMath(
            from: sourcePara,
            toBodyParagraphIndex: 0,
            position: .atEnd,
            rPrMode: .omathOnly
        )
        try DocxWriter.write(target, to: firstURL)
        var firstReload = try DocxReader.read(from: firstURL)
        try assertExplicitOff(firstReload, stage: "first reload")

        firstReload.markPartDirty("word/document.xml")
        try DocxWriter.write(firstReload, to: secondURL)
        let secondReload = try DocxReader.read(from: secondURL)
        try assertExplicitOff(secondReload, stage: "dirty resave")
    }

    func testSplicedInlineOMathCanBeUsedAsAnotherSpliceSource() throws {
        let sourcePara = try parseParagraph(xml: Self.sourceInlineRunOMath)
        var firstTarget = makeDocument(with: try parseParagraph(xml: Self.targetEmpty))
        try firstTarget.spliceOMath(from: sourcePara, toBodyParagraphIndex: 0, position: .atEnd)
        guard case .paragraph(let firstResult) = firstTarget.body.children[0] else { XCTFail(); return }
        XCTAssertTrue(
            firstResult.flattenedDisplayText().contains("t"),
            "spliced inline OMath must remain visible to pre-save flattened text and anchor lookup"
        )

        var secondTarget = makeDocument(with: try parseParagraph(xml: Self.targetEmpty))
        XCTAssertNoThrow(
            try secondTarget.spliceOMath(from: firstResult, toBodyParagraphIndex: 0, position: .atEnd)
        )
        guard case .paragraph(let secondResult) = secondTarget.body.children[0] else { XCTFail(); return }
        XCTAssertEqual(secondResult.runs.filter { omathXML(in: $0) != nil }.count, 1)
    }

    func testStrictNamespacePolicyRecognizesSplicedInlineOMathTarget() throws {
        let sourcePara = try parseParagraph(xml: Self.sourceMMLPrefixOMath)
        var target = makeDocument(with: try parseParagraph(xml: Self.targetEmpty))

        try target.spliceOMath(
            from: sourcePara,
            toBodyParagraphIndex: 0,
            position: .atEnd,
            namespacePolicy: .lenient
        )
        guard case .paragraph(let rawElementTarget) = target.body.children[0] else { XCTFail(); return }

        XCTAssertNoThrow(
            try target.spliceOMath(
                from: sourcePara,
                toBodyParagraphIndex: 0,
                position: .atEnd,
                namespacePolicy: .strict
            ),
            "strict policy should observe the existing mml prefix in the spliced Run carrier: \(rawElementTarget)"
        )
    }

    func testVendorOMathNamedRawXMLRemainsExactReplacement() {
        let vendorXML = #"<x:oMath xmlns:x="urn:vendor:not-omml"><x:data/></x:oMath>"#
        var run = Run(text: "")
        run.rawXML = vendorXML
        run.properties.bold = true

        XCTAssertEqual(
            run.toXML(),
            vendorXML,
            "same local name in a non-OMML namespace must keep generic rawXML replacement semantics"
        )

        let paragraphXML = Paragraph(runs: [run]).toXML()
        XCTAssertEqual(
            paragraphXML,
            "<w:p>\(vendorXML)</w:p>",
            "Paragraph.emitRun must use the same namespace-aware classification"
        )
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
        let omathRuns = resultPara.runs.filter { omathXML(in: $0) != nil }

        // Should have α + β (original) + t (spliced) = 3 OMath runs total.
        XCTAssertEqual(omathRuns.count, 3,
            "Expected 3 OMath runs (α, β preserved + t spliced); got \(omathRuns.count)")

        // Verify each glyph is present.
        let allOMathContent = omathRuns.compactMap { omathXML(in: $0) }.joined()
        XCTAssertTrue(allOMathContent.contains("<m:t>α</m:t>"), "Expected α preserved")
        XCTAssertTrue(allOMathContent.contains("<m:t>β</m:t>"), "Expected β preserved")
        XCTAssertTrue(allOMathContent.contains("<m:t>t</m:t>"), "Expected t spliced")
    }
}
