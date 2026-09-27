import XCTest
@testable import OOXMLSwift

/// Unit tests for `Operation.setRunFormatInRange` (PsychQuant/macdoc#232).
///
/// `set_bold` in `docx apply` manifests only has a paragraph anchor + a text
/// range (`substring`) — never a Run ElementID — so the pre-existing
/// `Operation.setRunFormat` (which requires its target to already resolve to
/// a `<w:r>`) always failed with `malformedOp("setRunFormat target must be
/// <w:r>")`. This op resolves `substring` (optionally the Nth `instance`) to
/// a text range inside the target paragraph, splitting run(s) as needed so
/// ONLY the matched text is reformatted.
///
/// Coverage:
/// 1. `testRangeInsideSingleRun_splitsIntoThree` — range strictly inside one
///    run's text; the run is split into prefix/matched-middle/suffix and
///    only the middle is bolded.
/// 2. `testRangeEqualsWholeRun_noSplitNeeded` — the substring exactly equals
///    one run's full text; that run becomes bold as a single node (no empty
///    prefix/suffix fragments).
/// 3. `testRangeSpansMultipleRuns` — the match starts partway through one
///    run and ends partway through a later run, with a fully-enclosed run
///    in between; all three pieces end up bold, unmatched text does not.
/// 4. `testSecondInstanceTargetsSecondOccurrence` — `instance: 2` skips the
///    first occurrence.
/// 5. `testSubstringNotFoundThrowsWithReason` — clear `reason`, not just an
///    opaque error code.
/// 6. `testInstanceBeyondAvailableOccurrencesThrows`
/// 7. `testOriginalRunAttributesPreservedAcrossSplit` — italic + a custom
///    font + an unmodeled `<w:rPr>` child on the original run survive on
///    every split-out fragment (prefix, matched middle, suffix).
/// 8. `testBoldFalseWritesExplicitOff` — `value: false` writes
///    `<w:b w:val="0"/>` rather than removing `<w:b>` (ooxml-swift#161).
/// 9. `testEmptySubstringThrows`, `testInstanceLessThanOneThrows`,
///    `testNonParagraphTargetThrows`, `testUnsupportedFormatFieldThrows` —
///    the four `malformedOp` guard clauses.
/// 10. `testJSONLRoundTrip` — `Operation.setRunFormatInRange` survives
///     `OperationLog.encodeJSONL()` / `decodeJSONL()` byte-for-byte.
/// 11. `testSaveAndReloadPreservesResult` — after materialize, serializing
///     the tree (`XmlTreeWriter.serialize`) and re-parsing it
///     (`XmlTreeReader.parse`) reproduces the same run/text/bold structure.
/// 12. `testReducerErrorReasonSurvivesLocalizedDescription` — the
///     macdoc#232 second complaint: `ReducerError.localizedDescription` (as
///     used by `WordDocument+Apply.swift`'s `operationLogFailure` wrapping)
///     must include the `reason` string, not bridge to a generic
///     "(OOXMLSwift.ReducerError error N.)" NSError message.
final class OperationReducerSetRunFormatInRangeTests: XCTestCase {

    // MARK: - Test helpers

    /// Builds `<w:body><w:p>[<w:r><w:t>text</w:t></w:r> for each text in
    /// `texts`]</w:p></w:body>` and returns the tree + the paragraph's
    /// ElementID. Each run gets its own `libraryUUID` so individual runs are
    /// independently addressable if a test needs to (not required by the new
    /// op itself, which addresses the paragraph).
    private func makeParagraph(runs texts: [String]) -> (XmlTree, paragraphID: ElementID) {
        let paraUUID = UUID()
        var runNodes: [XmlNode] = []
        for text in texts {
            let wt = XmlNode.element(prefix: "w", localName: "t", children: [XmlNode.text(text)])
            let wr = XmlNode.element(prefix: "w", localName: "r", children: [wt])
            wr.libraryUUID = UUID()
            runNodes.append(wr)
        }
        let wp = XmlNode.element(prefix: "w", localName: "p", children: runNodes)
        wp.libraryUUID = paraUUID
        let body = XmlNode.element(prefix: "w", localName: "body", children: [wp])
        let doc = XmlNode.element(prefix: "w", localName: "document", children: [body])
        return (XmlTree.synthesized(root: doc), ElementID(libraryUUID: paraUUID))
    }

    /// Returns the `<w:r>` children of the paragraph identified by
    /// `paragraphID`, in document order.
    private func runsOf(_ tree: XmlTree, paragraphID: ElementID) -> [XmlNode] {
        guard let paragraph = OperationReducer.findNode(elementID: paragraphID, in: tree) else {
            return []
        }
        return paragraph.children.filter { $0.kind == .element && $0.localName == "r" }
    }

    /// Flattened `<w:t>` text of a `<w:r>` node.
    private func text(of run: XmlNode) -> String {
        run.children
            .filter { $0.kind == .element && $0.localName == "t" }
            .flatMap { $0.children.filter { $0.kind == .text }.map(\.textContent) }
            .joined()
    }

    private func rPr(of run: XmlNode) -> XmlNode? {
        run.children.first { $0.kind == .element && $0.localName == "rPr" }
    }

    private func boldElement(of run: XmlNode) -> XmlNode? {
        rPr(of: run)?.children.first { $0.kind == .element && $0.localName == "b" }
    }

    private func isBold(_ run: XmlNode) -> Bool {
        guard let b = boldElement(of: run) else { return false }
        // Naked <w:b/> (no w:val, or w:val truthy) = on.
        guard let val = b.attributes.first(where: { $0.prefix == "w" && $0.localName == "val" })?.value else {
            return true
        }
        return !["0", "false", "off"].contains(val)
    }

    private func isItalic(_ run: XmlNode) -> Bool {
        rPr(of: run)?.children.contains { $0.kind == .element && $0.localName == "i" } ?? false
    }

    // MARK: - 1. Range inside a single run's interior

    func testRangeInsideSingleRun_splitsIntoThree() throws {
        let (base, paragraphID) = makeParagraph(runs: ["Hello world today"])

        var log = OperationLog()
        log.append(
            .setRunFormatInRange(
                target: paragraphID, substring: "world", instance: 1,
                format: RunFormatPayload(bold: true)),
            source: .swift)

        let result = try OperationReducer.materialize(log: log, base: base)
        let runs = runsOf(result, paragraphID: paragraphID)

        XCTAssertEqual(runs.map(text(of:)), ["Hello ", "world", " today"])
        XCTAssertFalse(isBold(runs[0]), "prefix unchanged")
        XCTAssertTrue(isBold(runs[1]), "matched middle is bold")
        XCTAssertFalse(isBold(runs[2]), "suffix unchanged")
    }

    /// A fragment that starts or ends with whitespace needs
    /// `xml:space="preserve"` on its `<w:t>`; without it Word drops the edge
    /// whitespace and the words on either side of the split run together
    /// ("Hello world today" rendered as "Helloworldtoday"). The source run
    /// did not need the attribute, so it cannot simply be inherited.
    func testSplitFragmentsWithEdgeWhitespacePreserveIt() throws {
        let (base, paragraphID) = makeParagraph(runs: ["Hello world today"])

        var log = OperationLog()
        log.append(
            .setRunFormatInRange(
                target: paragraphID, substring: "world", instance: 1,
                format: RunFormatPayload(bold: true)),
            source: .swift)

        let result = try OperationReducer.materialize(log: log, base: base)
        let runs = runsOf(result, paragraphID: paragraphID)
        func space(_ run: XmlNode) -> String? {
            run.children.first { $0.kind == .element && $0.localName == "t" }?
                .attributeValue(prefix: "xml", localName: "space")
        }
        XCTAssertEqual(runs.map(text(of:)), ["Hello ", "world", " today"])
        XCTAssertEqual(space(runs[0]), "preserve", "\"Hello \" ends with a space")
        XCTAssertEqual(space(runs[2]), "preserve", "\" today\" starts with a space")
    }

    // MARK: - 2. Range equals a whole run — no split needed

    func testRangeEqualsWholeRun_noSplitNeeded() throws {
        let (base, paragraphID) = makeParagraph(runs: ["Hello", "world", "today"])

        var log = OperationLog()
        log.append(
            .setRunFormatInRange(
                target: paragraphID, substring: "world", instance: 1,
                format: RunFormatPayload(bold: true)),
            source: .swift)

        let result = try OperationReducer.materialize(log: log, base: base)
        let runs = runsOf(result, paragraphID: paragraphID)

        XCTAssertEqual(runs.map(text(of:)), ["Hello", "world", "today"],
                       "no empty prefix/suffix fragments introduced")
        XCTAssertFalse(isBold(runs[0]))
        XCTAssertTrue(isBold(runs[1]))
        XCTAssertFalse(isBold(runs[2]))
    }

    // MARK: - 3. Range spans multiple runs

    func testRangeSpansMultipleRuns() throws {
        // Runs: "foo" | "bar" | "baz" | "qux". Match "obarbazq" spans:
        // tail of "foo" (from offset 2 -> "o"), all of "bar", all of "baz",
        // head of "qux" (up to offset 1 -> "q").
        let (base, paragraphID) = makeParagraph(runs: ["foo", "bar", "baz", "qux"])

        var log = OperationLog()
        log.append(
            .setRunFormatInRange(
                target: paragraphID, substring: "obarbazq", instance: 1,
                format: RunFormatPayload(bold: true)),
            source: .swift)

        let result = try OperationReducer.materialize(log: log, base: base)
        let runs = runsOf(result, paragraphID: paragraphID)

        XCTAssertEqual(runs.map(text(of:)), ["fo", "o", "bar", "baz", "q", "ux"])
        XCTAssertFalse(isBold(runs[0]), "\"fo\" prefix of first run unchanged")
        XCTAssertTrue(isBold(runs[1]), "\"o\" matched tail of first run")
        XCTAssertTrue(isBold(runs[2]), "\"bar\" fully-enclosed run")
        XCTAssertTrue(isBold(runs[3]), "\"baz\" fully-enclosed run")
        XCTAssertTrue(isBold(runs[4]), "\"q\" matched head of last run")
        XCTAssertFalse(isBold(runs[5]), "\"ux\" suffix of last run unchanged")
    }

    // MARK: - 4. instance disambiguates repeated substrings

    func testSecondInstanceTargetsSecondOccurrence() throws {
        let (base, paragraphID) = makeParagraph(runs: ["cat sat on the cat mat"])

        var log = OperationLog()
        log.append(
            .setRunFormatInRange(
                target: paragraphID, substring: "cat", instance: 2,
                format: RunFormatPayload(bold: true)),
            source: .swift)

        let result = try OperationReducer.materialize(log: log, base: base)
        let runs = runsOf(result, paragraphID: paragraphID)

        XCTAssertEqual(runs.map(text(of:)), ["cat sat on the ", "cat", " mat"])
        XCTAssertFalse(isBold(runs[0]), "first \"cat\" occurrence untouched")
        XCTAssertTrue(isBold(runs[1]), "second \"cat\" occurrence is bold")
        XCTAssertFalse(isBold(runs[2]))
    }

    // MARK: - 5/6. Not-found error paths

    func testSubstringNotFoundThrowsWithReason() {
        let (base, paragraphID) = makeParagraph(runs: ["Hello world"])

        var log = OperationLog()
        log.append(
            .setRunFormatInRange(
                target: paragraphID, substring: "xyz", instance: 1,
                format: RunFormatPayload(bold: true)),
            source: .swift)

        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: base)) { error in
            guard case ReducerError.malformedOp(_, let reason) = error else {
                XCTFail("Expected .malformedOp, got \(error)")
                return
            }
            XCTAssertTrue(reason.contains("xyz"), "reason should name the missing substring: \(reason)")
            XCTAssertTrue(reason.contains("not found"), "reason: \(reason)")
        }
    }

    func testInstanceBeyondAvailableOccurrencesThrows() {
        let (base, paragraphID) = makeParagraph(runs: ["cat sat on the mat"])

        var log = OperationLog()
        log.append(
            .setRunFormatInRange(
                target: paragraphID, substring: "cat", instance: 2,
                format: RunFormatPayload(bold: true)),
            source: .swift)

        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: base)) { error in
            guard case ReducerError.malformedOp(_, let reason) = error else {
                XCTFail("Expected .malformedOp, got \(error)")
                return
            }
            XCTAssertTrue(reason.contains("occurrence"), "reason: \(reason)")
        }
    }

    // MARK: - 7. Original run attributes preserved across the split

    func testOriginalRunAttributesPreservedAcrossSplit() throws {
        let (base, paragraphID) = makeParagraph(runs: ["Hello world today"])
        let run = runsOf(base, paragraphID: paragraphID)[0]

        // Seed pre-existing formatting: italic + a custom font + an
        // unmodeled <w:rPr> child (<w:lang>, not represented by
        // RunFormatPayload at all) that must survive verbatim.
        let rPrNode = XmlNode.element(prefix: "w", localName: "rPr")
        let iNode = XmlNode.element(prefix: "w", localName: "i")
        let rFonts = XmlNode.element(prefix: "w", localName: "rFonts")
        rFonts.setAttribute(prefix: "w", localName: "ascii", value: "Georgia")
        let lang = XmlNode.element(prefix: "w", localName: "lang")
        lang.setAttribute(prefix: "w", localName: "val", value: "ja-JP")
        rPrNode.children = [rFonts, iNode, lang]
        run.children = [rPrNode] + run.children

        var log = OperationLog()
        log.append(
            .setRunFormatInRange(
                target: paragraphID, substring: "world", instance: 1,
                format: RunFormatPayload(bold: true)),
            source: .swift)

        let result = try OperationReducer.materialize(log: log, base: base)
        let runs = runsOf(result, paragraphID: paragraphID)
        XCTAssertEqual(runs.map(text(of:)), ["Hello ", "world", " today"])

        for (index, run) in runs.enumerated() {
            XCTAssertTrue(isItalic(run), "run \(index) should keep italic from the original")
            let fonts = rPr(of: run)?.children.first { $0.localName == "rFonts" }
            XCTAssertEqual(
                fonts?.attributes.first { $0.localName == "ascii" }?.value, "Georgia",
                "run \(index) should keep the custom font")
            let langNode = rPr(of: run)?.children.first { $0.localName == "lang" }
            XCTAssertEqual(
                langNode?.attributes.first { $0.localName == "val" }?.value, "ja-JP",
                "run \(index) should keep the unmodeled <w:lang> child verbatim")
        }
        XCTAssertFalse(isBold(runs[0]))
        XCTAssertTrue(isBold(runs[1]), "matched middle additionally gets bold, on top of preserved italic/font")
        XCTAssertFalse(isBold(runs[2]))
    }

    // MARK: - 8. bold: false writes explicit off (ooxml-swift#161 alignment)

    func testBoldFalseWritesExplicitOff() throws {
        let (base, paragraphID) = makeParagraph(runs: ["Hello world today"])

        var log = OperationLog()
        log.append(
            .setRunFormatInRange(
                target: paragraphID, substring: "world", instance: 1,
                format: RunFormatPayload(bold: false)),
            source: .swift)

        let result = try OperationReducer.materialize(log: log, base: base)
        let runs = runsOf(result, paragraphID: paragraphID)
        let matched = runs[1]
        let bElement = try XCTUnwrap(boldElement(of: matched), "explicit off still writes a <w:b> element")
        let val = bElement.attributes.first { $0.prefix == "w" && $0.localName == "val" }?.value
        XCTAssertEqual(val, "0", "false must write explicit off, not remove <w:b> (ooxml-swift#161)")
    }

    // MARK: - 9. Guard clauses

    func testEmptySubstringThrows() {
        let (base, paragraphID) = makeParagraph(runs: ["Hello world"])
        var log = OperationLog()
        log.append(
            .setRunFormatInRange(
                target: paragraphID, substring: "", instance: 1,
                format: RunFormatPayload(bold: true)),
            source: .swift)
        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: base)) { error in
            guard case ReducerError.malformedOp(_, let reason) = error else {
                XCTFail("Expected .malformedOp, got \(error)"); return
            }
            XCTAssertTrue(reason.contains("empty"), "reason: \(reason)")
        }
    }

    func testInstanceLessThanOneThrows() {
        let (base, paragraphID) = makeParagraph(runs: ["Hello world"])
        var log = OperationLog()
        log.append(
            .setRunFormatInRange(
                target: paragraphID, substring: "world", instance: 0,
                format: RunFormatPayload(bold: true)),
            source: .swift)
        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: base)) { error in
            guard case ReducerError.malformedOp(_, let reason) = error else {
                XCTFail("Expected .malformedOp, got \(error)"); return
            }
            XCTAssertTrue(reason.contains("instance"), "reason: \(reason)")
        }
    }

    func testNonParagraphTargetThrows() {
        // Target a Run (not a paragraph) — this is exactly issue #232's
        // original failure shape, now with a specific reason instead of a
        // bare error code.
        let (base, paragraphID) = makeParagraph(runs: ["Hello world"])
        let runID = ElementID(node: runsOf(base, paragraphID: paragraphID)[0])
            ?? { fatalError("run must have a libraryUUID") }()

        var log = OperationLog()
        log.append(
            .setRunFormatInRange(
                target: runID, substring: "world", instance: 1,
                format: RunFormatPayload(bold: true)),
            source: .swift)
        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: base)) { error in
            guard case ReducerError.malformedOp(_, let reason) = error else {
                XCTFail("Expected .malformedOp, got \(error)"); return
            }
            XCTAssertTrue(reason.contains("<w:p>"), "reason: \(reason)")
        }
    }

    func testUnsupportedFormatFieldThrows() {
        let (base, paragraphID) = makeParagraph(runs: ["Hello world"])
        var log = OperationLog()
        log.append(
            .setRunFormatInRange(
                target: paragraphID, substring: "world", instance: 1,
                format: RunFormatPayload(italic: true)),
            source: .swift)
        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: base)) { error in
            guard case ReducerError.malformedOp(_, let reason) = error else {
                XCTFail("Expected .malformedOp, got \(error)"); return
            }
            XCTAssertTrue(reason.contains("bold"), "reason should say only bold is supported: \(reason)")
        }
    }

    // MARK: - 10. JSONL round-trip

    func testJSONLRoundTrip() throws {
        var log = OperationLog()
        log.append(
            .setRunFormatInRange(
                target: ElementID(libraryUUID: UUID()),
                substring: "hello \"world\"",
                instance: 2,
                format: RunFormatPayload(bold: false)),
            source: .swift)

        let data = log.encodeJSONL()
        let decoded = try OperationLog.decodeJSONL(data)

        XCTAssertEqual(decoded.entries.count, 1)
        XCTAssertEqual(decoded.entries[0].op, log.entries[0].op)

        // Second round-trip (encode the decoded log again) must be
        // byte-identical — the JSONL contract's "byte-equal round-trip"
        // requirement.
        XCTAssertEqual(decoded.encodeJSONL(), data)
    }

    // MARK: - 11. Save (serialize) and reload preserves the result

    /// Structural lookup (NOT by ElementID): `libraryUUID` is an in-memory
    /// convenience that is never written to the XML bytes, so it does not
    /// survive a serialize→reparse round trip. The save/reload test below
    /// therefore finds "the paragraph" by tree position, exactly as a real
    /// caller would after re-opening a saved .docx and re-deriving IDs from
    /// `w14:paraId` / structural position rather than a stale in-memory UUID.
    private func firstParagraphRuns(_ tree: XmlTree) -> [XmlNode] {
        var found: XmlNode?
        func walk(_ node: XmlNode) {
            guard found == nil else { return }
            if node.kind == .element && node.localName == "p" {
                found = node
                return
            }
            for child in node.children { walk(child) }
        }
        walk(tree.root)
        return found?.children.filter { $0.kind == .element && $0.localName == "r" } ?? []
    }

    func testSaveAndReloadPreservesResult() throws {
        let (base, paragraphID) = makeParagraph(runs: ["Hello world today"])

        var log = OperationLog()
        log.append(
            .setRunFormatInRange(
                target: paragraphID, substring: "world", instance: 1,
                format: RunFormatPayload(bold: true)),
            source: .swift)

        let result = try OperationReducer.materialize(log: log, base: base)

        let serialized = try XmlTreeWriter.serialize(result)
        let reloaded = try XmlTreeReader.parse(serialized)

        let runs = firstParagraphRuns(reloaded)
        XCTAssertEqual(runs.map(text(of:)), ["Hello ", "world", " today"])
        XCTAssertFalse(isBold(runs[0]))
        XCTAssertTrue(isBold(runs[1]))
        XCTAssertFalse(isBold(runs[2]))
    }

    // MARK: - 12. ReducerError.localizedDescription carries `reason`

    func testReducerErrorReasonSurvivesLocalizedDescription() {
        let opID = UUID()
        let error = ReducerError.malformedOp(
            opID: opID, reason: "setRunFormat target must be <w:r>")
        XCTAssertTrue(
            (error as Error).localizedDescription.contains("setRunFormat target must be <w:r>"),
            "localizedDescription must surface `reason`, got: \((error as Error).localizedDescription)")
    }
}
