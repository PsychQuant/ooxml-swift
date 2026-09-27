import XCTest
@testable import OOXMLSwift

/// R3 (RED→GREEN) coverage for the two MEDIUM findings the R2 adversarial
/// review flagged as safe-but-worth-fixing-before-release (PsychQuant/
/// macdoc#232, `review-ox-232.md`'s "R2 複審" section, findings (b) and
/// (c)):
///
/// - (b): `<w:proofErr>` and the other genuinely zero-width, textless
///   structural markers (bookmark/comment-range/perm/move-range start/end)
///   made an otherwise-ordinary, visually-contiguous match fail, because R2
///   treated EVERY non-`<w:r>` direct child as a hard boundary. Real Word
///   documents with technical terms, names, or non-English words routinely
///   contain `<w:proofErr>`; this op exists specifically for real Word
///   documents.
/// - (c): when a match failed because of a boundary that contributed ZERO
///   characters to the diagnostic text (the review's example: `<w:proofErr>`
///   itself, before it was whitelisted here), the error message degraded to
///   a generic "not found" instead of naming the responsible boundary.
///
/// Also covers a related, closely-adjacent real-Word-output shape flagged
/// in the same review round: `<w:lastRenderedPageBreak/>` (Word's layout
/// engine stamps this into whatever run happens to span a rendered page
/// boundary, in essentially every multi-page real `.docx`) must not make
/// its run's own text unsearchable, and must not be duplicated across a
/// split.
final class OperationReducerSetRunFormatInRangeR3Tests: XCTestCase {

    // MARK: - Shared helpers

    private func run(_ text: String) -> XmlNode {
        let wt = XmlNode.element(prefix: "w", localName: "t", children: [XmlNode.text(text)])
        let wr = XmlNode.element(prefix: "w", localName: "r", children: [wt])
        wr.libraryUUID = UUID()
        return wr
    }

    private func paragraph(children: [XmlNode]) -> (XmlTree, ElementID) {
        let paraUUID = UUID()
        let wp = XmlNode.element(prefix: "w", localName: "p", children: children)
        wp.libraryUUID = paraUUID
        let body = XmlNode.element(prefix: "w", localName: "body", children: [wp])
        let doc = XmlNode.element(prefix: "w", localName: "document", children: [body])
        return (XmlTree.synthesized(root: doc), ElementID(libraryUUID: paraUUID))
    }

    private func paragraphChildren(_ tree: XmlTree, paragraphID: ElementID) -> [XmlNode] {
        guard let paragraph = OperationReducer.findNode(elementID: paragraphID, in: tree) else {
            return []
        }
        return paragraph.children
    }

    private func text(of run: XmlNode) -> String {
        run.children
            .filter { $0.kind == .element && $0.localName == "t" }
            .flatMap { $0.children.filter { $0.kind == .text }.map(\.textContent) }
            .joined()
    }

    private func isBold(_ run: XmlNode) -> Bool {
        guard let rPr = run.children.first(where: { $0.kind == .element && $0.localName == "rPr" })
        else { return false }
        return rPr.children.contains { $0.kind == .element && $0.localName == "b" }
    }

    private func reason(of error: Error) -> String {
        guard case ReducerError.malformedOp(_, let reason) = error else {
            XCTFail("Expected .malformedOp, got \(error)")
            return ""
        }
        return reason
    }

    // MARK: - (b) <w:proofErr> is transparent

    /// The review report's exact repro: "The " <proofErr spellStart/>
    /// "Fourier" <proofErr spellEnd/> " transform is linear". Searching
    /// "Fourier transform" (spanning the proofErr-wrapped word into the
    /// next) must now SUCCEED — and the proofErr markers must survive,
    /// unduplicated, at their original position.
    func testProofErrWrappedTextIsFindableAndMarkerStaysInPlace() throws {
        let proofStart = XmlNode.element(
            prefix: "w", localName: "proofErr",
            attributes: [XmlAttribute(prefix: "w", localName: "type", value: "spellStart")])
        let proofEnd = XmlNode.element(
            prefix: "w", localName: "proofErr",
            attributes: [XmlAttribute(prefix: "w", localName: "type", value: "spellEnd")])

        let (base, paragraphID) = paragraph(children: [
            run("The "), proofStart, run("Fourier"), proofEnd, run(" transform is linear"),
        ])

        var log = OperationLog()
        log.append(
            .setRunFormatInRange(target: paragraphID, substring: "Fourier transform", instance: 1,
                                  format: RunFormatPayload(bold: true)),
            source: .swift)

        let result = try OperationReducer.materialize(log: log, base: base)
        let children = paragraphChildren(result, paragraphID: paragraphID)

        // Reconstructed visible text (markers contribute nothing) must be
        // byte-identical to the original.
        let runs = children.filter { $0.kind == .element && $0.localName == "r" }
        XCTAssertEqual(runs.map(text(of:)).joined(), "The Fourier transform is linear")

        // The two proofErr markers must both still be present, unduplicated,
        // with their original `w:type` attributes intact.
        let proofErrs = children.filter { $0.kind == .element && $0.localName == "proofErr" }
        XCTAssertEqual(proofErrs.count, 2, "both proofErr markers must survive, none duplicated")
        XCTAssertEqual(
            proofErrs.map { $0.attributeValue(prefix: "w", localName: "type") },
            ["spellStart", "spellEnd"],
            "markers keep their original attributes and relative order")

        // Bold lands on "Fourier" and the "transform" fragment; "The " and
        // " is linear" are untouched.
        XCTAssertTrue(runs.contains { text(of: $0) == "Fourier" && isBold($0) })
        XCTAssertTrue(runs.contains { text(of: $0) == " transform" && isBold($0) })
        XCTAssertTrue(runs.contains { text(of: $0) == "The " && !isBold($0) })
        XCTAssertTrue(runs.contains { text(of: $0) == " is linear" && !isBold($0) })

        // The markers still sit immediately before/after the "Fourier" run,
        // not moved or reordered.
        guard let fourierIdx = children.firstIndex(where: {
            $0.kind == .element && $0.localName == "r" && text(of: $0) == "Fourier"
        }) else {
            XCTFail("Fourier run missing"); return
        }
        XCTAssertEqual(children[fourierIdx - 1].localName, "proofErr")
        XCTAssertEqual(children[fourierIdx + 1].localName, "proofErr")
    }

    /// A match entirely on one side of a proofErr pair (not crossing it at
    /// all) is unaffected — sanity check this isn't accidentally broken.
    func testProofErrElsewhereDoesNotBlockAnUnrelatedMatch() throws {
        let proofStart = XmlNode.element(prefix: "w", localName: "proofErr",
            attributes: [XmlAttribute(prefix: "w", localName: "type", value: "spellStart")])
        let proofEnd = XmlNode.element(prefix: "w", localName: "proofErr",
            attributes: [XmlAttribute(prefix: "w", localName: "type", value: "spellEnd")])
        let (base, paragraphID) = paragraph(children: [
            run("The "), proofStart, run("Fourier"), proofEnd, run(" transform is linear"),
        ])
        var log = OperationLog()
        log.append(
            .setRunFormatInRange(target: paragraphID, substring: "linear", instance: 1,
                                  format: RunFormatPayload(bold: true)),
            source: .swift)
        let result = try OperationReducer.materialize(log: log, base: base)
        let runs = paragraphChildren(result, paragraphID: paragraphID)
            .filter { $0.kind == .element && $0.localName == "r" }
        XCTAssertTrue(runs.contains { text(of: $0) == "linear" && isBold($0) })
    }

    // MARK: - (b) bookmark spanning the match range

    /// A `<w:bookmarkStart>`/`<w:bookmarkEnd>` pair sitting strictly inside
    /// the matched range ("foo " + bookmark("bar") + " baz", searching the
    /// full "foo bar baz") must not block the match, and the bookmark pair
    /// must survive at its original position.
    func testBookmarkInsideMatchRangeWorksNormallyAndSurvives() throws {
        let bookmarkStart = XmlNode.element(
            prefix: "w", localName: "bookmarkStart",
            attributes: [
                XmlAttribute(prefix: "w", localName: "id", value: "0"),
                XmlAttribute(prefix: "w", localName: "name", value: "X"),
            ])
        let bookmarkEnd = XmlNode.element(
            prefix: "w", localName: "bookmarkEnd",
            attributes: [XmlAttribute(prefix: "w", localName: "id", value: "0")])

        let (base, paragraphID) = paragraph(children: [
            run("foo "), bookmarkStart, run("bar"), bookmarkEnd, run(" baz"),
        ])

        var log = OperationLog()
        log.append(
            .setRunFormatInRange(target: paragraphID, substring: "foo bar baz", instance: 1,
                                  format: RunFormatPayload(bold: true)),
            source: .swift)

        let result = try OperationReducer.materialize(log: log, base: base)
        let children = paragraphChildren(result, paragraphID: paragraphID)
        let runs = children.filter { $0.kind == .element && $0.localName == "r" }

        XCTAssertEqual(runs.map(text(of:)).joined(), "foo bar baz")
        XCTAssertTrue(runs.allSatisfy(isBold), "the entire matched range, including 'bar', is bold")

        let bookmarks = children.filter {
            $0.kind == .element && ($0.localName == "bookmarkStart" || $0.localName == "bookmarkEnd")
        }
        XCTAssertEqual(bookmarks.map(\.localName), ["bookmarkStart", "bookmarkEnd"],
                       "bookmark pair survives, unduplicated, in original order")
        XCTAssertEqual(bookmarks[0].attributeValue(prefix: "w", localName: "name"), "X")
    }

    // MARK: - (b) remaining closed-list markers are also transparent

    /// Every OTHER member of the closed whitelist (commentRange*, perm*,
    /// moveFromRange*, moveToRange*) is likewise transparent — a match
    /// spanning across each, one at a time, succeeds.
    func testAllRemainingWhitelistedMarkersAreTransparent() throws {
        let pairs: [(start: String, end: String)] = [
            ("commentRangeStart", "commentRangeEnd"),
            ("permStart", "permEnd"),
            ("moveFromRangeStart", "moveFromRangeEnd"),
            ("moveToRangeStart", "moveToRangeEnd"),
        ]
        for pair in pairs {
            let startMarker = XmlNode.element(
                prefix: "w", localName: pair.start,
                attributes: [XmlAttribute(prefix: "w", localName: "id", value: "1")])
            let endMarker = XmlNode.element(
                prefix: "w", localName: pair.end,
                attributes: [XmlAttribute(prefix: "w", localName: "id", value: "1")])
            let (base, paragraphID) = paragraph(children: [
                run("foo "), startMarker, run("bar"), endMarker, run(" baz"),
            ])
            var log = OperationLog()
            log.append(
                .setRunFormatInRange(target: paragraphID, substring: "foo bar baz", instance: 1,
                                      format: RunFormatPayload(bold: true)),
                source: .swift)
            XCTAssertNoThrow(try OperationReducer.materialize(log: log, base: base),
                             "\(pair.start)/\(pair.end) must be transparent to grouping")
        }
    }

    // MARK: - Remaining boundaries (ins/del/hyperlink/sdt/complex-run) unchanged

    /// A tracked insertion still blocks a crossing match — R2 behavior
    /// unchanged by R3's whitelist (only genuinely zero-width markers are
    /// exempted; <w:ins> wraps real, addressable content).
    func testTrackedInsertionStillBlocksACrossingMatch() throws {
        let insWrapped = XmlNode.element(prefix: "w", localName: "ins", children: [run("XXX")])
        let (base, paragraphID) = paragraph(children: [run("foo"), insWrapped, run("bar")])
        var log = OperationLog()
        log.append(
            .setRunFormatInRange(target: paragraphID, substring: "foobar", instance: 1,
                                  format: RunFormatPayload(bold: true)),
            source: .swift)
        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: base))
    }

    // MARK: - (c) zero-contribution boundary still named in the reason

    /// An EMPTY `<w:hyperlink>` (no runs inside, contributes zero
    /// characters to the diagnostic text) sitting between "foo" and "bar"
    /// must still be NAMED in the refusal reason when a match needs to
    /// cross it — not degrade to a generic "not found". This is the general
    /// mechanism (`VisualBoundary` position tracking) that also would have
    /// covered `<w:proofErr>` before it was whitelisted; proved here with a
    /// still-a-boundary element so the fix is demonstrated independently of
    /// the whitelist change.
    func testZeroContributionBoundaryIsStillNamedInReason() throws {
        let emptyHyperlink = XmlNode.element(
            prefix: "w", localName: "hyperlink",
            attributes: [XmlAttribute(prefix: "r", localName: "id", value: "rId1")],
            children: [])
        let (base, paragraphID) = paragraph(children: [run("foo"), emptyHyperlink, run("bar")])
        var log = OperationLog()
        log.append(
            .setRunFormatInRange(target: paragraphID, substring: "foobar", instance: 1,
                                  format: RunFormatPayload(bold: true)),
            source: .swift)
        XCTAssertThrowsError(try OperationReducer.materialize(log: log, base: base)) { error in
            XCTAssertTrue(self.reason(of: error).contains("<w:hyperlink>"),
                          "expected the empty hyperlink to be named, got: \(self.reason(of: error))")
        }
    }

    // MARK: - <w:lastRenderedPageBreak/> is simple and never duplicated

    private func runWithPageBreak(_ text: String) -> XmlNode {
        let pageBreak = XmlNode.element(prefix: "w", localName: "lastRenderedPageBreak")
        let wt = XmlNode.element(prefix: "w", localName: "t", children: [XmlNode.text(text)])
        let wr = XmlNode.element(prefix: "w", localName: "r", children: [pageBreak, wt])
        wr.libraryUUID = UUID()
        return wr
    }

    /// Real Word output shape: `<w:r><w:lastRenderedPageBreak/><w:t>
    /// continued text here</w:t></w:r>`. A 3-way split (prefix/middle/
    /// suffix) must produce the marker on EXACTLY ONE fragment (the first
    /// — "prefix", since the marker precedes the text) — never zero, never
    /// two.
    func testLastRenderedPageBreakSurvivesExactlyOnceAcrossAThreeWaySplit() throws {
        let (base, paragraphID) = paragraph(children: [runWithPageBreak("continued text here")])
        var log = OperationLog()
        log.append(
            .setRunFormatInRange(target: paragraphID, substring: "text", instance: 1,
                                  format: RunFormatPayload(bold: true)),
            source: .swift)

        let result = try OperationReducer.materialize(log: log, base: base)
        let runs = paragraphChildren(result, paragraphID: paragraphID)
            .filter { $0.kind == .element && $0.localName == "r" }

        XCTAssertEqual(runs.map(text(of:)), ["continued ", "text", " here"])
        XCTAssertTrue(runs.contains { text(of: $0) == "text" && isBold($0) })

        func hasPageBreak(_ run: XmlNode) -> Bool {
            run.children.contains { $0.kind == .element && $0.localName == "lastRenderedPageBreak" }
        }
        let withMarker = runs.filter(hasPageBreak)
        XCTAssertEqual(withMarker.count, 1, "lastRenderedPageBreak must appear on exactly one fragment")
        XCTAssertEqual(withMarker.first.map(text(of:)), "continued ",
                       "the marker (which precedes the text) stays on the FIRST fragment")
    }

    /// When the match starts at offset 0 (no prefix fragment produced), the
    /// marker stays on the matched "middle" piece — still the first
    /// fragment produced from that run — never duplicated onto the suffix.
    func testLastRenderedPageBreakStaysOnMatchedFragmentWhenMatchStartsAtOffsetZero() throws {
        let (base, paragraphID) = paragraph(children: [runWithPageBreak("continued text here")])
        var log = OperationLog()
        log.append(
            .setRunFormatInRange(target: paragraphID, substring: "continued", instance: 1,
                                  format: RunFormatPayload(bold: true)),
            source: .swift)

        let result = try OperationReducer.materialize(log: log, base: base)
        let runs = paragraphChildren(result, paragraphID: paragraphID)
            .filter { $0.kind == .element && $0.localName == "r" }

        XCTAssertEqual(runs.map(text(of:)), ["continued", " text here"])
        func hasPageBreak(_ run: XmlNode) -> Bool {
            run.children.contains { $0.kind == .element && $0.localName == "lastRenderedPageBreak" }
        }
        let withMarker = runs.filter(hasPageBreak)
        XCTAssertEqual(withMarker.count, 1)
        XCTAssertEqual(withMarker.first.map(text(of:)), "continued")
        XCTAssertTrue(withMarker.first.map(isBold) ?? false)
    }

    /// The page-break-bearing run's OWN text is now directly searchable —
    /// R2 treated it as "complex" (content = [lastRenderedPageBreak, t],
    /// count 2) and rejected any match touching it; R3 fixes this
    /// specifically, without touching how genuinely complex runs (tab/br/
    /// multiple <w:t>) are handled.
    func testPageBreakRunsOwnTextIsDirectlySearchableAcrossNeighbors() throws {
        let (base, paragraphID) = paragraph(children: [
            run("Some text "), runWithPageBreak("continued"), run(" more text"),
        ])
        var log = OperationLog()
        log.append(
            .setRunFormatInRange(target: paragraphID, substring: "text continued more", instance: 1,
                                  format: RunFormatPayload(bold: true)),
            source: .swift)
        XCTAssertNoThrow(try OperationReducer.materialize(log: log, base: base))
    }
}
