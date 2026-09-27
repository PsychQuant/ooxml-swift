// SetBoldInRangeTests.swift
// PsychQuant/macdoc#232 — paragraph+substring-addressed sibling of setBold.
//
// SCOPE: Validates OOXMLEdit.setBoldInRange → Operation.setRunFormatInRange
// emission (mirrors SetBoldTests.swift's coverage of the plain setBold
// case), plus an end-to-end WordDocument.apply reproduction of the exact
// failure macdoc#232 reported: a paragraph-targeted `set_bold` no longer
// throws `malformedOp("setRunFormat target must be <w:r>")`.

import XCTest
@testable import OOXMLSwift

final class SetBoldInRangeTests: XCTestCase {

    // MARK: - Emission: setBoldInRange → Operation.setRunFormatInRange

    func testSetBoldInRangeEmitsSetRunFormatInRangeWithBoldTrue() throws {
        let target = ElementID(libraryUUID: UUID())
        let edit = OOXMLEdit.setBoldInRange(target: target, substring: "world", value: true)

        let ops = try edit.operations()
        XCTAssertEqual(ops.count, 1, "setBoldInRange lowers to exactly one Operation")

        guard case .setRunFormatInRange(let opTarget, let substring, let instance, let payload) = ops[0] else {
            XCTFail("Expected Operation.setRunFormatInRange, got \(ops[0])")
            return
        }
        XCTAssertEqual(opTarget, target, "target: ElementID round-trips")
        XCTAssertEqual(substring, "world")
        XCTAssertEqual(instance, 1, "instance defaults to 1")
        XCTAssertEqual(payload.bold, true, "value: true lowers to RunFormatPayload.bold = true")
    }

    func testSetBoldInRangeFalseIsExplicitNotNil() throws {
        let target = ElementID(libraryUUID: UUID())
        let edit = OOXMLEdit.setBoldInRange(target: target, substring: "world", value: false)

        let ops = try edit.operations()
        guard case .setRunFormatInRange(_, _, _, let payload) = ops[0] else {
            XCTFail("Expected Operation.setRunFormatInRange")
            return
        }
        XCTAssertEqual(payload.bold, false,
                       "value: false → explicit RunFormatPayload.bold = false (NOT nil)")
        XCTAssertNotNil(payload.bold, "bold MUST be non-nil; nil means 'leave unchanged'")
    }

    func testSetBoldInRangeCustomInstancePassesThrough() throws {
        let target = ElementID(libraryUUID: UUID())
        let edit = OOXMLEdit.setBoldInRange(
            target: target, substring: "cat", value: true, instance: 2)

        let ops = try edit.operations()
        guard case .setRunFormatInRange(_, let substring, let instance, _) = ops[0] else {
            XCTFail("Expected Operation.setRunFormatInRange")
            return
        }
        XCTAssertEqual(substring, "cat")
        XCTAssertEqual(instance, 2)
    }

    func testSetBoldInRangeDoesNotTouchOtherFormatFields() throws {
        let target = ElementID(libraryUUID: UUID())
        let edit = OOXMLEdit.setBoldInRange(target: target, substring: "world", value: true)

        let ops = try edit.operations()
        guard case .setRunFormatInRange(_, _, _, let payload) = ops[0] else {
            XCTFail("Expected Operation.setRunFormatInRange")
            return
        }
        XCTAssertNil(payload.italic)
        XCTAssertNil(payload.underline)
        XCTAssertNil(payload.fontSizeHalfPoints)
        XCTAssertNil(payload.color)
    }

    func testSetBoldInRangeLowerReturnsSelf() {
        let target = ElementID(libraryUUID: UUID())
        let edit = OOXMLEdit.setBoldInRange(target: target, substring: "world", value: true)

        let lowered = edit.lower()
        XCTAssertEqual(lowered.count, 1, "OOXMLEdit.lower() is identity (returns [self])")
        XCTAssertEqual(lowered[0], edit, "OOXMLEdit.lower() returns self verbatim")
    }

    // MARK: - End-to-end: WordDocument.apply reproduces & fixes macdoc#232

    /// Reproduces the exact shape macdoc#232 reported: `EditPlanner` compiled
    /// `set_bold` into a paragraph-targeted edit, which `OOXMLEdit.setBold`
    /// (→ `Operation.setRunFormat`, which requires a `<w:r>` target) always
    /// rejected with `malformedOp`. `setBoldInRange` targets the SAME
    /// paragraph ElementID and succeeds, only bolding the substring match.
    func testSetBoldInRangeMutatesOnlyMatchedRunViaWordDocumentApply() throws {
        let paraUUID = UUID()
        let wt = XmlNode.element(prefix: "w", localName: "t", children: [XmlNode.text("Hello world today")])
        let wr = XmlNode.element(prefix: "w", localName: "r", children: [wt])
        let wp = XmlNode.element(prefix: "w", localName: "p", children: [wr])
        wp.libraryUUID = paraUUID
        let body = XmlNode.element(prefix: "w", localName: "body", children: [wp])
        let root = XmlNode.element(prefix: "w", localName: "document", children: [body])

        var doc = WordDocument()
        doc.xmlTrees["word/document.xml"] = XmlTree.synthesized(root: root)

        let paragraphID = ElementID(libraryUUID: paraUUID)

        // Sanity: this is exactly the call shape that used to throw.
        // `.setBold(target: paragraphID, value: true)` would fail here with
        // `malformedOp("setRunFormat target must be <w:r>")` because
        // paragraphID resolves to a <w:p>, not a <w:r>.
        XCTAssertThrowsError(try doc.apply(OOXMLEdit.setBold(target: paragraphID, value: true))) { error in
            guard case EditError.operationLogFailure(let underlying) = error else {
                XCTFail("Expected EditError.operationLogFailure, got \(error)")
                return
            }
            XCTAssertTrue(underlying.contains("setRunFormat target must be <w:r>"),
                          "confirms the ORIGINAL macdoc#232 failure mode, reason now visible: \(underlying)")
        }

        // The fix: setBoldInRange on the SAME paragraph target succeeds.
        let edit = OOXMLEdit.setBoldInRange(target: paragraphID, substring: "world", value: true)
        let result = try doc.apply(edit)

        guard let tree = result.xmlTrees["word/document.xml"],
              let paragraph = OperationReducer.findNode(elementID: paragraphID, in: tree) else {
            XCTFail("Expected paragraph in result")
            return
        }
        let runs = paragraph.children.filter { $0.kind == .element && $0.localName == "r" }
        func text(_ run: XmlNode) -> String {
            run.children
                .filter { $0.kind == .element && $0.localName == "t" }
                .flatMap { $0.children.filter { $0.kind == .text }.map(\.textContent) }
                .joined()
        }
        func isBold(_ run: XmlNode) -> Bool {
            guard let rPr = run.children.first(where: { $0.kind == .element && $0.localName == "rPr" })
            else { return false }
            return rPr.children.contains { $0.kind == .element && $0.localName == "b" }
        }

        XCTAssertEqual(runs.map(text), ["Hello ", "world", " today"])
        XCTAssertFalse(isBold(runs[0]))
        XCTAssertTrue(isBold(runs[1]), "only the matched substring's run fragment is bold")
        XCTAssertFalse(isBold(runs[2]))

        // Immutable apply: original doc's tree is untouched.
        let originalParagraph = doc.xmlTrees["word/document.xml"].flatMap {
            OperationReducer.findNode(elementID: paragraphID, in: $0)
        }
        XCTAssertEqual(
            originalParagraph?.children.filter { $0.kind == .element && $0.localName == "r" }.count, 1,
            "original doc still has exactly one (unsplit) run")
    }
}
