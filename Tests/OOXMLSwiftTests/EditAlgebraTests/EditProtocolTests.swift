// EditProtocolTests.swift
// EditAlgebra — Phase 2 of ooxml-edit-isomorphism-foundation (PsychQuant/macdoc#99)
//
// §1.2 + §1.3 of #105 tasks — protocol conformance smoke tests.
// Confirms scaffold is wired correctly: OOXMLEdit + WordEdit conform to
// Edit; EditError is throwable; per-case construction works.
//
// End-to-end behavior tests live in DocumentApplyTests (§2.x) and
// FullyFaithfulFunctorTests (§8.x).

import XCTest
@testable import OOXMLSwift

final class EditProtocolTests: XCTestCase {

    // MARK: - §1.2: OOXMLEdit conformance + construction

    func testOOXMLEditEnumExists() {
        let paraID = ElementID(libraryUUID: UUID())
        let runID = ElementID(libraryUUID: UUID())

        // All 5 canonical cases per design.md Decision 1 mapping table
        let cases: [OOXMLEdit] = [
            .insertParagraph(after: paraID, content: "hello", styleId: nil),
            .insertParagraphBefore(before: paraID, content: "world", styleId: "Heading1"),
            .setBold(target: runID, value: true),
            .insertHyperlink(target: runID, href: URL(string: "https://example.com")!, displayText: nil),
            .removeParagraph(target: paraID),
        ]

        XCTAssertEqual(cases.count, 5, "All 5 canonical OOXMLEdit cases should construct")

        // Edit conformance: each case can be type-erased to `any Edit`
        let edits: [any Edit] = cases
        XCTAssertEqual(edits.count, 5)
    }

    func testOOXMLEditLowerReturnsIdentity() {
        let runID = ElementID(libraryUUID: UUID())
        let edit = OOXMLEdit.setBold(target: runID, value: true)

        let lowered = edit.lower()
        XCTAssertEqual(lowered.count, 1, "OOXMLEdit.lower() is identity — returns [self]")
        XCTAssertEqual(lowered[0], edit, "Lowered element equals self")
    }

    func testOOXMLEditApplyThrowsNotImplementedInScaffold() {
        // Phase 2 §1 scaffold throws notImplemented for apply.
        // §2 (Document.apply wiring) makes this pass.
        let document = WordDocument()
        let edit = OOXMLEdit.removeParagraph(target: ElementID(libraryUUID: UUID()))

        XCTAssertThrowsError(try edit.apply(to: document)) { error in
            guard case EditError.notImplemented(let message) = error else {
                XCTFail("Expected .notImplemented, got \(error)")
                return
            }
            XCTAssertTrue(message.contains("§2"), "Error message references §2 task: \(message)")
        }
    }

    // MARK: - §1.3: WordEdit conformance + construction

    func testWordEditEnumExists() {
        let r1 = ElementID(libraryUUID: UUID())
        let r2 = ElementID(libraryUUID: UUID())
        let p1 = ElementID(libraryUUID: UUID())

        let range = WordRange(startRun: r1, startOffset: 0, endRun: r2, endOffset: 5)
        let paraRef = ParagraphRef(p1)

        // All 3 canonical cases per design.md Decision 2 mapping table
        let cases: [WordEdit] = [
            .applyBold(range: range),
            .applyLink(range: range, url: URL(string: "https://example.com")!),
            .applyInsertParagraph(after: paraRef, content: "hello"),
        ]

        XCTAssertEqual(cases.count, 3, "All 3 canonical WordEdit cases should construct")

        let edits: [any Edit] = cases
        XCTAssertEqual(edits.count, 3)
    }

    func testWordEditLowerReturnsEmptyInScaffold() {
        // Phase 2 §1 scaffold returns []. §7 lands per-case lower() bodies.
        // Naturality tests in §9 will catch when stub returns persist past §7.
        let range = WordRange(
            startRun: ElementID(libraryUUID: UUID()),
            startOffset: 0,
            endRun: ElementID(libraryUUID: UUID()),
            endOffset: 5
        )
        let edit = WordEdit.applyBold(range: range)
        XCTAssertTrue(edit.lower().isEmpty, "Scaffold WordEdit.lower() returns [] (TODO marker)")
    }

    func testWordRangeIsEquatable() {
        let r1 = ElementID(libraryUUID: UUID())
        let r2 = ElementID(libraryUUID: UUID())

        let range1 = WordRange(startRun: r1, startOffset: 0, endRun: r2, endOffset: 5)
        let range2 = WordRange(startRun: r1, startOffset: 0, endRun: r2, endOffset: 5)
        let range3 = WordRange(startRun: r1, startOffset: 0, endRun: r2, endOffset: 6)

        XCTAssertEqual(range1, range2)
        XCTAssertNotEqual(range1, range3)
    }

    // MARK: - EditError tests

    func testEditErrorCasesAreThrowable() {
        let runID = ElementID(libraryUUID: UUID())

        let errors: [EditError] = [
            .pathNotFound(runID),
            .preserveViolation(part: "word/document.xml", expected: "abc", actual: "xyz"),
            .unsupportedOperation("setBold requires Run target"),
            .notImplemented("§1 scaffold"),
            .operationLogFailure(underlying: "reducer failed"),
        ]
        XCTAssertEqual(errors.count, 5)

        for error in errors {
            XCTAssertThrowsError(try { throw error }())
        }
    }

    func testEditErrorEquatable() {
        let runID = ElementID(libraryUUID: UUID())
        let a: EditError = .pathNotFound(runID)
        let b: EditError = .pathNotFound(runID)
        XCTAssertEqual(a, b)

        let c: EditError = .pathNotFound(ElementID(libraryUUID: UUID()))
        XCTAssertNotEqual(a, c)
    }

    // §5 (insertHyperlink + wrapWithHyperlink emission) shipped in
    // macdoc#110 — no OOXMLEdit case stubs remain at the operations()
    // layer. The stub-mechanism test that asserted insertHyperlink
    // throws notImplemented has been deleted (per the original
    // "DELETE once §5 ships" comment).
}
