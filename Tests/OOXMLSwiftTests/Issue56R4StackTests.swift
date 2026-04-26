import XCTest
@testable import OOXMLSwift

/// Issue #56 R5 stack-completion regression tests.
///
/// Each test in this file targets one of the 6 P0 + 7 P1 findings from R4
/// verify (BLOCK verdict, 6 reviewers). Per the design Decision: every test
/// SHALL exercise full save→re-read roundtrip via `roundtrip(_:)` from
/// `Helpers/RoundtripHelper.swift` so the writer path participates in the
/// regression coverage. The R3 stack's all-in-memory test pattern was the
/// proven blind spot of the R2→R3→R4 cycle.
///
/// Tests are added per-task as the R5 stack proceeds:
/// - §2 (P0 #1): mixed-content revision wrapper across all parts
/// - §3 (P0 #2): SDT position ≥ 1 reader assignment
/// - §4 (P0 #3): XML attribute escape sweep
/// - §5 (P0 #4): block-level SDT typed Revision propagation
/// - §6 (P0 #5): Document.replaceText container-symmetric surface walk
/// - §7 (P0 #6): container parser w:tbl capture
/// - §8 (P1 batch)
final class Issue56R4StackTests: XCTestCase {

    // MARK: - Tests added per task — see commit history per fix(#56-r5-p0-N)

    func testPlaceholder_RemovedAfterFirstRealTestLands() {
        // Stub to keep the suite buildable while §2-§8 tests are being added
        // task-by-task. The first real R5 test (task 2.1) replaces this.
        XCTAssertTrue(true)
    }
}

// MARK: - Allow-list audit table for emit sites NOT routed through escapeXMLAttribute
//
// Per Issue #56 R5 stack-completion spec `xml-attribute-escape` Requirement 3:
// "Issue56R4StackTests SHALL include an allow-list audit table for emit sites
// NOT routed through escapeXMLAttribute". This is the explicit allow-list of
// exemptions, NOT a deny-list claiming "all sites covered". A reviewer can
// verify by checking that every emit site is either (a) routed through
// `escapeXMLAttribute(_:)` from `XMLAttributeEscape.swift`, or (b) named here
// with rationale.
//
// Format: `<file>:<line(s)>` — <rationale>
//
// Initial allow-list (populated as §4 sweep proceeds):
// - (none — every emit site SHALL be routed through escapeXMLAttribute or
//   listed here with rationale before §4.13 is marked done)
