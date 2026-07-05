import XCTest
@testable import OOXMLSwift

/// word-aligned-state-sync Phase 3 tasks 4.3–4.6 — sync primitives:
/// - 4.3 `ConflictReport` on overlapping mutations
/// - 4.4 `SyncPolicy` (`.abortOnConflict` / `.swiftWins` / `.wordWins` /
///   `.askUser(handler:)`, "Decision 7: Conflict policy is opt-in and explicit")
/// - 4.5 file watcher contract (mtime fast-path + SHA-256 confirmation)
/// - 4.6 Word lock-file interaction (`~$<filename>.docx`)
final class SyncPrimitivesTests: XCTestCase {

    private let pidA = ElementID(rawString: "w14:paraId=AAA")
    private let pidB = ElementID(rawString: "w14:paraId=BBB")

    private func pendingSwiftEntry(_ op: OOXMLSwift.Operation) -> LogEntry {
        LogEntry(opID: UUID(), op: op, source: .swift, timestamp: Date())
    }

    // MARK: - 4.3 Conflict detection

    func testOverlappingTextEditRaisesConflict() throws {
        // Spec scenario "Overlapping text edit raises conflict".
        let swiftEntry = pendingSwiftEntry(.setText(target: pidA, text: "swift_text"))
        let wordOps: [OOXMLSwift.Operation] = [.setText(target: pidA, text: "word_text")]

        let report = SyncConflict.detect(pendingSwiftOps: [swiftEntry], wordOps: wordOps)

        XCTAssertEqual(report.entries.count, 1)
        let entry = report.entries[0]
        XCTAssertEqual(entry.elementID, pidA)
        XCTAssertEqual(entry.swiftOpID, swiftEntry.opID)
        guard case .setText(_, let wordText) = entry.wordOp else {
            return XCTFail("conflict entry must carry the Word-inferred op")
        }
        XCTAssertEqual(wordText, "word_text")
    }

    func testNonOverlappingOpsProduceEmptyReport() throws {
        let swiftEntry = pendingSwiftEntry(.setText(target: pidA, text: "x"))
        let wordOps: [OOXMLSwift.Operation] = [.setText(target: pidB, text: "y")]

        let report = SyncConflict.detect(pendingSwiftOps: [swiftEntry], wordOps: wordOps)
        XCTAssertTrue(report.entries.isEmpty)
    }

    // MARK: - 4.4 SyncPolicy resolution

    func testAbortOnConflictThrowsStructuredError() throws {
        let swiftEntry = pendingSwiftEntry(.setText(target: pidA, text: "swift_text"))
        let wordOps: [OOXMLSwift.Operation] = [.setText(target: pidA, text: "word_text")]

        XCTAssertThrowsError(
            try SyncConflict.resolve(wordOps: wordOps, pendingSwiftOps: [swiftEntry],
                                     policy: .abortOnConflict)
        ) { error in
            guard case SyncError.conflict(let report) = error else {
                return XCTFail("expected SyncError.conflict, got \(error)")
            }
            XCTAssertEqual(report.entries.count, 1)
        }
    }

    func testSwiftWinsDropsConflictingWordOpsKeepsRest() throws {
        // Spec scenario "swiftWins drops Word's conflicting ops".
        let swiftEntry = pendingSwiftEntry(.setText(target: pidA, text: "swift_text"))
        let wordOps: [OOXMLSwift.Operation] = [
            .setText(target: pidA, text: "word_text"),      // conflicts
            .setText(target: pidB, text: "independent"),     // does not
        ]

        let kept = try SyncConflict.resolve(
            wordOps: wordOps, pendingSwiftOps: [swiftEntry], policy: .swiftWins)

        XCTAssertEqual(kept.count, 1)
        guard case .setText(let target, _) = kept[0] else {
            return XCTFail("expected surviving setText")
        }
        XCTAssertEqual(target, pidB, "only the non-conflicting Word op survives swiftWins")
    }

    func testWordWinsKeepsAllWordOps() throws {
        let swiftEntry = pendingSwiftEntry(.setText(target: pidA, text: "swift_text"))
        let wordOps: [OOXMLSwift.Operation] = [.setText(target: pidA, text: "word_text")]

        let kept = try SyncConflict.resolve(
            wordOps: wordOps, pendingSwiftOps: [swiftEntry], policy: .wordWins)
        XCTAssertEqual(kept.count, 1,
                       "wordWins keeps the conflicting Word op (append-order last-write-wins)")
    }

    func testAskUserHandlerDecidesPerElement() throws {
        // Spec scenario "askUser handler decides per element".
        let swiftA = pendingSwiftEntry(.setText(target: pidA, text: "swift_a"))
        let swiftB = pendingSwiftEntry(.setText(target: pidB, text: "swift_b"))
        let wordOps: [OOXMLSwift.Operation] = [
            .setText(target: pidA, text: "word_a"),
            .setText(target: pidB, text: "word_b"),
        ]
        var handlerSawEntries = 0

        let kept = try SyncConflict.resolve(
            wordOps: wordOps, pendingSwiftOps: [swiftA, swiftB],
            policy: .askUser(handler: { report in
                handlerSawEntries = report.entries.count
                // take Word's version for A, keep Swift's for B
                return [self.pidA: .takeWord, self.pidB: .takeSwift]
            }))

        XCTAssertEqual(handlerSawEntries, 2, "handler must see the full report")
        XCTAssertEqual(kept.count, 1)
        guard case .setText(let target, let text) = kept[0] else {
            return XCTFail("expected setText")
        }
        XCTAssertEqual(target, pidA)
        XCTAssertEqual(text, "word_a", ".takeWord keeps the Word op; .takeSwift drops it")
    }

    // MARK: - 4.5 File watcher contract

    func testMtimeOnlyTouchWithoutContentChangeIsIgnored() throws {
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("watch-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try Data("stable content".utf8).write(to: url)

        var detector = try DocxChangeDetector(url: url)
        // Touch: bump mtime without changing bytes.
        try FileManager.default.setAttributes(
            [.modificationDate: Date().addingTimeInterval(5)], ofItemAtPath: url.path)

        XCTAssertFalse(try detector.poll(),
                       "mtime-only change with identical content hash must not trigger")
    }

    func testContentChangeTriggers() throws {
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent("watch-\(UUID().uuidString).docx")
        defer { try? FileManager.default.removeItem(at: url) }
        try Data("before".utf8).write(to: url)

        var detector = try DocxChangeDetector(url: url)
        try Data("after — different bytes".utf8).write(to: url)

        XCTAssertTrue(try detector.poll(), "content hash change must trigger")
        XCTAssertFalse(try detector.poll(),
                       "after reporting once, the new state becomes the baseline")
    }

    // MARK: - 4.6 Word lock-file interaction

    func testLockFileURLDerivation() {
        let docx = URL(fileURLWithPath: "/tmp/thesis/report.docx")
        XCTAssertEqual(WordLock.lockFileURL(for: docx).lastPathComponent, "~$report.docx")
        XCTAssertEqual(WordLock.lockFileURL(for: docx).deletingLastPathComponent().path,
                       "/tmp/thesis")
    }

    func testIsLockedDetectsSpecLiteralLockFile() throws {
        let dir = FileManager.default.temporaryDirectory
            .appendingPathComponent("lock-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: dir) }
        let docx = dir.appendingPathComponent("report.docx")
        try Data("d".utf8).write(to: docx)

        XCTAssertFalse(WordLock.isLockedByWord(docx))
        try Data("l".utf8).write(to: dir.appendingPathComponent("~$report.docx"))
        XCTAssertTrue(WordLock.isLockedByWord(docx))
    }

    func testIsLockedDetectsWordMinusTwoVariant() throws {
        // Real Word for Mac derives the owner-file name by dropping the
        // first two characters of long filenames (8.3-era legacy):
        // `mydocument.docx` → `~$document.docx`.
        let dir = FileManager.default.temporaryDirectory
            .appendingPathComponent("lock-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: dir) }
        let docx = dir.appendingPathComponent("mydocument.docx")
        try Data("d".utf8).write(to: docx)

        try Data("l".utf8).write(to: dir.appendingPathComponent("~$document.docx"))
        XCTAssertTrue(WordLock.isLockedByWord(docx),
                      "the minus-two-characters Word owner-file variant must also be detected")
    }
}
