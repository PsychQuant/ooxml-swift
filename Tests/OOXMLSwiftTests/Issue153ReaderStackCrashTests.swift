import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Regression tests for [PsychQuant/ooxml-swift#153](https://github.com/PsychQuant/ooxml-swift/issues/153).
///
/// The issue reports `DocxReader.read` segfaulting in a **debug** build on a
/// real, deeply-nested-table document, reached through `macdoc docx apply`'s
/// async call path — not through a direct, main-thread call to `read` (the
/// coordinator's own follow-up correction on the issue narrows the repro to
/// that specific call context). Per #195 (the analogous WRITE-side crash,
/// `Table`/`TableCell`/`Paragraph` serialization), the mechanism is a
/// stack overflow: Swift concurrency's cooperative-pool worker threads have
/// a stack far smaller than the main thread's 8 MB (this suite uses 512 KB,
/// the same convention #195 established), and a `-Onone` (debug) build's
/// function frame is the SUM of every local declared anywhere in that
/// function's body — it does not shrink as the function progresses, and it
/// does not matter whether a given branch actually runs at runtime.
///
/// Two independent, compounding sources of stack pressure were found and
/// fixed here (see the accompanying fix report for full measurements):
///
/// 1. `DocxReader.read(from:wireTreeBackedViews:)` itself used to be one
///    ~830-line function — reading headers, footers, footnotes, endnotes,
///    core.xml, comments.xml, commentsExtended.xml, and more, ALL as
///    sequential inline blocks in the SAME function body. Even a minimal
///    document with NONE of those optional parts present still paid for
///    every one of their locals: measured baseline stack requirement before
///    the fix was north of 300 KB just to complete a ONE-level, no-optional-
///    parts read, before any table-nesting recursion even began. Split into
///    thirteen `@inline(never)` phase functions (`readHeaders`,
///    `readFooters`, `readFootnotesPart`, `readEndnotesPart`,
///    `readCorePropertiesPart`, `readCommentsPart`,
///    `linkCommentParagraphIndices`, `propagateAllRevisions`,
///    `readCommentsExtendedPart`, `calibrateNextBookmarkId`,
///    `wireTreeBackedBodyViews`, `sweepRemainingXMLParts`,
///    `readStylesPart`/`readNumberingPart`/`readSettingsPart`/
///    `parseBodyAndSectionProperties`), each phase's locals are freed the
///    moment that phase returns instead of persisting for the rest of the
///    read — measured baseline (depth 0) dropped from ~320 KB to ~210 KB.
/// 2. `parseTable`/`parseTableRow`/`parseTableCell`'s own per-level
///    recursive frames (the actual reader-side analogue of #195's
///    write-side finding): `parseTableCell`'s paragraph-parsing loop
///    declared a `para: Paragraph` (~25-stored-property struct) loop
///    variable that, under `-Onone`, stayed on the stack for the entire
///    `for nestedEl in element.elements(forName: "w:tbl")` recursion
///    below it — i.e. on every currently-open nesting level's frame, all
///    the way down. Likewise `parseTable`'s table-properties/conditional-
///    styles block (several locals, including a per-`tblStylePr` loop
///    variable) stayed live through the row-parsing loop that recurses
///    into nested tables. Both moved into `@inline(never)` helpers
///    (`cellPropertiesAndParagraphs`, `parseTablePropertiesBlock`) that
///    return already-assembled values, freeing those slots before the
///    recursive call begins — the same technique #195 used for
///    `TableCell.toXML(depth:)`'s `bodyAndTrailingParagraphXML` and
///    `Table.toXML(depth:)`'s `tblGridXML`. This dropped the per-level
///    cost from ~32 KB/level to ~24 KB/level.
///
/// Measured thresholds (debug build, binary-search on `Thread.stackSize`,
/// full `DocxReader.read` including ZIP unzip, relationships, and body
/// parsing — not an isolated helper call; "depth N" = N levels of table
/// nesting below the outermost table, i.e. the deepest `parseTable` depth
/// argument reached is N):
///
/// | Scenario | Before fix | After fix | Margin at 512 KB (after) |
/// |---|---:|---:|---:|
/// | depth 0 (single table, no optional parts) | safe from ~320 KB | safe from ~208 KB | ~59% |
/// | depth 5 (reader's own `MAX_TABLE_NEST_DEPTH` ceiling) | safe from ~496 KB (crashes below — as little as ~4% margin at an assumed 512 KB production stack) | safe from ~336 KB | ~34% |
///
/// Depth 5 is the deepest a document can ever reach through the normal
/// `DocxReader.read` entry point — `parseTable`'s own `MAX_TABLE_NEST_DEPTH`
/// guard throws (cleanly, no crash) before parsing anything deeper, so a
/// depth-5 fixture with real, source-parsed paragraph content in every cell
/// (not API-constructed) is the worst case this reader's own guard rail
/// allows to reach `DocxWriter`/caller code. The pre-fix depth-5 number
/// (~496 KB) is the direct explanation for #153 itself: on an assumed
/// 512 KB production cooperative-pool stack, that leaves only ~16 KB of
/// margin BEFORE accounting for whatever `Executor.apply`'s own async call
/// chain consumes above `DocxReader.read` in the original report — which
/// the issue's own follow-up correction could not isolate further
/// ("Something about the call context in Executor.apply ... is required").
final class Issue153ReaderStackCrashTests: XCTestCase {

    // MARK: - Fixture construction

    /// Builds a minimal `.docx` whose `word/document.xml` is exactly
    /// `documentXML`.
    private func buildFixture(documentXML: String) throws -> URL {
        let staging = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue153-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: staging, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: staging) }

        func write(_ content: String, to relativePath: String) throws {
            let url = staging.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(), withIntermediateDirectories: true)
            try content.write(to: url, atomically: true, encoding: .utf8)
        }

        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="xml" ContentType="application/xml"/>
                <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>
            """, to: "[Content_Types].xml")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """, to: "_rels/.rels")
        try write(documentXML, to: "word/document.xml")

        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue153-\(UUID().uuidString).docx")
        do {
            let archive = try Archive(url: docxURL, accessMode: .create)
            let base = staging.resolvingSymlinksInPath().path
            let enumerator = FileManager.default.enumerator(
                at: staging, includingPropertiesForKeys: [.isDirectoryKey])!
            for case let fileURL as URL in enumerator {
                let isDir = (try? fileURL.resourceValues(forKeys: [.isDirectoryKey]))?.isDirectory ?? false
                if isDir { continue }
                let entry = String(fileURL.resolvingSymlinksInPath().path.dropFirst(base.count + 1))
                try archive.addEntry(with: entry, fileURL: fileURL, compressionMethod: .deflate)
            }
        } catch {
            try? FileManager.default.removeItem(at: docxURL)
            throw error
        }
        return docxURL
    }

    /// `levels` levels of table-in-single-cell nesting, one real
    /// (source-parsed, not API-constructed) paragraph of text per level —
    /// matches what a real document's reader path exercises: every run
    /// gets a source `position` from `DocxReader`, so cell content routes
    /// through the same code paths a genuine nested-table document would.
    private func makeNLevelNestedTableXML(levels: Int) -> String {
        var open = ""
        var close = ""
        for i in 0..<levels {
            open += "<w:tbl><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr>"
            open += "<w:p><w:r><w:t>L\(i)</w:t></w:r></w:p>"
            close = "<w:p/></w:tc></w:tr></w:tbl>" + close
        }
        return """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
            <w:p><w:r><w:t>Unrelated</w:t></w:r></w:p>\
            \(open)<w:p><w:r><w:t>Leaf</w:t></w:r></w:p>\(close)\
            <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>\
            </w:body></w:document>
            """
    }

    // MARK: - Stack-constrained execution

    /// Runs `body` on a `Thread` with the given stack size and asserts it
    /// completes within `timeout` seconds. A crash inside `body` kills the
    /// whole test PROCESS (matching the debug-build `SIGSEGV`/`SIGBUS` this
    /// issue is about) — `swift test` reports the executable exiting with
    /// an unexpected signal, not a normal assertion failure. There is no
    /// in-process way to catch that and convert it to a soft `XCTFail` —
    /// the crash IS the evidence. Same helper shape as
    /// `Issue195TwoLevelNestedTableCrashTests.runOnStackConstrainedThread`.
    private func runOnStackConstrainedThread(
        stackKB: Int, timeout: TimeInterval = 30, _ body: @escaping () throws -> Void
    ) throws {
        let sem = DispatchSemaphore(value: 0)
        var caught: Error?
        let thread = Thread {
            do { try body() } catch { caught = error }
            sem.signal()
        }
        thread.stackSize = stackKB * 1024
        thread.start()
        let waitResult = sem.wait(timeout: .now() + timeout)
        XCTAssertEqual(waitResult, .success, "stack=\(stackKB)KB: timed out (hang)")
        if let caught { throw caught }
    }

    // MARK: - Modeling the call context the issue's own correction could not isolate
    //
    // R2 (independent review follow-up): the issue's own correction comment
    // narrows the repro to a SPECIFIC call context — "macdoc docx apply
    // reaches the reader through Executor.apply, inside swift-argument-
    // parser's async main. Reading the file in isolation is not the same
    // path, and only the full path crashes... Something about the call
    // context in Executor.apply (async task stack, or the deeper frame
    // chain below it) is required." The tests above call `DocxReader.read`
    // directly at the top of a freshly-started 512KB thread — the FULL
    // budget is available when `read` starts. They do not model whatever
    // `Executor.apply` and everything above it (swift-argument-parser's
    // async main, manifest/argument processing, …) consumes BEFORE ever
    // calling into `DocxReader.read`.
    //
    // `Executor.apply` cannot be invoked from here: it lives in the
    // downstream `docx-workflow-swift` package (per the issue's own
    // backtrace: "Executor.apply(manifest:baselineURL:outputURL:warnHandler:)
    // (docx-workflow-swift)"). `docx-workflow-swift` depends ON ooxml-swift
    // (this package's `Package.swift` declares exactly one dependency,
    // ZIPFoundation — there is no `Executor` type anywhere in this
    // checkout); adding a test-only dependency the other direction would be
    // circular and is out of this fix's scope regardless.
    //
    // Two alternative, constructible-from-here models of the SAME two
    // hypotheses the issue's correction raises:
    //
    // 1. "async task stack" — `testFiveLevelNestingThroughSwiftConcurrencyTaskDoesNotCrash`
    //    runs the read through a REAL `Task { }` (Swift concurrency's actual
    //    cooperative-thread-pool mechanism — the same kind of async context
    //    "swift-argument-parser's async main" is), not a manually-sized
    //    `Thread`, so it exercises the genuine runtime-determined stack
    //    Swift concurrency worker threads get, whatever that is on this
    //    toolchain/OS, rather than the 512KB figure this suite otherwise
    //    assumes by convention (from #195).
    // 2. "the deeper frame chain below it" — `withSyntheticCallerFrames`
    //    below models an UNKNOWN caller chain (whatever `Executor.apply` and
    //    everything above it does before reaching `DocxReader.read`) as a
    //    tunable number of additional, genuinely stack-consuming recursive
    //    frames, run on the SAME 512KB thread BEFORE the read call — so the
    //    fresh-512KB assumption the other tests make is no longer free.
    //    `testFiveLevelNestingWithSyntheticCallerOverheadDoesNotCrash` pins
    //    a frame count comfortably inside the measured safe range (see its
    //    doc comment for the bisected threshold) as a regression floor.

    /// Recurses `framesRemaining` times, each frame holding a genuinely
    /// stack-resident ~256-byte local (32 `Double`s — under `-Onone` this is
    /// not optimized away, the same "frame size is the sum of every local"
    /// mechanism #153/#195 measured elsewhere), before calling `body` at the
    /// bottom. A synthetic, tunable stand-in for however much unknown
    /// caller-chain stack `Executor.apply` (and whatever calls it) actually
    /// costs in the real crash — a quantity the issue's own correction
    /// explicitly could not isolate. `@inline(never)` so the compiler cannot
    /// fold the recursion away or reuse a frame.
    @inline(never)
    private func withSyntheticCallerFrames<T>(
        _ framesRemaining: Int, _ body: () throws -> T
    ) rethrows -> T {
        guard framesRemaining > 0 else { return try body() }
        let filler = (1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0,
                      9.0, 10.0, 11.0, 12.0, 13.0, 14.0, 15.0, 16.0,
                      17.0, 18.0, 19.0, 20.0, 21.0, 22.0, 23.0, 24.0,
                      25.0, 26.0, 27.0, 28.0, 29.0, 30.0, 31.0, 32.0)
        _ = filler.0
        return try withSyntheticCallerFrames(framesRemaining - 1, body)
    }

    /// Model 1: the read happens inside a REAL `Task { }` — Swift
    /// concurrency's actual cooperative-thread-pool mechanism, the same
    /// KIND of async context `Executor.apply`'s caller (swift-argument-
    /// parser's async `main()`) runs in — rather than a manually-sized
    /// `Thread`. This does not let the test choose the stack size (Swift
    /// concurrency worker threads' stack size is runtime/OS-determined, not
    /// test-controllable), so it is a complement to, not a replacement for,
    /// the explicit-512KB tests above: it exercises the genuine mechanism
    /// the issue names ("async task stack") on whatever this toolchain
    /// actually provides.
    ///
    /// Honest result, checked against the PRE-fix `read()` (the ~830-line
    /// monolith) on this toolchain/OS: this test passes even without the
    /// fix — this toolchain's actual cooperative-pool worker stack is
    /// evidently larger than the 512KB this suite otherwise assumes by
    /// convention, so "async task stack" alone (on THIS platform) does not
    /// reproduce the crash; it takes real caller-chain depth too (see
    /// Model 2 below, which DOES fail pre-fix). Kept anyway: it is the
    /// actual mechanism the issue names, pins present/future behavior
    /// under genuine `Task` execution, and would matter on a
    /// platform/toolchain where the cooperative pool's stack is smaller.
    func testFiveLevelNestingThroughSwiftConcurrencyTaskDoesNotCrash() throws {
        let xml = makeNLevelNestedTableXML(levels: 6)
        let url = try buildFixture(documentXML: xml)
        defer { try? FileManager.default.removeItem(at: url) }

        let sem = DispatchSemaphore(value: 0)
        var caught: Error?
        var readCount: Int?
        Task {
            do {
                var doc = try DocxReader.read(from: url)
                if case .table(let table) = doc.body.children[1] {
                    readCount = self.countNestedTables(table)
                }
                doc.close()
            } catch {
                caught = error
            }
            sem.signal()
        }
        XCTAssertEqual(sem.wait(timeout: .now() + 30), .success, "timed out (hang)")
        if let caught { throw caught }
        XCTAssertEqual(readCount, 6, "all 6 nesting levels must survive a read done inside a real Task")
    }

    /// Model 2: bisected threshold (debug build, 512KB thread, depth-5
    /// fixture) — synthetic caller-chain overhead beyond which even the
    /// FIXED code crashes: safe at 600 frames, crashes at 650 (binary
    /// search step 50). That is a substantial amount of headroom above the
    /// ~176KB (512KB total − ~336KB depth-5's own measured floor) this
    /// suite's other numbers would predict, consistent with each frame
    /// costing meaningfully less than the naive "32 Doubles = 256 bytes"
    /// estimate once call overhead and register allocation are accounted
    /// for. This test pins 400 frames — comfortably (33%) below the
    /// measured 600-frame safe boundary — as a regression floor: if the
    /// fix's own frame-splitting ever regresses and the depth-5 margin
    /// shrinks, this is the test that would catch it FIRST, before the
    /// plain (no synthetic overhead) variant above does, because it no
    /// longer assumes the full 512KB is available fresh when
    /// `DocxReader.read` starts.
    func testFiveLevelNestingWithSyntheticCallerOverheadDoesNotCrash() throws {
        let xml = makeNLevelNestedTableXML(levels: 6)
        let url = try buildFixture(documentXML: xml)
        defer { try? FileManager.default.removeItem(at: url) }

        var doc: WordDocument?
        try runOnStackConstrainedThread(stackKB: 512) {
            try self.withSyntheticCallerFrames(400) {
                doc = try DocxReader.read(from: url)
            }
        }
        var document = try XCTUnwrap(doc)
        defer { document.close() }
        guard case .table(let table) = document.body.children[1] else {
            XCTFail("expected table at body.children[1]")
            return
        }
        XCTAssertEqual(countNestedTables(table), 6)
    }

    // MARK: - Tests

    /// The reader's own worst legitimate case: 6 levels of table-in-cell
    /// nesting (the outermost table plus 5 nested — reaching depth 5, the
    /// `MAX_TABLE_NEST_DEPTH` ceiling), real source-parsed paragraph content
    /// at every level, full `DocxReader.read` (not an isolated helper call)
    /// on a 512KB thread. Before the fix this crashed (SIGBUS/SIGSEGV) at
    /// every stack size this suite tests, including 512KB itself; after the
    /// fix it completes with comfortable headroom (measured safe from
    /// ~336KB, ~34% margin at 512KB — see the fix report for the full
    /// measurement table).
    func testFiveLevelNestingFullReadSmallStackDoesNotCrash() throws {
        let xml = makeNLevelNestedTableXML(levels: 6)
        let url = try buildFixture(documentXML: xml)
        defer { try? FileManager.default.removeItem(at: url) }

        var doc: WordDocument?
        try runOnStackConstrainedThread(stackKB: 512) {
            doc = try DocxReader.read(from: url)
        }
        var document = try XCTUnwrap(doc)
        defer { document.close() }

        guard case .table(let table) = document.body.children[1] else {
            XCTFail("expected table at body.children[1]")
            return
        }
        XCTAssertEqual(countNestedTables(table), 6, "all 6 nesting levels (depths 0...5, the MAX_TABLE_NEST_DEPTH ceiling) must be present")
        for i in 0..<6 {
            XCTAssertTrue(containsText(table, "L\(i)"), "missing L\(i)")
        }
        XCTAssertTrue(containsText(table, "Leaf"))
    }

    /// Same fixture and stack size, but with `wireTreeBackedViews: true`
    /// (che-word-mcp's `open_document` op-log path) — exercises the
    /// tree-wiring phase (`wireTreeBackedBodyViews`) too.
    func testFiveLevelNestingTreeBackedReadSmallStackDoesNotCrash() throws {
        let xml = makeNLevelNestedTableXML(levels: 6)
        let url = try buildFixture(documentXML: xml)
        defer { try? FileManager.default.removeItem(at: url) }

        var doc: WordDocument?
        try runOnStackConstrainedThread(stackKB: 512) {
            doc = try DocxReader.read(from: url, wireTreeBackedViews: true)
        }
        var document = try XCTUnwrap(doc)
        defer { document.close() }

        guard case .table(let table) = document.body.children[1] else {
            XCTFail("expected table at body.children[1]")
            return
        }
        XCTAssertEqual(countNestedTables(table), 6)
    }

    /// One level of nesting (the common case — most real documents never
    /// approach the reader's depth ceiling at all) with NO optional parts
    /// (styles/numbering/settings/headers/footers/footnotes/endnotes/
    /// comments all absent) on a 512KB thread. This is the "fixed cost"
    /// floor: before the fix, `read()`'s own ~830-line monolithic body
    /// already needed north of 300KB for this minimal case, before any
    /// table recursion. After the fix, safe from ~208KB — comfortable
    /// (~59%) margin even for documents nowhere near the nesting ceiling.
    func testSingleLevelMinimalDocumentSmallStackDoesNotCrash() throws {
        let xml = makeNLevelNestedTableXML(levels: 1)
        let url = try buildFixture(documentXML: xml)
        defer { try? FileManager.default.removeItem(at: url) }

        var doc: WordDocument?
        try runOnStackConstrainedThread(stackKB: 512) {
            doc = try DocxReader.read(from: url)
        }
        var document = try XCTUnwrap(doc)
        defer { document.close() }
        guard case .table(let table) = document.body.children[1] else {
            XCTFail("expected table at body.children[1]")
            return
        }
        XCTAssertTrue(containsText(table, "L0"))
    }

    /// Boundary behavior lock-down: 7 levels of nesting (the outermost table
    /// plus 6 nested — reaching depth 6, one past `MAX_TABLE_NEST_DEPTH`)
    /// must still fail with the existing, clean `invalidDocx` refusal — NOT
    /// a crash — confirming the frame-splitting fix did not accidentally
    /// weaken or remove `parseTable`'s own depth guard.
    func testSevenLevelNestingStillRefusesCleanlyNotCrash() throws {
        let xml = makeNLevelNestedTableXML(levels: 7)
        let url = try buildFixture(documentXML: xml)
        defer { try? FileManager.default.removeItem(at: url) }

        XCTAssertThrowsError(try DocxReader.read(from: url)) { error in
            guard case WordError.invalidDocx(let message) = error else {
                XCTFail("expected WordError.invalidDocx, got \(error)")
                return
            }
            XCTAssertTrue(message.contains("5"), "expected message to name the depth limit, got: \(message)")
        }
    }

    // MARK: - Helpers

    private func countNestedTables(_ table: Table) -> Int {
        var count = 1
        for row in table.rows {
            for cell in row.cells {
                for nested in cell.nestedTables {
                    count += countNestedTables(nested)
                }
            }
        }
        return count
    }

    private func containsText(_ table: Table, _ needle: String) -> Bool {
        for row in table.rows {
            for cell in row.cells {
                for para in cell.paragraphs where para.getText().contains(needle) {
                    return true
                }
                for nested in cell.nestedTables where containsText(nested, needle) {
                    return true
                }
            }
        }
        return false
    }
}
