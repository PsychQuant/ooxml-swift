import Foundation
import XCTest
@testable import OOXMLSwift

/// #152 — shared probe for the reader's shared extraction namespace
/// (`ZipHelper.readerNamespace`, aka `che-word-mcp`, under the process temp
/// directory).
///
/// The namespace is process-shared: a second `swift test` run, or a live MCP
/// server, extracts into the SAME directory concurrently, so a bare
/// before/after `Set` diff of directory names flaps. Every leak test
/// therefore writes a per-test marker into the fixture and only counts a
/// newly-appeared entry that is recognisably OURS as its own leak — `owns`
/// below is the recognition predicate, since different callers recognise
/// "theirs" differently (a marker inside `word/document.xml`, an
/// unreferenced marker media part, marker bytes anywhere inside the
/// extracted tree or a leaked private zip copy).
///
/// This logic previously existed three times, nearly verbatim, inside
/// `SidecarStoreTests`, `SyncOrchestratorTests` and `DocxWriterWriteDataTests`
/// (surfaced by #152's "test suite leaves ~292 reader namespace directories
/// per run" investigation). It is now a single implementation the three
/// suites — and any new leak test — share.
enum ReaderNamespaceLeakProbe {

    /// The shared reader-namespace root under the process temp directory.
    static var root: URL {
        FileManager.default.temporaryDirectory
            .appendingPathComponent(ZipHelper.readerNamespace)
    }

    /// The current entry names directly under `root`. Missing/unreadable
    /// root reads as empty (matches the three call sites' prior behavior via
    /// `try?`), not a hard error — this is a best-effort snapshot, not an
    /// assertion that the namespace exists.
    static func snapshot() -> Set<String> {
        Set((try? FileManager.default.contentsOfDirectory(atPath: root.path)) ?? [])
    }

    /// Entries that appeared since `before` AND are recognisably ours, per
    /// `owns`. Use this directly when neither convenience form below fits.
    static func leaked(since before: Set<String>, owns: (String) -> Bool) -> [String] {
        snapshot().subtracting(before).filter(owns).sorted()
    }

    /// Convenience form: ownership means "the extracted `word/document.xml`
    /// under this entry contains `marker`" (SidecarStoreTests' shape).
    static func leakedByDocumentMarker(since before: Set<String>, marker: String) -> [String] {
        leaked(since: before) { name in
            let part = root.appendingPathComponent(name).appendingPathComponent("word/document.xml")
            guard let data = try? Data(contentsOf: part) else { return false }
            return String(decoding: data, as: UTF8.self).contains(marker)
        }
    }

    /// Convenience form: ownership means "this entry's
    /// `word/media/<marker>.bin` unreferenced part exists" (SyncOrchestratorTests'
    /// shape — an unreferenced media part invisible to the document model,
    /// but present on disk purely to name the extracted archive).
    static func leakedByMediaMarkerFile(since before: Set<String>, marker: String) -> [String] {
        leaked(since: before) { name in
            FileManager.default.fileExists(atPath: root.appendingPathComponent(name)
                .appendingPathComponent("word/media/\(marker).bin").path)
        }
    }
}

extension XCTestCase {

    /// #152 — opens `url` through `DocxReader.read` and unconditionally
    /// frees its `che-word-mcp` extraction directory in `addTeardownBlock`
    /// (XCTest runs registered teardown blocks, in reverse registration
    /// order, after the test method returns and before `tearDown()`).
    ///
    /// Forgetting a manual `.close()` is the concrete failure mode #152
    /// investigated (a full `swift test` run left ~292 directories behind
    /// because tests opened documents through `DocxReader` and never called
    /// `close()`). A test that reads through this helper instead cannot
    /// forget — the close is registered before the read call returns.
    ///
    /// The returned `WordDocument` is an independent value copy (it's a
    /// struct); the teardown block holds its OWN copy captured at read time.
    /// Both copies' `preservedArchive` point at the same on-disk temp
    /// directory URL, so closing either one frees the shared resource —
    /// calling `.close()` again on the caller's copy after teardown has run
    /// is a safe no-op (`WordDocument.close()` is documented idempotent).
    @discardableResult
    func readAutoClosing(from url: URL, wireTreeBackedViews: Bool = false) throws -> WordDocument {
        var doc = try DocxReader.read(from: url, wireTreeBackedViews: wireTreeBackedViews)
        addTeardownBlock {
            doc.close()
        }
        return doc
    }
}
