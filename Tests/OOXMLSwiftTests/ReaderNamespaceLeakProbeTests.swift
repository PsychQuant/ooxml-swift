import XCTest
@testable import OOXMLSwift

/// #152 — regression coverage for the `readAutoClosing` teardown-registered
/// close helper (`Tests/OOXMLSwiftTests/Helpers/ReaderNamespaceLeakProbe.swift`).
///
/// The concrete failure #152 investigated: a full `swift test` run left
/// ~292 `che-word-mcp` reader-namespace directories behind because tests
/// opened documents through `DocxReader` and never called `close()`. The fix
/// sketch was a test helper that opens a document and registers `close()` in
/// `addTeardownBlock`. This file pins that the helper actually reclaims the
/// directory — not merely that it compiles.
final class ReaderNamespaceLeakProbeTests: XCTestCase {

    private func makeMarkedFixture(marker: String) throws -> URL {
        let dir = FileManager.default.temporaryDirectory
            .appendingPathComponent("leak-probe-fixture-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: marker))
        let url = dir.appendingPathComponent("fixture.docx")
        try DocxWriter.write(doc, to: url)
        return url
    }

    /// `readAutoClosing` must have actually run its registered `close()` by
    /// the time ITS OWN teardown block fires — not merely scheduled it.
    ///
    /// XCTest runs teardown blocks in LIFO order (reverse of registration),
    /// after the test method body returns and before `tearDown()`. This
    /// assertion is itself registered via `addTeardownBlock` BEFORE the
    /// `readAutoClosing` call below, so in LIFO order it runs LAST — i.e.
    /// strictly after `readAutoClosing`'s own close-on-teardown block has
    /// already executed. That ordering is what lets this test observe "the
    /// close already happened" rather than "the close was merely scheduled".
    func testReadAutoClosingReclaimsNamespaceEntryByTeardown() throws {
        let marker = "leak-probe-marker-\(UUID().uuidString)"
        let url = try makeMarkedFixture(marker: marker)
        addTeardownBlock {
            try? FileManager.default.removeItem(at: url.deletingLastPathComponent())
        }

        addTeardownBlock {
            let stillPresent = ReaderNamespaceLeakProbe.leaked(since: []) { name in
                let part = ReaderNamespaceLeakProbe.root.appendingPathComponent(name)
                    .appendingPathComponent("word/document.xml")
                guard let data = try? Data(contentsOf: part) else { return false }
                return String(decoding: data, as: UTF8.self).contains(marker)
            }
            XCTAssertTrue(stillPresent.isEmpty,
                "readAutoClosing's teardown-registered close() must already have run and "
                    + "removed its extraction directory by the time this later-firing (LIFO) "
                    + "teardown block checks; still present: \(stillPresent)")
        }

        // Sanity: readAutoClosing must actually have created a namespace
        // entry in the first place (proves the probe below isn't vacuously
        // passing because nothing was ever extracted).
        let before = ReaderNamespaceLeakProbe.snapshot()
        _ = try readAutoClosing(from: url)
        let immediatelyAfter = ReaderNamespaceLeakProbe.leakedByDocumentMarker(since: before, marker: marker)
        XCTAssertEqual(immediatelyAfter.count, 1,
            "readAutoClosing must have extracted exactly one namespace entry for this fixture "
                + "before its close() teardown block runs")
    }
}
