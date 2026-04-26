import XCTest
@testable import OOXMLSwift

/// Unit tests for `DocxReader.parseParagraph` revision parsing.
///
/// Per the `docx-reader-top-level-revisions` change (Decision: Testability via
/// Internal Visibility), these tests exercise the parser directly via
/// `@testable import OOXMLSwift` against hand-constructed `XMLElement`
/// instances, bypassing the full .docx ZIP pipeline.
final class RevisionParsingTests: XCTestCase {

    // MARK: - Helpers

    /// Wrap the given inner XML in a `<w:p xmlns:w="...">...</w:p>` element
    /// and return the parsed root element. Traps on invalid XML.
    private func paragraphElement(inner: String) -> XMLElement {
        let wrapped = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        \(inner)
        </w:p>
        """
        let doc = try! XMLDocument(xmlString: wrapped, options: [])
        return doc.rootElement()!
    }

    /// Call `DocxReader.parseParagraph` on the given element with empty
    /// relationships / styles / numbering.
    private func parse(_ element: XMLElement) throws -> Paragraph {
        try DocxReader.parseParagraph(
            from: element,
            relationships: RelationshipsCollection(),
            styles: [],
            numbering: Numbering()
        )
    }

    /// Capture anything written to stderr during `block`.
    ///
    /// Carefully orders the fd dance so `readDataToEndOfFile()` does not
    /// deadlock: stderr (fd 2) is restored BEFORE reading from the pipe, so
    /// the pipe's write end is no longer held by the process. Then we close
    /// the owning `FileHandle` reference and read to EOF.
    private func captureStderr(_ block: () throws -> Void) rethrows -> String {
        fflush(stderr)
        let saved = dup(STDERR_FILENO)
        let pipe = Pipe()
        dup2(pipe.fileHandleForWriting.fileDescriptor, STDERR_FILENO)

        try block()

        fflush(stderr)
        // Restore stderr first — this releases the write end of the pipe
        // that fd 2 was pointing at, so readDataToEndOfFile will see EOF.
        dup2(saved, STDERR_FILENO)
        close(saved)
        pipe.fileHandleForWriting.closeFile()
        let data = pipe.fileHandleForReading.readDataToEndOfFile()
        return String(data: data, encoding: .utf8) ?? ""
    }

    // MARK: - moveFrom

    func testParsesMoveFromRevision() throws {
        let el = paragraphElement(inner: """
        <w:moveFrom w:id="3" w:author="Alice" w:date="2026-04-16T12:00:00Z">
            <w:r><w:t>moved source</w:t></w:r>
        </w:moveFrom>
        """)
        let paragraph = try parse(el)

        XCTAssertEqual(paragraph.revisions.count, 1)
        let rev = paragraph.revisions[0]
        XCTAssertEqual(rev.id, 3)
        XCTAssertEqual(rev.type, .moveFrom)
        XCTAssertEqual(rev.author, "Alice")
        XCTAssertEqual(rev.originalText, "moved source")
        XCTAssertNil(rev.newText)
    }

    func testMoveFromConcatenatesMultipleRuns() throws {
        let el = paragraphElement(inner: """
        <w:moveFrom w:id="7" w:author="Alice" w:date="2026-04-16T12:00:00Z">
            <w:r><w:t xml:space="preserve">first </w:t></w:r>
            <w:r><w:t>second</w:t></w:r>
        </w:moveFrom>
        """)
        let paragraph = try parse(el)

        XCTAssertEqual(paragraph.revisions.count, 1)
        XCTAssertEqual(paragraph.revisions[0].originalText, "first second")
    }

    // MARK: - moveTo

    func testParsesMoveToRevision() throws {
        let el = paragraphElement(inner: """
        <w:moveTo w:id="4" w:author="Alice" w:date="2026-04-16T12:00:00Z">
            <w:r><w:t>moved destination</w:t></w:r>
        </w:moveTo>
        """)
        let paragraph = try parse(el)

        XCTAssertEqual(paragraph.revisions.count, 1)
        let rev = paragraph.revisions[0]
        XCTAssertEqual(rev.id, 4)
        XCTAssertEqual(rev.type, .moveTo)
        XCTAssertEqual(rev.author, "Alice")
        XCTAssertEqual(rev.newText, "moved destination")
        XCTAssertNil(rev.originalText)
    }

    // MARK: - ins / del regression

    func testExistingInsDelStillParsed() throws {
        let el = paragraphElement(inner: """
        <w:ins w:id="1" w:author="Alice" w:date="2026-04-16T10:00:00Z">
            <w:r><w:t>inserted</w:t></w:r>
        </w:ins>
        <w:del w:id="2" w:author="Bob" w:date="2026-04-16T11:00:00Z">
            <w:r><w:delText>deleted</w:delText></w:r>
        </w:del>
        """)
        let paragraph = try parse(el)

        XCTAssertEqual(paragraph.revisions.count, 2)

        let insertion = paragraph.revisions[0]
        XCTAssertEqual(insertion.type, .insertion)
        XCTAssertEqual(insertion.author, "Alice")
        XCTAssertEqual(insertion.newText, "inserted")

        let deletion = paragraph.revisions[1]
        XCTAssertEqual(deletion.type, .deletion)
        XCTAssertEqual(deletion.author, "Bob")
        XCTAssertEqual(deletion.originalText, "deleted")
    }

    // MARK: - Within-paragraph order aggregation

    /// Covers the **Revision aggregation preserves revision order** requirement
    /// at the within-paragraph level. Cross-paragraph aggregation via
    /// `DocxReader.read(from:)` step 10 walks body.children in order and is
    /// covered indirectly by `DocxReaderIntegrationTests` round-trip tests —
    /// the aggregation step is unchanged by this change and is a simple index
    /// enumeration, so a dedicated cross-paragraph test here would duplicate
    /// that coverage.
    func testRevisionAggregationPreservesOrderWithinParagraph() throws {
        let el = paragraphElement(inner: """
        <w:ins w:id="10" w:author="Alice" w:date="2026-04-16T10:00:00Z">
            <w:r><w:t>first</w:t></w:r>
        </w:ins>
        <w:moveTo w:id="11" w:author="Alice" w:date="2026-04-16T10:01:00Z">
            <w:r><w:t>second</w:t></w:r>
        </w:moveTo>
        """)
        let paragraph = try parse(el)

        XCTAssertEqual(paragraph.revisions.count, 2)
        XCTAssertEqual(paragraph.revisions[0].type, .insertion)
        XCTAssertEqual(paragraph.revisions[1].type, .moveTo)
    }

    // MARK: - Debug logging

    func testDebugLoggingDisabledProducesNoOutput() throws {
        DocxReader.debugLoggingEnabled = false
        let el = paragraphElement(inner: "<w:customElement/>")

        let captured = try captureStderr {
            _ = try self.parse(el)
        }

        XCTAssertEqual(captured, "")
    }

    func testDebugLoggingEnabledEmitsOneLinePerUnknownElement() throws {
        DocxReader.debugLoggingEnabled = true
        defer { DocxReader.debugLoggingEnabled = false }

        let el = paragraphElement(inner: "<w:customElement/>")

        let captured = try captureStderr {
            _ = try self.parse(el)
        }

        // v0.19.0+ (PsychQuant/che-word-mcp#56) Phase 4: log message updated
        // from "skipped" → "captured" because unknown elements are now
        // preserved on `Paragraph.unrecognizedChildren` for round-trip
        // survival, not silently dropped. Position is the source-order index
        // assigned during the paragraph child walk.
        // v0.19.5+ (#56 R5 P0 #2): position counter starts at 1 (was 0) so
        // position 0 is the unambiguous "API-built sentinel" semantic.
        XCTAssertEqual(
            captured,
            "DocxReader.parseParagraph: captured unmodeled element customElement at position 1\n"
        )
    }
}
