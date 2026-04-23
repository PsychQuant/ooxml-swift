import XCTest
@testable import OOXMLSwift

/// Phase B of `che-word-mcp-insert-crash-autosave-fix` (closes #41 prerequisite).
///
/// Confirms that `DocxReader.read(from:)` produces deterministic output across
/// repeated invocations against the same source bytes. This is a prerequisite
/// for `recover_from_autosave` — without parsing determinism, the recovered
/// in-memory state may differ from the pre-crash state in subtle ways that
/// the user cannot detect.
///
/// Spec: `openspec/changes/che-word-mcp-insert-crash-autosave-fix/specs/ooxml-content-insertion-primitives/spec.md`
/// Requirement: "ooxml-swift IO layer is fully serial; parallel primitives forbidden"
/// Scenario: "DocxReader.read uses serial chunk parsing" (deterministic-ordering clause)
final class DocxReaderDeterminismTests: XCTestCase {

    private var tempDir: URL!

    override func setUp() {
        super.setUp()
        tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("DocxReaderDeterminism-\(UUID().uuidString)")
        try? FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
    }

    override func tearDown() {
        try? FileManager.default.removeItem(at: tempDir)
        super.tearDown()
    }

    /// Build a fixture with > 256 paragraphs so the body parser exercises the
    /// chunked path (pre-v0.13.3 was the parallel chunked path; v0.13.3+ is
    /// the serial chunked path). Determinism is required either way.
    func testRepeatedReadsProduceIdenticalParagraphSequence() throws {
        var doc = WordDocument()
        for i in 0..<300 {
            doc.appendParagraph(Paragraph(text: "paragraph #\(i) — fixture content for determinism check"))
        }
        let url = tempDir.appendingPathComponent("determinism.docx")
        try DocxWriter.write(doc, to: url)

        var firstChildrenCount: Int?
        var firstParaTexts: [String]?

        for iter in 0..<5 {
            var loadedDoc = try DocxReader.read(from: url)
            defer { loadedDoc.close() }

            let count = loadedDoc.body.children.count
            let paraTexts = loadedDoc.body.children.compactMap { child -> String? in
                if case .paragraph(let p) = child {
                    return p.runs.map { $0.text }.joined()
                }
                return nil
            }

            if let expectedCount = firstChildrenCount {
                XCTAssertEqual(count, expectedCount,
                               "iter \(iter): body.children.count SHALL be deterministic across reads")
            } else {
                firstChildrenCount = count
            }

            if let expectedTexts = firstParaTexts {
                XCTAssertEqual(paraTexts, expectedTexts,
                               "iter \(iter): paragraph text sequence SHALL be deterministic across reads")
            } else {
                firstParaTexts = paraTexts
            }
        }

        XCTAssertEqual(firstChildrenCount, 300,
                       "Fixture should yield exactly 300 paragraphs")
    }
}
