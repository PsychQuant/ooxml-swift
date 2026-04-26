import Foundation
import XCTest
@testable import OOXMLSwift

/// Writes the given `WordDocument` to a unique temp `.docx`, reads it back,
/// and returns the re-read instance. Cleans up the temp file via `defer`.
///
/// Required by Issue #56 R5 stack-completion (`ooxml-roundtrip-fidelity`
/// Requirement: "Issue #56 R5 stack regression tests SHALL exercise full save
/// then re-read roundtrip"). Every R5 regression test SHALL run its assertions
/// against the result of this helper, not against the in-memory document, so
/// that the writer path participates in the regression coverage.
func roundtrip(_ document: WordDocument, file: StaticString = #file, line: UInt = #line) throws -> WordDocument {
    let tmpURL = FileManager.default.temporaryDirectory
        .appendingPathComponent("r5-roundtrip-\(UUID().uuidString).docx")
    try DocxWriter.write(document, to: tmpURL)
    defer { try? FileManager.default.removeItem(at: tmpURL) }
    return try DocxReader.read(from: tmpURL)
}
