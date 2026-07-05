import Foundation
import XCTest
@testable import OOXMLSwift

/// word-aligned-state-sync Phase 3 task 4.2 (+ 4.9 regression pin) —
/// `ooxml-word-sync` Requirement "Word-import diff via element identity
/// matching" ("Decision 6: Word-import diff via structural element-identity
/// matching").
///
/// Scenarios pinned verbatim from the spec:
/// - Matched-by-ID element with text change produces SetText
/// - rsid-only difference produces empty op set (4.9)
/// - New paragraph in Word produces InsertParagraphAfter
final class WordImportDiffTests: XCTestCase {

    private func tree(_ bodyChildren: String) throws -> XmlTree {
        let xml = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>\(bodyChildren)</w:body></w:document>
            """
        return try XmlTreeReader.parse(Data(xml.utf8))
    }

    // MARK: - Spec scenario: matched-by-ID text change → SetText

    func testMatchedByIDTextChangeProducesSetText() throws {
        let snapshot = try tree(
            #"<w:p w14:paraId="ABC"><w:r><w:t>original</w:t></w:r></w:p>"#)
        let current = try tree(
            #"<w:p w14:paraId="ABC"><w:r><w:t>modified</w:t></w:r></w:p>"#)

        let diff = WordImport.diff(snapshot: snapshot, current: current)

        XCTAssertEqual(diff.operations.count, 1)
        guard case .setText(let target, let text) = diff.operations.first else {
            return XCTFail("expected SetText, got \(diff.operations)")
        }
        XCTAssertEqual(target.raw, "w14:paraId=ABC",
                       "op must reference the paragraph by its stable paraId")
        XCTAssertEqual(text, "modified")
    }

    // MARK: - Spec scenario (4.9): rsid-only difference → empty op set

    func testRsidOnlyDifferenceProducesEmptyOpSet() throws {
        let snapshot = try tree(
            #"<w:p w14:paraId="ABC" w:rsidR="00AAA111" w:rsidRDefault="00AAA111"><w:r><w:t>stable</w:t></w:r></w:p><w:p w14:paraId="DEF" w:rsidR="00BBB222"><w:r><w:t>also stable</w:t></w:r></w:p>"#)
        let current = try tree(
            #"<w:p w14:paraId="ABC" w:rsidR="00CCC333" w:rsidRDefault="00CCC333"><w:r><w:t>stable</w:t></w:r></w:p><w:p w14:paraId="DEF" w:rsidR="00DDD444"><w:r><w:t>also stable</w:t></w:r></w:p>"#)

        let diff = WordImport.diff(snapshot: snapshot, current: current)

        XCTAssertTrue(diff.operations.isEmpty,
                      "rsid renumbering with no content change must produce an empty op set")
        XCTAssertTrue(diff.unrepresentedChanges.isEmpty,
                      "rsid noise must not surface as unrepresented changes either")
    }

    // MARK: - Spec scenario: new paragraph → InsertParagraphAfter

    func testNewParagraphProducesInsertParagraphAfter() throws {
        let snapshot = try tree(
            #"<w:p w14:paraId="P1"><w:r><w:t>first</w:t></w:r></w:p><w:p w14:paraId="P2"><w:r><w:t>second</w:t></w:r></w:p>"#)
        let current = try tree(
            #"<w:p w14:paraId="P1"><w:r><w:t>first</w:t></w:r></w:p><w:p w14:paraId="PNEW"><w:r><w:t>brand new</w:t></w:r></w:p><w:p w14:paraId="P2"><w:r><w:t>second</w:t></w:r></w:p>"#)

        let diff = WordImport.diff(snapshot: snapshot, current: current)

        XCTAssertEqual(diff.operations.count, 1)
        guard case .insertParagraphAfter(let after, let payload) = diff.operations.first else {
            return XCTFail("expected InsertParagraphAfter, got \(diff.operations)")
        }
        XCTAssertEqual(after.raw, "w14:paraId=P1",
                       "insert must anchor on the preceding matched paragraph")
        XCTAssertEqual(payload.text, "brand new")
    }

    func testNewFirstParagraphProducesInsertParagraphBefore() throws {
        let snapshot = try tree(
            #"<w:p w14:paraId="P1"><w:r><w:t>first</w:t></w:r></w:p>"#)
        let current = try tree(
            #"<w:p w14:paraId="PNEW"><w:r><w:t>prologue</w:t></w:r></w:p><w:p w14:paraId="P1"><w:r><w:t>first</w:t></w:r></w:p>"#)

        let diff = WordImport.diff(snapshot: snapshot, current: current)

        XCTAssertEqual(diff.operations.count, 1)
        guard case .insertParagraphBefore(let before, let payload) = diff.operations.first else {
            return XCTFail("expected InsertParagraphBefore, got \(diff.operations)")
        }
        XCTAssertEqual(before.raw, "w14:paraId=P1")
        XCTAssertEqual(payload.text, "prologue")
    }

    // MARK: - Removal

    func testRemovedParagraphProducesRemoveParagraph() throws {
        let snapshot = try tree(
            #"<w:p w14:paraId="P1"><w:r><w:t>keep</w:t></w:r></w:p><w:p w14:paraId="P2"><w:r><w:t>drop me</w:t></w:r></w:p>"#)
        let current = try tree(
            #"<w:p w14:paraId="P1"><w:r><w:t>keep</w:t></w:r></w:p>"#)

        let diff = WordImport.diff(snapshot: snapshot, current: current)

        XCTAssertEqual(diff.operations.count, 1)
        guard case .removeParagraph(let id) = diff.operations.first else {
            return XCTFail("expected RemoveParagraph, got \(diff.operations)")
        }
        XCTAssertEqual(id.raw, "w14:paraId=P2")
    }

    // MARK: - No-stable-ID fallback: structural fingerprint matching

    func testParagraphsWithoutStableIDsMatchByFingerprint() throws {
        // Two identical-content paragraphs without paraIds: fingerprint
        // matching pairs them positionally; no ops result.
        let snapshot = try tree(
            #"<w:p><w:r><w:t>alpha</w:t></w:r></w:p><w:p><w:r><w:t>beta</w:t></w:r></w:p>"#)
        let current = try tree(
            #"<w:p><w:r><w:t>alpha</w:t></w:r></w:p><w:p><w:r><w:t>beta</w:t></w:r></w:p>"#)

        let diff = WordImport.diff(snapshot: snapshot, current: current)
        XCTAssertTrue(diff.operations.isEmpty,
                      "identical no-ID paragraphs must pair by fingerprint with no ops")
    }

    // MARK: - Word-sourced attribution is the importer's job (integration below)

    func testFormattingOnlyChangeIsReportedNotSilentlyDropped() throws {
        // MVP scope: matched paragraph whose fingerprint differs but whose
        // text is unchanged (formatting-only change) cannot be represented
        // as SetText. It must surface in unrepresentedChanges — loud, not
        // silently swallowed.
        let snapshot = try tree(
            #"<w:p w14:paraId="ABC"><w:r><w:t>same text</w:t></w:r></w:p>"#)
        let current = try tree(
            #"<w:p w14:paraId="ABC"><w:r><w:rPr><w:b/></w:rPr><w:t>same text</w:t></w:r></w:p>"#)

        let diff = WordImport.diff(snapshot: snapshot, current: current)

        XCTAssertTrue(diff.operations.isEmpty,
                      "formatting-only change has no representable op in the MVP")
        XCTAssertEqual(diff.unrepresentedChanges.map(\.raw), ["w14:paraId=ABC"],
                       "formatting-only change must be reported as unrepresented")
    }
}
