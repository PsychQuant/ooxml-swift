import XCTest
@testable import WordDSLSwift

/// Phase 7 entry-point placeholder test for the `mdocx-syntax` Spectra change.
///
/// Verifies the WordDSLSwift module compiles and every top-level DSL type
/// declared in `openspec/specs/mdocx-grammar/spec.md` is reachable. The test
/// does not exercise behavior — actual DSL semantics are implemented in
/// word-aligned-state-sync Phase 7.
final class WordDSLSwiftCompilesTests: XCTestCase {

    func testTopLevelDSLTypesAreReachable() {
        _ = WordDocument.self
        _ = Section.self
        _ = Paragraph.self
        _ = Run.self
        _ = Tab.self
        _ = Break.self
        _ = NoBreakHyphen.self
        _ = Hyperlink.self
        _ = Bookmark.self
        _ = Table.self
        _ = TableRow.self
        _ = TableCell.self
        _ = WordComponent.self
        _ = WordBuilder.self
    }
}
