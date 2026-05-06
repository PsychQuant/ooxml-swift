import XCTest
@testable import OOXMLSwift

/// TDD scaffold for **task 1.3** of `sibling-types-tree-projection-impl`
/// (Spectra change `sibling-types-tree-projection-impl`,
/// `word-aligned-state-sync` Phase 1 task 2.4, target ooxml-swift v0.31.1).
///
/// Spec: `openspec/changes/sibling-types-tree-projection-impl/specs/ooxml-typed-views-tree-projection/spec.md`
///
/// ## Status
///
/// **Tests are written GREEN-from-the-start** per Decision 7 (no `#if false`
/// gate phase). Implementation lives in
/// `Sources/OOXMLSwift/Models/Section.swift`.
///
/// ## Decisions encoded
///
/// | Test | Decision pinned |
/// |------|-----------------|
/// | `testTreeBackedSectionProperties_constructorTakesXmlNode` | Decision 1: `SectionProperties(xmlNode:)` constructor exists; legacy convenience init stays |
/// | `testTreeBackedSectionProperties_idFallsBackToLibraryUUID` | Decision 2: `id` falls back to `lib:<UUID>` when no native stable ID |
/// | `testTreeBackedSectionProperties_identityEqualityNotContentEquality` | Decision 5: tree-backed equality is identity-based (`===`), not content-based |
/// | `testTreeBackedSectionProperties_phase1StubReturnsDefaultPageSize` | Decision 6: Phase 1 stub — structured fields are NOT parsed from `<w:sectPr>` children |
/// | `testLegacyDetachedSectionProperties_stillCompiles` | Migration safety: legacy convenience init still works; `id == nil` for detached |

final class SectionPropertiesTreeProjectionTests: XCTestCase {

    // MARK: - Helpers

    /// Build a `<w:sectPr>` xmlNode, optionally including a `<w:pgSz>` child.
    /// Used by the Phase 1 stub test to demonstrate that the child is NOT
    /// parsed into `pageSize` — the field stays at the `SectionProperties()`
    /// default.
    private func makeSectPrNode(withPgSz: Bool = false) -> XmlNode {
        var children: [XmlNode] = []
        if withPgSz {
            let pgSz = XmlNode.element(prefix: "w", localName: "pgSz")
            pgSz.setAttribute(prefix: "w", localName: "w", value: "12240")
            pgSz.setAttribute(prefix: "w", localName: "h", value: "15840")
            children.append(pgSz)
        }
        return XmlNode.element(prefix: "w", localName: "sectPr", children: children)
    }

    // MARK: - Constructor + identity

    /// **Decision pinned (Decision 1)**: `SectionProperties(xmlNode:)` is the
    /// tree-backed entry point. Constructing succeeds for any well-formed
    /// `<w:sectPr>` node (semantic validation is the caller's responsibility).
    func testTreeBackedSectionProperties_constructorTakesXmlNode() {
        let node = makeSectPrNode()
        let sp = SectionProperties(xmlNode: node)
        XCTAssertNotNil(sp.xmlNode,
            "SectionProperties(xmlNode:) MUST retain the wrapped XmlNode reference")
        XCTAssertTrue(sp.xmlNode === node,
            "Wrapped xmlNode MUST be the exact reference passed to the initializer")
    }

    /// **Decision pinned (Decision 2)**: when no native OOXML stable-ID
    /// attribute is present, `id` falls back to `"lib:<UUID>"` from
    /// `XmlNode.libraryUUID`. `<w:sectPr>` does not natively carry
    /// `w14:paraId` / `w:bookmarkId`, so the library UUID fallback is the
    /// primary identity source for sections.
    func testTreeBackedSectionProperties_idFallsBackToLibraryUUID() {
        let node = makeSectPrNode()
        node.libraryUUID = UUID(uuidString: "AABBCCDD-1234-5678-9999-FFEEDDCCBBAA")
        let sp = SectionProperties(xmlNode: node)
        XCTAssertEqual(sp.id, "lib:AABBCCDD-1234-5678-9999-FFEEDDCCBBAA",
            "SectionProperties.id MUST fall back to lib:<libraryUUID> when no native stable ID exists")
    }

    // MARK: - Identity equality (Decision 5)

    /// **Decision pinned (Decision 5)**: equality on tree-backed
    /// `SectionProperties` is identity-based (`===` on the wrapped xmlNode),
    /// not content-based. Two sections wrapping the SAME xmlNode are equal;
    /// two wrapping different xmlNodes are NOT equal even if the children
    /// match. Op-log addresses sections by id (== identity); content equality
    /// across different elements would silently merge log entries.
    func testTreeBackedSectionProperties_identityEqualityNotContentEquality() {
        let nodeA = makeSectPrNode()
        let nodeB = makeSectPrNode()

        let spA1 = SectionProperties(xmlNode: nodeA)
        let spA2 = SectionProperties(xmlNode: nodeA)
        let spB = SectionProperties(xmlNode: nodeB)

        XCTAssertEqual(spA1, spA2,
            "Two SectionProperties wrapping the SAME xmlNode MUST be equal (identity equality)")
        XCTAssertNotEqual(spA1, spB,
            "Two SectionProperties wrapping DIFFERENT xmlNodes MUST NOT be equal — id is the equality key")
    }

    // MARK: - Phase 1 stub (Decision 6)

    /// **Decision pinned (Decision 6)**: tree-backed `SectionProperties` is
    /// identity-only in Phase 1. The 12+ structured fields are NOT parsed from
    /// `<w:sectPr>` children — they return the same defaults `SectionProperties()`
    /// would give. This test wraps a `<w:sectPr>` with an explicit
    /// `<w:pgSz w:w="12240" w:h="15840"/>` child and asserts that
    /// `sp.pageSize == PageSize.letter` (the default), NOT a parsed value.
    ///
    /// Reader continues to produce detached `SectionProperties` (with structured
    /// fields populated by the existing parser) in this release, so all
    /// existing call sites that go through Reader are unaffected. Tree-backed
    /// `SectionProperties` is opt-in for downstream library code that explicitly
    /// constructs `SectionProperties(xmlNode:)`. A separate change
    /// (`section-properties-tree-walking-impl`) will add the tree-walking parsers.
    func testTreeBackedSectionProperties_phase1StubReturnsDefaultPageSize() {
        let node = makeSectPrNode(withPgSz: true)
        let sp = SectionProperties(xmlNode: node)
        XCTAssertEqual(sp.pageSize, PageSize.letter,
            "Phase 1 stub: tree-backed SectionProperties.pageSize MUST equal PageSize.letter (the SectionProperties() default), NOT a value parsed from the <w:pgSz> xmlNode child")
    }

    // MARK: - Legacy detached mode (back-compat)

    /// **Migration safety**: existing legacy convenience initializer
    /// `SectionProperties(...)` continues to compile and behave identically.
    /// Detached values have no xmlNode, so `id` is `nil`. che-word-mcp's
    /// production tests depend on this surface; breaking it in Phase 1 is
    /// forbidden.
    func testLegacyDetachedSectionProperties_stillCompiles() {
        let sp = SectionProperties()
        XCTAssertNil(sp.xmlNode,
            "Legacy detached SectionProperties MUST have nil xmlNode")
        XCTAssertNil(sp.id,
            "Detached SectionProperties (no xmlNode) has no stable id; id MUST be nil")
        XCTAssertEqual(sp.pageSize, PageSize.letter,
            "Legacy detached SectionProperties() MUST keep its default pageSize behavior")
    }
}
