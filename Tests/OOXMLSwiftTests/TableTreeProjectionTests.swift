import XCTest
@testable import OOXMLSwift

/// Tests for **task 1.2** of Spectra change `sibling-types-tree-projection-impl`
/// (target ooxml-swift v0.31.1, follow-up to v0.31.0 paragraph tree projection).
///
/// Spec: `openspec/changes/sibling-types-tree-projection-impl/specs/ooxml-typed-views-tree-projection/spec.md`
///
/// ## What this file pins
///
/// Each test pins ONE design decision from the spec. The tests are surgical
/// so the three Table-family structs (`Table`, `TableRow`, `TableCell`) can be
/// reviewed in isolation from `Run` and `SectionProperties` (sibling tasks
/// 1.1 and 1.3).
///
/// | Test | Decision pinned |
/// |------|-----------------|
/// | `testTreeBackedTable_constructorTakesXmlNode` | Constructor surface (`Table(xmlNode:)`) |
/// | `testTreeBackedTableRow_constructorTakesXmlNode` | Constructor surface (`TableRow(xmlNode:)`) |
/// | `testTreeBackedTableCell_constructorTakesXmlNode` | Constructor surface (`TableCell(xmlNode:)`) |
/// | `testTreeBackedTable_idFallsBackToLibraryUUID` | Identity model (Decision 2: `lib:` UUID fallback) |
/// | `testTreeBackedTable_rowsCountReflectsTreeChildren` | Tree-walking getter (Decision 3: live view, no caching) |
/// | `testTreeBackedTableRow_cellsReflectsTreeChildren` | Tree-walking getter for `TableRow.cells` |
/// | `testTreeBackedTableCell_paragraphsReturnsTreeBackedParagraphs` | Cross-type composition (`TableCell.paragraphs` → `Paragraph(xmlNode:)`) |
/// | `testTreeBackedTable_identityEqualityNotContentEquality` | Identity equality (Decision 5: `===` on tree-backed, content on detached) |
/// | `testLegacyTable_detachedConstructorsStillCompile` | Migration safety (Decision: legacy constructors preserved) |
final class TableTreeProjectionTests: XCTestCase {

    // MARK: - Helpers

    /// Build a `<w:tbl>` xmlNode with `rows` × `cellsPerRow` shape; each cell
    /// contains exactly one empty `<w:p>` to mirror the OOXML invariant that
    /// every `<w:tc>` has at least one paragraph child.
    private func makeTblNode(rows: Int, cellsPerRow: Int) -> XmlNode {
        let trChildren: [XmlNode] = (0..<rows).map { _ in
            let tcChildren: [XmlNode] = (0..<cellsPerRow).map { _ in
                XmlNode.element(prefix: "w", localName: "tc",
                    children: [XmlNode.element(prefix: "w", localName: "p", children: [])])
            }
            return XmlNode.element(prefix: "w", localName: "tr", children: tcChildren)
        }
        return XmlNode.element(prefix: "w", localName: "tbl", children: trChildren)
    }

    /// Build a `<w:tr>` xmlNode with the given number of empty `<w:tc>` cells.
    private func makeTrNode(cellsPerRow: Int) -> XmlNode {
        let tcChildren: [XmlNode] = (0..<cellsPerRow).map { _ in
            XmlNode.element(prefix: "w", localName: "tc",
                children: [XmlNode.element(prefix: "w", localName: "p", children: [])])
        }
        return XmlNode.element(prefix: "w", localName: "tr", children: tcChildren)
    }

    /// Build a `<w:tc>` xmlNode with `paragraphCount` `<w:p>` children. Optionally
    /// stamps a `w14:paraId` attribute on each `<w:p>` so the returned tree-backed
    /// Paragraphs have non-nil `id`.
    private func makeTcNode(paragraphCount: Int, paraIdPrefix: String? = nil) -> XmlNode {
        let pChildren: [XmlNode] = (0..<paragraphCount).map { i in
            let p = XmlNode.element(prefix: "w", localName: "p", children: [])
            if let prefix = paraIdPrefix {
                p.setAttribute(prefix: "w14", localName: "paraId", value: "\(prefix)\(i)")
            }
            return p
        }
        return XmlNode.element(prefix: "w", localName: "tc", children: pChildren)
    }

    // MARK: - Constructor + identity

    /// **Decision pinned**: `Table(xmlNode:)` is the tree-backed entry point;
    /// constructing succeeds for any well-formed `<w:tbl>` node.
    func testTreeBackedTable_constructorTakesXmlNode() {
        let node = makeTblNode(rows: 2, cellsPerRow: 3)
        let t = Table(xmlNode: node)
        XCTAssertNotNil(t.xmlNode, "Table(xmlNode:) MUST retain the wrapped xmlNode reference")
    }

    /// **Decision pinned**: `TableRow(xmlNode:)` is the tree-backed entry
    /// point; constructing succeeds for any well-formed `<w:tr>` node.
    func testTreeBackedTableRow_constructorTakesXmlNode() {
        let node = makeTrNode(cellsPerRow: 4)
        let r = TableRow(xmlNode: node)
        XCTAssertNotNil(r.xmlNode, "TableRow(xmlNode:) MUST retain the wrapped xmlNode reference")
    }

    /// **Decision pinned**: `TableCell(xmlNode:)` is the tree-backed entry
    /// point; constructing succeeds for any well-formed `<w:tc>` node.
    func testTreeBackedTableCell_constructorTakesXmlNode() {
        let node = makeTcNode(paragraphCount: 1)
        let c = TableCell(xmlNode: node)
        XCTAssertNotNil(c.xmlNode, "TableCell(xmlNode:) MUST retain the wrapped xmlNode reference")
    }

    /// **Decision pinned (Decision 2)**: when no native OOXML stable-ID
    /// attribute is present on the `<w:tbl>` element, `table.id` falls back
    /// to `"lib:<libraryUUID>"`. `<w:tbl>` does not natively carry
    /// `w14:paraId`, so the `lib:` fallback is the dominant id source for
    /// the Table family.
    func testTreeBackedTable_idFallsBackToLibraryUUID() {
        let node = makeTblNode(rows: 1, cellsPerRow: 1)
        node.libraryUUID = UUID(uuidString: "550E8400-E29B-41D4-A716-446655440000")
        let t = Table(xmlNode: node)
        XCTAssertEqual(t.id, "lib:550E8400-E29B-41D4-A716-446655440000",
                       "Table.id MUST fall back to lib:<libraryUUID> when no native stable ID exists")
    }

    // MARK: - Tree-walking getters (Decision 3)

    /// **Decision pinned (Decision 3)**: `table.rows.count` reflects the
    /// current count of `<w:tr>` children. Mutating `xmlNode.children` MUST
    /// be observable through subsequent reads without reconstructing the
    /// `Table` value (live view, no caching).
    func testTreeBackedTable_rowsCountReflectsTreeChildren() {
        let node = makeTblNode(rows: 3, cellsPerRow: 2)
        let t = Table(xmlNode: node)
        XCTAssertEqual(t.rows.count, 3,
                       "rows.count MUST reflect <w:tr> direct children of the wrapped xmlNode")

        // Mutate the tree directly (not via Table) — view must reflect immediately.
        let extraRow = XmlNode.element(prefix: "w", localName: "tr",
            children: [XmlNode.element(prefix: "w", localName: "tc",
                children: [XmlNode.element(prefix: "w", localName: "p", children: [])])])
        node.children.append(extraRow)
        XCTAssertEqual(t.rows.count, 4,
                       "rows is a live view; tree mutation must be observable without reconstructing the Table")
    }

    /// **Decision pinned (Decision 3)**: `tableRow.cells.count` reflects
    /// the current count of `<w:tc>` children of the wrapped `<w:tr>`.
    func testTreeBackedTableRow_cellsReflectsTreeChildren() {
        let node = makeTrNode(cellsPerRow: 5)
        let r = TableRow(xmlNode: node)
        XCTAssertEqual(r.cells.count, 5,
                       "cells.count MUST reflect <w:tc> direct children of the wrapped xmlNode")

        // Drop one cell directly on the tree; the view must reflect.
        node.children.removeLast()
        XCTAssertEqual(r.cells.count, 4,
                       "cells is a live view; tree mutation must be observable")
    }

    /// **Decision pinned (Decision 3 + cross-type composition)**:
    /// `tableCell.paragraphs` walks `<w:p>` children and returns
    /// `Paragraph(xmlNode:)` per child — i.e. the returned Paragraphs are
    /// themselves tree-backed (uses the v0.31.0 `Paragraph(xmlNode:)`
    /// constructor). When the underlying `<w:p>` carries a `w14:paraId`,
    /// `paragraph.id` MUST be non-nil for every returned Paragraph.
    func testTreeBackedTableCell_paragraphsReturnsTreeBackedParagraphs() {
        let node = makeTcNode(paragraphCount: 2, paraIdPrefix: "PID")
        let c = TableCell(xmlNode: node)
        XCTAssertEqual(c.paragraphs.count, 2,
                       "paragraphs.count MUST reflect <w:p> direct children of the wrapped <w:tc>")
        for p in c.paragraphs {
            XCTAssertNotNil(p.id,
                "Each returned Paragraph MUST be tree-backed; with w14:paraId stamped, paragraph.id MUST be non-nil")
        }
    }

    // MARK: - Identity equality (Decision 5)

    /// **Decision pinned (Decision 5)**: equality on tree-backed `Table`s is
    /// identity-based (`===` on the wrapped xmlNode reference); two Tables
    /// wrapping different xmlNodes with byte-identical content are NOT equal.
    func testTreeBackedTable_identityEqualityNotContentEquality() {
        let nodeA = makeTblNode(rows: 2, cellsPerRow: 2)
        let nodeB = makeTblNode(rows: 2, cellsPerRow: 2)

        let tA1 = Table(xmlNode: nodeA)
        let tA2 = Table(xmlNode: nodeA)
        let tB = Table(xmlNode: nodeB)

        XCTAssertEqual(tA1, tA2,
                       "Two Tables wrapping the SAME xmlNode MUST be equal (identity equality)")
        XCTAssertNotEqual(tA1, tB,
            "Two Tables wrapping DIFFERENT xmlNodes with identical content MUST NOT be equal — id is the equality key")
    }

    // MARK: - Legacy detached mode (Decision: API surface preserved)

    /// **Decision pinned**: existing legacy constructors keep working
    /// byte-equivalent for detached values. che-word-mcp's 297 tests depend
    /// on this surface; breaking it in Phase 1 is forbidden.
    func testLegacyTable_detachedConstructorsStillCompile() {
        let t = Table(rowCount: 2, columnCount: 3)
        XCTAssertEqual(t.rows.count, 2, "Legacy Table(rowCount:columnCount:) MUST keep its observable behavior")
        XCTAssertEqual(t.rows.first?.cells.count, 3,
                       "Each synthesized row MUST contain the requested number of cells")
        XCTAssertNil(t.id, "Detached table (no xmlNode) has no stable id; id MUST be nil")

        // Other legacy constructors stay compileable too.
        let t2 = Table(rows: [TableRow(cells: [TableCell()])], properties: TableProperties())
        XCTAssertEqual(t2.rows.count, 1)
        XCTAssertNil(t2.id)

        let r = TableRow(cells: [TableCell(), TableCell()])
        XCTAssertEqual(r.cells.count, 2)
        XCTAssertNil(r.id)

        let cEmpty = TableCell()
        XCTAssertEqual(cEmpty.paragraphs.count, 1, "TableCell() MUST seed exactly one empty paragraph")
        XCTAssertNil(cEmpty.id)

        let cWithText = TableCell(text: "hi")
        XCTAssertEqual(cWithText.paragraphs.count, 1)
        XCTAssertNil(cWithText.id)
    }
}
