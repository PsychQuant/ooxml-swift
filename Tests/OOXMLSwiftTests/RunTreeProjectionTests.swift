import XCTest
@testable import OOXMLSwift

/// Tree-backed `Run` view tests for **task 1.1** of Spectra change
/// `sibling-types-tree-projection-impl` (`word-aligned-state-sync` Phase 1
/// task 2.2, target ooxml-swift v0.31.1).
///
/// Spec: `openspec/changes/sibling-types-tree-projection-impl/specs/`
/// `ooxml-typed-views-tree-projection/spec.md`
///
/// Decisions pinned (Run subset of the change):
///   1. Same struct-with-class-reference pattern as Paragraph
///   2. `id: String?` derives from `XmlNode.stableID` with `lib:` UUID fallback
///   3. Tree-walking primary content getter for `text`
///   4. Phase 1 stub setter mutates the tree directly (preserves non-`<w:t>` siblings)
///   5. Identity-based `Equatable`
///   7. GREEN-from-the-start (no `#if false` gate)
///
/// Each test pins ONE design decision, mirroring the Paragraph test class
/// shipped in v0.31.0 (`ParagraphTreeProjectionTests`).
final class RunTreeProjectionTests: XCTestCase {

    // MARK: - Helpers

    /// Build a `<w:r [w:id="..."]>` xmlNode wrapping the supplied texts.
    /// Each text becomes a direct `<w:t>text</w:t>` child of the `<w:r>`.
    private func makeRunNode(wId: String? = nil, texts: [String]) -> XmlNode {
        let textChildren: [XmlNode] = texts.map { text in
            XmlNode.element(prefix: "w", localName: "t",
                            children: [XmlNode.text(text)])
        }
        let r = XmlNode.element(prefix: "w", localName: "r",
                                children: textChildren)
        if let wId = wId {
            r.setAttribute(prefix: "w", localName: "id", value: wId)
        }
        return r
    }

    // MARK: - Constructor + identity

    /// **Decision 1**: `Run(xmlNode:)` is the tree-backed entry point.
    /// Constructing succeeds without throwing for any well-formed `<w:r>` node.
    func testTreeBackedRun_constructorTakesXmlNode() {
        let node = makeRunNode(texts: ["Hello"])
        let run = Run(xmlNode: node)
        XCTAssertNotNil(run.xmlNode,
            "Run(xmlNode:) MUST retain the wrapped XmlNode as the source of truth")
    }

    /// **Decision 2**: `run.id` reads `w:id` (Run revision ID) via
    /// `XmlNode.stableID`. `<w:r>` does not natively carry `w14:paraId`;
    /// `XmlNode.stableID` resolves `w:id` for revision-tracking runs.
    func testTreeBackedRun_idDerivesFromWId() {
        let node = makeRunNode(wId: "42", texts: ["Hello"])
        let run = Run(xmlNode: node)
        XCTAssertEqual(run.id, "w:id=42",
            "Run.id MUST be the XmlNode stableID format when w:id is present")
    }

    /// **Decision 2 (fallback)**: when no native stable ID attribute is
    /// present, identity falls back to `XmlNode.libraryUUID` (assigned by the
    /// reader for nodes without native ID). Without a fallback, op log cannot
    /// address runs from older docs.
    func testTreeBackedRun_idFallsBackToLibraryUUID() {
        let node = makeRunNode(texts: ["No w:id"])
        node.libraryUUID = UUID(uuidString: "550E8400-E29B-41D4-A716-446655440000")
        let run = Run(xmlNode: node)
        XCTAssertEqual(run.id, "lib:550E8400-E29B-41D4-A716-446655440000",
            "Run.id MUST fall back to lib:<libraryUUID> when no native stable ID exists")
    }

    // MARK: - Getter routing (tree → typed view)

    /// **Decision 3**: `run.text` is a computed property that walks the
    /// `<w:t>` direct children of the wrapped `<w:r>` and concatenates their
    /// text content. It is NOT a stored property cached at construction time.
    func testTreeBackedRun_textGetterConcatenatesWtChildren() {
        let node = makeRunNode(texts: ["Hello", " ", "World"])
        let run = Run(xmlNode: node)
        XCTAssertEqual(run.text, "Hello World",
            "run.text MUST concatenate <w:t> direct children's text in document order")
    }

    // MARK: - Setter routing (typed view → tree, Phase 1 stub)

    /// **Decision 4**: writing `run.text` replaces the wrapped `<w:r>`'s
    /// `<w:t>` direct children with a single new `<w:t>X</w:t>` element while
    /// PRESERVING every non-`<w:t>` sibling (`<w:rPr>`, `<w:tab>`, `<w:br>`,
    /// `<w:drawing>`). This differs from Paragraph's setter (which wipes all
    /// children) — the spec explicitly says "exactly one `<w:t>New</w:t>`
    /// element among any non-`<w:t>` siblings preserved".
    func testTreeBackedRun_textSetterMutatesTreePreservingSiblings() {
        // Start with <w:r><w:rPr><w:b/></w:rPr><w:t>Old</w:t></w:r>.
        let bold = XmlNode.element(prefix: "w", localName: "b")
        let rPr = XmlNode.element(prefix: "w", localName: "rPr", children: [bold])
        let oldText = XmlNode.element(prefix: "w", localName: "t",
                                      children: [XmlNode.text("Old")])
        let runNode = XmlNode.element(prefix: "w", localName: "r",
                                      children: [rPr, oldText])

        var run = Run(xmlNode: runNode)
        run.text = "New"

        XCTAssertEqual(run.text, "New",
            "Setter must update the tree so subsequent getter reads see the new value")

        // Verify the tree retains exactly one <w:t>New</w:t> AND the <w:rPr> sibling.
        let wtChildren = runNode.children.filter { $0.kind == .element && $0.localName == "t" }
        XCTAssertEqual(wtChildren.count, 1,
            "Setter MUST collapse the <w:t> children to exactly one")
        let wtText = wtChildren.first?.children.first { $0.kind == .text }?.textContent
        XCTAssertEqual(wtText, "New",
            "The single surviving <w:t> MUST contain the new text")

        let rPrSiblings = runNode.children.filter { $0.kind == .element && $0.localName == "rPr" }
        XCTAssertEqual(rPrSiblings.count, 1,
            "Setter MUST preserve non-<w:t> siblings (<w:rPr> survives)")
        XCTAssertTrue(rPrSiblings.first?.children.contains { $0.localName == "b" } ?? false,
            "Surviving <w:rPr> MUST still carry its original <w:b/> child")
    }

    /// **Decision 4**: any setter MUST mark the underlying xmlNode dirty so
    /// `XmlTreeWriter` re-serializes that sub-tree from typed fields rather
    /// than re-emitting the original source bytes.
    func testTreeBackedRun_setterMarksNodeDirty() {
        let node = makeRunNode(texts: ["Old"])
        // XmlNode initializes isDirty=true for synthesized nodes (no sourceRange).
        // The contract: after a setter call, isDirty MUST be true regardless of prior state.
        var run = Run(xmlNode: node)
        run.text = "New"
        XCTAssertTrue(node.isDirty,
            "Setter MUST mark the xmlNode dirty so the writer re-serializes from typed fields")
    }

    // MARK: - Identity equality (op log addressing)

    /// **Decision 5**: equality on tree-backed `Run`s is identity-based, not
    /// content-based. Two runs wrapping the SAME xmlNode are equal; two
    /// wrapping different xmlNodes with byte-identical content are NOT equal.
    /// Op log addresses runs by id (== identity), so content-equal-but-
    /// different-id runs being `==` would silently merge log entries that
    /// target different elements.
    func testTreeBackedRun_identityEqualityNotContentEquality() {
        let nodeA = makeRunNode(wId: "id-A", texts: ["Same content"])
        let nodeB = makeRunNode(wId: "id-B", texts: ["Same content"])

        let rA1 = Run(xmlNode: nodeA)
        let rA2 = Run(xmlNode: nodeA)
        let rB = Run(xmlNode: nodeB)

        XCTAssertEqual(rA1, rA2,
            "Two Runs wrapping the SAME xmlNode MUST be equal (identity equality)")
        XCTAssertNotEqual(rA1, rB,
            "Two Runs wrapping DIFFERENT xmlNodes with identical content MUST NOT be equal — id is the equality key")
    }

    // MARK: - Legacy detached mode (back-compat)

    /// **Decision 1 (legacy preservation)**: existing legacy constructor
    /// `Run(text:, properties:)` continues to compile and behave identically.
    /// che-word-mcp's 297 production tests depend on this surface; breaking
    /// it in v0.31.1 is forbidden.
    func testLegacyRun_detachedConstructorStillCompiles() {
        let run = Run(text: "x")
        XCTAssertEqual(run.text, "x",
            "Legacy detached Run(text:) MUST keep its observable text behavior")
        XCTAssertNil(run.id,
            "Detached run (no xmlNode) has no stable id; id MUST be nil")

        // Detached content equality preserved (mode 2 of Decision 5).
        let rA = Run(text: "same", properties: RunProperties())
        let rB = Run(text: "same", properties: RunProperties())
        XCTAssertEqual(rA, rB,
            "Two detached Runs with identical legacy fields MUST be content-equal")
    }
}
