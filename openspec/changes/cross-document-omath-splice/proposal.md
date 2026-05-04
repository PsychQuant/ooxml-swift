## Why

`WordDocument` has no API to copy `<m:oMath>` (Office Math Markup Language) blocks verbatim from one document to another. The only existing OMath insertion API (`insertEquation(at:latex:displayMode:)`) routes through the deprecated `MathEquation` LaTeX-to-OMML simplifier, which collapses structured OMML (fractions, subscripts, n-ary operators) into flat text runs and is unsuitable for cross-document content rescue.

The 郭嘉員 thesis rescue pipeline (`kiki830621/collaboration_guo_analysis`) needs to splice 522 inline OMath blocks (Greek letters, statistical results like `α=0.1001`, `β=0.8755`, `p<0.001`) from `_raw.docx` into `碩士論文-rescue-swift-v317.docx` after a prior image-insertion pipeline silently dropped them. Hardcoding 510+ LaTeX strings is impractical and fundamentally cannot reproduce the original OMML structure. Verbatim XML splice is the only viable approach. See [PsychQuant/ooxml-swift#57](https://github.com/PsychQuant/ooxml-swift/issues/57) for full diagnostic context.

## What Changes

- Add `WordDocument.spliceOMath(from:toBodyParagraphIndex:position:omathIndex:rPrMode:namespacePolicy:)` — single-OMath verbatim splice between paragraphs (low-level, full caller control)
- Add `WordDocument.spliceParagraphOMath(from:toBodyParagraphIndex:rPrMode:namespacePolicy:)` — paragraph-level batch (high-level convenience for cross-doc rescue; auto-derives anchor from each OMath's source-text-context)
- Add `OMathSplicePosition` enum: `.atStart` / `.atEnd` / `.afterText(_, instance:, options:)` / `.beforeText(_, instance:, options:)` — mirrors existing `InsertLocation` anchor pattern
- Add `OMathSpliceRpRMode` enum: `.full` (default) / `.omathOnly` / `.discard` — controls source-Run rPr (font/size/lang) propagation to target
- Add `OMathSpliceNamespacePolicy` enum: `.lenient` (default) / `.strict` — controls behavior on namespace prefix vs URI mismatch
- Add `OMathSpliceError` enum: 6 cases covering the failure taxonomy
- Carrier preservation: source's `Run.rawXML` OMath splices into target as `Run.rawXML`; source's direct-child `unrecognizedChildren` OMath splices into target as `unrecognizedChildren` (visual semantics — inline stays inline, display stays display)
- Mid-paragraph splice via anchor-Run split (does not touch the other 12 position-indexed paragraph carriers — isolated blast radius; mirrors `replaceText` pattern)

## Non-Goals (optional)

(Documented in design.md Goals/Non-Goals section.)

## Capabilities

### New Capabilities

- `omath-splice`: Cross-document verbatim copy of `<m:oMath>` XML blocks between `WordDocument` paragraphs, preserving carrier shape (inline vs direct-child), source rPr, and namespace declarations. Includes single-OMath low-level API and paragraph-level batch API.

### Modified Capabilities

(none — additive feature; existing OMath round-trip behavior at issues #85, #92, and #99 through #103 is unaffected)

## Impact

- **Affected specs**: new `specs/omath-splice/spec.md`
- **Affected code**:
  - `Sources/OOXMLSwift/Models/Document.swift` — public API surface (~150 LOC additions)
  - `Sources/OOXMLSwift/Models/OMathSplice.swift` — new file, types + helpers + carrier extraction logic (~200 LOC)
  - `Sources/OOXMLSwift/Models/Paragraph.swift` — possibly extract `splitRun(at:)` helper if not already factored from existing `replaceText` path
  - `Tests/OOXMLSwiftTests/OMathSpliceTests.swift` — new test file (~300 LOC, 8+ test cases including round-trip + Word-render fixture)
- **Downstream consumers**:
  - `PsychQuant/che-word-mcp#160` — MCP tool wrapper (`splice_omath_from_source`) tracking this API (filed as sister concern from #57 diagnose)
  - `kiki830621/collaboration_guo_analysis` rescue Phase 7 — primary consumer; will adopt this API after release
- **Version**: bump to `0.24.0` (additive, no breaking changes to existing API)
- **MathEquation deprecation**: independent track (#58 — sister concern); this API does not depend on `MathEquation`
