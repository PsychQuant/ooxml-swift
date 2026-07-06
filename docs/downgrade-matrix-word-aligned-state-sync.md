# word-aligned-state-sync downgrade matrix（task 6.8）

Decision 8 staged the migration across six releases so each tag is
independently adoptable — and independently *downgradable*. This matrix
records what consumers lose (and what to watch for) when pinning back
from each tag.

| From | To | Safe? | What you lose / caveats |
|------|----|-------|--------------------------|
| v0.30.0（XmlNode tree foundation） | v0.29.x | ✅ | Tree layer disappears; typed model behavior identical（tree was purely additive; no consumer-visible behavior change） |
| v0.31.x（typed views tree-backed, opt-in wire） | v0.30.0 | ✅ | `wireTreeBackedViews:` parameter and `elementID` on typed views vanish; default read path identical（wire was opt-in default-false） |
| v0.32.0（op log + JSONL sidecars） | v0.31.x | ✅ | `saveWithSidecars`/`openWithSidecars` APIs vanish. Existing sidecar files on disk are ignored by older versions（plain files next to the docx — inert, safe to leave or delete） |
| v0.33.0（SyncOrchestrator + Word import diff） | v0.32.0 | ✅ | Sync layer vanishes; sidecars written by 0.33 still load（same JSONL wire）. `snapshot.json` files carrying the 0.33 `documentXML` field decode fine on 0.32（optional field, ignored） |
| v0.33.1（§4b OOXML-mirror authoring ops） | v0.33.0 | ⚠️ | Sidecar logs containing the 8 new op_types（`appendParagraph`, `setRuns`, `defineStyle`, `beginComponent`/`endComponent`, `insertTab`/`insertBreak`/`insertNoBreakHyphen`）decode as `unknown` on 0.33.0 — preserved byte-equal（forward-compat rule）but NOT replayable there |
| v0.34.0（.mdocx transcoder + WordDSLSwift） | v0.33.1 | ✅ | Transcoder + DSL module vanish; `.mdocx.swift` files become plain unexecuted sources. Op wire unchanged |
| v0.34.1（moveNode `sourceNode` wire fix） | v0.34.0 or older | ⚠️ | Sidecar lines with `moveNode` written by 0.34.1 use the `sourceNode` field — older decoders look for `source` and hit the envelope collision（the very bug 0.34.1 fixed）: decode does not throw, but the moved-element ID reads wrong. moveNode lines written BY older versions were already corrupt（ID lost at encode）. If your logs contain moveNode, do not downgrade past 0.34.1 |
| v1.0.0（legacy typed-only IO removed — pending） | v0.34.x | ⚠️ | Planned breaking: `rawChildren` fields and typed-only IO paths removed; code depending on them must pin ≤0.34.x until migrated |

**General rules**

- All 0.30–0.34 releases are wire-compatible in the forward direction:
  newer readers always decode older sidecars.
- Backward（downgrade）direction is safe except where an op_type or
  field name did not exist yet — those decode via the `unknown` /
  forward-compat path（preserved, not replayable）.
- Sidecars are plain files and opt-in（spec-frozen Q1）: the zero-risk
  downgrade path for any version is "stop opting in" — the docx itself
  never carries sync metadata.
