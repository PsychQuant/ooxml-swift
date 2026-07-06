# word-aligned-state-sync v1.0 benchmarks（tasks 6.6 / 6.7）

Measured 2026-07-06 on Apple Silicon (arm64, debug build), ooxml-swift
v0.34.1. Harness: `Tests/OOXMLSwiftTests/V1BenchmarkTests.swift`
(gated `RUN_BENCHMARKS=1` — CI never runs these).

## 6.7 Typed-view read performance（200-paragraph fixture）

DSL-built 200-paragraph docx, 30 iterations per mode. `get_paragraphs`
in che-word-mcp is `DocxReader.read` + body enumeration, so read()
dominates the tool's cost.

| Mode | mean read() |
|------|-------------|
| `wireTreeBackedViews: false`（pre） | 23.15 ms |
| `wireTreeBackedViews: true`（post） | 22.42 ms |
| **Overhead** | **-3.1%（noise — zero cost）** |

**Verdict**: the tree wiring adds no measurable latency to the typed
read path. The performance-regression risk from the design's risk
section did not materialize.

## 6.6 Tree memory cost（real-world large document）

Fixture: real 1.6 MB docx（`20260505v.docx`, 1,198 paragraphs, 12 XML
parts; main part `word/document.xml` = 3.9 MB XML, 165,122 nodes).

| Measurement | Value |
|-------------|-------|
| XmlTree alone（document.xml parsed standalone） | **38.4 MB** |
| Full `read(wireTreeBackedViews: true)` RSS delta | 177.6 MB |

The tree itself sits **under the 50 MB risk bar**（~10× the XML byte
size — one `XmlNode` class instance plus heap Strings per node is the
expected shape). The remaining ~140 MB of the full-read delta is the
typed model, ZIP/parse transients, and allocator watermark — a one-shot
read peak, not tree residency.

### Documented mitigation（risk-bar second branch）

For workloads where even ~40 MB of tree residency per open document is
too much（long-lived MCP server holding many documents）:

1. **Release early** — `WordDocument.close()` + letting the value leave
   scope frees the trees today; hold one document at a time.
2. **Lazy per-part tree parse**（v1.x follow-up） — the Phase 1 generic
   sweep parses every XML part eagerly; parsing a part's tree only when
   that part is first read/edited cuts residency to the parts actually
   touched（typical edit sessions touch 1-3 of the 12+ parts）.
3. **Slice-backed nodes**（v2 architecture option） — back nodes with
   offsets into the original XML buffer instead of per-node Strings;
   ~10× reduction potential, deferred until real demand.

