# Fixture 15: `reverse-cli-roundtrip`

## Covers

`mdocx-grammar` Requirement **"Reverse CLI shape — macdoc word reverse"** (Requirement 15, special-shape).

## What this fixture pins

The `macdoc word reverse <docx> --to-mdocx <out>` CLI: given a docx, produce DSL source that, when re-executed, produces a docx byte-equal to the input.

## Files in this fixture

Standard four required files PLUS one optional Phase B comparison golden:

- `reverse-cli-roundtrip.mdocx.swift` — INPUT script. Phase B runner executes this, captures the produced docx, then runs `macdoc word reverse <docx>` against it.
- `reverse-cli-roundtrip.docx` — the docx that the input script SHOULD produce.
- `reverse-cli-roundtrip.normalized.docx` — byte-equal to .docx.
- `reverse-cli-roundtrip.expected-source.mdocx.swift` — the EXPECTED output of `macdoc word reverse`. Phase B compares the actual reverse output against this file after canonicalization.
- `README.md` — this file.

## Canonicalization rules

The reverse output may differ from the input source in these allowed ways (canonicalization normalizes both sides before comparison):

- **Comment content**: input may have narrative comments; reverse output has only auto-generated marker comments.
- **Indentation**: 4-space indent uniformly.
- **Parameter order on `Paragraph(id:, style:, ...)`**: reverse output emits parameters in spec-defined order regardless of input order.
- **Trailing whitespace and blank line counts**: normalized to one blank line between top-level declarations.

## Edge case captured

Round-trip identity check at the source level. Phase B also exercises `macdoc word reverse` with `--from-oplog` flag (replays the oplog to obtain current state) and without it (reverse-engineers from docx parts only).

## What to inspect when this fixture fails

| Failure | Likely cause |
|---------|--------------|
| Phase B `reverse-source equivalence` | Reverse CLI lost an element (e.g., dropped Section wrapper); OR added unexpected content; OR canonicalization rules don't match what the implementation produces — update either implementation or `expected-source.mdocx.swift` to align |
| Phase B `docx byte-diff` (after re-execution) | The reverse-generated source produced a different docx than the input — round-trip identity broken at the docx layer |
