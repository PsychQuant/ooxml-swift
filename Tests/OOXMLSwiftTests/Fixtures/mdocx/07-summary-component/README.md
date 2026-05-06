# Fixture 07: `summary-component`

## Covers

`mdocx-grammar` Requirement **"Component-aware op log via BeginComponent and EndComponent"** (Requirement 7).

## What this fixture pins

A user-defined `WordComponent` (here: `Summary`) wraps its body in `BeginComponent`/`EndComponent` envelope ops. The envelope is metadata that survives the op log (so the reverse transcoder reconstructs the component invocation) but produces NO element in the final OOXML output.

## Files

- `summary-component.mdocx.swift` — defines `Summary: WordComponent` and uses it once.
- `summary-component.docx` — the OOXML output: ONE `<w:p>` with the inner content. **No** marker elements for BeginComponent/EndComponent.
- `summary-component.normalized.docx` — byte-equal to the .docx (no normalizable content).
- `summary-component.oplog.jsonl` — the expected op log: `BeginComponent` + `InsertParagraph` + `SetRuns` + `EndComponent` in order.

## Why `.oplog.jsonl` is included

This fixture is one of two (the other is fixture 09) that exercises op-log structure beyond what's visible in the docx. Phase B of the runner compares the produced op log byte-equal against this file when `activatePhaseB == true`.

## What to inspect when this fixture fails

| Failure | Likely cause |
|---------|--------------|
| Phase B `docx byte-diff` | Implementation emitted marker elements (e.g., processing instructions or comments) for the component envelope — incorrect; envelope is op-log only |
| Phase B `oplog byte-diff` | Missing BeginComponent/EndComponent ops; OR wrong order; OR wrong `id` value (must match `Summary(id: "ch1-summary")`) |
