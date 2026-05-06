# Fixture 14: `atomic-three-file-save`

## Covers

`mdocx-grammar` Requirement **"save(to:) atomic three-file write"** (Requirement 14, special-shape).

## What this fixture pins

`WordDocument.save(to: URL)` writes three files atomically as one logical state:

1. `<name>.docx` — the OOXML container.
2. `<name>.docx.oplog.jsonl` — append-only operation history.
3. `<name>.docx.snapshot.json` — current XmlNode tree snapshot used by `WordImport` for diff.

On failure of any of the three writes, the file system MUST be left in the state before `save(to:)` was called (no partial output).

## Files in this fixture

Standard four required files PLUS two optional Phase B comparison goldens:

- `atomic-three-file-save.mdocx.swift` — calls `try document.save(to: ...)`.
- `atomic-three-file-save.docx` — the expected primary file (one paragraph, one section).
- `atomic-three-file-save.normalized.docx` — byte-equal to .docx.
- `atomic-three-file-save.oplog.jsonl` — expected op log content (3 ops: InsertSection, InsertParagraph, SetRuns).
- `atomic-three-file-save.snapshot.json` — expected tree snapshot (schemaVersion 1; document → body → paragraph + sectPr structure).
- `README.md` — this file.

## Edge case captured

Phase B verifies all three files are produced byte-equal to the goldens. Additionally, the failure-case scenario (forcing a write error on file 2 leaves files 1 and 3 unchanged) is exercised by the runner's Phase B Requirement 14 path.

The `.snapshot.json` schema is provisional (Phase A: any valid JSON; Phase B activation will pin the exact schema). Current shape mirrors the XmlNode tree with `kind`/`prefix`/`localName`/`attributes`/`children` fields and stable `stableID` annotations.

## What to inspect when this fixture fails

| Failure | Likely cause |
|---------|--------------|
| Phase B `oplog byte-diff` | Implementation emitted ops in wrong order; OR omitted InsertSection (treated section as implicit); OR added implementation-detail ops not in the spec |
| Phase B `snapshot byte-diff` | Tree snapshot schema diverged from this fixture's shape — either the fixture needs updating to match the implementation's chosen schema (which then becomes the locked snapshot format), or the implementation diverges from the locked format |
| Phase B `docx byte-diff` | Standard issues: paraId not normalized, sectPr position wrong, etc. |
| Phase B failure-case test | Implementation didn't roll back atomically — found docx written but oplog missing, or vice versa |
