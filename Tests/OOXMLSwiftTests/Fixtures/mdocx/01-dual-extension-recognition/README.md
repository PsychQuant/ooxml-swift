# Fixture 01: `dual-extension-recognition`

## Covers

`mdocx-grammar` Requirement **"File extension and dual-extension pattern"** (Requirement 1, canonical at `openspec/specs/mdocx-grammar/spec.md`).

## What this fixture pins

The simplest possible `.mdocx.swift` script:

- Filename ends in `.mdocx.swift` (the dual-extension pattern from Decision 8 of the grammar spec).
- Body uses the minimum DSL surface: `WordDocument { Section(id:) { Paragraph(id:) { "..." } } }`.
- The corresponding hand-crafted `<slug>.docx` is the smallest valid OOXML document containing exactly that one paragraph, one section, one paragraph ID. No RSID attributes, no theme, no settings keys, single monotonic paraId.

## Edge case captured

This fixture is the **smoke baseline**: the first fixture authored, used by Phase A infrastructure validation (task 1.4) to verify that the runner's directory walk + naming check + file-set check + Swift compile-pass tokenization all work end-to-end against a real corpus directory.

If everything else in the corpus breaks, this fixture should still pass Phase A.

## What to inspect when this fixture fails

| Failure | Likely cause |
|---------|--------------|
| `directory layout: name does not match <NN>[<letter>]-<slug> pattern` | Directory was renamed |
| `file set: missing <slug>.docx` | The hand-crafted docx file was deleted; rebuild via the Python script in this README |
| `file set: missing <slug>.normalized.docx` | The pre-normalized golden was deleted; for this fixture it equals the raw `.docx` byte-for-byte (no normalizable content), so `cp <slug>.docx <slug>.normalized.docx` restores it |
| `Swift compile-pass: file is empty` | `dual-extension-recognition.mdocx.swift` was emptied |
| `Swift compile-pass: no import declaration` | Someone removed the `import WordDSLSwift` line |
| `Swift compile-pass: unbalanced braces` | Syntax error introduced into the .mdocx.swift |
| Phase B: `docx byte-diff` (only when `activatePhaseB == true`) | The new `WordDSLSwift` implementation's output diverges from the golden — most likely the implementation made a different choice for paraId numbering, sectPr placement, or whitespace, all of which need to match the hand-crafted golden bytes |

## How the docx golden was built

```python
import zipfile
fixed_time = (2026, 1, 1, 0, 0, 0)
entries = sorted([
    ("[Content_Types].xml", content_types_xml),
    ("_rels/.rels", root_rels_xml),
    ("word/_rels/document.xml.rels", doc_rels_xml),
    ("word/document.xml", document_xml),
])
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
    for name, content in entries:
        info = zipfile.ZipInfo(name, date_time=fixed_time)
        info.compress_type = zipfile.ZIP_DEFLATED
        info.create_system = 3  # Unix; suppresses extras
        zf.writestr(info, content)
```

The same script is used as the template for fixtures 02-13. Fixed timestamp `(2026, 1, 1, 0, 0, 0)` plus Unix `create_system=3` keeps ZIP framing deterministic; sorted entry order matches what `MdocxFixtureNormalizer` emits.

## Why `<slug>.normalized.docx` equals `<slug>.docx` byte-for-byte here

The hand-crafted docx for this fixture deliberately omits every field the normalizer strips (no `w:rsid*` attributes, no `word/theme/theme1.xml`, no `w:rsids` / `w:zoom` / `w:proofState` / `w:defaultTabStop` in settings — the docx has no `word/settings.xml` at all, since none of the four default keys are needed). The single `w14:paraId` is already in the monotonic `00000001` form the normalizer would re-number to. So normalize is a no-op and the two files are byte-identical.

This property is fixture-specific. Fixtures 02+ that exercise normalization rules will have `<slug>.docx` and `<slug>.normalized.docx` diverge.
