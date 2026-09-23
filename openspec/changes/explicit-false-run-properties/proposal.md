## Why

`RunProperties` collapses OOXML on/off properties from three source states (absent, on, explicit off) into plain `Bool`. `DocxReader` therefore reads `<w:b w:val="0"/>` as true, and a later typed write silently turns untouched text bold. The same representation makes a caller's explicit `false` formatting patch indistinguishable from an omitted field. See PsychQuant/ooxml-swift#115 and parent PsychQuant/macdoc#173.

## What Changes

- Preserve absent/on/off presence for typed run booleans while keeping the public Bool read surface source-compatible.
- Parse all ST_OnOff spellings for bold, italic, strikethrough, and noProof.
- Emit explicit false as a semantic off value and preserve true as the canonical naked element.
- Make `RunProperties.merge(with:)` apply explicit false but leave an untouched field unspecified.
- Track underline assignment presence so a nil underline patch can explicitly remove an existing underline.
- Add reader/writer, unrelated-mutation, merge, and compatibility regressions.

## Capabilities

### New Capabilities

- `run-property-on-off`: Lossless typed on/off run-property preservation and patch semantics.

## Impact

- `Sources/OOXMLSwift/Models/Run.swift`
- `Sources/OOXMLSwift/IO/DocxReader.swift`
- `Tests/OOXMLSwiftTests/RunPropertyOnOffTests.swift`
- Downstream consumer: PsychQuant/che-word-mcp#197
