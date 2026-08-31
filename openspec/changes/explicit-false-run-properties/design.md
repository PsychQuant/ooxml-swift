## Context

ECMA-376 on/off elements use element presence plus optional `w:val`. Absence means “not directly specified”; a present element with omitted/true/on/1 means on; a present element with false/off/0 means explicitly off. The current typed model stores only a Bool and its additive merge checks only true.

## Goals / Non-Goals

**Goals:**

- Preserve explicit off values through read → unrelated mutation → write.
- Keep existing callers able to read and assign `properties.bold` as Bool.
- Distinguish fresh/unspecified patches from explicit false assignments.
- Apply the same mechanism to typed peer booleans and underline removal.

**Non-Goals:**

- Typing every raw CT_RPr boolean such as bCs/iCs/vanish.
- Changing the tree reducer's existing `RunFormatPayload` contract.
- Byte-preserving which lexical off spelling was used; semantic equivalence is sufficient.

## Decisions

### Keep Bool accessors and store assignment presence separately

`bold`, `italic`, `strikethrough`, and `noProof` remain Bool. Each property records whether it was explicitly assigned. A fresh `RunProperties()` has no specified fields; parsing a present element records true or false; external assignment also records presence. This avoids the source break of changing public fields to `Bool?`.

### Merge by presence, not truthiness

`merge(with:)` applies a boolean whenever the patch marks it specified, including false. An untouched fresh patch leaves the base value unchanged. Underline gains the same assignment-presence bit so assigning nil can remove it.

### Canonical semantic emission

Specified true emits a naked element; specified false emits `w:val="0"`. The reader accepts `0`, `false`, and `off` as false and omitted, `1`, `true`, and `on` as true. Unknown lexical values follow OOXML's default-on behavior rather than being silently treated as off.

## Risks / Trade-offs

- Synthesized equality will distinguish absent false from explicit false; this is intentional because they have different write/merge semantics.
- Property observers do not run for initialization; custom initializers and the reader must mark presence explicitly.
- Partial formatting must never initialize every boolean as specified false.
