## ADDED Requirements

### Requirement: Typed run on/off properties preserve explicit false

The system SHALL distinguish absent, explicitly on, and explicitly off states for typed boolean run properties while retaining Bool read access. Bold, italic, strikethrough, and noProof SHALL parse OOXML ST_OnOff values and SHALL serialize explicit false as a semantic off value.

#### Scenario: Unrelated mutation preserves explicit-off bold

- **GIVEN** a source run contains `<w:b w:val="0"/>`
- **WHEN** a different paragraph is mutated and the document is saved
- **THEN** the untouched run remains explicitly non-bold after save/reopen
- **AND** it SHALL NOT be emitted as naked `<w:b/>`

#### Scenario: ST_OnOff lexical forms are accepted

- **WHEN** a typed on/off element uses `0`, `false`, or `off`
- **THEN** its Bool accessor is false and its state is explicitly specified
- **WHEN** the element omits `w:val` or uses `1`, `true`, or `on`
- **THEN** its Bool accessor is true and its state is explicitly specified

#### Scenario: Parallel copy and revision paths preserve explicit off

- **GIVEN** typed run properties contain explicit-off bold, italic, strikethrough, noProof, or underline
- **WHEN** the run is copied through the tree-backed append/run-payload path or serialized as an `rPrChange` previous format
- **THEN** each specified off value remains explicit in emitted OOXML
- **AND** it SHALL NOT collapse to absence

#### Scenario: WordprocessingML namespace aliases are accepted

- **WHEN** a run uses a non-`w` prefix bound to the WordprocessingML namespace and carries explicit-off properties
- **THEN** the typed reader recognizes the elements and values by namespace URI
- **AND** the Bool accessors remain false and specified

### Requirement: Run-property merge distinguishes omission from explicit false

A fresh `RunProperties()` patch SHALL leave existing boolean and underline properties unchanged. Explicit assignment of false SHALL disable the corresponding boolean. Explicit assignment of nil underline SHALL remove an existing underline.

#### Scenario: Explicit false clears bold without clearing omitted peers

- **GIVEN** a base run is bold, italic, and underlined
- **AND** a patch explicitly assigns `bold = false` but does not assign italic or underline
- **WHEN** the patch is merged
- **THEN** bold is false
- **AND** italic and underline remain unchanged

#### Scenario: Explicit nil removes underline

- **GIVEN** a base run has single underline
- **WHEN** a patch explicitly assigns `underline = nil`
- **THEN** the merged run has no underline
