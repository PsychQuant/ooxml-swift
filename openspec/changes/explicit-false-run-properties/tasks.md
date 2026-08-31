## 1. Contract and RED

- [x] 1.1 Freeze absent/on/off, merge, source-compatibility, and scope decisions in proposal/design/spec.
- [x] 1.2 Add RED unit tests for ST_OnOff parsing/emission and explicit-false merge; frozen RED produced 12 assertion failures across 5 tests.
- [x] 1.3 Add RED DOCX regression: untouched `<w:b w:val="0"/>` survives an unrelated paragraph mutation and save/reopen.

## 2. Implementation

- [x] 2.1 Implement **Typed run on/off properties preserve explicit false** via **Keep Bool accessors and store assignment presence separately** for typed boolean and underline fields.
- [x] 2.2 Implement **Canonical semantic emission**: parse ST_OnOff lexical values and emit canonical explicit false.
- [x] 2.3 Implement **Run-property merge distinguishes omission from explicit false** via **Merge by presence, not truthiness**, including explicit underline removal.

## 3. Verification

- [x] 3.1 Run focused Run/reader/writer/on-off/OMath suites: 143/143 PASS.
- [x] 3.2 Run full `swift test`: 1,448 tests, 31 skips, 0 failures; Spectra strict/analyze clean.
- [x] 3.3 Verify PsychQuant/che-word-mcp#197 against frozen `5c44ef5`: focused 2/2 and full 347 tests with 12 skips, 0 failures; draft PR #198 pins the exact revision pending release.
- [x] 3.4 Close verify R1 blockers with RED→GREEN coverage: `.omathOnly` preserves absent vs explicit-off bold/italic; non-standard `no` follows the frozen unknown/default-on rule; authoritative diff hygiene is clean.
