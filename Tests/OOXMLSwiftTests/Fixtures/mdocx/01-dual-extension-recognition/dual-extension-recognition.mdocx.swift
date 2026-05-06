// Fixture 01: dual-extension-recognition
//
// Pins `mdocx-grammar` Requirement: "File extension and dual-extension pattern"
// (canonical at openspec/specs/mdocx-grammar/spec.md).
//
// Edge case: the simplest possible .mdocx.swift script. Demonstrates the
// `.mdocx.swift` dual-extension pattern (filename ends in `.mdocx.swift`
// so the Swift toolchain treats it as Swift source while the `.mdocx`
// segment signals macdoc DSL routing).
//
// What this fixture establishes:
//   1. The dual-extension filename `<slug>.mdocx.swift` parses as Swift.
//   2. The minimal valid WordDocument shape compiles against WordDSLSwift.
//   3. The corresponding hand-crafted docx (one paragraph, one section,
//      single stable paraId, no RSIDs / theme / settings noise) is the
//      simplest possible normalized form — byte-equal pre/post normalize.
//
// Phase A status: this file's compile-pass check uses lightweight
// tokenization (non-empty + `import` keyword + balanced braces). Phase B
// (after WordDSLSwift implementation lands) will execute the script and
// byte-diff the output against `dual-extension-recognition.normalized.docx`.

import WordDSLSwift

let document = WordDocument {
    Section(id: "main") {
        Paragraph(id: "p1") {
            "Smoke fixture for dual-extension recognition"
        }
    }
}
