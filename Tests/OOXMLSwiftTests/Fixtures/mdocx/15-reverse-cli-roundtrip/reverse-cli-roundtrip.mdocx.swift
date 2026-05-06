// Fixture 15: reverse-cli-roundtrip
// Pins `mdocx-grammar` Requirement: "Reverse CLI shape — macdoc word reverse".
// Edge case: macdoc word reverse <docx> --to-mdocx <out.mdocx.swift> produces DSL source that,
// when re-executed, produces a docx byte-equal to the input. This file is the INPUT script that
// the test runs to produce the docx; sister file expected-source.mdocx.swift is the OUTPUT
// the reverse CLI is expected to produce when run against that docx.

import WordDSLSwift

let document = WordDocument {
    Section(id: "main") {
        Paragraph(id: "p1") {
            "Reverse CLI roundtrip content."
        }
    }
}
