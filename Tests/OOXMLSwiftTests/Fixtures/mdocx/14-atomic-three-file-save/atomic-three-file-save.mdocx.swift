// Fixture 14: atomic-three-file-save
// Pins `mdocx-grammar` Requirement: "save(to:) atomic three-file write".
// Edge case: WordDocument.save(to:) writes three files atomically as one logical state:
// <name>.docx + <name>.docx.oplog.jsonl + <name>.docx.snapshot.json. Failure of any leaves
// the filesystem in pre-call state. Forces a write-error scenario tested in README.

import Foundation
import WordDSLSwift

let document = WordDocument {
    Section(id: "main") {
        Paragraph(id: "p1") {
            "Atomic save smoke content."
        }
    }
}

let outURL = URL(fileURLWithPath: "/tmp/mdocx-fixture-14-out.docx")
try document.save(to: outURL)
