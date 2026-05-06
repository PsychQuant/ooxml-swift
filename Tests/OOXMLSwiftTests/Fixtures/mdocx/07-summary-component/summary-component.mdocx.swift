// Fixture 07: summary-component
// Pins `mdocx-grammar` Requirement: "Component-aware op log via BeginComponent and EndComponent".
// Edge case: a user-defined Summary component expands to one styled Paragraph + one Run.
// The op log brackets the expansion with BeginComponent/EndComponent envelope (visible in
// summary-component.oplog.jsonl). The .docx contains ONLY the expanded Paragraph — no marker
// elements survive serialisation (BeginComponent/EndComponent are op-log metadata only).

import WordDSLSwift

struct Summary: WordComponent {
    let id: String
    let body: () -> Paragraph

    init(id: String, @WordBuilder body: @escaping () -> Paragraph) {
        self.id = id
        self.body = body
    }
}

let document = WordDocument {
    Section(id: "main") {
        Summary(id: "ch1-summary") {
            Paragraph(id: "sum-frame", style: .summaryFrame) {
                "note text"
            }
        }
    }
}
