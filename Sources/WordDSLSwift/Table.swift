// Table.swift
// mdocx-grammar: "Table grammar mirrors OOXML three-layer structure"
// (`Table` ↔ <w:tbl>). v0.34: three-layer DSL types with mandatory ids
// compile (fixtures 10a/10b); op emission awaits an appendTable authoring
// op + reducer support — buildLog throws loudly rather than dropping.
public struct Table {
    public let id: String
    public let rows: [TableRow]
    public init(id: String, @WordBuilder content: () -> [TableRow]) {
        self.id = id
        self.rows = content()
    }
}
