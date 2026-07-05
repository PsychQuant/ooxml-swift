// TableRow.swift — `TableRow` ↔ <w:tr> (see Table.swift).
public struct TableRow {
    public let id: String
    public let cells: [TableCell]
    public init(id: String, @WordBuilder content: () -> [TableCell]) {
        self.id = id
        self.cells = content()
    }
}
