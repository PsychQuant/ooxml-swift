// TableCell.swift — `TableCell` ↔ <w:tc> (see Table.swift).
public struct TableCell {
    public let id: String
    public let paragraphs: [Paragraph]
    public init(id: String, @WordBuilder content: () -> [Paragraph]) {
        self.id = id
        self.paragraphs = content()
    }
}
