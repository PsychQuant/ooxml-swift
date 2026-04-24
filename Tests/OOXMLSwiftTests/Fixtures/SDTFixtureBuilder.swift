import Foundation
@testable import OOXMLSwift

/// Programmatic fixture builder for SDT parser and round-trip tests.
///
/// Produces an in-memory `.docx` containing one SDT per supported type,
/// plus a nested Group→PlainText case and a block-level SDT case.
/// Follows the project convention established by
/// `RoundTripFidelityTests` (build fixtures via DocxWriter rather than
/// committing binary `.docx` files to the repo).
///
/// Created for change `che-word-mcp-content-controls-read-write` task 1.1.
/// Note: the original task text references a "Word-generated" .docx file.
/// We use programmatic generation to match project convention; a human
/// Word-open validation step lives in task 3.5's round-trip test.
enum SDTFixtureBuilder {

    /// Builds an 11-SDT fixture and returns the raw .docx bytes.
    ///
    /// SDTs included, in insertion order (all use stable tag strings so
    /// tests can locate them without depending on auto-allocated ids):
    /// - `richText` tag="intro"
    /// - `plainText` tag="client_name"
    /// - `picture` tag="logo"
    /// - `date` tag="issue_date"
    /// - `dropDownList` tag="priority" with items [H=High, M=Medium, L=Low]
    /// - `comboBox` tag="category" with items [A=Alpha, B=Beta]
    /// - `checkbox` tag="acceptance"
    /// - `bibliography` tag="references"
    /// - `citation` tag="cite_a"
    /// - `group` tag="address_block" containing nested `plainText` tag="city"
    /// - `repeatingSection` tag="line_items" with 3 plain items
    static func build() throws -> Data {
        var doc = WordDocument()

        // Anchor paragraph so body is non-empty before first SDT insert.
        doc.insertParagraph(Paragraph(text: "SDT Fixture Header"), at: 0)

        try doc.insertContentControl(
            .richText(tag: "intro", alias: "Introduction", content: "Rich text content here."),
            at: 1
        )
        try doc.insertContentControl(
            .plainText(tag: "client_name", alias: "Client Name", content: "ACME Corp"),
            at: 2
        )
        try doc.insertContentControl(
            .picture(tag: "logo", alias: "Company Logo"),
            at: 3
        )
        try doc.insertContentControl(
            .date(tag: "issue_date", alias: "Issue Date"),
            at: 4
        )

        // Dropdown — construct directly since no factory yet.
        let dropdownSdt = StructuredDocumentTag(
            id: 10001,
            tag: "priority",
            alias: "Priority",
            type: .dropDownList
        )
        try doc.insertContentControl(
            ContentControl(sdt: dropdownSdt, content: "High"),
            at: 5
        )

        // ComboBox — same pattern.
        let comboSdt = StructuredDocumentTag(
            id: 10002,
            tag: "category",
            alias: "Category",
            type: .comboBox
        )
        try doc.insertContentControl(
            ContentControl(sdt: comboSdt, content: "Alpha"),
            at: 6
        )

        // Checkbox — uses w14 extended namespace.
        let checkboxSdt = StructuredDocumentTag(
            id: 10003,
            tag: "acceptance",
            alias: "Acceptance",
            type: .checkbox
        )
        try doc.insertContentControl(
            ContentControl(sdt: checkboxSdt, content: ""),
            at: 7
        )

        // Bibliography SDT.
        let bibSdt = StructuredDocumentTag(
            id: 10004,
            tag: "references",
            alias: "References",
            type: .bibliography
        )
        try doc.insertContentControl(
            ContentControl(sdt: bibSdt, content: ""),
            at: 8
        )

        // Citation SDT.
        let citationSdt = StructuredDocumentTag(
            id: 10005,
            tag: "cite_a",
            alias: "Citation A",
            type: .citation
        )
        try doc.insertContentControl(
            ContentControl(sdt: citationSdt, content: "Smith, 2024"),
            at: 9
        )

        // Group SDT — content is the serialized XML of a nested plain-text SDT.
        let nestedPlainTextSdt = StructuredDocumentTag(
            id: 10007,
            tag: "city",
            alias: "City",
            type: .plainText
        )
        let nestedPlainText = ContentControl(sdt: nestedPlainTextSdt, content: "Taipei")
        let groupSdt = StructuredDocumentTag(
            id: 10006,
            tag: "address_block",
            alias: "Address Block",
            type: .group
        )
        try doc.insertContentControl(
            ContentControl(sdt: groupSdt, content: nestedPlainText.toXML()),
            at: 10
        )

        // Repeating section — uses its own code path with 3 items.
        let repeatingSection = RepeatingSection(
            tag: "line_items",
            alias: "Line Items",
            items: [
                RepeatingSectionItem(content: "Item A"),
                RepeatingSectionItem(content: "Item B"),
                RepeatingSectionItem(content: "Item C"),
            ],
            allowInsertDeleteSections: true,
            sectionTitle: "Line Items"
        )
        try doc.insertRepeatingSection(repeatingSection, at: 11)

        return try DocxWriter.writeData(doc)
    }

    /// Builds the fixture and writes it to `url`.
    static func build(to url: URL) throws {
        let data = try build()
        try data.write(to: url)
    }
}
