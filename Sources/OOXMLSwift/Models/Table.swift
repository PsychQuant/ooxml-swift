import Foundation

/// 表格 (Table) - Word 文件中的表格結構
public struct Table: Equatable {
    /// v0.31.1+ (Spectra `sibling-types-tree-projection-impl`,
    /// `word-aligned-state-sync` Phase 1 task 2.3): when non-nil, this Table
    /// is a tree-backed view over the wrapped `<w:tbl>` element. Getters walk
    /// `xmlNode.children` at access time; setters are Phase 1 ghost-writes to
    /// the legacy buffer below (proper tree-mutating routing arrives with the
    /// op-log path in Phase 2). When nil (legacy detached mode), getters /
    /// setters operate on the stored properties below. `XmlNode` is a class,
    /// so two value-copies of the same tree-backed Table share the same
    /// underlying tree state.
    public var xmlNode: XmlNode?

    /// Legacy stored backing for `rows` used in detached mode.
    /// In tree-backed mode the public `rows` accessor walks `xmlNode.children`
    /// instead and this buffer is ignored. Renamed from the previous public
    /// `rows` stored property in v0.31.1 (`sibling-types-tree-projection-impl`).
    internal var _legacyRows: [TableRow] = []
    public var properties: TableProperties

    /// v0.17.0+ (#49): conditional formatting blocks (firstRow / lastRow / banded etc.)
    /// emitted as `<w:tblStylePr w:type="...">` inside `<w:tblPr>`.
    public var conditionalStyles: [TableConditionalStyle] = []

    /// v0.17.0+ (#49): explicit `<w:tblInd>` table-level left indent (twips).
    public var tableIndent: Int?

    /// v0.17.0+ (#49): explicit table layout mode (`<w:tblLayout w:type>`).
    /// nil means inherit from parent / default.
    public var explicitLayout: TableLayout?

    public init(rows: [TableRow] = [], properties: TableProperties = TableProperties()) {
        self._legacyRows = rows
        self.properties = properties
    }

    /// 便利初始化器：建立指定行列的空表格
    public init(rowCount: Int, columnCount: Int, properties: TableProperties = TableProperties()) {
        self.properties = properties
        self._legacyRows = (0..<rowCount).map { _ in
            TableRow(cells: (0..<columnCount).map { _ in TableCell() })
        }
    }

    /// v0.31.1+ Tree-backed initializer. Wraps an existing `<w:tbl>` xmlNode
    /// so getters walk its children and setters mutate the tree directly.
    ///
    /// The legacy stored fields (`_legacyRows`, `properties`,
    /// `conditionalStyles`, `tableIndent`, `explicitLayout`) are initialized
    /// to their empty defaults; in tree-backed mode they are shadowed by
    /// computed accessors that read from `xmlNode.children`. Callers MUST
    /// NOT rely on those stored fields when `xmlNode != nil`.
    ///
    /// Semantic validation (asserting the node is a `<w:tbl>` element) is
    /// left to callers; this initializer accepts any element xmlNode so
    /// unit tests can synthesize fixtures without paying the schema-check
    /// cost.
    public init(xmlNode: XmlNode) {
        self.xmlNode = xmlNode
        self.properties = TableProperties()
    }

    /// v0.31.1+ Mode-aware view of `<w:tr>` rows.
    ///
    /// - Tree-backed: walks `xmlNode.children` at every access and returns
    ///   one `TableRow(xmlNode:)` per `<w:tr>` direct child in document
    ///   order. No caching — re-walks on every access so direct tree
    ///   mutations are observable immediately.
    /// - Detached: returns the legacy stored buffer.
    ///
    /// Setter writes to `_legacyRows` in both modes (Phase 1 ghost-write
    /// per Decision 4 of `sibling-types-tree-projection-impl`). In
    /// tree-backed mode the write is not visible through the getter;
    /// proper tree-mutating routing arrives with the op-log path in
    /// Phase 2.
    public var rows: [TableRow] {
        get {
            guard let node = xmlNode else { return _legacyRows }
            return node.children.compactMap { child -> TableRow? in
                guard child.kind == .element, child.localName == "tr" else { return nil }
                return TableRow(xmlNode: child)
            }
        }
        set { _legacyRows = newValue }
    }

    /// v0.31.1+ Stable identifier for this table.
    ///
    /// - Tree-backed: returns `xmlNode.stableID` if any OOXML stable-ID
    ///   attribute is present; otherwise falls back to `"lib:<UUID>"` when
    ///   the reader assigned a library-generated UUID; otherwise `nil`.
    /// - Detached (legacy): always returns `nil`.
    ///
    /// Note: `<w:tbl>` does not natively carry `w14:paraId`; its stable IDs
    /// come from `w:id` (revision tracking), `r:id` (relationship), or the
    /// `lib:` UUID fallback assigned by the reader.
    public var id: String? {
        guard let node = xmlNode else { return nil }
        if let stable = node.stableID { return stable }
        if let lib = node.libraryUUID { return "lib:\(lib.uuidString)" }
        return nil
    }

    /// 取得表格純文字
    public func getText() -> String {
        return rows.map { row in
            row.cells.map { cell in
                cell.getText()
            }.joined(separator: "\t")
        }.joined(separator: "\n")
    }

}

// MARK: - Equatable (mode-aware identity vs content)

extension Table {
    /// v0.31.1+ Custom Equatable replacing auto-synthesized conformance per
    /// `sibling-types-tree-projection-impl` Decision 5.
    ///
    /// Behavior depends on the storage mode of both sides:
    ///
    /// 1. **Both tree-backed**: identity equality on the wrapped `xmlNode`
    ///    reference (`===`). Op-log addresses tables by id (== identity);
    ///    content equality on different elements would silently merge log
    ///    entries that target different tables.
    /// 2. **Both detached**: content equality across the legacy stored
    ///    fields, preserving pre-v0.31.1 auto-synthesized behavior.
    /// 3. **Mixed (one tree-backed, one detached)**: always `false`. The
    ///    two storage modes are not interchangeable; comparing across them
    ///    is almost certainly a caller mistake worth surfacing.
    public static func == (lhs: Table, rhs: Table) -> Bool {
        switch (lhs.xmlNode, rhs.xmlNode) {
        case let (a?, b?):
            return a === b
        case (nil, nil):
            return contentEquals(lhs, rhs)
        default:
            return false
        }
    }

    /// Detached-mode content equality across all legacy stored fields.
    /// Mirrors what auto-synthesized `Equatable` would have compared.
    private static func contentEquals(_ lhs: Table, _ rhs: Table) -> Bool {
        return lhs._legacyRows == rhs._legacyRows
            && lhs.properties == rhs.properties
            && lhs.conditionalStyles == rhs.conditionalStyles
            && lhs.tableIndent == rhs.tableIndent
            && lhs.explicitLayout == rhs.explicitLayout
    }
}

// MARK: - Table Row

/// 表格行
public struct TableRow: Equatable {
    /// v0.31.1+ Tree-backed view marker; see `Table.xmlNode` for the same
    /// contract.
    public var xmlNode: XmlNode?

    /// Legacy stored backing for `cells` used in detached mode.
    /// In tree-backed mode the public `cells` accessor walks
    /// `xmlNode.children` instead and this buffer is ignored.
    internal var _legacyCells: [TableCell] = []
    public var properties: TableRowProperties

    public init(cells: [TableCell] = [], properties: TableRowProperties = TableRowProperties()) {
        self._legacyCells = cells
        self.properties = properties
    }

    /// v0.31.1+ Tree-backed initializer wrapping an existing `<w:tr>` xmlNode.
    /// Legacy stored fields (`_legacyCells`, `properties`) initialize to
    /// their empty defaults; in tree-backed mode they are shadowed by the
    /// computed `cells` accessor.
    public init(xmlNode: XmlNode) {
        self.xmlNode = xmlNode
        self.properties = TableRowProperties()
    }

    /// v0.31.1+ Mode-aware view of `<w:tc>` cells.
    ///
    /// - Tree-backed: walks `xmlNode.children` at every access and returns
    ///   one `TableCell(xmlNode:)` per `<w:tc>` direct child in document
    ///   order.
    /// - Detached: returns the legacy stored buffer.
    ///
    /// Setter writes to `_legacyCells` in both modes (Phase 1 ghost-write).
    public var cells: [TableCell] {
        get {
            guard let node = xmlNode else { return _legacyCells }
            return node.children.compactMap { child -> TableCell? in
                guard child.kind == .element, child.localName == "tc" else { return nil }
                return TableCell(xmlNode: child)
            }
        }
        set { _legacyCells = newValue }
    }

    /// v0.31.1+ Stable identifier for this row. Same fallback chain as
    /// `Table.id`: `xmlNode.stableID` → `"lib:<UUID>"` → nil.
    public var id: String? {
        guard let node = xmlNode else { return nil }
        if let stable = node.stableID { return stable }
        if let lib = node.libraryUUID { return "lib:\(lib.uuidString)" }
        return nil
    }
}

extension TableRow {
    /// v0.31.1+ Custom Equatable replacing auto-synthesized conformance per
    /// `sibling-types-tree-projection-impl` Decision 5.
    public static func == (lhs: TableRow, rhs: TableRow) -> Bool {
        switch (lhs.xmlNode, rhs.xmlNode) {
        case let (a?, b?):
            return a === b
        case (nil, nil):
            return contentEquals(lhs, rhs)
        default:
            return false
        }
    }

    private static func contentEquals(_ lhs: TableRow, _ rhs: TableRow) -> Bool {
        return lhs._legacyCells == rhs._legacyCells
            && lhs.properties == rhs.properties
    }
}

/// 表格行屬性
public struct TableRowProperties: Equatable {
    public var height: Int?                // 行高 (twips)
    public var heightRule: HeightRule?     // 行高規則
    public var isHeader: Bool = false      // 是否為表頭行（每頁重複）
    public var cantSplit: Bool = false     // 禁止跨頁分割

    public init() {}
}

/// 行高規則
public enum HeightRule: String, Codable {
    case auto = "auto"
    case exact = "exact"
    case atLeast = "atLeast"
}

// MARK: - Table Cell

/// 表格儲存格
public struct TableCell: Equatable {
    /// v0.31.1+ Tree-backed view marker; see `Table.xmlNode` for the same
    /// contract.
    public var xmlNode: XmlNode?

    /// Legacy stored backing for `paragraphs` used in detached mode.
    /// In tree-backed mode the public `paragraphs` accessor walks
    /// `xmlNode.children` instead and this buffer is ignored.
    internal var _legacyParagraphs: [Paragraph] = []
    public var properties: TableCellProperties

    /// Legacy stored backing for `nestedTables` used in detached mode.
    /// In tree-backed mode the public `nestedTables` accessor walks
    /// `xmlNode.children` instead and this buffer is ignored.
    ///
    /// v0.17.0+ (#49): nested tables inside this cell (depth-limited to 5
    /// at parser layer per design.md decision 1). Distinct from `paragraphs`
    /// so the writer can emit `<w:tbl>` siblings of `<w:p>` correctly.
    internal var _legacyNestedTables: [Table] = []

    public init() {
        self._legacyParagraphs = [Paragraph()]
        self.properties = TableCellProperties()
    }

    public init(paragraphs: [Paragraph], properties: TableCellProperties = TableCellProperties()) {
        self._legacyParagraphs = paragraphs.isEmpty ? [Paragraph()] : paragraphs
        self.properties = properties
    }

    /// 便利初始化器：用文字建立儲存格
    public init(text: String) {
        self._legacyParagraphs = [Paragraph(text: text)]
        self.properties = TableCellProperties()
    }

    /// v0.31.1+ Tree-backed initializer wrapping an existing `<w:tc>` xmlNode.
    /// Legacy stored fields (`_legacyParagraphs`, `properties`,
    /// `_legacyNestedTables`) initialize to their empty defaults; in
    /// tree-backed mode they are shadowed by the computed `paragraphs` /
    /// `nestedTables` accessors.
    public init(xmlNode: XmlNode) {
        self.xmlNode = xmlNode
        self.properties = TableCellProperties()
    }

    /// v0.31.1+ Mode-aware view of `<w:p>` paragraphs.
    ///
    /// - Tree-backed: walks `xmlNode.children` at every access and returns
    ///   one `Paragraph(xmlNode:)` per `<w:p>` direct child in document
    ///   order. The returned `Paragraph` values are themselves tree-backed
    ///   (uses the v0.31.0 `Paragraph(xmlNode:)` constructor).
    /// - Detached: returns the legacy stored buffer.
    ///
    /// Setter writes to `_legacyParagraphs` in both modes (Phase 1
    /// ghost-write).
    public var paragraphs: [Paragraph] {
        get {
            guard let node = xmlNode else { return _legacyParagraphs }
            return node.children.compactMap { child -> Paragraph? in
                guard child.kind == .element, child.localName == "p" else { return nil }
                return Paragraph(xmlNode: child)
            }
        }
        set { _legacyParagraphs = newValue }
    }

    /// v0.31.1+ Mode-aware view of `<w:tbl>` nested tables.
    ///
    /// - Tree-backed: walks `xmlNode.children` at every access and returns
    ///   one `Table(xmlNode:)` per `<w:tbl>` direct child in document
    ///   order.
    /// - Detached: returns the legacy stored buffer.
    ///
    /// Setter writes to `_legacyNestedTables` in both modes (Phase 1
    /// ghost-write).
    public var nestedTables: [Table] {
        get {
            guard let node = xmlNode else { return _legacyNestedTables }
            return node.children.compactMap { child -> Table? in
                guard child.kind == .element, child.localName == "tbl" else { return nil }
                return Table(xmlNode: child)
            }
        }
        set { _legacyNestedTables = newValue }
    }

    /// v0.31.1+ Stable identifier for this cell. Same fallback chain as
    /// `Table.id`: `xmlNode.stableID` → `"lib:<UUID>"` → nil.
    public var id: String? {
        guard let node = xmlNode else { return nil }
        if let stable = node.stableID { return stable }
        if let lib = node.libraryUUID { return "lib:\(lib.uuidString)" }
        return nil
    }

    /// 取得儲存格純文字
    public func getText() -> String {
        return paragraphs.map { $0.getText() }.joined(separator: "\n")
    }
}

extension TableCell {
    /// v0.31.1+ Custom Equatable replacing auto-synthesized conformance per
    /// `sibling-types-tree-projection-impl` Decision 5.
    public static func == (lhs: TableCell, rhs: TableCell) -> Bool {
        switch (lhs.xmlNode, rhs.xmlNode) {
        case let (a?, b?):
            return a === b
        case (nil, nil):
            return contentEquals(lhs, rhs)
        default:
            return false
        }
    }

    private static func contentEquals(_ lhs: TableCell, _ rhs: TableCell) -> Bool {
        return lhs._legacyParagraphs == rhs._legacyParagraphs
            && lhs.properties == rhs.properties
            && lhs._legacyNestedTables == rhs._legacyNestedTables
    }
}

/// 表格儲存格屬性
public struct TableCellProperties: Equatable {
    public var width: Int?                     // 寬度 (twips)
    public var widthType: WidthType?           // 寬度類型
    public var verticalAlignment: CellVerticalAlignment?
    public var gridSpan: Int?                  // 水平合併（跨幾欄）
    public var verticalMerge: VerticalMerge?   // 垂直合併
    public var borders: CellBorders?           // 邊框
    public var shading: CellShading?           // 底色

    public init() {}
}

/// 寬度類型
public enum WidthType: String, Codable {
    case auto = "auto"
    case dxa = "dxa"        // twips
    case pct = "pct"        // 百分比 (50 = 50%)
    case nil_ = "nil"       // 無寬度
}

/// 儲存格垂直對齊
public enum CellVerticalAlignment: String, Codable {
    case top = "top"
    case center = "center"
    case bottom = "bottom"
}

/// 垂直合併
public enum VerticalMerge: String, Codable {
    case restart = "restart"    // 合併的第一個儲存格
    case `continue` = "continue" // 被合併的儲存格
}

/// 儲存格邊框
public struct CellBorders: Equatable {
    public var top: Border?
    public var bottom: Border?
    public var left: Border?
    public var right: Border?

    /// v0.17.0+ (#49): diagonal borders.
    /// `tl2br` = top-left to bottom-right (`<w:tl2br>`)
    /// `tr2bl` = top-right to bottom-left (`<w:tr2bl>`)
    public var tl2br: Border?
    public var tr2bl: Border?

    public init(top: Border? = nil, bottom: Border? = nil, left: Border? = nil, right: Border? = nil,
                tl2br: Border? = nil, tr2bl: Border? = nil) {
        self.top = top
        self.bottom = bottom
        self.left = left
        self.right = right
        self.tl2br = tl2br
        self.tr2bl = tr2bl
    }

    /// 便利方法：建立四邊相同邊框
    public static func all(_ border: Border) -> CellBorders {
        CellBorders(top: border, bottom: border, left: border, right: border)
    }
}

/// 邊框
public struct Border: Equatable {
    public var style: BorderStyle
    public var size: Int           // 1/8 點
    public var color: String       // RGB hex

    public init(style: BorderStyle = .single, size: Int = 4, color: String = "000000") {
        self.style = style
        self.size = size
        self.color = color
    }
}

/// 邊框樣式
public enum BorderStyle: String, Codable {
    case single = "single"
    case double = "double"
    case dotted = "dotted"
    case dashed = "dashed"
    case thick = "thick"
    case nil_ = "nil"           // 無邊框
}

/// 儲存格底色
public struct CellShading: Equatable {
    public var fill: String            // 背景色 RGB hex
    public var color: String?          // 前景色（用於圖案）
    public var pattern: ShadingPattern?

    public init(fill: String, color: String? = nil, pattern: ShadingPattern? = nil) {
        self.fill = fill
        self.color = color
        self.pattern = pattern
    }

    /// 便利方法：純色背景
    public static func solid(_ color: String) -> CellShading {
        CellShading(fill: color, pattern: .clear)
    }

    /// 產生 XML 字串（供段落屬性使用）
    public func toXML() -> String {
        // v0.19.5+ (#56 R5 P0 #3): caller-controlled fill / color routed
        // through escapeXMLAttribute (MCP `set_paragraph_shading`,
        // `set_table_conditional_style`).
        var attrs = ["w:fill=\"\(escapeXMLAttribute(fill))\""]

        if let pattern = pattern {
            attrs.insert("w:val=\"\(pattern.rawValue)\"", at: 0)
        } else {
            attrs.insert("w:val=\"clear\"", at: 0)
        }

        if let color = color {
            attrs.append("w:color=\"\(escapeXMLAttribute(color))\"")
        }

        return "<w:shd \(attrs.joined(separator: " "))/>"
    }
}

/// 底色圖案
public enum ShadingPattern: String, Codable {
    case clear = "clear"
    case solid = "solid"
    case horzStripe = "horzStripe"
    case vertStripe = "vertStripe"
    case diagStripe = "diagStripe"
}

// MARK: - Table Properties

/// 表格屬性
public struct TableProperties: Equatable {
    public var width: Int?                     // 表格寬度
    public var widthType: WidthType?
    public var alignment: Alignment?           // 表格對齊
    public var borders: TableBorders?          // 表格邊框
    public var cellMargins: TableCellMargins?  // 預設儲存格邊距
    public var layout: TableLayout?            // 版面配置

    public init() {}
}

/// 表格邊框
public struct TableBorders: Equatable {
    public var top: Border?
    public var bottom: Border?
    public var left: Border?
    public var right: Border?
    public var insideH: Border?    // 內部水平線
    public var insideV: Border?    // 內部垂直線

    public init() {}

    /// 便利方法：建立全邊框
    public static func all(_ border: Border) -> TableBorders {
        var borders = TableBorders()
        borders.top = border
        borders.bottom = border
        borders.left = border
        borders.right = border
        borders.insideH = border
        borders.insideV = border
        return borders
    }
}

/// 表格儲存格邊距
public struct TableCellMargins: Equatable {
    public var top: Int?
    public var bottom: Int?
    public var left: Int?
    public var right: Int?

    public init(top: Int? = nil, bottom: Int? = nil, left: Int? = nil, right: Int? = nil) {
        self.top = top
        self.bottom = bottom
        self.left = left
        self.right = right
    }

    /// 便利方法：四邊相同邊距
    public static func all(_ margin: Int) -> TableCellMargins {
        TableCellMargins(top: margin, bottom: margin, left: margin, right: margin)
    }
}

/// 表格版面配置
public enum TableLayout: String, Codable {
    case fixed = "fixed"        // 固定欄寬
    case autofit = "autofit"    // 自動調整
}

// MARK: - v0.17.0+ (#49) Conditional Formatting

/// `<w:tblStylePr w:type>` discriminator — region of the table that the
/// formatting applies to.
public enum TableConditionalStyleType: String, Codable {
    case firstRow = "firstRow"
    case lastRow = "lastRow"
    case firstCol = "firstCol"
    case lastCol = "lastCol"
    case bandedRows = "band1Horz"   // OOXML uses band1Horz for odd rows
    case bandedCols = "band1Vert"   // OOXML uses band1Vert for odd cols
    case neCell = "neCell"
    case nwCell = "nwCell"
    case seCell = "seCell"
    case swCell = "swCell"
}

/// Properties applied to a conditional formatting region. Sparse — only
/// non-nil fields are emitted.
public struct TableConditionalStyleProperties: Equatable {
    public var bold: Bool?
    public var italic: Bool?
    public var color: String?           // RGB hex
    public var backgroundColor: String? // RGB hex (becomes `<w:shd>`)
    public var fontSize: Int?           // 半點 (half-points) — caller passes pt*2

    public init(
        bold: Bool? = nil,
        italic: Bool? = nil,
        color: String? = nil,
        backgroundColor: String? = nil,
        fontSize: Int? = nil
    ) {
        self.bold = bold
        self.italic = italic
        self.color = color
        self.backgroundColor = backgroundColor
        self.fontSize = fontSize
    }
}

/// One `<w:tblStylePr>` block: which region + the formatting to apply.
public struct TableConditionalStyle: Equatable {
    public var type: TableConditionalStyleType
    public var properties: TableConditionalStyleProperties

    public init(type: TableConditionalStyleType, properties: TableConditionalStyleProperties) {
        self.type = type
        self.properties = properties
    }

    /// Render to OOXML `<w:tblStylePr>` block.
    func toXML() -> String {
        var rPrParts: [String] = []
        if properties.bold == true { rPrParts.append("<w:b/>") }
        if properties.italic == true { rPrParts.append("<w:i/>") }
        if let c = properties.color {
            rPrParts.append("<w:color w:val=\"\(escapeXMLAttribute(c))\"/>")
        }
        if let sz = properties.fontSize {
            rPrParts.append("<w:sz w:val=\"\(sz)\"/>")
        }

        var tcPrParts: [String] = []
        if let bg = properties.backgroundColor {
            tcPrParts.append("<w:shd w:val=\"clear\" w:fill=\"\(escapeXMLAttribute(bg))\"/>")
        }

        var xml = "<w:tblStylePr w:type=\"\(type.rawValue)\">"
        if !rPrParts.isEmpty {
            xml += "<w:rPr>" + rPrParts.joined() + "</w:rPr>"
        }
        if !tcPrParts.isEmpty {
            xml += "<w:tcPr>" + tcPrParts.joined() + "</w:tcPr>"
        }
        xml += "</w:tblStylePr>"
        return xml
    }
}

// MARK: - XML 生成

extension Table {
    /// 轉換為 OOXML XML 字串
    public func toXML() -> String {
        var xml = "<w:tbl>"

        // Table Properties (extended in v0.17.0+ to inject tblInd / explicit
        // layout / conditional styles into the existing tblPr block)
        xml += extendedTablePropertiesXML()

        // Table Grid (欄位定義)
        xml += "<w:tblGrid>"
        if let firstRow = rows.first {
            for cell in firstRow.cells {
                let width = cell.properties.width ?? 2000
                xml += "<w:gridCol w:w=\"\(width)\"/>"
            }
        }
        xml += "</w:tblGrid>"

        // Rows
        for row in rows {
            xml += row.toXML()
        }

        xml += "</w:tbl>"
        return xml
    }
}

extension Table {
    /// v0.17.0+ (#49): emit Table-level extensions (tblInd, conditional styles,
    /// explicit layout) into the existing tblPr block. Wrapper around
    /// `TableProperties.toXML` with table-level fields injected.
    func extendedTablePropertiesXML() -> String {
        let baseXML = properties.toXML()
        var injected = ""
        // Inject ind / explicit layout / conditional styles inside <w:tblPr>.
        if let indent = tableIndent {
            injected += "<w:tblInd w:w=\"\(indent)\" w:type=\"dxa\"/>"
        }
        if let layout = explicitLayout {
            injected += "<w:tblLayout w:type=\"\(layout.rawValue)\"/>"
        }
        for cs in conditionalStyles {
            injected += cs.toXML()
        }
        guard !injected.isEmpty else { return baseXML }
        // Inject before closing </w:tblPr>.
        if let range = baseXML.range(of: "</w:tblPr>") {
            return baseXML.replacingCharacters(in: range, with: injected + "</w:tblPr>")
        }
        return baseXML
    }
}

extension TableProperties {
    public func toXML() -> String {
        var parts: [String] = ["<w:tblPr>"]

        // 寬度
        if let width = width {
            let type = widthType?.rawValue ?? "dxa"
            parts.append("<w:tblW w:w=\"\(width)\" w:type=\"\(type)\"/>")
        }

        // 對齊
        if let alignment = alignment {
            parts.append("<w:jc w:val=\"\(alignment.rawValue)\"/>")
        }

        // 版面配置
        if let layout = layout {
            parts.append("<w:tblLayout w:type=\"\(layout.rawValue)\"/>")
        }

        // 邊框
        if let borders = borders {
            parts.append(borders.toXML())
        }

        // 儲存格邊距
        if let margins = cellMargins {
            parts.append("<w:tblCellMar>")
            if let top = margins.top { parts.append("<w:top w:w=\"\(top)\" w:type=\"dxa\"/>") }
            if let bottom = margins.bottom { parts.append("<w:bottom w:w=\"\(bottom)\" w:type=\"dxa\"/>") }
            if let left = margins.left { parts.append("<w:left w:w=\"\(left)\" w:type=\"dxa\"/>") }
            if let right = margins.right { parts.append("<w:right w:w=\"\(right)\" w:type=\"dxa\"/>") }
            parts.append("</w:tblCellMar>")
        }

        parts.append("</w:tblPr>")
        return parts.joined()
    }
}

extension TableBorders {
    public func toXML() -> String {
        var parts: [String] = ["<w:tblBorders>"]

        if let top = top { parts.append(top.toXML(name: "top")) }
        if let bottom = bottom { parts.append(bottom.toXML(name: "bottom")) }
        if let left = left { parts.append(left.toXML(name: "left")) }
        if let right = right { parts.append(right.toXML(name: "right")) }
        if let insideH = insideH { parts.append(insideH.toXML(name: "insideH")) }
        if let insideV = insideV { parts.append(insideV.toXML(name: "insideV")) }

        parts.append("</w:tblBorders>")
        return parts.joined()
    }
}

extension Border {
    public func toXML(name: String) -> String {
        // v0.19.5+ (#56 R5 P0 #3): caller-controlled color routed through escape.
        return "<w:\(name) w:val=\"\(style.rawValue)\" w:sz=\"\(size)\" w:color=\"\(escapeXMLAttribute(color))\"/>"
    }
}

extension TableRow {
    public func toXML() -> String {
        var xml = "<w:tr>"

        // Row Properties
        if properties.height != nil || properties.isHeader || properties.cantSplit {
            xml += "<w:trPr>"
            if let height = properties.height {
                let rule = properties.heightRule?.rawValue ?? "auto"
                xml += "<w:trHeight w:val=\"\(height)\" w:hRule=\"\(rule)\"/>"
            }
            if properties.isHeader {
                xml += "<w:tblHeader/>"
            }
            if properties.cantSplit {
                xml += "<w:cantSplit/>"
            }
            xml += "</w:trPr>"
        }

        // Cells
        for cell in cells {
            xml += cell.toXML()
        }

        xml += "</w:tr>"
        return xml
    }
}

extension TableCell {
    public func toXML() -> String {
        var xml = "<w:tc>"

        // Cell Properties
        xml += properties.toXML()

        // Paragraphs (每個儲存格至少需要一個段落)
        if paragraphs.isEmpty {
            xml += Paragraph().toXML()
        } else {
            for para in paragraphs {
                xml += para.toXML()
            }
        }

        // v0.17.0+ (#49): nested tables emit as siblings of paragraphs
        for nested in nestedTables {
            xml += nested.toXML()
        }

        // OOXML requires the cell to end with a paragraph after a nested table
        // (Word silently appends one if missing). Append a trailing empty
        // paragraph so re-saving produces compliant output.
        if !nestedTables.isEmpty {
            xml += "<w:p/>"
        }

        xml += "</w:tc>"
        return xml
    }
}

extension TableCellProperties {
    public func toXML() -> String {
        var parts: [String] = ["<w:tcPr>"]

        // 寬度
        if let width = width {
            let type = widthType?.rawValue ?? "dxa"
            parts.append("<w:tcW w:w=\"\(width)\" w:type=\"\(type)\"/>")
        }

        // 水平合併
        if let gridSpan = gridSpan, gridSpan > 1 {
            parts.append("<w:gridSpan w:val=\"\(gridSpan)\"/>")
        }

        // 垂直合併
        if let vMerge = verticalMerge {
            parts.append("<w:vMerge w:val=\"\(vMerge.rawValue)\"/>")
        }

        // 垂直對齊
        if let vAlign = verticalAlignment {
            parts.append("<w:vAlign w:val=\"\(vAlign.rawValue)\"/>")
        }

        // 邊框
        if let borders = borders {
            parts.append("<w:tcBorders>")
            if let top = borders.top { parts.append(top.toXML(name: "top")) }
            if let bottom = borders.bottom { parts.append(bottom.toXML(name: "bottom")) }
            if let left = borders.left { parts.append(left.toXML(name: "left")) }
            if let right = borders.right { parts.append(right.toXML(name: "right")) }
            // v0.17.0+ (#49): diagonal borders
            if let tl2br = borders.tl2br { parts.append(tl2br.toXML(name: "tl2br")) }
            if let tr2bl = borders.tr2bl { parts.append(tr2bl.toXML(name: "tr2bl")) }
            parts.append("</w:tcBorders>")
        }

        // 底色
        if let shading = shading {
            // v0.19.5+ (#56 R5 P0 #3): caller-controlled fill / color routed
            // through escape (MCP `set_paragraph_shading` for cell shading).
            var attrs = "w:fill=\"\(escapeXMLAttribute(shading.fill))\""
            if let color = shading.color { attrs += " w:color=\"\(escapeXMLAttribute(color))\"" }
            if let pattern = shading.pattern { attrs += " w:val=\"\(pattern.rawValue)\"" }
            parts.append("<w:shd \(attrs)/>")
        }

        parts.append("</w:tcPr>")
        return parts.joined()
    }
}
