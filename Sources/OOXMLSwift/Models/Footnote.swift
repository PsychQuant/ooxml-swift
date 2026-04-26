import Foundation

// MARK: - Footnote

/// 腳註（出現在頁面底部）
public struct Footnote: Equatable {
    public var id: Int                 // 腳註唯一 ID
    public var text: String            // 腳註文字
    public var paragraphIndex: Int     // 腳註附加的段落索引

    /// v0.19.5+ (#56 R5 P0 #6): canonical storage for footnote body children.
    /// See `Header.bodyChildren` for rationale.
    public var bodyChildren: [BodyChild] = []

    public init(id: Int, text: String, paragraphIndex: Int) {
        self.id = id
        self.text = text
        self.paragraphIndex = paragraphIndex
    }

    /// v0.19.5+ (#56 R5 P0 #6): backward-compatible computed view of
    /// `.paragraph` cases inside `bodyChildren`. Same semantics as
    /// `Header.paragraphs`.
    public var paragraphs: [Paragraph] {
        get {
            bodyChildren.compactMap { if case .paragraph(let p) = $0 { return p } else { return nil } }
        }
        set {
            var result: [BodyChild] = []
            var newIter = newValue.makeIterator()
            for child in bodyChildren {
                switch child {
                case .paragraph:
                    if let np = newIter.next() {
                        result.append(.paragraph(np))
                    }
                case .table, .contentControl:
                    result.append(child)
                }
            }
            while let np = newIter.next() {
                result.append(.paragraph(np))
            }
            bodyChildren = result
        }
    }
}

// MARK: - Footnote XML Generation

extension Footnote {
    /// 產生 footnotes.xml 中的單一腳註 XML
    /// v0.19.5+ (#56 R5 P1 #2): when `paragraphs` is non-empty (Reader-loaded
    /// footnote), emit from the typed paragraph collection so any mutation to
    /// inner runs / hyperlinks / fieldSimples / revisions survives round-trip.
    /// Falls back to the legacy single-text-run template only for API-built
    /// footnotes constructed via `Footnote(id:text:paragraphIndex:)` without
    /// further paragraph mutation.
    func toXML() -> String {
        let inner: String
        if !bodyChildren.isEmpty {
            // v0.19.5+ (#56 R5 P0 #6): emit from bodyChildren so direct-child
            // tables round-trip in addition to paragraphs.
            // v0.19.5+ (#56 R5-CONT P1 #9): .contentControl now emits via
            // shared helper instead of being dropped — verify R5 P1 #9.
            inner = bodyChildren.map { child -> String in
                return DocxWriter.xmlForBodyChild(child)
            }.joined()
        } else {
            inner = """
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="FootnoteText"/>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:rStyle w:val="FootnoteReference"/>
                    </w:rPr>
                    <w:footnoteRef/>
                </w:r>
                <w:r>
                    <w:t xml:space="preserve"> \(escapeXML(text))</w:t>
                </w:r>
            </w:p>
            """
        }
        return """
        <w:footnote w:id="\(id)">
            \(inner)
        </w:footnote>
        """
    }

    /// 產生文件中的腳註參照標記
    func toReferenceXML() -> String {
        return """
        <w:r>
            <w:rPr>
                <w:rStyle w:val="FootnoteReference"/>
            </w:rPr>
            <w:footnoteReference w:id="\(id)"/>
        </w:r>
        """
    }

    private func escapeXML(_ string: String) -> String {
        return string
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
            .replacingOccurrences(of: "'", with: "&apos;")
    }
}

// MARK: - Footnotes Collection

/// 腳註集合
public struct FootnotesCollection: Equatable {
    public var footnotes: [Footnote] = []

    /// v0.19.2+ (#56 follow-up F4): captured `<w:footnotes>` root attributes
    /// from the source `footnotes.xml`. Empty when API-built — fallback
    /// emits `xmlns:w` + `xmlns:r` only (sufficient for default footnotes
    /// without VML / smart-art / extension elements).
    public var rootAttributes: [String: String] = [:]

    /// v0.19.5+ (#56 R5-CONT P1 #8): captured
    /// `word/_rels/footnotes.xml.rels` (per-collection relationships).
    /// Hyperlink rIds inside any footnote resolve here, NOT in
    /// `document.xml.rels`. See `Header.relationships` for full rationale.
    public var relationships: RelationshipsCollection = RelationshipsCollection()

    public init(footnotes: [Footnote] = [], rootAttributes: [String: String] = [:]) {
        self.footnotes = footnotes
        self.rootAttributes = rootAttributes
    }

    /// 取得下一個可用的腳註 ID
    mutating func nextFootnoteId() -> Int {
        // 腳註 ID 從 1 開始（0 和 -1 是保留的分隔符）
        let maxId = footnotes.map { $0.id }.max() ?? 0
        return max(1, maxId + 1)
    }

    /// 產生完整的 footnotes.xml 內容
    func toXML() -> String {
        var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        xml += ContainerRootTag.render(elementName: "w:footnotes", attributes: rootAttributes)
        xml += """
            <w:footnote w:type="separator" w:id="-1">
                <w:p>
                    <w:r>
                        <w:separator/>
                    </w:r>
                </w:p>
            </w:footnote>
            <w:footnote w:type="continuationSeparator" w:id="0">
                <w:p>
                    <w:r>
                        <w:continuationSeparator/>
                    </w:r>
                </w:p>
            </w:footnote>
        """

        for footnote in footnotes {
            xml += footnote.toXML()
        }

        xml += "</w:footnotes>"
        return xml
    }

    /// Content Type for footnotes.xml
    public static let contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"

    /// Relationship type for footnotes
    public static let relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
}

// MARK: - Endnote

/// 尾註（出現在文件結尾）
public struct Endnote: Equatable {
    public var id: Int                 // 尾註唯一 ID
    public var text: String            // 尾註文字
    public var paragraphIndex: Int     // 尾註附加的段落索引

    /// v0.19.5+ (#56 R5 P0 #6): canonical storage. See Header.bodyChildren.
    public var bodyChildren: [BodyChild] = []

    public init(id: Int, text: String, paragraphIndex: Int) {
        self.id = id
        self.text = text
        self.paragraphIndex = paragraphIndex
    }

    /// v0.19.5+ (#56 R5 P0 #6): backward-compatible computed view.
    public var paragraphs: [Paragraph] {
        get {
            bodyChildren.compactMap { if case .paragraph(let p) = $0 { return p } else { return nil } }
        }
        set {
            var result: [BodyChild] = []
            var newIter = newValue.makeIterator()
            for child in bodyChildren {
                switch child {
                case .paragraph:
                    if let np = newIter.next() {
                        result.append(.paragraph(np))
                    }
                case .table, .contentControl:
                    result.append(child)
                }
            }
            while let np = newIter.next() {
                result.append(.paragraph(np))
            }
            bodyChildren = result
        }
    }
}

// MARK: - Endnote XML Generation

extension Endnote {
    /// 產生 endnotes.xml 中的單一尾註 XML
    /// v0.19.5+ (#56 R5 P1 #2): emit from `paragraphs` when populated; same
    /// rationale as `Footnote.toXML` above.
    func toXML() -> String {
        let inner: String
        if !bodyChildren.isEmpty {
            // v0.19.5+ (#56 R5 P0 #6): emit from bodyChildren — see Footnote.toXML.
            // v0.19.5+ (#56 R5-CONT P1 #9): .contentControl now emits via
            // shared helper instead of being dropped — verify R5 P1 #9.
            inner = bodyChildren.map { child -> String in
                return DocxWriter.xmlForBodyChild(child)
            }.joined()
        } else {
            inner = """
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="EndnoteText"/>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:rStyle w:val="EndnoteReference"/>
                    </w:rPr>
                    <w:endnoteRef/>
                </w:r>
                <w:r>
                    <w:t xml:space="preserve"> \(escapeXML(text))</w:t>
                </w:r>
            </w:p>
            """
        }
        return """
        <w:endnote w:id="\(id)">
            \(inner)
        </w:endnote>
        """
    }

    /// 產生文件中的尾註參照標記
    func toReferenceXML() -> String {
        return """
        <w:r>
            <w:rPr>
                <w:rStyle w:val="EndnoteReference"/>
            </w:rPr>
            <w:endnoteReference w:id="\(id)"/>
        </w:r>
        """
    }

    private func escapeXML(_ string: String) -> String {
        return string
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
            .replacingOccurrences(of: "'", with: "&apos;")
    }
}

// MARK: - Endnotes Collection

/// 尾註集合
public struct EndnotesCollection: Equatable {
    public var endnotes: [Endnote] = []

    /// v0.19.2+ (#56 follow-up F4): captured `<w:endnotes>` root attributes
    /// from the source `endnotes.xml`. See `FootnotesCollection.rootAttributes`
    /// for full semantics.
    public var rootAttributes: [String: String] = [:]

    /// v0.19.5+ (#56 R5-CONT P1 #8): captured `word/_rels/endnotes.xml.rels`.
    /// See `Header.relationships` for full rationale.
    public var relationships: RelationshipsCollection = RelationshipsCollection()

    public init(endnotes: [Endnote] = [], rootAttributes: [String: String] = [:]) {
        self.endnotes = endnotes
        self.rootAttributes = rootAttributes
    }

    /// 取得下一個可用的尾註 ID
    mutating func nextEndnoteId() -> Int {
        // 尾註 ID 從 1 開始（0 和 -1 是保留的分隔符）
        let maxId = endnotes.map { $0.id }.max() ?? 0
        return max(1, maxId + 1)
    }

    /// 產生完整的 endnotes.xml 內容
    func toXML() -> String {
        var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        xml += ContainerRootTag.render(elementName: "w:endnotes", attributes: rootAttributes)
        xml += """
            <w:endnote w:type="separator" w:id="-1">
                <w:p>
                    <w:r>
                        <w:separator/>
                    </w:r>
                </w:p>
            </w:endnote>
            <w:endnote w:type="continuationSeparator" w:id="0">
                <w:p>
                    <w:r>
                        <w:continuationSeparator/>
                    </w:r>
                </w:p>
            </w:endnote>
        """

        for endnote in endnotes {
            xml += endnote.toXML()
        }

        xml += "</w:endnotes>"
        return xml
    }

    /// Content Type for endnotes.xml
    public static let contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"

    /// Relationship type for endnotes
    public static let relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
}

// MARK: - Errors

public enum FootnoteError: Error, LocalizedError {
    case notFound(Int)
    case invalidParagraphIndex(Int)

    public var errorDescription: String? {
        switch self {
        case .notFound(let id):
            return "Footnote with id \(id) not found"
        case .invalidParagraphIndex(let index):
            return "Invalid paragraph index: \(index)"
        }
    }
}

public enum EndnoteError: Error, LocalizedError {
    case notFound(Int)
    case invalidParagraphIndex(Int)

    public var errorDescription: String? {
        switch self {
        case .notFound(let id):
            return "Endnote with id \(id) not found"
        case .invalidParagraphIndex(let index):
            return "Invalid paragraph index: \(index)"
        }
    }
}
