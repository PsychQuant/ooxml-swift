import Foundation

// MARK: - Footer

/// 頁尾
public struct Footer: Equatable {
    public var id: String           // 關係 ID (如 "rId11")

    /// v0.19.5+ (#56 R5 P0 #6): canonical storage. See Header.bodyChildren.
    public var bodyChildren: [BodyChild] = []

    public var type: HeaderFooterType
    public var pageNumberFormat: PageNumberFormat?  // 頁碼格式（如果是頁碼頁尾）
    public var pageNumberAlignment: ParagraphAlignment  // 頁碼對齊方式

    /// Archive file path the footer was read from (v0.13.0+).
    /// `DocxReader.read()` populates this from the relationship `Target`
    /// attribute (e.g., `"footer3.xml"`). `nil` for newly-built footers
    /// — `fileName` then falls back to type-based default.
    ///
    /// **Security (v0.13.5+, che-word-mcp#55)**: setter validates against
    /// path traversal. Invalid values silently coerced to `nil`. See
    /// `Header.originalFileName` for full rationale.
    public var originalFileName: String? {
        didSet {
            originalFileName = Self.sanitizeOriginalFileName(originalFileName)
        }
    }

    /// v0.19.2+ (#56 follow-up F4): captured `<w:ftr>` root attributes from the
    /// source `footer*.xml`. Empty when API-built — fallback in `toXML()` uses
    /// the hardcoded 5-namespace template (`w`/`r`/`v`/`o`/`w10`). See
    /// `Header.rootAttributes` for full semantics.
    public var rootAttributes: [String: String] = [:]

    /// v0.19.5+ (#56 R5-CONT P1 #8): captured `word/_rels/footer*.xml.rels`
    /// (per-container relationships). See `Header.relationships` for
    /// rationale — same per-container rels gap pre-fix.
    public var relationships: RelationshipsCollection = RelationshipsCollection()

    public init(id: String, paragraphs: [Paragraph] = [], type: HeaderFooterType = .default, pageNumberFormat: PageNumberFormat? = nil, pageNumberAlignment: ParagraphAlignment = .center, originalFileName: String? = nil, rootAttributes: [String: String] = [:]) {
        self.id = id
        self.bodyChildren = paragraphs.map { .paragraph($0) }
        self.type = type
        self.pageNumberFormat = pageNumberFormat
        self.pageNumberAlignment = pageNumberAlignment
        self.originalFileName = Self.sanitizeOriginalFileName(originalFileName)
        self.rootAttributes = rootAttributes
    }

    /// v0.19.5+ (#56 R5 P0 #6): backward-compatible computed view. See
    /// `Header.paragraphs` for full semantics.
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
                case .table, .contentControl, .bookmarkMarker, .rawBlockElement:
                    // Non-paragraph children (incl. body-level markers per #58) pass through unchanged.
                    result.append(child)
                }
            }
            while let np = newIter.next() {
                result.append(.paragraph(np))
            }
            bodyChildren = result
        }
    }

    /// Sanitize per #55 security baseline. See `Header.sanitizeOriginalFileName`.
    private static func sanitizeOriginalFileName(_ candidate: String?) -> String? {
        guard let candidate = candidate else { return nil }
        if isSafeRelativeOOXMLPath(candidate) {
            return candidate
        }
        FileHandle.standardError.write(
            Data("Warning: Footer.originalFileName rejected unsafe path '\(candidate)' (#55 security baseline); falling back to type-based default\n".utf8)
        )
        return nil
    }

    /// 建立含單一文字的頁尾
    public static func withText(_ text: String, id: String, type: HeaderFooterType = .default) -> Footer {
        var para = Paragraph(text: text)
        para.properties.alignment = .center
        return Footer(id: id, paragraphs: [para], type: type)
    }

    /// 建立含頁碼的頁尾
    public static func withPageNumber(id: String, alignment: ParagraphAlignment = .center, format: PageNumberFormat = .simple, type: HeaderFooterType = .default) -> Footer {
        // 儲存格式資訊，讓 DocxWriter 能夠正確生成 XML
        return Footer(id: id, paragraphs: [], type: type, pageNumberFormat: format, pageNumberAlignment: alignment)
    }
}

// MARK: - Page Number Format

/// 頁碼格式
public enum PageNumberFormat: Equatable {
    case simple              // "1"
    case pageOfTotal         // "Page 1 of 10"
    case withDash            // "- 1 -"
    case withText(String)    // "第 1 頁" (使用 # 作為頁碼佔位符，如 "第#頁")
}

// MARK: - XML Generation

extension Footer {
    /// 轉換為完整的 footer.xml 內容
    /// v0.19.5+ (#56 R5 P0 #6): emit from bodyChildren — see Header.toXML.
    func toXML() -> String {
        var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        xml += ContainerRootTag.render(elementName: "w:ftr", attributes: rootAttributes)
        // v0.19.5+ (#56 R5-CONT P1 #9): route through shared helper so
        // .contentControl emits as <w:sdt>...</w:sdt> instead of being
        // silently dropped — see Header.toXML.
        for child in bodyChildren {
            xml += DocxWriter.xmlForBodyChild(child)
        }
        if bodyChildren.isEmpty {
            xml += "<w:p/>"
        }
        xml += "</w:ftr>"
        return xml
    }

    /// 轉換為含頁碼的頁尾 XML
    func toXMLWithPageNumber(format: PageNumberFormat = .simple, alignment: ParagraphAlignment = .center) -> String {
        var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        xml += ContainerRootTag.render(elementName: "w:ftr", attributes: rootAttributes)
        xml += "<w:p><w:pPr><w:jc w:val=\"\(alignment.rawValue)\"/></w:pPr>"

        // 根據格式添加內容
        switch format {
        case .simple:
            xml += pageFieldXML()

        case .pageOfTotal:
            xml += "<w:r><w:t xml:space=\"preserve\">Page </w:t></w:r>"
            xml += pageFieldXML()
            xml += "<w:r><w:t xml:space=\"preserve\"> of </w:t></w:r>"
            xml += numPagesFieldXML()

        case .withDash:
            xml += "<w:r><w:t xml:space=\"preserve\">- </w:t></w:r>"
            xml += pageFieldXML()
            xml += "<w:r><w:t xml:space=\"preserve\"> -</w:t></w:r>"

        case .withText(let template):
            let parts = template.components(separatedBy: "#")
            if parts.count >= 1 && !parts[0].isEmpty {
                xml += "<w:r><w:t xml:space=\"preserve\">\(escapeXML(parts[0]))</w:t></w:r>"
            }
            xml += pageFieldXML()
            if parts.count >= 2 && !parts[1].isEmpty {
                xml += "<w:r><w:t xml:space=\"preserve\">\(escapeXML(parts[1]))</w:t></w:r>"
            }
        }

        xml += "</w:p></w:ftr>"
        return xml
    }

    /// PAGE 欄位 XML
    private func pageFieldXML() -> String {
        return """
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>1</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
        """
    }

    /// NUMPAGES 欄位 XML
    private func numPagesFieldXML() -> String {
        return """
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>1</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
        """
    }

    /// XML 跳脫
    private func escapeXML(_ string: String) -> String {
        return string
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
            .replacingOccurrences(of: "'", with: "&apos;")
    }

    /// 取得檔案名稱
    /// v0.13.0+: returns `originalFileName` when present (preserves multi-instance
    /// same-type files like `footer1.xml`/`footer2.xml`/.../`footer4.xml`).
    /// Falls back to type-based default for newly-built footers.
    public var fileName: String {
        if let original = originalFileName {
            return original
        }
        switch type {
        case .default: return "footer1.xml"
        case .first: return "footerFirst.xml"
        case .even: return "footerEven.xml"
        }
    }

    /// 取得關係類型
    public static var relationshipType: String {
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
    }

    /// 取得內容類型
    public static var contentType: String {
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
    }
}
