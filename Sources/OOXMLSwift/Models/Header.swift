import Foundation

// MARK: - Header

/// 頁首
public struct Header: Equatable {
    public var id: String           // 關係 ID (如 "rId10")

    /// v0.19.5+ (#56 R5 P0 #6): canonical storage for header body children
    /// in source order. Captures both `<w:p>` (as `.paragraph`) and `<w:tbl>`
    /// (as `.table`) direct children of the source `<w:hdr>`. Pre-R5 the
    /// container parser only kept paragraphs, silently dropping table
    /// children — see DocumentWalker, calibration, and revision walker
    /// recursion that all now route through this collection.
    public var bodyChildren: [BodyChild] = []

    public var type: HeaderFooterType

    /// Archive file path the header was read from (v0.13.0+).
    /// `DocxReader.read()` populates this from the relationship `Target`
    /// attribute (e.g., `"header4.xml"`). `nil` for newly-built headers
    /// — `fileName` then falls back to type-based default.
    ///
    /// **Security (v0.13.5+, che-word-mcp#55)**: setter validates against
    /// path traversal (`..`, absolute paths, URL-encoded escapes, control
    /// chars). Invalid values are silently coerced to `nil`, falling back to
    /// the type-based default fileName. Same validation applies to the
    /// initializer parameter.
    public var originalFileName: String? {
        didSet {
            // didSet is not called during init; init runs sanitize itself.
            originalFileName = Self.sanitizeOriginalFileName(originalFileName)
        }
    }

    /// v0.19.2+ (#56 follow-up F4): captured `<w:hdr>` root attributes from the
    /// source `header*.xml` (every `xmlns:*` declaration plus `mc:Ignorable`
    /// and any vendor / unmodeled attributes). Empty when the header is
    /// API-built — `toXML()` then falls back to the hardcoded 5-namespace
    /// template that Word's exported headers minimally need (`w`/`r`/`v`/`o`/`w10`).
    /// Source-loaded headers (e.g., NTPU thesis with VML watermarks declaring
    /// `mc`/`wp`/`w14`/`w15`) round-trip every declaration verbatim.
    public var rootAttributes: [String: String] = [:]

    /// v0.19.5+ (#56 R5-CONT P1 #8): captured `word/_rels/header*.xml.rels`
    /// (per-container relationships). Hyperlink rIds inside this header
    /// resolve here, NOT in `document.xml.rels`. Pre-fix the model carried
    /// no per-container rels storage so any URL update via
    /// `Document.updateHyperlink(url:)` only landed in document-scope
    /// `hyperlinkReferences` and never reached the actual `header*.xml.rels`
    /// file → URL change silently doesn't persist for container hyperlinks.
    /// Empty for API-built headers; populated by DocxReader; emitted back
    /// by DocxWriter when dirty.
    public var relationships: RelationshipsCollection = RelationshipsCollection()

    public init(id: String, paragraphs: [Paragraph] = [], type: HeaderFooterType = .default, originalFileName: String? = nil, rootAttributes: [String: String] = [:]) {
        self.id = id
        self.bodyChildren = paragraphs.map { .paragraph($0) }
        self.type = type
        self.originalFileName = Self.sanitizeOriginalFileName(originalFileName)
        self.rootAttributes = rootAttributes
    }

    /// v0.19.5+ (#56 R5 P0 #6): backward-compatible computed view of the
    /// `.paragraph` cases inside `bodyChildren`. Get extracts in source order;
    /// set replaces every `.paragraph` slot in `bodyChildren` (in order),
    /// preserving relative positions of `.table` siblings. If the new array
    /// is longer than the current paragraph count, extras append to the end.
    /// Existing call sites that mutate `header.paragraphs[i].xxx` continue to
    /// work via Swift's modify accessor on a `var` computed property.
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

    /// Sanitize a candidate `originalFileName` per #55 security baseline.
    /// Returns the original value when safe; nil when traversal/absolute/
    /// control-char detected. Logs rejections to stderr for observability.
    private static func sanitizeOriginalFileName(_ candidate: String?) -> String? {
        guard let candidate = candidate else { return nil }
        if isSafeRelativeOOXMLPath(candidate) {
            return candidate
        }
        FileHandle.standardError.write(
            Data("Warning: Header.originalFileName rejected unsafe path '\(candidate)' (#55 security baseline); falling back to type-based default\n".utf8)
        )
        return nil
    }

    /// 建立含單一文字的頁首
    public static func withText(_ text: String, id: String, type: HeaderFooterType = .default) -> Header {
        var para = Paragraph(text: text)
        para.properties.alignment = .center
        return Header(id: id, paragraphs: [para], type: type)
    }

    /// 建立含頁碼的頁首
    public static func withPageNumber(id: String, alignment: ParagraphAlignment = .center, type: HeaderFooterType = .default) -> Header {
        // 頁碼會在 XML 生成時處理
        var para = Paragraph()
        para.properties.alignment = alignment
        return Header(id: id, paragraphs: [para], type: type)
    }
}

// MARK: - Header/Footer Type

/// 頁首/頁尾類型
public enum HeaderFooterType: String {
    case `default` = "default"  // 預設（奇數頁/所有頁）
    case first = "first"        // 首頁
    case even = "even"          // 偶數頁
}

// MARK: - XML Generation

extension Header {
    /// 轉換為完整的 header.xml 內容
    ///
    /// **VML watermark / OLE preservation chain (v0.14.0+, che-word-mcp#52)**:
    /// `Header.toXML()` → `Paragraph.toXML()` → `Run.toXML()`. The Run-layer
    /// `rawElements` carrier (per `ooxml-header-footer-raw-element-preservation`
    /// capability) emits unknown OOXML elements verbatim after typed children.
    /// `<w:hdr>` declares `xmlns:v` / `xmlns:o` / `xmlns:w10` so descendant
    /// `<v:shape>` / `<o:lock>` / `<w10:wrap>` resolve when the saved
    /// `header*.xml` is re-read.
    func toXML() -> String {
        var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        xml += ContainerRootTag.render(elementName: "w:hdr", attributes: rootAttributes)
        // v0.19.5+ (#56 R5 P0 #6): emit from bodyChildren so direct-child
        // tables round-trip.
        // v0.19.5+ (#56 R5-CONT P1 #9): .contentControl now routes through
        // the shared `DocxWriter.xmlForBodyChild` helper (recursive SDT
        // serialization including nested children). Pre-fix the case
        // arm was `break` which silently dropped any block-level SDT held
        // in `bodyChildren` — verify R5 P1 #9 / Logic L6.
        for child in bodyChildren {
            xml += DocxWriter.xmlForBodyChild(child)
        }
        // 如果沒有段落且沒有任何 bodyChildren，加一個空段落
        if bodyChildren.isEmpty {
            xml += "<w:p/>"
        }
        xml += "</w:hdr>"
        return xml
    }

    /// 轉換為含頁碼的頁首 XML
    func toXMLWithPageNumber() -> String {
        var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        xml += ContainerRootTag.render(elementName: "w:hdr", attributes: rootAttributes)
        xml += "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr>"
        xml += pageFieldXML()
        xml += "</w:p></w:hdr>"
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

    /// 取得檔案名稱
    /// v0.13.0+: returns `originalFileName` when present (preserves multi-instance
    /// same-type files like `header1.xml`/`header2.xml`/.../`header6.xml`).
    /// Falls back to type-based default for newly-built headers.
    public var fileName: String {
        if let original = originalFileName {
            return original
        }
        switch type {
        case .default: return "header1.xml"
        case .first: return "headerFirst.xml"
        case .even: return "headerEven.xml"
        }
    }

    /// 取得關係類型
    public static var relationshipType: String {
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    }

    /// 取得內容類型
    public static var contentType: String {
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
    }
}
