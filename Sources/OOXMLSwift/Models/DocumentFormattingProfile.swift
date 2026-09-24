import Foundation
import ZIPFoundation

/// The caller supplies creation intent; archive presence does not imply it.
public enum DocumentFormattingContext: Sendable { case newDocument, existingDocument }

public enum DocumentFormattingProfileError: Error, Equatable, LocalizedError {
    case unsupportedVersion(Int)
    case missingRequiredFormatting(String)
    case invalidSnapshot(String)
    case unsupportedFormatting(String)
    case unsupportedNumbering
    /// Two main-part relationships of one implicit Type (styles, theme or
    /// fontTable) resolve to different parts: a malformed package, refused
    /// on import, apply and write before anything is written
    /// (PsychQuant/macdoc#196). The payload is the relationship Type suffix.
    case duplicateRelationship(String)
    /// A main-part styles/theme/fontTable relationship whose Target resolves
    /// to a path ending in `/` (for example `styles.xml/.`), which is not an
    /// OPC part name (PsychQuant/macdoc#196). The payload is the Target as
    /// written.
    case invalidRelationshipTarget(String)

    public var errorDescription: String? {
        switch self {
        case .unsupportedVersion(let version): return "不支援的格式快照版本：\(version)"
        case .missingRequiredFormatting(let field):
            // PsychQuant/macdoc#212: a known missing field carries its fix.
            let fields = field.components(separatedBy: ", ")
            let fixes = Self.completenessFixes.filter { fields.contains($0.field) }.map(\.fix)
            return (["格式快照缺少必要資料：\(field)"] + fixes).joined(separator: "\n")
        case .invalidSnapshot(let reason): return "格式快照無效：\(reason)"
        case .unsupportedFormatting(let field): return "無法安全保留範本格式：\(field)"
        case .unsupportedNumbering:
            return "範本含有清單編號：numbering part 定義了 abstractNum／num／numPicBullet，或某個樣式以非零 numId 引用編號。"
                + "格式 profile 第一版不支援編號，整份範本無法匯入（不會只略過編號）。"
                + "請改用不含清單編號的範本，或改用 inherit（沿用文件本身的格式）。"
        case .invalidRelationshipTarget(let target): return "relationship Target「\(target)」解析後以 / 結尾，不是合法的 OPC part 名稱"
        case .duplicateRelationship(let type): return "文件套件內同一 Type（\(type)）出現多筆 Target 不同的 Relationship，屬於不合法輸入"
        }
    }

    /// PsychQuant/macdoc#212. `missingRequiredFormatting` payloads naming a
    /// docDefaults field the first profile version requires, with how to
    /// make it explicit. Word's implicit defaults are never filled in on the
    /// user's behalf: guessing a size or spacing Word would render
    /// differently is worse than refusing. A payload lists several fields
    /// separated by ", ".
    static let defaultSizeField = "docDefaults/rPrDefault/rPr/sz"
    static let paragraphDefaultsField = "docDefaults/pPrDefault"
    static let completenessFixes: [(field: String, fix: String)] = [
        (defaultSizeField,
         "範本沒有明確寫出預設字級（styles.xml 的 docDefaults/rPrDefault/rPr/sz）。Word 開啟時會自行套用隱含的預設值，"
            + "但格式 profile 不代為猜測。修正方式：在 Word 開啟這份範本，於「字型」對話框選好字級後按「設為預設值」"
            + "（Mac 版為「預設值…」），選擇只套用到這份文件（即範本本身），存檔後重新匯入；"
            + "或直接在 styles.xml 的 <w:rPrDefault><w:rPr> 內加入 <w:sz w:val=\"21\"/>（單位是半點，21 即 10.5pt）。"),
        (paragraphDefaultsField,
         "範本沒有段落預設值（styles.xml 的 docDefaults/pPrDefault）。修正方式：在 Word 開啟這份範本，於「段落」對話框按"
            + "「設為預設值」（Mac 版為「預設值…」），選擇只套用到這份文件，存檔後重新匯入；"
            + "或直接在 styles.xml 的 <w:docDefaults> 內加入 <w:pPrDefault><w:pPr/></w:pPrDefault>。")
    ]
}

/// A versioned, formatting-only value. No source filename, package paths,
/// relationships, document text or application configuration is retained.
/// Decode and application both validate, including namespace-expanded XML.
public struct DocumentFormattingProfile: Codable, Equatable, Sendable {
    public enum Kind: String, Codable, Sendable { case inherit, official }
    public let schemaVersion: Int
    public let kind: Kind
    public let stylesXML: String?
    public let sectionXML: String?
    public let themeXML: String?
    public let fontsXML: String?

    public static let inherit = DocumentFormattingProfile(kind: .inherit)

    private init(kind: Kind, stylesXML: String? = nil, sectionXML: String? = nil,
                 themeXML: String? = nil, fontsXML: String? = nil) {
        schemaVersion = 1
        self.kind = kind
        self.stylesXML = stylesXML
        self.sectionXML = sectionXML
        self.themeXML = themeXML
        self.fontsXML = fontsXML
    }

    private enum CodingKeys: String, CodingKey {
        case schemaVersion, kind, stylesXML, sectionXML, themeXML, fontsXML
    }

    public init(from decoder: Decoder) throws {
        let c = try decoder.container(keyedBy: CodingKeys.self)
        schemaVersion = try c.decode(Int.self, forKey: .schemaVersion)
        kind = try c.decode(Kind.self, forKey: .kind)
        stylesXML = try c.decodeIfPresent(String.self, forKey: .stylesXML)
        sectionXML = try c.decodeIfPresent(String.self, forKey: .sectionXML)
        themeXML = try c.decodeIfPresent(String.self, forKey: .themeXML)
        fontsXML = try c.decodeIfPresent(String.self, forKey: .fontsXML)
        try validate()
    }

    /// Upper bound for one formatting part read from a template, enforced on
    /// both the declared and the actually inflated size.
    static let maxFormattingPartBytes = 4 * 1024 * 1024

    /// Reads only the formatting parts the main part's relationships name,
    /// directly from the ZIP in memory; archive-controlled paths are never
    /// extracted to disk. Unsupported numbering fails explicitly.
    ///
    /// PsychQuant/macdoc#213: styles, theme and fontTable are the parts the
    /// implicit relationships of `word/_rels/document.xml.rels` resolve to
    /// (the lexical OPC rules of `normalizedRelationshipTarget`), not fixed
    /// paths. No styles relationship, or a relationship whose part is
    /// absent, is refused; theme and fontTable without a relationship are
    /// absent, even if an orphan part sits at the default path. The main
    /// part itself is still `word/document.xml`.
    ///
    /// PsychQuant/macdoc#214: every part read here must be UTF-8 and must not
    /// declare another encoding (`ProfileXML.requireUTF8`).
    public static func importOfficial(from templateURL: URL) throws -> Self {
        let archive = try Archive(url: templateURL, accessMode: .read)
        func read(_ path: String) throws -> XmlNode? {
            let matches = archive.filter { $0.path == path && $0.type == .file }
            guard matches.count <= 1 else { throw DocumentFormattingProfileError.invalidSnapshot("duplicate part") }
            guard let entry = matches.first else { return nil }
            // The declared size is only a cheap precheck: a deflated entry
            // inflates for as long as its compressed stream lasts, whatever
            // the metadata says, so the cap is enforced on the actual bytes
            // (PsychQuant/macdoc#194).
            guard entry.uncompressedSize <= UInt64(maxFormattingPartBytes) else {
                throw DocumentFormattingProfileError.invalidSnapshot("formatting part too large")
            }
            var bytes = Data()
            _ = try archive.extract(entry) { chunk in
                guard bytes.count + chunk.count <= maxFormattingPartBytes else {
                    throw DocumentFormattingProfileError.invalidSnapshot("formatting part too large")
                }
                bytes.append(chunk)
            }
            try ProfileXML.requireUTF8(bytes, part: path)
            return try ProfileXML.parseRejectingDTD(bytes)
        }
        // PsychQuant/macdoc#196: a template whose relationships name two
        // different parts for one implicit Type is refused rather than
        // half-honoured.
        let relationships = try read(ProfileXML.mainRelationshipsPart)
        if let relationships { try ProfileXML.validateImplicitRelationships(in: relationships) }
        func related(_ type: String, required: Bool) throws -> XmlNode? {
            guard let part = try ProfileXML.implicitPart(of: type, in: relationships) else {
                if required {
                    throw DocumentFormattingProfileError.missingRequiredFormatting("\(ProfileXML.mainRelationshipsPart) 沒有 \(type) relationship")
                }
                return nil
            }
            guard let node = try read(part.name) else {
                throw DocumentFormattingProfileError.missingRequiredFormatting(
                    "\(type) relationship 的 Target「\(ProfileXML.excerpt(part.target, limit: 200))」指向不存在的 part \(ProfileXML.excerpt(part.name, limit: 200))")
            }
            return node
        }
        // Numbering is only a refusal check; nothing of it is stored. The
        // related part and the default path are both checked, so the wider
        // check can refuse more templates but never changes a snapshot.
        var numberingParts = [try related("numbering", required: false)]
        if try ProfileXML.implicitPart(of: "numbering", in: relationships)?.name != ProfileXML.defaultNumberingPart {
            numberingParts.append(try read(ProfileXML.defaultNumberingPart))
        }
        for numbering in numberingParts.compactMap({ $0 }) {
            guard numbering.namespaceURI == ProfileXML.w, numbering.localName == "numbering" else {
                throw DocumentFormattingProfileError.invalidSnapshot("numbering root")
            }
            if ProfileXML.walk(numbering).contains(where: { $0.namespaceURI == ProfileXML.w && ["abstractNum", "num", "numPicBullet"].contains($0.localName) }) {
                throw DocumentFormattingProfileError.unsupportedNumbering
            }
        }
        let rawStyles = try related("styles", required: true)!
        let styles = try ProfileXML.clean(rawStyles, root: "styles")
        guard let document = try read("word/document.xml") else {
            throw DocumentFormattingProfileError.missingRequiredFormatting("word/document.xml")
        }
        guard document.namespaceURI == ProfileXML.w, document.localName == "document",
              let body = ProfileXML.child(document, "body"),
              let section = body.children.last(where: { $0.kind == .element }),
              section.namespaceURI == ProfileXML.w, section.localName == "sectPr" else {
            throw DocumentFormattingProfileError.missingRequiredFormatting("final body sectPr")
        }
        let safeSection = try ProfileXML.clean(section, root: "sectPr", inherited: ProfileXML.namespaceScope(document))
        let theme = try related("theme", required: false).map { try ProfileXML.clean($0, root: "theme") }
        let fonts = try related("fontTable", required: false).map { try ProfileXML.clean($0, root: "fonts") }
        let value = Self(kind: .official, stylesXML: try ProfileXML.string(styles),
                         sectionXML: try ProfileXML.string(safeSection),
                         themeXML: try theme.map(ProfileXML.string), fontsXML: try fonts.map(ProfileXML.string))
        try value.validate()
        return value
    }

    public func validate() throws {
        guard schemaVersion == 1 else { throw DocumentFormattingProfileError.unsupportedVersion(schemaVersion) }
        if kind == .inherit {
            guard stylesXML == nil, sectionXML == nil, themeXML == nil, fontsXML == nil else {
                throw DocumentFormattingProfileError.invalidSnapshot("inherit has payload")
            }
            return
        }
        guard let stylesXML, let sectionXML else {
            throw DocumentFormattingProfileError.missingRequiredFormatting("styles/section")
        }
        let styles = try ProfileXML.checked(stylesXML, root: "styles")
        try ProfileXML.validateStyles(styles)
        // PsychQuant/macdoc#212: name exactly the missing docDefaults fields
        // (all of them, so one fix round is enough); a present but unusable
        // size is invalid data, not missing data. An empty pPrDefault is
        // accepted.
        let defaults = ProfileXML.child(styles, "docDefaults")
        var missing: [String] = []
        if let size = defaults.flatMap({ ProfileXML.child($0, "rPrDefault") }).flatMap({ ProfileXML.child($0, "rPr") })
            .flatMap({ ProfileXML.child($0, "sz") }) {
            let raw = ProfileXML.value(size, "val") ?? ""
            guard let halfPoints = Int(raw), halfPoints > 0 else {
                throw DocumentFormattingProfileError.invalidSnapshot("\(DocumentFormattingProfileError.defaultSizeField) 必須是正整數（半點），實際為「\(ProfileXML.excerpt(raw))」")
            }
        } else {
            missing.append(DocumentFormattingProfileError.defaultSizeField)
        }
        if defaults.flatMap({ ProfileXML.child($0, "pPrDefault") }) == nil {
            missing.append(DocumentFormattingProfileError.paragraphDefaultsField)
        }
        guard missing.isEmpty else {
            throw DocumentFormattingProfileError.missingRequiredFormatting(missing.joined(separator: ", "))
        }
        let section = try ProfileXML.checked(sectionXML, root: "sectPr")
        for (name, attrs) in [("pgSz", ["w", "h"]), ("pgMar", ["top", "right", "bottom", "left", "header", "footer", "gutter"])] {
            guard let node = ProfileXML.child(section, name) else { throw DocumentFormattingProfileError.missingRequiredFormatting(name) }
            for attr in attrs {
                guard let raw = ProfileXML.value(node, attr), let number = Int(raw),
                      name != "pgSz" || number > 0,
                      ["top", "bottom"].contains(attr) || number >= 0 else {
                    throw DocumentFormattingProfileError.invalidSnapshot("\(name)/\(attr)")
                }
            }
        }
        if let themeXML {
            let theme = try ProfileXML.checked(themeXML, root: "theme")
            guard let elements = ProfileXML.child(theme, "themeElements", ns: ProfileXML.a),
                  let scheme = ProfileXML.child(elements, "fontScheme", ns: ProfileXML.a),
                  ProfileXML.child(scheme, "majorFont", ns: ProfileXML.a) != nil,
                  ProfileXML.child(scheme, "minorFont", ns: ProfileXML.a) != nil else {
                throw DocumentFormattingProfileError.missingRequiredFormatting("theme fontScheme")
            }
        } else if ProfileXML.walk(styles).contains(where: { $0.attributes.contains(where: { $0.localName.hasSuffix("Theme") || $0.localName.hasPrefix("theme") }) }) {
            throw DocumentFormattingProfileError.missingRequiredFormatting("theme used by styles")
        }
        if let fontsXML { _ = try ProfileXML.checked(fontsXML, root: "fonts") }
    }
}

/// All snapshot XML is rebuilt from explicit element/attribute vocabularies.
/// Names are resolved using namespace URI, including aliased attributes.
internal enum ProfileXML {
    static let w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    static let a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    static let w14 = "http://schemas.microsoft.com/office/word/2010/wordml"
    static let wordAttributes: [String: Set<String>] = [
        "styles": [], "docDefaults": [], "rPrDefault": [], "pPrDefault": [], "rPr": [], "pPr": [],
        "style": ["type", "styleId", "default", "customStyle"],
        "name": ["val"], "basedOn": ["val"], "next": ["val"], "link": ["val"],
        "qFormat": ["val"], "hidden": ["val"], "semiHidden": ["val"], "unhideWhenUsed": ["val"], "uiPriority": ["val"],
        "autoRedefine": ["val"], "locked": ["val"], "personal": ["val"], "personalCompose": ["val"], "personalReply": ["val"],
        "rFonts": ["ascii", "hAnsi", "eastAsia", "cs", "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme", "hint"],
        "sz": ["val"], "szCs": ["val"], "b": ["val"], "bCs": ["val"], "i": ["val"], "iCs": ["val"],
        "caps": ["val"], "smallCaps": ["val"], "strike": ["val"], "dstrike": ["val"], "outline": ["val"], "shadow": ["val"],
        "emboss": ["val"], "imprint": ["val"], "noProof": ["val"], "snapToGrid": ["val"], "vanish": ["val"], "webHidden": ["val"],
        "color": ["val", "themeColor", "themeTint", "themeShade"], "highlight": ["val"],
        "u": ["val", "color", "themeColor", "themeTint", "themeShade"], "vertAlign": ["val"],
        "spacing": ["val", "before", "after", "beforeLines", "afterLines", "beforeAutospacing", "afterAutospacing", "line", "lineRule"],
        "w": ["val"], "kern": ["val"], "position": ["val"], "lang": ["val", "eastAsia", "bidi"],
        "rtl": ["val"], "cs": ["val"], "eastAsianLayout": ["id", "combine", "combineBrackets", "vert", "vertCompress"],
        "keepNext": ["val"], "keepLines": ["val"], "pageBreakBefore": ["val"], "widowControl": ["val"],
        "suppressLineNumbers": ["val"], "suppressAutoHyphens": ["val"], "contextualSpacing": ["val"], "bidi": ["val"],
        "adjustRightInd": ["val"], "autoSpaceDE": ["val"], "autoSpaceDN": ["val"], "kinsoku": ["val"],
        "wordWrap": ["val"], "overflowPunct": ["val"], "topLinePunct": ["val"], "mirrorIndents": ["val"],
        "jc": ["val"], "outlineLvl": ["val"], "textAlignment": ["val"], "textboxTightWrap": ["val"],
        "ind": ["left", "right", "start", "end", "firstLine", "hanging", "leftChars", "rightChars", "firstLineChars", "hangingChars"],
        "tabs": [], "tab": ["val", "leader", "pos"], "numPr": [], "numId": ["val"], "ilvl": ["val"],
        "pBdr": [], "bdr": ["val", "sz", "space", "color", "themeColor", "themeTint", "themeShade", "shadow", "frame"],
        "shd": ["val", "color", "fill", "themeColor", "themeFill", "themeTint", "themeShade", "themeFillTint", "themeFillShade"],
        "sectPr": [], "type": ["val"], "pgSz": ["w", "h", "code", "orient"], "pgMar": ["top", "right", "bottom", "left", "header", "footer", "gutter"],
        "cols": ["num", "space", "equalWidth", "sep"], "docGrid": ["type", "linePitch", "charSpace"],
        "fonts": [], "font": ["name"], "altName": ["val"], "charset": ["val"], "family": ["val"], "pitch": ["val"],
        "panose1": ["val"], "sig": ["usb0", "usb1", "usb2", "usb3", "csb0", "csb1"],
        "tblPr": [], "tblInd": ["w", "type"], "tblCellMar": []
    ]
    static let drawingElements: Set<String> = Set("theme themeElements clrScheme dk1 lt1 dk2 lt2 accent1 accent2 accent3 accent4 accent5 accent6 hlink folHlink sysClr srgbClr scrgbClr hslClr schemeClr prstClr tint shade comp inv gray alpha alphaOff alphaMod hue hueOff hueMod sat satOff satMod lum lumOff lumMod red redOff redMod green greenOff greenMod blue blueOff blueMod gamma invGamma fontScheme majorFont minorFont latin ea cs font fmtScheme fillStyleLst solidFill noFill gradFill gsLst gs lin path fillToRect tileRect lnStyleLst ln prstDash custDash ds round bevel miter headEnd tailEnd effectStyleLst effectStyle effectLst outerShdw innerShdw glow softEdge reflection blur backgroundFillStyleLst bgFillStyleLst scene3d camera lightRig rot bevelT bevelB sp3d".split(separator: " ").map(String.init))
    static let drawingAttributes: Set<String> = Set("name val lastClr r g b hue sat lum typeface panose pitchFamily charset script pos ang scaled path l t r b w cap cmpd algn lim type len dist dir sx sy kx ky algn rotWithShape blurRad rad stA stPos endA endPos fadeDir prst fov zoom lat lon rev rig dir extrusionH contourW z prstMaterial h".split(separator: " ").map(String.init))
    static let discarded: Set<String> = ["latentStyles", "rsid", "aliases", "headerReference", "footerReference", "printerSettings", "sectPrChange", "pPrChange", "rPrChange", "ins", "del", "moveFrom", "moveTo", "embedRegular", "embedBold", "embedItalic", "embedBoldItalic", "extLst", "objectDefaults", "extraClrSchemeLst", "custClrLst"]
    static let wordChildren: [String: Set<String>] = [
        "styles": ["docDefaults", "style"], "docDefaults": ["rPrDefault", "pPrDefault"],
        "rPrDefault": ["rPr"], "pPrDefault": ["pPr"],
        "style": ["name", "basedOn", "next", "link", "qFormat", "hidden", "semiHidden", "unhideWhenUsed", "uiPriority", "autoRedefine", "locked", "personal", "personalCompose", "personalReply", "pPr", "rPr", "tblPr"],
        "rPr": Set("rFonts sz szCs b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color highlight u vertAlign spacing w kern position lang rtl cs eastAsianLayout bdr shd".split(separator: " ").map(String.init)),
        "pPr": Set("keepNext keepLines pageBreakBefore widowControl suppressLineNumbers suppressAutoHyphens contextualSpacing bidi adjustRightInd autoSpaceDE autoSpaceDN kinsoku wordWrap overflowPunct topLinePunct mirrorIndents jc outlineLvl textAlignment textboxTightWrap ind tabs numPr pBdr shd spacing snapToGrid".split(separator: " ").map(String.init)),
        "tabs": ["tab"], "numPr": ["numId", "ilvl"], "pBdr": ["top", "bottom", "left", "right", "between", "bar"],
        "sectPr": ["type", "pgSz", "pgMar", "cols", "docGrid"], "fonts": ["font"],
        "font": ["altName", "charset", "family", "pitch", "panose1", "sig"],
        "tblPr": ["tblInd", "tblCellMar"], "tblCellMar": ["top", "bottom", "left", "right"]
    ]

    static func parse(_ data: Data) throws -> XmlNode {
        do { return try XmlTreeReader.parse(data).root }
        catch { throw DocumentFormattingProfileError.invalidSnapshot("malformed XML") }
    }
    static func parse(_ string: String) throws -> XmlNode { try parse(Data(string.utf8)) }
    static func string(_ node: XmlNode) throws -> String {
        let copy = node.deepClone()
        for item in walk(copy) { item.sourceRange = nil }
        return String(decoding: try XmlTreeWriter.serialize(.synthesized(root: copy)), as: UTF8.self)
    }
    static func walk(_ node: XmlNode) -> [XmlNode] { [node] + node.children.flatMap(walk) }
    static func child(_ node: XmlNode, _ name: String, ns: String = w) -> XmlNode? {
        node.children.first { $0.kind == .element && $0.namespaceURI == ns && $0.localName == name }
    }
    static func value(_ node: XmlNode, _ attr: String) -> String? { node.attributeValue(prefix: "w", localName: attr) }
    static func namespaceScope(_ node: XmlNode) -> [String: String] {
        var result: [String: String] = [:]
        for attr in node.attributes { if let prefix = attr.declaredNamespacePrefix { result[prefix] = attr.value } }
        return result
    }
    /// Normalize Word names on a target tree without filtering target-owned
    /// extensions. Resolve attribute namespaces before changing any binding.
    static func canonicalWordTree(_ source: XmlNode) -> XmlNode {
        let root = source.deepClone()
        var occupied = Set(walk(source).flatMap { $0.attributes.compactMap(\.declaredNamespacePrefix) })
        var foreignAliases: [String: String] = [:]
        func normalizedPrefix(_ prefix: String, uri: String) -> String {
            if uri == w { return "w" }
            guard prefix == "w" else { return prefix }
            if let alias = foreignAliases[uri] { return alias }
            var index = 1
            while occupied.contains("profileTarget\(index)") { index += 1 }
            let alias = "profileTarget\(index)"
            occupied.insert(alias)
            foreignAliases[uri] = alias
            return alias
        }
        // `declared` tracks, for the OUTPUT tree being built, which prefix→URI
        // bindings an ancestor has already written as an `xmlns:` attribute.
        // It is distinct from `scope`/`originalScope` (which track the INPUT
        // tree's bindings, used for `normalizedPrefix`/attribute-prefix
        // resolution): a node only needs to declare a binding once for its
        // whole subtree to inherit it via ordinary XML namespace scoping.
        // Without this, every element that carries the `w` (or an aliased
        // target) namespace re-declares it, bloating `word/styles.xml`
        // linearly with the element count (#195).
        func visit(_ node: XmlNode, scope: [String: String], declared: [String: String]) {
            guard node.kind == .element else { return }
            var originalScope = scope
            originalScope.merge(namespaceScope(node)) { _, new in new }
            var bindings: [String: String] = [:]
            if let uri = node.namespaceURI {
                if uri == w { node.prefix = "w"; bindings["w"] = w }
                else if let prefix = node.prefix {
                    let mapped = normalizedPrefix(prefix, uri: uri)
                    node.prefix = mapped
                    bindings[mapped] = uri
                }
            }
            node.attributes = node.attributes.map { attr in
                guard !attr.isNamespaceDeclaration, let prefix = attr.prefix,
                      let uri = originalScope[prefix] else { return attr }
                var copy = attr
                let mapped = normalizedPrefix(prefix, uri: uri)
                copy.prefix = mapped
                bindings[mapped] = uri
                return copy
            }
            var childDeclared = declared
            for (prefix, uri) in bindings.sorted(by: { $0.key < $1.key }) {
                node.attributes.removeAll { $0.declaredNamespacePrefix == prefix }
                guard declared[prefix] != uri else { continue }
                node.attributes.append(XmlAttribute(prefix: "xmlns", localName: prefix, value: uri))
                childDeclared[prefix] = uri
            }
            for child in node.children { visit(child, scope: originalScope, declared: childDeclared) }
        }
        visit(root, scope: ["xml": "http://www.w3.org/XML/1998/namespace"], declared: [:])
        return root
    }
    static func clean(_ node: XmlNode, root: String, inherited: [String: String] = [:], strict: Bool = false) throws -> XmlNode {
        let ns = root == "theme" ? a : w
        guard node.namespaceURI == ns, node.localName == root else { throw DocumentFormattingProfileError.invalidSnapshot("wrong \(root) root") }
        func rebuild(_ input: XmlNode, scope: [String: String], parent: String? = nil) throws -> XmlNode? {
            guard input.kind == .element else {
                if strict && (input.kind != .text || !input.textContent.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty) {
                    throw DocumentFormattingProfileError.invalidSnapshot("non-formatting content")
                }
                return nil
            }
            var scope = scope
            scope.merge(namespaceScope(input)) { _, new in new }
            let name = input.localName
            if input.namespaceURI == w14, name == "ligatures", parent == "rPr" {
                let attrs = input.attributes.filter { !$0.isNamespaceDeclaration }
                guard attrs.count == 1, let attr = attrs.first, attr.localName == "val", attr.prefix.flatMap({ scope[$0] }) == w14,
                      ["none", "standard", "contextual", "historical", "discretional", "standardContextual", "standardHistorical", "contextualHistorical", "standardDiscretional", "contextualDiscretional", "historicalDiscretional", "standardContextualHistorical", "standardContextualDiscretional", "standardHistoricalDiscretional", "contextualHistoricalDiscretional", "all"].contains(attr.value),
                      input.children.allSatisfy({ $0.kind == .text && $0.textContent.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty }) else {
                    throw DocumentFormattingProfileError.unsupportedFormatting("ligatures")
                }
                return .element(prefix: "w14", localName: name, namespaceURI: w14, attributes: [
                    XmlAttribute(prefix: "xmlns", localName: "w14", value: w14), XmlAttribute(prefix: "w14", localName: "val", value: attr.value)])
            }
            if discarded.contains(name) {
                if strict { throw DocumentFormattingProfileError.invalidSnapshot("unsafe element \(name)") }
                return nil
            }
            guard input.namespaceURI == ns else {
                throw DocumentFormattingProfileError.unsupportedFormatting("foreign namespace \(name)")
            }
            if ns == w, let parent, !(wordChildren[parent] ?? []).contains(name) {
                throw DocumentFormattingProfileError.unsupportedFormatting("\(parent)/\(name)")
            }
            let allowed: Set<String>
            if ns == a {
                guard drawingElements.contains(name) else { throw DocumentFormattingProfileError.unsupportedFormatting(name) }
                allowed = drawingAttributes
            } else if ["top", "bottom", "left", "right", "between", "bar"].contains(name) {
                allowed = parent == "tblCellMar" ? ["w", "type"] : wordAttributes["bdr"]!
            } else {
                guard let names = wordAttributes[name] else { throw DocumentFormattingProfileError.unsupportedFormatting(name) }
                allowed = names
            }
            var attrs: [XmlAttribute] = []
            var seenAttributes: Set<String> = []
            for attr in input.attributes where !attr.isNamespaceDeclaration {
                let uri = attr.prefix.flatMap { scope[$0] }
                guard seenAttributes.insert((uri ?? "") + ":" + attr.localName).inserted else {
                    throw DocumentFormattingProfileError.invalidSnapshot("duplicate attribute")
                }
                let correctNS = ns == a ? attr.prefix == nil : uri == w
                guard correctNS, allowed.contains(attr.localName) else {
                    if strict { throw DocumentFormattingProfileError.invalidSnapshot("unsafe attribute \(attr.localName)") }
                    if attr.localName.hasPrefix("rsid") || uri == "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                        || (uri == "http://schemas.openxmlformats.org/markup-compatibility/2006" && attr.localName == "Ignorable") { continue }
                    throw DocumentFormattingProfileError.unsupportedFormatting("\(name)/\(attr.localName)")
                }
                attrs.append(XmlAttribute(prefix: ns == w ? "w" : nil, localName: attr.localName, value: attr.value))
            }
            if name == "numId", attrs.first(where: { $0.localName == "val" })?.value != "0" {
                throw DocumentFormattingProfileError.unsupportedNumbering
            }
            if name == "cols", (attrs.first(where: { $0.localName == "equalWidth" })?.value == "0") {
                throw DocumentFormattingProfileError.unsupportedFormatting("unequal columns")
            }
            let children = try input.children.compactMap { try rebuild($0, scope: scope, parent: name) }
            if ns == w {
                var seen: Set<String> = []
                for child in children {
                    let repeatable = (name == "styles" && child.localName == "style")
                        || (name == "fonts" && child.localName == "font")
                        || (name == "tabs" && child.localName == "tab")
                    let key = (child.namespaceURI ?? "") + ":" + child.localName
                    if !repeatable, !seen.insert(key).inserted {
                        throw DocumentFormattingProfileError.invalidSnapshot("duplicate singleton \(name)/\(child.localName)")
                    }
                }
            }
            return .element(prefix: ns == w ? "w" : "a", localName: name, namespaceURI: ns,
                            attributes: attrs, children: children)
        }
        let result = try rebuild(node, scope: inherited)!
        result.attributes.insert(XmlAttribute(prefix: "xmlns", localName: ns == w ? "w" : "a", value: ns), at: 0)
        return result
    }
    /// Snapshot payloads (decode and apply) pass the same encoding gate as
    /// imported parts; import itself always writes `encoding="UTF-8"`.
    static func checked(_ xml: String, root: String) throws -> XmlNode {
        let data = Data(xml.utf8)
        try requireUTF8(data, part: "格式快照 \(snapshotKeys[root] ?? root)")
        return try clean(parseRejectingDTD(data), root: root, strict: true)
    }
    static let snapshotKeys = ["styles": "stylesXML", "sectPr": "sectionXML", "theme": "themeXML", "fonts": "fontsXML"]

    /// At most `limit` characters of template-controlled text for a message.
    static func excerpt(_ text: String, limit: Int = 40) -> String {
        text.count > limit ? String(text.prefix(limit)) + "…" : text
    }

    // MARK: - PsychQuant/macdoc#214 encoding gate

    /// The profile path persists what it parses. `XmlTreeReader` decodes
    /// every byte as UTF-8 and skips the XML declaration, while Word decodes
    /// by the declared `encoding`, so a part declaring ISO-8859-1 or
    /// Shift_JIS whose bytes happen to be valid UTF-8 would be stored as
    /// text Word never shows. Profile import and snapshot payloads are
    /// therefore accepted only when:
    /// - the bytes are not UTF-16/UTF-32 (a BOM, or a NUL in the first two
    ///   bytes) — UTF-16 stays refused as in 3.10.0 (no charset expansion),
    ///   now with a message naming it;
    /// - the bytes are valid UTF-8 without NUL (an optional UTF-8 BOM), so
    ///   nothing is stored with U+FFFD replacement characters;
    /// - the XML declaration, located the way `XmlTreeReader` locates it, is
    ///   absent, has no `encoding`, or declares `UTF-8` (ASCII
    ///   case-insensitive). A declaration this gate cannot read is refused.
    /// `DocxReader` keeps its own, wider decoding for ordinary documents.
    static func requireUTF8(_ data: Data, part: String) throws {
        let bytes = [UInt8](data)
        if bytes.starts(with: [0xFE, 0xFF]) || bytes.starts(with: [0xFF, 0xFE])
            || (bytes.count >= 2 && (bytes[0] == 0 || bytes[1] == 0)) {
            throw DocumentFormattingProfileError.invalidSnapshot("\(part) 是 UTF-16／UTF-32 編碼，格式 profile 只接受 UTF-8")
        }
        guard !bytes.contains(0), String(bytes: bytes, encoding: .utf8) != nil else {
            throw DocumentFormattingProfileError.invalidSnapshot("\(part) 不是合法的 UTF-8（含無法解碼的位元組或 NUL），格式 profile 只接受 UTF-8")
        }
        if let declared = try declaredEncoding(bytes, part: part), declared.lowercased() != "utf-8" {
            throw DocumentFormattingProfileError.invalidSnapshot("\(part) 的 XML 宣告編碼為「\(excerpt(declared))」，格式 profile 只接受 UTF-8")
        }
    }

    /// The `encoding` of the XML declaration, checked against the XML 1.0
    /// `XMLDecl` grammar: `<?xml`, then `version` (`1.` digits), optional
    /// `encoding` (`EncName`) and optional `standalone` (`yes`/`no`), in that
    /// order, each preceded by whitespace, quoted with `"` or `'`, up to `?>`.
    /// Anything else — a missing version, unknown or repeated names, wrong
    /// order, `Encoding` in another case, no whitespace between
    /// pseudo-attributes — is refused, since the tree reader skips the
    /// declaration without validating it. Whitespace before the declaration
    /// is not XML 1.0 but `XmlTreeReader.skipProlog` tolerates it, so it is
    /// tolerated here too and the declaration after it is still checked. The
    /// rest of the prolog (comments and processing instructions the tree
    /// reader also skips) may not contain another declaration, including one
    /// spelled `<?XML`: such a part is not well-formed and Word refuses it.
    static func declaredEncoding(_ bytes: [UInt8], part: String) throws -> String? {
        func isSpace(_ byte: UInt8) -> Bool { byte == 0x20 || byte == 0x09 || byte == 0x0A || byte == 0x0D }
        func isLetter(_ byte: UInt8) -> Bool {
            (UInt8(ascii: "a")...UInt8(ascii: "z")).contains(byte) || (UInt8(ascii: "A")...UInt8(ascii: "Z")).contains(byte)
        }
        func isDigit(_ byte: UInt8) -> Bool { (UInt8(ascii: "0")...UInt8(ascii: "9")).contains(byte) }
        func starts(_ literal: String, at position: Int) -> Bool {
            let needle = Array(literal.utf8)
            return position + needle.count <= bytes.count && Array(bytes[position..<position + needle.count]) == needle
        }
        func find(_ literal: String, from position: Int) -> Int? {
            var cursor = position
            while cursor < bytes.count { if starts(literal, at: cursor) { return cursor }; cursor += 1 }
            return nil
        }
        func validValue(_ name: String, _ value: [UInt8]) -> Bool {
            switch name {
            case "version": return value.count > 2 && value.starts(with: Array("1.".utf8)) && value.dropFirst(2).allSatisfy(isDigit)
            case "encoding":
                return value.first.map(isLetter) == true && value.allSatisfy {
                    isLetter($0) || isDigit($0) || $0 == UInt8(ascii: ".") || $0 == UInt8(ascii: "_") || $0 == UInt8(ascii: "-")
                }
            default: return value == Array("yes".utf8) || value == Array("no".utf8)
            }
        }
        let unreadable = DocumentFormattingProfileError.invalidSnapshot("\(part) 的 XML 宣告無法解析")
        var index = bytes.starts(with: [0xEF, 0xBB, 0xBF]) ? 3 : 0
        while index < bytes.count, isSpace(bytes[index]) { index += 1 }
        var encoding: String?
        if starts("<?xml", at: index), index + 5 < bytes.count, isSpace(bytes[index + 5]) || bytes[index + 5] == UInt8(ascii: "?") {
            index += 5
            let names = ["version", "encoding", "standalone"]
            var nextAllowed = 0
            var values: [String: String] = [:]
            while true {
                let spaceStart = index
                while index < bytes.count, isSpace(bytes[index]) { index += 1 }
                if starts("?>", at: index) { index += 2; break }
                guard index > spaceStart, index < bytes.count else { throw unreadable }
                let nameStart = index
                while index < bytes.count, isLetter(bytes[index]) { index += 1 }
                let name = String(decoding: bytes[nameStart..<index], as: UTF8.self)
                guard let position = names.firstIndex(of: name), position >= nextAllowed else { throw unreadable }
                nextAllowed = position + 1
                while index < bytes.count, isSpace(bytes[index]) { index += 1 }
                guard index < bytes.count, bytes[index] == UInt8(ascii: "=") else { throw unreadable }
                index += 1
                while index < bytes.count, isSpace(bytes[index]) { index += 1 }
                guard index < bytes.count, bytes[index] == UInt8(ascii: "\"") || bytes[index] == UInt8(ascii: "'") else { throw unreadable }
                let quote = bytes[index]
                index += 1
                let valueStart = index
                while index < bytes.count, bytes[index] != quote, bytes[index] != UInt8(ascii: "<"), bytes[index] != UInt8(ascii: ">") { index += 1 }
                guard index < bytes.count, bytes[index] == quote, validValue(name, Array(bytes[valueStart..<index])) else { throw unreadable }
                values[name] = String(decoding: bytes[valueStart..<index], as: UTF8.self)
                index += 1
            }
            guard values["version"] != nil else { throw unreadable }
            encoding = values["encoding"]
        }
        while index < bytes.count {
            while index < bytes.count, isSpace(bytes[index]) { index += 1 }
            if starts("<!--", at: index) {
                guard let end = find("-->", from: index + 4) else { break }
                index = end + 3
            } else if starts("<?", at: index) {
                var targetEnd = index + 2
                while targetEnd < bytes.count, !isSpace(bytes[targetEnd]), bytes[targetEnd] != UInt8(ascii: "?") { targetEnd += 1 }
                if String(decoding: bytes[(index + 2)..<targetEnd], as: UTF8.self).lowercased() == "xml" {
                    throw DocumentFormattingProfileError.invalidSnapshot("\(part) 含有位置或大小寫不合法的 XML 宣告")
                }
                guard let end = find("?>", from: targetEnd) else { break }
                index = end + 2
            } else {
                break
            }
        }
        return encoding
    }

    /// Profile payloads and template formatting parts refuse any DTD, like
    /// DocxReader does for the parts it reads; the snapshot parser would
    /// otherwise skip it silently (PsychQuant/macdoc#196).
    static func parseRejectingDTD(_ data: Data) throws -> XmlNode {
        do { try DocxReader.rejectDTD(data, part: "formatting snapshot") }
        catch { throw DocumentFormattingProfileError.invalidSnapshot("DTD not allowed") }
        return try parse(data)
    }

    static let relationshipsNS = "http://schemas.openxmlformats.org/package/2006/relationships"
    static let officeRelationshipsNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    static let implicitRelationshipTypes = ["styles", "theme", "fontTable"]
    static let mainRelationshipsPart = "word/_rels/document.xml.rels"
    static let defaultNumberingPart = "word/numbering.xml"

    /// PsychQuant/macdoc#213. The ZIP item name (no leading `/`) of the part
    /// the main part's implicit relationship of `type` names, with the
    /// Target as written; nil when there is no such relationship (or no
    /// relationships part). Resolution is the lexical OPC normalization
    /// used by the duplicate check. Fails closed on an External target, a
    /// Target that resolves to a path ending in `/`, and registrations of
    /// one Type that name different parts.
    static func implicitPart(of type: String, in rels: XmlNode?) throws -> (target: String, name: String)? {
        guard let rels else { return nil }
        let registrations = rels.children.filter {
            $0.kind == .element && $0.namespaceURI == relationshipsNS && $0.localName == "Relationship"
                && $0.attributeValue(prefix: nil, localName: "Type") == officeRelationshipsNS + type
        }
        guard let first = registrations.first else { return nil }
        if registrations.contains(where: { $0.attributeValue(prefix: nil, localName: "TargetMode") == "External" }) {
            throw DocumentFormattingProfileError.invalidSnapshot("\(type) relationship 的 TargetMode 為 External，格式 part 必須在套件內")
        }
        let targets = registrations.map { $0.attributeValue(prefix: nil, localName: "Target") ?? "" }
        for target in targets where normalizedRelationshipTarget(target).hasSuffix("/") {
            throw DocumentFormattingProfileError.invalidRelationshipTarget(target)
        }
        let names = Set(targets.map(normalizedRelationshipTarget))
        guard names.count == 1, let name = names.first else { throw DocumentFormattingProfileError.duplicateRelationship(type) }
        return (first.attributeValue(prefix: nil, localName: "Target") ?? "", String(name.dropFirst()))
    }

    /// PsychQuant/macdoc#196 policy point 1. Resolves a main-part
    /// relationship Target against `/word/document.xml` (RFC 3986 §5.2) to
    /// the part name used for equivalence. Purely lexical — the host
    /// filesystem is never consulted, unlike `NSString.standardizingPath`,
    /// which resolves symlinks and strips `/private` and so merges distinct
    /// OPC part names on macOS.
    /// - Only escapes of RFC 3986 unreserved characters (ALPHA, DIGIT, `-`,
    ///   `.`, `_`, `~`) are decoded (§6.2.2.2): `styl%65s.xml` is
    ///   `styles.xml`. Every other escape — `%2F`, `%5C`, `%25`, … — stays
    ///   encoded with uppercase hex (§6.2.2.1), so an encoded `/` is data
    ///   inside a segment and `%252F` never collapses onto `%2F`.
    /// - `.` and `..` segments are removed (§5.2.4); a final `.` or `..`
    ///   leaves a trailing slash (`styles.xml/.` is `/word/styles.xml/`,
    ///   not the part `styles.xml`); empty segments are kept.
    /// - Case is not folded. This is the decided, locked behavior: a
    ///   case-different Target names a different part here.
    static func normalizedRelationshipTarget(_ target: String) -> String {
        let path = target.hasPrefix("/") ? target : "/word/" + target
        var segments: [String] = []
        var endsInDirectory = false
        for raw in path.split(separator: "/", omittingEmptySubsequences: false).dropFirst() {
            let segment = normalizedPercentEncoding(raw)
            endsInDirectory = segment == "." || segment == ".."
            switch segment {
            case ".": continue
            case "..": if !segments.isEmpty { segments.removeLast() }
            default: segments.append(segment)
            }
        }
        if endsInDirectory, !segments.isEmpty { segments.append("") }
        return "/" + segments.joined(separator: "/")
    }

    /// Decodes `%XX` only when it encodes an unreserved ASCII character and
    /// rewrites every other well-formed escape with uppercase hex; bytes
    /// that are not part of a well-formed escape are kept as written.
    private static func normalizedPercentEncoding(_ segment: Substring) -> String {
        func hex(_ byte: UInt8) -> UInt8? {
            switch byte {
            case UInt8(ascii: "0")...UInt8(ascii: "9"): return byte - UInt8(ascii: "0")
            case UInt8(ascii: "A")...UInt8(ascii: "F"): return byte - UInt8(ascii: "A") + 10
            case UInt8(ascii: "a")...UInt8(ascii: "f"): return byte - UInt8(ascii: "a") + 10
            default: return nil
            }
        }
        func isUnreserved(_ byte: UInt8) -> Bool {
            switch byte {
            case UInt8(ascii: "A")...UInt8(ascii: "Z"), UInt8(ascii: "a")...UInt8(ascii: "z"),
                 UInt8(ascii: "0")...UInt8(ascii: "9"), UInt8(ascii: "-"), UInt8(ascii: "."),
                 UInt8(ascii: "_"), UInt8(ascii: "~"): return true
            default: return false
            }
        }
        let bytes = Array(segment.utf8)
        var output: [UInt8] = []
        output.reserveCapacity(bytes.count)
        var index = 0
        while index < bytes.count {
            if bytes[index] == UInt8(ascii: "%"), index + 2 < bytes.count,
               let high = hex(bytes[index + 1]), let low = hex(bytes[index + 2]) {
                let value = high << 4 | low
                if isUnreserved(value) {
                    output.append(value)
                } else {
                    output.append(contentsOf: Array(String(format: "%%%02X", value).utf8))
                }
                index += 3
            } else {
                output.append(bytes[index])
                index += 1
            }
        }
        return String(decoding: output, as: UTF8.self)
    }

    /// PsychQuant/macdoc#196 policy points 1 and 2 for the main part's
    /// implicit styles/theme/fontTable relationships:
    /// - a Target that resolves to a path ending in `/` is not an OPC part
    ///   name (`invalidRelationshipTarget`);
    /// - ECMA-376 allows at most one implicit relationship of each of these
    ///   Types, so registrations of one Type that resolve to different parts
    ///   make the package malformed (`duplicateRelationship`) — fail closed
    ///   instead of letting a `.first` lookup pick one. Equivalent spellings
    ///   of one part are not a violation.
    static func validateImplicitRelationships(in rels: XmlNode) throws {
        for rel in implicitRelationshipTypes {
            let relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + rel
            let targets = rels.children.filter {
                $0.kind == .element && $0.namespaceURI == relationshipsNS && $0.localName == "Relationship"
                    && $0.attributeValue(prefix: nil, localName: "Type") == relationshipType
            }.map { $0.attributeValue(prefix: nil, localName: "Target") ?? "" }
            for target in targets where normalizedRelationshipTarget(target).hasSuffix("/") {
                throw DocumentFormattingProfileError.invalidRelationshipTarget(target)
            }
            guard Set(targets.map(normalizedRelationshipTarget)).count <= 1 else {
                throw DocumentFormattingProfileError.duplicateRelationship(rel)
            }
        }
    }

    static func validateStyles(_ root: XmlNode) throws {
        let styles = root.children.filter { $0.localName == "style" }
        var byID: [String: XmlNode] = [:]
        for style in styles {
            guard let id = value(style, "styleId"), !id.isEmpty, byID[id] == nil,
                  let type = value(style, "type"), StyleType(rawValue: type) != nil else {
                throw DocumentFormattingProfileError.invalidSnapshot("duplicate/missing style ID or type")
            }
            byID[id] = style
        }
        let defaults = styles.filter { value($0, "type") == "paragraph" && ["1", "true", "on"].contains(value($0, "default") ?? "") }
        guard defaults.count == 1 else { throw DocumentFormattingProfileError.missingRequiredFormatting("one default paragraph style") }
        for style in styles {
            for name in ["basedOn", "next", "link"] {
                for reference in style.children where reference.namespaceURI == w && reference.localName == name {
                    guard let target = value(reference, "val"), byID[target] != nil else { throw DocumentFormattingProfileError.invalidSnapshot("dangling \(name)") }
                }
            }
            var visited: Set<String> = []
            var current: XmlNode? = style
            while let node = current, let id = value(node, "styleId") {
                guard visited.insert(id).inserted else { throw DocumentFormattingProfileError.invalidSnapshot("basedOn cycle") }
                current = child(node, "basedOn").flatMap { value($0, "val") }.flatMap { byID[$0] }
            }
        }
    }
}
