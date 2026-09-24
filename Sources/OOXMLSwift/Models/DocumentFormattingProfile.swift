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

    public var errorDescription: String? {
        switch self {
        case .unsupportedVersion(let version): return "不支援的格式快照版本：\(version)"
        case .missingRequiredFormatting(let field): return "格式快照缺少必要資料：\(field)"
        case .invalidSnapshot(let reason): return "格式快照無效：\(reason)"
        case .unsupportedFormatting(let field): return "無法安全保留範本格式：\(field)"
        case .unsupportedNumbering: return "格式快照第一版不支援編號定義或非零 numId"
        case .duplicateRelationship(let type): return "文件套件內同一 Type（\(type)）出現多筆 Target 不同的 Relationship，屬於不合法輸入"
        }
    }
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

    /// Read only the fixed formatting parts directly from ZIP, never extract
    /// archive-controlled paths. Unsupported numbering fails explicitly.
    public static func importOfficial(from templateURL: URL) throws -> Self {
        let archive = try Archive(url: templateURL, accessMode: .read)
        func read(_ path: String, required: Bool = false) throws -> XmlNode? {
            let matches = archive.filter { $0.path == path && $0.type == .file }
            guard matches.count <= 1 else { throw DocumentFormattingProfileError.invalidSnapshot("duplicate part") }
            guard let entry = matches.first else {
                if required { throw DocumentFormattingProfileError.missingRequiredFormatting(path) }
                return nil
            }
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
            return try ProfileXML.parseRejectingDTD(bytes)
        }
        // PsychQuant/macdoc#196: formatting parts are read from fixed paths,
        // so a template whose relationships name two different parts for
        // one implicit Type is refused rather than half-honoured.
        if let relationships = try read("word/_rels/document.xml.rels") {
            try ProfileXML.rejectDuplicateImplicitRelationships(in: relationships)
        }
        if let numbering = try read("word/numbering.xml") {
            guard numbering.namespaceURI == ProfileXML.w, numbering.localName == "numbering" else {
                throw DocumentFormattingProfileError.invalidSnapshot("numbering root")
            }
            if ProfileXML.walk(numbering).contains(where: { $0.namespaceURI == ProfileXML.w && ["abstractNum", "num", "numPicBullet"].contains($0.localName) }) {
                throw DocumentFormattingProfileError.unsupportedNumbering
            }
        }
        let rawStyles = try read("word/styles.xml", required: true)!
        let styles = try ProfileXML.clean(rawStyles, root: "styles")
        let document = try read("word/document.xml", required: true)!
        guard document.namespaceURI == ProfileXML.w, document.localName == "document",
              let body = ProfileXML.child(document, "body"),
              let section = body.children.last(where: { $0.kind == .element }),
              section.namespaceURI == ProfileXML.w, section.localName == "sectPr" else {
            throw DocumentFormattingProfileError.missingRequiredFormatting("final body sectPr")
        }
        let safeSection = try ProfileXML.clean(section, root: "sectPr", inherited: ProfileXML.namespaceScope(document))
        let theme = try read("word/theme/theme1.xml").map { try ProfileXML.clean($0, root: "theme") }
        let fonts = try read("word/fontTable.xml").map { try ProfileXML.clean($0, root: "fonts") }
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
        guard let defaults = ProfileXML.child(styles, "docDefaults"),
              let r = ProfileXML.child(defaults, "rPrDefault").flatMap({ ProfileXML.child($0, "rPr") }),
              let size = ProfileXML.child(r, "sz").flatMap({ ProfileXML.value($0, "val") }).flatMap(Int.init), size > 0,
              ProfileXML.child(defaults, "pPrDefault") != nil else {
            throw DocumentFormattingProfileError.missingRequiredFormatting("docDefaults/rPrDefault/sz and pPrDefault")
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
    static func checked(_ xml: String, root: String) throws -> XmlNode { try clean(parseRejectingDTD(Data(xml.utf8)), root: root, strict: true) }

    /// Profile payloads and template formatting parts refuse any DTD, like
    /// DocxReader does for the parts it reads; the snapshot parser would
    /// otherwise skip it silently (PsychQuant/macdoc#196).
    static func parseRejectingDTD(_ data: Data) throws -> XmlNode {
        do { try DocxReader.rejectDTD(data, part: "formatting snapshot") }
        catch { throw DocumentFormattingProfileError.invalidSnapshot("DTD not allowed") }
        return try parse(data)
    }

    static let relationshipsNS = "http://schemas.openxmlformats.org/package/2006/relationships"
    static let implicitRelationshipTypes = ["styles", "theme", "fontTable"]

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
    /// - `.` and `..` segments are removed (§5.2.4); empty segments are kept.
    /// - Case is not folded. This is the decided, locked behavior: a
    ///   case-different Target names a different part here.
    static func normalizedRelationshipTarget(_ target: String) -> String {
        let path = target.hasPrefix("/") ? target : "/word/" + target
        var segments: [String] = []
        for raw in path.split(separator: "/", omittingEmptySubsequences: false).dropFirst() {
            switch normalizedPercentEncoding(raw) {
            case ".": continue
            case "..": if !segments.isEmpty { segments.removeLast() }
            case let segment: segments.append(segment)
            }
        }
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

    /// PsychQuant/macdoc#196 policy point 2. ECMA-376 allows at most one
    /// implicit styles/theme/fontTable relationship from the main document
    /// part. Registrations of one Type that resolve to different parts make
    /// the package malformed: fail closed instead of letting a `.first`
    /// lookup pick one. Equivalent spellings of one part are not a violation.
    static func rejectDuplicateImplicitRelationships(in rels: XmlNode) throws {
        for rel in implicitRelationshipTypes {
            let relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + rel
            let targets = rels.children.filter {
                $0.kind == .element && $0.namespaceURI == relationshipsNS && $0.localName == "Relationship"
                    && $0.attributeValue(prefix: nil, localName: "Type") == relationshipType
            }.map { normalizedRelationshipTarget($0.attributeValue(prefix: nil, localName: "Target") ?? "") }
            guard Set(targets).count <= 1 else { throw DocumentFormattingProfileError.duplicateRelationship(rel) }
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
