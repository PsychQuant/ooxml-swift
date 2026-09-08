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

    public var errorDescription: String? {
        switch self {
        case .unsupportedVersion(let version): return "不支援的格式快照版本：\(version)"
        case .missingRequiredFormatting(let field): return "格式快照缺少必要資料：\(field)"
        case .invalidSnapshot(let reason): return "格式快照無效：\(reason)"
        case .unsupportedFormatting(let field): return "無法安全保留範本格式：\(field)"
        case .unsupportedNumbering: return "格式快照第一版不支援編號定義或非零 numId"
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
            guard entry.uncompressedSize <= 4 * 1024 * 1024 else {
                throw DocumentFormattingProfileError.invalidSnapshot("formatting part too large")
            }
            var bytes = Data()
            _ = try archive.extract(entry) { bytes.append($0) }
            return try ProfileXML.parse(bytes)
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
        "sectPr": [], "pgSz": ["w", "h", "orient"], "pgMar": ["top", "right", "bottom", "left", "header", "footer", "gutter"],
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
        "sectPr": ["pgSz", "pgMar", "cols", "docGrid"], "fonts": ["font"],
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
            return .element(prefix: ns == w ? "w" : "a", localName: name, namespaceURI: ns,
                            attributes: attrs, children: try input.children.compactMap { try rebuild($0, scope: scope, parent: name) })
        }
        let result = try rebuild(node, scope: inherited)!
        result.attributes.insert(XmlAttribute(prefix: "xmlns", localName: ns == w ? "w" : "a", value: ns), at: 0)
        return result
    }
    static func checked(_ xml: String, root: String) throws -> XmlNode { try clean(parse(xml), root: root, strict: true) }

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
                if let reference = child(style, name) {
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
