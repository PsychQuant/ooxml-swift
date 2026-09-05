import Foundation
import ZIPFoundation

/// One image relationship, qualified by the part whose `.rels` declares it.
/// Relationship ids are scoped per part in OPC (`word/_rels/header1.xml.rels`
/// and `word/_rels/document.xml.rels` may both declare `rId4`), so a bare id
/// is not an identity.
public struct ImageRelationshipRef: Equatable, Hashable, Sendable {
    /// Package path of the part that owns the relationship, e.g. `word/document.xml`.
    public let part: String
    /// The relationship `Id` **as an XML parser delivers it**: entity
    /// references resolved and attribute-value whitespace normalized, i.e. the
    /// same string `DocxReader` sees (#137). `Id="rId&#54;"` is `rId6` here.
    public let id: String
    /// `"<part>:<id>"` — display form for messages. Not an identity: a part
    /// path and an id may both contain `:`, so compare `ImageRelationshipRef`
    /// values (this type is `Hashable`), never their qualified strings.
    public var qualified: String { "\(part):\(id)" }

    public init(part: String, id: String) {
        self.part = part
        self.id = id
    }
}

/// Post-serialization consistency report for embedded images (#175,
/// PsychQuant/macdoc#175).
///
/// The #175 failure mode was a package whose image relationships and media
/// entries existed while the body `<w:drawing>` that should reference them was
/// silently missing — every status channel reported success, and only opening
/// the file revealed the loss. This report gives save paths a cheap,
/// package-level check to turn that silence into an error.
///
/// **Coverage (read this before relying on it):**
/// - Declarations are taken from **every** `word/_rels/<part>.rels`, and each
///   part's declared image relationships are compared against references
///   found in **that part only** (`word/<part>`). Header/footer/footnote
///   images are therefore covered, and a `rId4` referenced by `header1.xml`
///   cannot mask an orphan `rId4` declared by `document.xml.rels`.
/// - Nested parts (`word/charts/…`, `word/diagrams/…`) are included: any
///   `<dir>/_rels/<name>.rels` under `word/` is compared against `<dir>/<name>`.
/// - Scanning is done by `XMLParser` — the same libxml2 `DocxReader` reads
///   with (v3.7.0, #137/#138) — **with namespace processing on**, so an
///   undeclared prefix is a parse error here exactly as it is in the reader's
///   `XMLDocument` (a 7-byte `<zz:x/>` in a part makes both refuse it).
///   Attribute values arrive decoded and whitespace-normalized, and comments
///   and CDATA are structural rather than textual. `Id` / `Type` are matched
///   by exact attribute name and `Relationship` by local name — the two
///   lookups `DocxReader` makes. A reference is an attribute whose prefix
///   resolves to the OOXML relationships namespace (transitional or strict)
///   with local name `embed` / `link` / `id`; a same-named attribute in any
///   other namespace is not a reference and cannot satisfy a declaration.
/// - Part names are compared case-insensitively (OPC part names are), so a
///   package spelled `Word/…` is scanned like `word/…` — the same way the
///   reader finds it after extraction to a case-insensitive file system.
/// - Parts that are not UTF-8 (a UTF-16 BOM, or NUL bytes in the first
///   bytes) are refused, as the reader refuses them; a UTF-8 BOM is fine.
/// - Nesting deeper than `maxElementDepth` (1024, the reader's `XmlTreeReader` limit) is refused,
///   as the reader refuses it.
/// - Before a part reaches the parser it passes `DocxReader.rejectDTD` (the
///   package-wide DTD policy) and a linear byte pre-check that refuses the
///   inputs libxml2 itself handles super-linearly: a comment containing `--`
///   (illegal XML 1.0 anyway), an unterminated comment or CDATA section, and
///   a start tag with more than `maxAttributesPerElement` attributes. Cost is
///   therefore bounded by package size; the package size itself is not
///   bounded here (PsychQuant/ooxml-swift#130).
/// - A part that is present but refused or unreadable is listed in
///   `unparsableParts`. Its own declarations and references are discarded —
///   unknown is not the same as missing, so it yields no orphans — **and
///   `isConsistent` is false**: an unreadable part is not a verdict of
///   consistency, and treating it as one would let a crafted package hide a
///   real orphan by corrupting an unrelated part. A part that is absent
///   (its `.rels` exists, the part does not) is knowable: its relationships
///   are unreferenced and are reported as orphans, as in 3.6.x.
/// - When an archive carries two entries with the same name, the first is
///   used; the second is invisible.
/// - No size limits are applied to the package (PsychQuant/ooxml-swift#130).
///
/// The authoritative failure signal is `orphanImageRelationshipRefs`. Raw
/// counts are informational — they can legitimately diverge (two
/// relationships may share one media file; a header image is not a body
/// drawing).
public struct ImageConsistencyReport: Equatable, Sendable {
    /// `<w:drawing>` elements in `word/document.xml` (body only). Counts every
    /// drawing, including charts, shapes and text boxes — a drawing is not an
    /// image, so this number does not decide whether a package has images.
    public let bodyDrawingCount: Int
    /// Image relationships declared across every `word/_rels/*.rels` part.
    public let imageRelationshipCount: Int
    /// Entries under `word/media/`.
    public let mediaEntryCount: Int
    /// Orphan ids declared by `word/_rels/document.xml.rels` (bare ids —
    /// kept for callers written against 3.6.0). Subset of the refs below.
    public let orphanImageRelationshipIds: [String]
    /// Every image relationship no reference in its own part points at — the
    /// #175 signature, across all parts.
    public let orphanImageRelationshipRefs: [ImageRelationshipRef]
    /// Every image relationship each part declares, in declaration order
    /// (v3.7.0). A consumer reconciling a listing against the package needs
    /// this: a relationship whose media file is missing, or whose target is
    /// external, never reaches `WordDocument.images` and would otherwise be
    /// invisible.
    public let declaredImageRelationshipRefs: [ImageRelationshipRef]
    /// Relationship ids declared more than once within one part, any type
    /// (v3.7.0). OPC forbids this; `DocxWriter` refuses to serialize such a
    /// document (#139), so a consumer can name the defect before trying.
    public let duplicateRelationshipRefs: [ImageRelationshipRef]
    /// Package paths that were present but could not be scanned, sorted
    /// (v3.7.0): a `.rels` part or a content part, each named by its own path.
    /// Their declarations and references are unknown, so they contribute no
    /// orphans — and they make `isConsistent` false.
    public let unparsableParts: [String]

    /// True when every declared image relationship is referenced by its own
    /// part, **every part could be scanned, and no part declares an id
    /// twice**. Never true for a package with an unreadable part ("no verdict"
    /// is not "consistent") nor for one `DocxWriter` would refuse to
    /// serialize. `orphanImageRelationshipIds` (the 3.6.0 bare-id view) is
    /// empty for an unreadable document part — read this property, not that
    /// list, to decide anything.
    public var isConsistent: Bool {
        orphanImageRelationshipRefs.isEmpty && unparsableParts.isEmpty && duplicateRelationshipRefs.isEmpty
    }

    public init(bodyDrawingCount: Int,
                imageRelationshipCount: Int,
                mediaEntryCount: Int,
                orphanImageRelationshipIds: [String],
                orphanImageRelationshipRefs: [ImageRelationshipRef],
                declaredImageRelationshipRefs: [ImageRelationshipRef] = [],
                duplicateRelationshipRefs: [ImageRelationshipRef] = [],
                unparsableParts: [String] = []) {
        self.bodyDrawingCount = bodyDrawingCount
        self.imageRelationshipCount = imageRelationshipCount
        self.mediaEntryCount = mediaEntryCount
        self.orphanImageRelationshipIds = orphanImageRelationshipIds
        self.orphanImageRelationshipRefs = orphanImageRelationshipRefs
        self.declaredImageRelationshipRefs = declaredImageRelationshipRefs
        self.duplicateRelationshipRefs = duplicateRelationshipRefs
        self.unparsableParts = unparsableParts
    }
}

public enum PackageInspector {

    private static let documentPart = "word/document.xml"

    /// Inspect a serialized .docx package for image consistency.
    ///
    /// - Parameter packageData: the bytes a writer produced (e.g.
    ///   `DocxWriter.writeData(_:)` output, or a file read back from disk).
    public static func imageConsistencyReport(of packageData: Data) throws -> ImageConsistencyReport {
        let archive = try Archive(data: packageData, accessMode: .read)

        // OPC part names are case-insensitive; first entry per name wins.
        var entryByLowerPath: [String: Entry] = [:]
        var orderedLowerPaths: [String] = []
        for entry in archive {
            let lower = entry.path.lowercased()
            if entryByLowerPath[lower] == nil { entryByLowerPath[lower] = entry; orderedLowerPaths.append(lower) }
        }
        func entryData(_ lowerPath: String) throws -> Data? {
            guard let entry = entryByLowerPath[lowerPath] else { return nil }
            var data = Data()
            _ = try archive.extract(entry) { data.append($0) }
            return data
        }

        var declared: [ImageRelationshipRef] = []
        var duplicates: [ImageRelationshipRef] = []
        var referencedByPart: [String: Set<String>] = [:]
        var unparsable: Set<String> = []
        var unparsableContentParts: Set<String> = []
        var documentDrawings = 0
        var documentScanned = false

        for path in orderedLowerPaths {
            // OPC: the relationships of part `<dir>/<name>` live at
            // `<dir>/_rels/<name>.rels` — so `word/charts/chart1.xml` is served by
            // `word/charts/_rels/chart1.xml.rels`, NOT `word/_rels/charts/…`
            // (R3 codex F5 corrected the earlier formula).
            guard path.hasPrefix("word/"), path.hasSuffix(".rels"),
                  let relsDirRange = path.range(of: "/_rels/", options: .backwards) else { continue }
            let dir = String(path[..<relsDirRange.lowerBound])
            let name = String(path[relsDirRange.upperBound...].dropLast(".rels".count))
            guard name.hasSuffix(".xml"), !name.contains("/") else { continue }
            let part = dir + "/" + name

            let rels = scanRels((try entryData(path)) ?? Data(), part: path)
            if rels.parsed {
                declared.append(contentsOf: rels.imageIds.map { ImageRelationshipRef(part: part, id: $0) })
                duplicates.append(contentsOf: rels.duplicateIds.map { ImageRelationshipRef(part: part, id: $0) })
            } else {
                // Whatever the parser delivered before it failed is not a
                // declaration list, it is a prefix of one. Discard it.
                unparsable.insert(path)
            }

            if referencedByPart[part] == nil {
                // A part that is *absent* is not a part that is *unreadable*:
                // nothing was lost in translation, the content simply is not
                // there, and its relationships are unreferenced in the #175
                // sense. Keep the pre-3.7.0 verdict for that case and reserve
                // "no verdict" for XML this scanner could not read.
                if let data = try entryData(part) {
                    let content = scanPart(data, part: part, countDrawingsAsIn: part == documentPart)
                    if !content.parsed { unparsable.insert(part); unparsableContentParts.insert(part) }
                    referencedByPart[part] = content.referenced
                    if part == documentPart { documentDrawings = content.drawingCount; documentScanned = true }
                } else {
                    referencedByPart[part] = []
                }
            }
        }

        // `word/document.xml` is scanned above only when its rels part exists.
        if !documentScanned, let data = try entryData(documentPart) {
            let content = scanPart(data, part: documentPart, countDrawingsAsIn: true)
            if !content.parsed { unparsable.insert(documentPart) }
            documentDrawings = content.drawingCount
        }

        let orphans = declared.filter {
            !unparsableContentParts.contains($0.part) && !(referencedByPart[$0.part] ?? []).contains($0.id)
        }
        let media = orderedLowerPaths.filter { $0.hasPrefix("word/media/") }.count   // unique names, like every other count here

        return ImageConsistencyReport(
            bodyDrawingCount: documentDrawings,
            imageRelationshipCount: declared.count,
            mediaEntryCount: media,
            orphanImageRelationshipIds: orphans.filter { $0.part == documentPart }.map(\.id),
            orphanImageRelationshipRefs: orphans,
            declaredImageRelationshipRefs: declared,
            duplicateRelationshipRefs: duplicates,
            unparsableParts: unparsable.sorted())
    }

    // MARK: - Scanning (XMLParser — the parser the reader uses, #137/#138)

    /// Start tags with more attributes than this are refused before parsing:
    /// libxml2's attribute handling is quadratic in the per-element count
    /// (measured: 32k attributes ≈ 0.3 s, doubling ≈ ×3.5), and no OOXML
    /// element carries more than a few dozen. Total attributes across many
    /// elements are linear and unbounded here.
    public static let maxAttributesPerElement = 4096

    /// Mirrors `XmlTreeReader`'s element-depth limit: the reader refuses a
    /// part nested deeper than this, so the inspector must not parse it.
    public static let maxElementDepth = 1024

    /// Why a part was refused before the parser saw it, or `nil` when it may
    /// be parsed. Linear in the part's bytes; visits each byte once.
    ///
    /// Refuses exactly the inputs on which libxml2 is not linear: a comment
    /// that contains `--` (its error-recovery path is quadratic in the number
    /// of `--` occurrences — the `<!--<!--<!--…` and `<!-- ---- … -->` shapes
    /// are both this), an unterminated comment or CDATA section (recovery
    /// again), and an over-wide start tag (see `maxAttributesPerElement`).
    /// Everything else is handed to the parser as-is.
    static func linearPrecheckFailure(_ data: Data) -> String? {
        let b = [UInt8](data); let n = b.count; var i = 0
        // Encoding: the reader is a UTF-8 reader (it accepts a UTF-8 BOM and
        // nothing wider); refuse what it would refuse rather than transcode
        // into a verdict it could never reach.
        if n >= 2, (b[0] == 0xFF && b[1] == 0xFE) || (b[0] == 0xFE && b[1] == 0xFF) { return "not UTF-8 (UTF-16 byte-order mark)" }
        if n >= 4, b[0] == 0 || b[1] == 0 || b[2] == 0 || b[3] == 0 { return "not UTF-8 (NUL bytes in the first four bytes)" }
        var depth = 0
        @inline(__always) func at(_ k: Int, _ lit: StaticString) -> Bool {
            let u = lit.utf8Start; let c = lit.utf8CodeUnitCount
            guard k + c <= n else { return false }
            for o in 0..<c where b[k + o] != u[o] { return false }
            return true
        }
        while i < n {
            guard b[i] == 0x3C /* < */ else { i += 1; continue }
            if at(i, "<!--") {
                var j = i + 4
                while true {
                    guard j + 1 < n else { return "unterminated comment" }
                    if b[j] == 0x2D && b[j + 1] == 0x2D {          // "--"
                        if j + 2 < n && b[j + 2] == 0x3E { i = j + 3; break }   // "-->"
                        return "comment contains \"--\" (illegal in XML 1.0; libxml2 recovers from it quadratically)"
                    }
                    j += 1
                }
            } else if at(i, "<![CDATA[") {
                var j = i + 9
                while true {
                    guard j + 2 < n else { return "unterminated CDATA section" }
                    if b[j] == 0x5D && b[j + 1] == 0x5D && b[j + 2] == 0x3E { i = j + 3; break }   // "]]>"
                    j += 1
                }
            } else if at(i, "<?") {
                var j = i + 2
                while j + 1 < n && !(b[j] == 0x3F && b[j + 1] == 0x3E) { j += 1 }               // "?>"
                i = j + 2
            } else {
                let isEnd = i + 1 < n && b[i + 1] == 0x2F                       // "</"
                var j = i + 1; var quote: UInt8 = 0; var equals = 0; var last: UInt8 = 0
                while j < n {
                    let c = b[j]
                    if quote != 0 { if c == quote { quote = 0 } }
                    else if c == 0x22 || c == 0x27 { quote = c }
                    else if c == 0x3D { equals += 1; if equals > maxAttributesPerElement { return "start tag with more than \(maxAttributesPerElement) attributes" } }
                    else if c == 0x3E { break }
                    if c > 0x20 { last = c }
                    j += 1
                }
                if isEnd { depth -= 1 }
                else if last != 0x2F && !(i + 1 < n && b[i + 1] == 0x21) {   // start tag, not self-closing, not "<!…"
                    depth += 1
                    if depth > maxElementDepth { return "nesting deeper than \(maxElementDepth) elements (the reader's limit)" }
                }
                i = j + 1
            }
        }
        return nil
    }

    /// Image relationship ids, ids declared more than once (any type), and
    /// whether the part could be scanned. `Id` and `Type` are looked up by
    /// exact attribute name, as `DocxReader` does (`attribute(forName:)`);
    /// `xmlns:Id`, `r:Id` or `p:Type` are not those attributes.
    static func scanRels(_ relsXML: Data, part: String = "") -> (imageIds: [String], duplicateIds: [String], parsed: Bool, allIds: [String]) {
        guard admissible(relsXML, part: part) else { return ([], [], false, []) }
        let delegate = RelsScanner()
        let parsed = makeParser(relsXML, delegate: delegate).parse()
        var seen = Set<String>(), flagged = Set<String>(), dupes: [String] = []
        for id in delegate.allIds where !seen.insert(id).inserted && flagged.insert(id).inserted { dupes.append(id) }
        return (delegate.imageIds, dupes, parsed, delegate.allIds)
    }

    static func scanPart(_ partXML: Data, part: String = "", countDrawingsAsIn countDrawings: Bool = false) -> (referenced: Set<String>, drawingCount: Int, parsed: Bool) {
        guard admissible(partXML, part: part) else { return ([], 0, false) }
        let delegate = PartScanner(countsDrawings: countDrawings)
        let parsed = makeParser(partXML, delegate: delegate).parse()
        return (delegate.referenced, delegate.drawingCount, parsed)
    }

    /// Namespace processing on: an undeclared prefix is a parse error, as it
    /// is for the reader's `XMLDocument`; prefixes are reported so scanners
    /// can resolve attribute namespaces; external entities never resolve.
    private static func makeParser(_ data: Data, delegate: XMLParserDelegate) -> XMLParser {
        let parser = XMLParser(data: data)
        parser.shouldProcessNamespaces = true
        parser.shouldReportNamespacePrefixes = true
        parser.shouldResolveExternalEntities = false
        parser.delegate = delegate
        return parser
    }

    static let relationshipsNamespaces: Set<String> = [
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "http://purl.oclc.org/ooxml/officeDocument/relationships",
    ]
    static let wordprocessingNamespaces: Set<String> = [
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "http://purl.oclc.org/ooxml/wordprocessingml/main",
    ]

    /// The two gates every part passes before the parser: the package-wide
    /// DTD policy (`DocxReader.rejectDTD` — one policy, one place) and the
    /// linear pre-check above.
    private static func admissible(_ data: Data, part: String) -> Bool {
        guard (try? DocxReader.rejectDTD(data, part: part)) != nil else { return false }
        return linearPrecheckFailure(data) == nil
    }

    /// String conveniences (tests and callers holding text).
    static func imageRelationshipIds(inRels relsXML: String) -> [String] {
        scanRels(Data(relsXML.utf8)).imageIds
    }

    static func referencedRelationshipIds(inPart partXML: String) -> Set<String> {
        scanPart(Data(partXML.utf8)).referenced
    }

}

/// Tracks in-scope namespace prefixes (namespace processing is on, so the
/// parser reports every `xmlns:…` mapping as it opens and closes).
private class NamespaceTrackingScanner: NSObject, XMLParserDelegate {
    private var scopes: [String: [String]] = [:]

    func parser(_ parser: XMLParser, didStartMappingPrefix prefix: String, toURI namespaceURI: String) {
        scopes[prefix, default: []].append(namespaceURI)
    }
    func parser(_ parser: XMLParser, didEndMappingPrefix prefix: String) {
        _ = scopes[prefix]?.popLast()
    }
    func namespace(ofAttribute qualifiedName: String) -> String? {
        guard let colon = qualifiedName.firstIndex(of: ":") else { return nil }     // unprefixed attributes have no namespace
        return scopes[String(qualifiedName[..<colon])]?.last
    }
    static func localName(_ qualified: String) -> Substring {
        guard let colon = qualified.firstIndex(of: ":") else { return Substring(qualified) }
        return qualified[qualified.index(after: colon)...]
    }
}

/// Collects `<Relationship>` declarations. Comments and CDATA are reported to
/// the delegate as their own events and are therefore never mistaken for
/// markup — the two blind spots of the pre-3.7.0 regex scan. Elements are
/// matched by local name and attributes by exact name: the two lookups
/// `DocxReader.parseRelationshipsFile` makes (`local-name()='Relationship'`,
/// `attribute(forName: "Id")`).
private final class RelsScanner: NamespaceTrackingScanner {
    var imageIds: [String] = []
    var allIds: [String] = []

    func parser(_ parser: XMLParser, didStartElement elementName: String, namespaceURI: String?,
                qualifiedName: String?, attributes: [String: String]) {
        guard elementName == "Relationship", let id = attributes["Id"] else { return }   // namespace processing on → local name
        allIds.append(id)
        if let type = attributes["Type"], type.hasSuffix("/image") { imageIds.append(id) }
    }
}

/// Collects relationship references — attributes in the OOXML relationships
/// namespace with local name `embed` / `link` / `id` — and, for
/// `word/document.xml`, `<w:drawing>` elements (by namespace, not prefix).
private final class PartScanner: NamespaceTrackingScanner {
    let countsDrawings: Bool
    var referenced: Set<String> = []
    var drawingCount = 0

    init(countsDrawings: Bool) {
        self.countsDrawings = countsDrawings
    }

    func parser(_ parser: XMLParser, didStartElement elementName: String, namespaceURI: String?,
                qualifiedName: String?, attributes: [String: String]) {
        if countsDrawings, elementName == "drawing", let ns = namespaceURI, PackageInspector.wordprocessingNamespaces.contains(ns) {
            drawingCount += 1
        }
        for (key, value) in attributes {
            guard let ns = namespace(ofAttribute: key), PackageInspector.relationshipsNamespaces.contains(ns) else { continue }
            switch Self.localName(key) {
            case "embed", "link", "id": referenced.insert(value)
            default: break
            }
        }
    }
}
