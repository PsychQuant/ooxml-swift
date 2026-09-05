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
///   with (v3.7.0, #137/#138). Attribute values therefore arrive decoded and
///   whitespace-normalized, comments and CDATA are structural rather than
///   textual, and every part is scanned once (no regex, no backtracking).
/// - A part XML rejects is listed in `unparsableParts` and produces **no
///   orphans**: unknown is not the same as missing, and reporting an orphan
///   there would refuse saves on files nothing is wrong with.
/// - When an archive carries two entries with the same name, the first is
///   used (ZIPFoundation subscript semantics); the second is invisible.
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
    /// Package paths XML could not parse, sorted (v3.7.0). Their declarations
    /// and references are unknown, so they contribute no orphans.
    public let unparsableParts: [String]

    /// True when every declared image relationship is referenced by its own part.
    public var isConsistent: Bool { orphanImageRelationshipRefs.isEmpty }

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

        func entryData(_ path: String) throws -> Data? {
            guard let entry = archive[path] else { return nil }
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

        for entry in archive {
            let path = entry.path
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

            let rels = scanRels((try entryData(path)) ?? Data())
            if !rels.parsed { unparsable.insert(path) }
            declared.append(contentsOf: rels.imageIds.map { ImageRelationshipRef(part: part, id: $0) })
            duplicates.append(contentsOf: rels.duplicateIds.map { ImageRelationshipRef(part: part, id: $0) })

            if referencedByPart[part] == nil {
                let content = scanPart((try entryData(part)) ?? Data(), countDrawingsAsIn: part == documentPart)
                if !content.parsed { unparsable.insert(part); unparsableContentParts.insert(part) }
                referencedByPart[part] = content.referenced
                if part == documentPart { documentDrawings = content.drawingCount; documentScanned = true }
            }
        }

        // `word/document.xml` is scanned above only when its rels part exists.
        if !documentScanned, let data = try entryData(documentPart) {
            let content = scanPart(data, countDrawingsAsIn: true)
            if !content.parsed { unparsable.insert(documentPart) }
            documentDrawings = content.drawingCount
        }

        let orphans = declared.filter {
            !unparsableContentParts.contains($0.part) && !(referencedByPart[$0.part] ?? []).contains($0.id)
        }
        let media = archive.filter { $0.path.hasPrefix("word/media/") }.count

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

    /// Image relationship ids, ids declared more than once (any type), and
    /// whether the XML parsed. Namespace processing stays off, so element and
    /// attribute names arrive exactly as written and an undeclared prefix is
    /// not an error — the same tolerance the previous attribute scan had,
    /// without its entity and comment blind spots.
    static func scanRels(_ relsXML: Data) -> (imageIds: [String], duplicateIds: [String], parsed: Bool) {
        let delegate = RelsScanner()
        let parser = XMLParser(data: relsXML)
        parser.shouldResolveExternalEntities = false
        parser.delegate = delegate
        let parsed = parser.parse()
        var seen = Set<String>(), dupes: [String] = []
        for id in delegate.allIds where !seen.insert(id).inserted && !dupes.contains(id) { dupes.append(id) }
        return (delegate.imageIds, dupes, parsed)
    }

    static func scanPart(_ partXML: Data, countDrawingsAsIn countDrawings: Bool = false) -> (referenced: Set<String>, drawingCount: Int, parsed: Bool) {
        let delegate = PartScanner(countsDrawings: countDrawings)
        let parser = XMLParser(data: partXML)
        parser.shouldResolveExternalEntities = false
        parser.delegate = delegate
        let parsed = parser.parse()
        return (delegate.referenced, delegate.drawingCount, parsed)
    }

    /// String conveniences (tests and callers holding text).
    static func imageRelationshipIds(inRels relsXML: String) -> [String] {
        scanRels(Data(relsXML.utf8)).imageIds
    }

    static func referencedRelationshipIds(inPart partXML: String) -> Set<String> {
        scanPart(Data(partXML.utf8)).referenced
    }

    fileprivate static func localName(_ qualified: String) -> Substring {
        guard let colon = qualified.lastIndex(of: ":") else { return Substring(qualified) }
        return qualified[qualified.index(after: colon)...]
    }
}

/// Shared delegate behaviour: **a part that declares anything in a DTD is
/// refused.** OPC forbids DTDs in package parts, no writer emits one, and a
/// document type is where parser-differential tricks live — an entity that one
/// XML API expands in an attribute value and another does not would put this
/// scanner and `DocxReader` back into disagreement, which is the whole of
/// #137. Refusing costs nothing on real packages and makes the agreement
/// structural. (Entity-expansion bombs are separately refused by libxml2
/// itself; external entities are never resolved.)
private class PackageXMLScanner: NSObject, XMLParserDelegate {
    func parser(_ parser: XMLParser, foundInternalEntityDeclarationWithName name: String, value: String?) {
        parser.abortParsing()
    }
    func parser(_ parser: XMLParser, foundExternalEntityDeclarationWithName name: String, publicID: String?, systemID: String?) {
        parser.abortParsing()
    }
    func parser(_ parser: XMLParser, foundUnparsedEntityDeclarationWithName name: String, publicID: String?, systemID: String?, notationName: String?) {
        parser.abortParsing()
    }
    func parser(_ parser: XMLParser, foundAttributeDeclarationWithName attributeName: String, forElement elementName: String, type: String?, defaultValue: String?) {
        parser.abortParsing()
    }
    func parser(_ parser: XMLParser, foundElementDeclarationWithName elementName: String, model: String) {
        parser.abortParsing()
    }
    func parser(_ parser: XMLParser, foundNotationDeclarationWithName name: String, publicID: String?, systemID: String?) {
        parser.abortParsing()
    }
}

/// Collects `<Relationship>` declarations. Comments and CDATA are reported to
/// the delegate as their own events and are therefore never mistaken for
/// markup — the two blind spots of the pre-3.7.0 regex scan.
private final class RelsScanner: PackageXMLScanner {
    var imageIds: [String] = []
    var allIds: [String] = []

    func parser(_ parser: XMLParser, didStartElement elementName: String, namespaceURI: String?,
                qualifiedName: String?, attributes: [String: String]) {
        guard PackageInspector.localName(elementName) == "Relationship" else { return }
        var id: String?, type: String?
        for (key, value) in attributes {
            switch PackageInspector.localName(key) {
            case "Id": id = value
            case "Type": type = value
            default: break
            }
        }
        guard let id else { return }
        allIds.append(id)
        if let type, type.hasSuffix("/image") { imageIds.append(id) }
    }
}

/// Collects relationship references (`*:embed` / `*:link` / `*:id`) and, for
/// `word/document.xml`, `<w:drawing>` elements.
private final class PartScanner: PackageXMLScanner {
    let countsDrawings: Bool
    var referenced: Set<String> = []
    var drawingCount = 0

    init(countsDrawings: Bool) {
        self.countsDrawings = countsDrawings
        super.init()
    }

    func parser(_ parser: XMLParser, didStartElement elementName: String, namespaceURI: String?,
                qualifiedName: String?, attributes: [String: String]) {
        if countsDrawings, elementName == "w:drawing" { drawingCount += 1 }
        for (key, value) in attributes where key.contains(":") {
            switch PackageInspector.localName(key) {
            case "embed", "link", "id": referenced.insert(value)
            default: break
            }
        }
    }
}
