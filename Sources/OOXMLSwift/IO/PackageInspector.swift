import Foundation
import ZIPFoundation

/// One image relationship, qualified by the part whose `.rels` declares it.
/// Relationship ids are scoped per part in OPC (`word/_rels/header1.xml.rels`
/// and `word/_rels/document.xml.rels` may both declare `rId4`), so a bare id
/// is not an identity.
public struct ImageRelationshipRef: Equatable, Hashable, Sendable {
    /// Package path of the part that owns the relationship, e.g. `word/document.xml`.
    public let part: String
    /// The relationship `Id` as declared in that part's `.rels`.
    public let id: String
    /// `"<part>:<id>"` — stable string form for messages and set diffs.
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
///   `word/_rels/<path>.rels` is compared against `word/<path>`.
/// - It is a comment-stripped attribute scan, not an OPC/XML parser: it
///   accepts both quote styles and any namespace prefix on `embed`/`link`/`id`,
///   but does not resolve entities or validate the XML.
/// - When an archive carries two entries with the same name, the first is
///   used (ZIPFoundation subscript semantics); the second is invisible.
/// - No size limits are applied to the package (PsychQuant/ooxml-swift#130).
///
/// The authoritative failure signal is `orphanImageRelationshipRefs`. Raw
/// counts are informational — they can legitimately diverge (two
/// relationships may share one media file; a header image is not a body
/// drawing).
public struct ImageConsistencyReport: Equatable, Sendable {
    /// `<w:drawing>` occurrences in `word/document.xml` (body only).
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

    /// True when every declared image relationship is referenced by its own part.
    public var isConsistent: Bool { orphanImageRelationshipRefs.isEmpty }
}

public enum PackageInspector {

    private static let documentPart = "word/document.xml"

    /// Inspect a serialized .docx package for image consistency.
    ///
    /// - Parameter packageData: the bytes a writer produced (e.g.
    ///   `DocxWriter.writeData(_:)` output, or a file read back from disk).
    public static func imageConsistencyReport(of packageData: Data) throws -> ImageConsistencyReport {
        let archive = try Archive(data: packageData, accessMode: .read)

        func entryText(_ path: String) throws -> String? {
            guard let entry = archive[path] else { return nil }
            var data = Data()
            _ = try archive.extract(entry) { data.append($0) }
            return String(decoding: data, as: UTF8.self)
        }

        var declared: [ImageRelationshipRef] = []
        var referencedByPart: [String: Set<String>] = [:]

        for entry in archive {
            let path = entry.path
            guard path.hasPrefix("word/_rels/"), path.hasSuffix(".rels") else { continue }
            let partName = String(path.dropFirst("word/_rels/".count).dropLast(".rels".count))
            guard partName.hasSuffix(".xml") else { continue }   // nested parts (charts/, diagrams/) included
            let part = "word/" + partName

            let relsXML = (try entryText(path)) ?? ""
            declared.append(contentsOf: imageRelationshipIds(inRels: relsXML).map {
                ImageRelationshipRef(part: part, id: $0)
            })
            if referencedByPart[part] == nil {
                referencedByPart[part] = referencedRelationshipIds(inPart: (try entryText(part)) ?? "")
            }
        }

        let orphans = declared.filter { !(referencedByPart[$0.part] ?? []).contains($0.id) }
        let documentXML = (try entryText(documentPart)) ?? ""
        let media = archive.filter { $0.path.hasPrefix("word/media/") }.count

        return ImageConsistencyReport(
            bodyDrawingCount: documentXML.components(separatedBy: "<w:drawing").count - 1,
            imageRelationshipCount: declared.count,
            mediaEntryCount: media,
            orphanImageRelationshipIds: orphans.filter { $0.part == documentPart }.map(\.id),
            orphanImageRelationshipRefs: orphans)
    }

    // MARK: - Scanning helpers (attribute-level, not a parser)

    private static let relationshipElement = try! NSRegularExpression(
        pattern: #"<Relationship\b(?:[^>"']|"[^"]*"|'[^']*')*>"#, options: [])   // a '>' inside a quoted value does not end the element
    private static let idAttribute = try! NSRegularExpression(
        pattern: #"\bId\s*=\s*(["'])([^"']*)\1"#, options: [])
    private static let typeAttribute = try! NSRegularExpression(
        pattern: #"\bType\s*=\s*(["'])([^"']*)\1"#, options: [])
    private static let referenceAttribute = try! NSRegularExpression(
        pattern: #"(?:^|[\s<])[A-Za-z_][\w.-]*:(?:embed|link|id)\s*=\s*(["'])([^"']*)\1"#, options: [])
    private static let xmlComment = try! NSRegularExpression(
        pattern: #"<!--.*?-->"#, options: [.dotMatchesLineSeparators])

    /// Ids of `<Relationship>` elements whose `Type` ends with `/image`.
    static func imageRelationshipIds(inRels relsXML: String) -> [String] {
        let relsXML = stripComments(relsXML)   // a commented-out declaration is not a declaration
        let ns = relsXML as NSString
        let whole = NSRange(location: 0, length: ns.length)
        return relationshipElement.matches(in: relsXML, range: whole).compactMap { m in
            let element = ns.substring(with: m.range)
            guard let type = firstCapture(typeAttribute, in: element), type.hasSuffix("/image"),
                  let id = firstCapture(idAttribute, in: element) else { return nil }
            return id
        }
    }

    /// Every `*:embed` / `*:link` / `*:id` attribute value in a part, with XML
    /// comments removed first so a commented-out reference cannot satisfy a
    /// declaration.
    static func stripComments(_ xml: String) -> String {
        xmlComment.stringByReplacingMatches(in: xml, range: NSRange(location: 0, length: (xml as NSString).length), withTemplate: "")
    }

    static func referencedRelationshipIds(inPart partXML: String) -> Set<String> {
        let stripped = stripComments(partXML)
        let ns = stripped as NSString
        let whole = NSRange(location: 0, length: ns.length)
        return Set(referenceAttribute.matches(in: stripped, range: whole).map { ns.substring(with: $0.range(at: 2)) })
    }

    private static func firstCapture(_ regex: NSRegularExpression, in text: String) -> String? {
        let ns = text as NSString
        guard let m = regex.firstMatch(in: text, range: NSRange(location: 0, length: ns.length)) else { return nil }
        return ns.substring(with: m.range(at: 2))
    }
}
