import Foundation
import ZIPFoundation

/// Post-serialization consistency report for embedded images (#175,
/// PsychQuant/macdoc#175).
///
/// The #175 failure mode was a package whose image relationships and media
/// entries existed while the body `<w:drawing>` that should reference them was
/// silently missing — every status channel reported success, and only opening
/// the file revealed the loss. This report gives save paths a cheap,
/// package-level check to turn that silence into an error.
///
/// The authoritative failure signal is `orphanImageRelationshipIds`: image
/// relationships declared in `word/_rels/document.xml.rels` that no
/// `r:embed`/`r:link` in any `word/**.xml` part references. Raw counts are
/// informational — they can legitimately diverge (images in headers/footers
/// are not body drawings; two relationships may share one media file).
public struct ImageConsistencyReport: Equatable, Sendable {
    /// `<w:drawing>` occurrences in `word/document.xml` (body only).
    public let bodyDrawingCount: Int
    /// Image relationships declared in `word/_rels/document.xml.rels`.
    public let imageRelationshipCount: Int
    /// Entries under `word/media/`.
    public let mediaEntryCount: Int
    /// Image relationship Ids no XML part references — the #175 signature.
    public let orphanImageRelationshipIds: [String]

    /// True when every declared image relationship is referenced somewhere.
    public var isConsistent: Bool { orphanImageRelationshipIds.isEmpty }
}

public enum PackageInspector {

    /// Inspect a serialized .docx package for image consistency.
    ///
    /// - Parameter packageData: the bytes a writer produced (e.g.
    ///   `DocxWriter.writeData(_:)` output, or a file read back from disk).
    public static func imageConsistencyReport(of packageData: Data) throws -> ImageConsistencyReport {
        let archive = try Archive(data: packageData, accessMode: .read)

        func entryText(_ path: String) throws -> String {
            guard let entry = archive[path] else { return "" }
            var data = Data()
            _ = try archive.extract(entry) { data.append($0) }
            return String(decoding: data, as: UTF8.self)
        }

        let documentXML = try entryText("word/document.xml")
        let relsXML = try entryText("word/_rels/document.xml.rels")

        // Image relationship ids from the document part's rels. Attribute
        // order within <Relationship> is not fixed, so capture Id and Type
        // independently per element.
        var imageRelIds: [String] = []
        for element in relsXML.components(separatedBy: "<Relationship ").dropFirst() {
            let tag = element.components(separatedBy: ">").first ?? element
            guard tag.contains("relationships/image") else { continue }
            guard let idRange = tag.range(of: #"Id="([^"]+)""#, options: .regularExpression) else { continue }
            let idAttr = String(tag[idRange])
            let id = idAttr
                .replacingOccurrences(of: "Id=\"", with: "")
                .replacingOccurrences(of: "\"", with: "")
            imageRelIds.append(id)
        }

        // Every XML part under word/ may reference an image (body, headers,
        // footers, footnotes...). Collect the referenced ids across all of
        // them so header/footer images are not misreported as orphans.
        var referencedIds: Set<String> = []
        for entry in archive where entry.path.hasPrefix("word/") && entry.path.hasSuffix(".xml") {
            let text = try entryText(entry.path)
            for match in ["r:embed=\"", "r:link=\"", "r:id=\""] {
                var search = text[text.startIndex...]
                while let range = search.range(of: match) {
                    let tail = search[range.upperBound...]
                    if let quote = tail.firstIndex(of: "\"") {
                        referencedIds.insert(String(tail[..<quote]))
                        search = tail[quote...]
                    } else { break }
                }
            }
        }

        let media = archive.filter { $0.path.hasPrefix("word/media/") }.count
        let orphans = imageRelIds.filter { !referencedIds.contains($0) }

        return ImageConsistencyReport(
            bodyDrawingCount: documentXML.components(separatedBy: "<w:drawing").count - 1,
            imageRelationshipCount: imageRelIds.count,
            mediaEntryCount: media,
            orphanImageRelationshipIds: orphans)
    }
}
