import Foundation
import ZIPFoundation

/// One image relationship, qualified by the part whose `.rels` declares it.
/// Relationship ids are scoped per part in OPC (`word/_rels/header1.xml.rels`
/// and `word/_rels/document.xml.rels` may both declare `rId4`), so a bare id
/// is not an identity.
public struct ImageRelationshipRef: Equatable, Hashable, Sendable {
    /// Package path of the part that owns the relationship: always
    /// `word/document.xml` for the document (the reader's name for it,
    /// whatever spelling the archive stored); other parts as `word/` + the
    /// sub-path the file system lists after extraction.
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
/// - Declarations are taken from every `<dir>/_rels/<name>.xml.rels` under
///   `word/` whose owner is an `.xml` part — present or absent — and each
///   part's declared image relationships are compared against references
///   found in **that part only** (`<dir>/<name>.xml`). Header/footer/footnote
///   and nested (`word/charts/…`) images are therefore covered, and a `rId4`
///   referenced by `header1.xml` cannot mask an orphan `rId4` declared by
///   `document.xml.rels`. A rels whose owner has another extension
///   (`vmlDrawing1.vml.rels`) is not reconciled (PsychQuant/ooxml-swift#144).
/// - Scanning is done by `XMLParser` — the same libxml2 `DocxReader` reads
///   with (v3.7.0, #137/#138) — **with namespace processing on**, and the
///   scanners refuse an undeclared prefix themselves (libxml2's SAX path
///   only records one; the reader's `XMLDocument` refuses, so a 7-byte
///   `<zz:x/>` in a part makes both refuse it).
/// - **A part is refused for these reasons and no others** (closed list):
///   a DTD (`DocxReader.rejectDTD`); bytes the linear pre-check rules out
///   (not UTF-8, comment containing `--`, unterminated comment or CDATA,
///   over-wide start tag, nesting past `maxElementDepth`); XML that
///   `XMLParser` cannot parse; an element or attribute prefix for which the
///   parser reports no `xmlns:` mapping — undeclared, or declared with an
///   empty URI (libxml2 reports no mapping for `xmlns:zz=""`; that is the
///   parser's report, not a rule of ours). `XMLParser` is otherwise **more
///   lenient than the reader's `XMLDocument` on namespace well-formedness**,
///   and that leniency is deliberate, not a gap to fill: a QName with two
///   colons or a trailing colon, a leading colon under a default namespace,
///   a prefix bound to an invalid URI, the `xml` prefix bound to the wrong
///   URI, and two prefixes for one URI carrying the same attribute all
///   parse here and are refused by the reader (verify R3 DA, twelve
///   classes; the ten shapes it built are pinned by tests — eight accepted,
///   the two empty-URI ones refused). Re-implementing libxml2's namespace layer in a delegate is
///   the anti-pattern this type exists to remove (#137); the reader's own
///   verdict is not a single predicate either — it parses only the document,
///   headers, footers, footnotes, endnotes, comments and their rels with
///   `XMLDocument`, and charts, settings and the rest with a lenient tree
///   reader. So: **`isConsistent` is a statement about relationships, never
///   "the reader can open this package"**. On every one of those shapes the
///   report's declarations and references are exactly what the parser
///   delivered (eight of the DA's ten packages report equal to the clean
///   control; the two empty-URI ones are refused and say so).
///   Attribute values arrive decoded and whitespace-normalized, and comments
///   and CDATA are structural rather than textual. `Id` / `Type` are matched
///   by exact attribute name and `Relationship` by local name — the two
///   lookups `DocxReader` makes. A reference is an attribute whose prefix
///   resolves to the OOXML relationships namespace (transitional or strict)
///   with local name `embed` / `link` / `id`; a same-named attribute in any
///   other namespace is not a reference and cannot satisfy a declaration.
/// - The package is extracted with the reader's own `ZipHelper.unzip` and
///   parts are read back from the file system, so part names mean exactly
///   what they mean to the reader: `Word/…` is `word/…` where the file
///   system folds case, `word/_rels/./x.rels` is `word/_rels/x.rels`, and two
///   entries that collide make extraction — and therefore the report —
///   fail (`WordError.invalidDocx`), as they make the reader fail.
/// - Parts that are not valid UTF-8 are refused (a UTF-8 BOM is fine). The
///   reader substitutes U+FFFD for a bad byte and carries on; refusing is
///   stricter, and it is the only way an id containing such a byte cannot
///   mean two different strings in the two places.
/// - Nesting deeper than `maxElementDepth` (1024, the reader's
///   `XmlTreeReader` limit, counted as it counts — every element, self-closing
///   ones included) is refused, as the reader refuses it.
/// - Before a part reaches the parser it passes `DocxReader.rejectDTD` (the
///   package-wide DTD policy) and a linear byte pre-check that refuses the
///   inputs libxml2 itself handles super-linearly: a comment containing `--`
///   (illegal XML 1.0 anyway), an unterminated comment or CDATA section, and
///   a start tag with more than `maxAttributesPerElement` attributes. What
///   remains is libxml2's own cost on the same bytes the reader parses —
///   including its namespace bookkeeping, which grows faster than linearly
///   in the number of `xmlns` declarations per element (a 4 MB package with
///   hundreds of thousands of them takes tens of seconds here and longer in
///   the reader). Neither the package size nor that axis is bounded here
///   (PsychQuant/ooxml-swift#130).
/// - **The inspector is deliberately stricter than the reader in four
///   places** (it refuses what the reader would open): a start tag wider than
///   `maxAttributesPerElement`, a comment containing `--`, bytes that are
///   not valid UTF-8 (the reader's rels path, `XMLDocument`, would even take
///   UTF-16), and — within the parts this type scans, i.e. the document and
///   every `.xml` part whose rels declares something — an undeclared prefix
///   in a part the reader never parses or never reaches (a header no section
///   references, a chart with a rels, …): the rule is applied to every
///   scanned part, not routed per part. A part without a declaring rels is
///   not scanned at all, so `settings.xml` (which almost never has one) is
///   not in that set. Cost caps and determinism rules, not reader parity.
/// - A part that is present but refused or unreadable is listed in
///   `unparsableParts`. An unreadable `.rels` contributes no declarations at
///   all; an unreadable content part keeps its rels' declarations in the
///   listing but contributes no references and yields no orphans — unknown
///   is not the same as missing — **and either makes `isConsistent` false**: an unreadable part is not a verdict of
///   consistency, and treating it as one would let a crafted package hide a
///   real orphan by corrupting an unrelated part. A part that is absent
///   (its `.rels` exists, the part does not) is knowable: its relationships
///   are unreferenced and are reported as orphans, as in 3.6.x.
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
    /// Files directly under `word/media/` (the directory entry itself, which
    /// some writers store, is not counted — v3.7.0; 3.6.x counted it).
    public let mediaEntryCount: Int
    /// Orphan ids declared by `word/_rels/document.xml.rels` (bare ids —
    /// kept for callers written against 3.6.0). Subset of the refs below.
    public let orphanImageRelationshipIds: [String]
    /// Every image relationship no reference in its own part points at — the
    /// #175 signature, across all parts.
    public let orphanImageRelationshipRefs: [ImageRelationshipRef]
    /// Every image relationship each part declares (v3.7.0): the document
    /// part first, then the other present parts in path order, then parts
    /// whose rels exists but which are themselves absent; within a part, in
    /// declaration order. A consumer reconciling a listing against the
    /// package needs this: a relationship whose media file is missing, or
    /// whose target is external, never reaches `WordDocument.images` and
    /// would otherwise be invisible.
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
    /// is not "consistent"). It says nothing about whether `DocxWriter` can
    /// serialize the package: a consistent package whose image uses `rId1`
    /// is still refused by the writer's fixed-slot collision (#140).
    /// `orphanImageRelationshipIds` (the 3.6.0 bare-id view) is empty for an
    /// unreadable document part — read this property, not that list, to
    /// decide anything.
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
    /// The package is **extracted the way `DocxReader` extracts it** —
    /// `ZipHelper.unzip`, the reader's own call, into the reader's own
    /// temporary location — and parts are then read back from the file
    /// system by constructed OPC addresses, the way the reader reaches them.
    /// Entry names are never interpreted here: whatever the file system does
    /// to them (fold case, collapse a `.` component, refuse a second entry
    /// that lands on an existing file) it does identically for the reader.
    /// Every remaining question about a listed name — is it the document, is
    /// it `<x>.xml`, is its directory `_rels`, is it `<x>.rels` — is answered
    /// by file identity (device + inode of the candidate name), not by
    /// folding strings: verify R3/R4 showed that `lowercased()` and
    /// `caseInsensitiveCompare` re-create the very mute switch (APFS folds
    /// U+017F to `s`; Swift does not) and that a lowercased archive index
    /// chose by archive order where the file system chooses by name.
    ///
    /// What is read: `word/document.xml` (always), every `word/**/*.xml`
    /// part **that has a rels** (a part without one declares nothing and is
    /// not read — reading it would let bytes the reader never opens refuse a
    /// package the reader opens), and every rels whose part is absent (its
    /// declarations are orphans, as in 3.6.x; they are listed after the
    /// present parts). A directory listing that fails is no verdict:
    /// `invalidDocx`. Part paths in the report are `word/` + the sub-path as
    /// the file system lists it; the document is always `word/document.xml`.
    ///
    /// **This function has a side effect and a cost that 3.6.x did not**: it
    /// writes the package's contents into `FileManager.default.temporaryDirectory`
    /// (a UUID-named directory under `ooxml-swift-inspector/` — the reader's
    /// `che-word-mcp/` namespace is not shared, so its own leak checks keep
    /// meaning; removed before returning, on every path) — peak disk use is the extracted size, plus
    /// the package size once more for the `Data` overload, and nothing caps
    /// either (PsychQuant/ooxml-swift#130). The extraction policy is
    /// `ZipHelper.unzip`'s, shared with the reader: an archive with a
    /// symbolic-link entry, a `..` path component or an absolute entry path
    /// is refused before anything is written (a link to `.` plus `..`
    /// components would otherwise write outside the directory — verify R4).
    ///
    /// A package the reader cannot extract — two entries whose names collide
    /// on the file system (`word/document.xml` twice, or once as
    /// `WORD/DOCUMENT.XML`), a path that escapes the directory, a corrupt
    /// member — throws `WordError.invalidDocx`, and so does one without
    /// `word/document.xml`: a package the reader cannot open has no
    /// consistency verdict, and returning one would be the "no verdict means
    /// consistent" switch this type refuses to be.
    ///
    /// - Parameter packageData: the bytes a writer produced (e.g.
    ///   `DocxWriter.writeData(_:)` output, or a file read back from disk).
    public static func imageConsistencyReport(of packageData: Data) throws -> ImageConsistencyReport {
        try imageConsistencyReport(extracting: { try ZipHelper.unzip(data: packageData, namespace: ZipHelper.inspectorNamespace) })
    }

    /// The same report for a package already on disk; nothing is copied.
    public static func imageConsistencyReport(ofPackageAt url: URL) throws -> ImageConsistencyReport {
        try imageConsistencyReport(extracting: { try ZipHelper.unzip(data: try Data(contentsOf: url), namespace: ZipHelper.inspectorNamespace) })
    }

    /// `listSubpaths` / `listDirectory` are the file-system listings the scan
    /// depends on; injectable so a test can make them fail (verify R5 DA: the
    /// "a failed listing is no verdict" rule had no test and could be
    /// reverted to `try? … ?? []` unnoticed).
    static func imageConsistencyReport(
        extracting extract: () throws -> URL,
        listSubpaths: (URL) throws -> [String] = { try FileManager.default.subpathsOfDirectory(atPath: $0.path) },
        listDirectory: (URL) throws -> [String] = { try FileManager.default.contentsOfDirectory(atPath: $0.path) }
    ) throws -> ImageConsistencyReport {
        let tempDir: URL
        do {
            tempDir = try extract()
        } catch {
            throw WordError.invalidDocx(
                "the package could not be extracted the way DocxReader extracts it (\(error.localizedDescription)); "
                + "a package the reader cannot open has no consistency verdict.")
        }
        defer { ZipHelper.cleanup(tempDir) }
        let fm = FileManager.default

        /// The file system's own answer to "is this the same file?": device
        /// and inode. Every other question about a listed name — is it the
        /// document, is it `<x>.xml`, does it live in `_rels`, is it
        /// `<x>.rels` — is asked by constructing the candidate name and
        /// comparing identities, never by folding strings ourselves (verify
        /// R3/R4: `lowercased()` and `caseInsensitiveCompare` are not the
        /// file system's rules). `nil` for anything that is not a regular file.
        struct FileIdentity: Hashable { let device: Int; let inode: Int }
        /// `nil` only when there is no such path. Any other failure to stat a
        /// path inside our own extraction directory is an I/O problem, not an
        /// absence, and absence must not be invented from it (verify R5).
        func attributes(_ url: URL) throws -> [FileAttributeKey: Any]? {
            do { return try fm.attributesOfItem(atPath: url.path) }
            catch let error as NSError where error.domain == NSCocoaErrorDomain && (error.code == NSFileReadNoSuchFileError || error.code == NSFileNoSuchFileError) { return nil }
            catch let error as NSError where error.domain == NSPOSIXErrorDomain && error.code == Int(ENOENT) { return nil }
            catch { throw WordError.invalidDocx("could not read file attributes under the extracted package (\(error.localizedDescription)); no consistency verdict.") }
        }
        func identity(_ fileURL: URL, ofType type: FileAttributeType = .typeRegular) throws -> FileIdentity? {
            guard let attrs = try attributes(fileURL), (attrs[.type] as? FileAttributeType) == type,
                  let device = (attrs[.systemNumber] as? NSNumber)?.intValue, let inode = (attrs[.systemFileNumber] as? NSNumber)?.intValue else { return nil }
            return FileIdentity(device: device, inode: inode)
        }
        /// A part as the file system serves it to the reader, or nil when
        /// there is no such regular file.
        func fileData(_ part: String) throws -> Data? {
            let fileURL = tempDir.appendingPathComponent(part)
            guard try identity(fileURL) != nil else { return nil }
            return try Data(contentsOf: fileURL)
        }
        /// Whether the file system serves `listed` under the name `<stem><suffix>`
        /// in the same directory — i.e. whether its name "ends with" `suffix`
        /// by the file system's rules, not ours. Returns the stem.
        func stem(of listed: URL, ifSuffixed suffix: String) throws -> String? {
            let name = listed.lastPathComponent
            guard name.count > suffix.count, let own = try identity(listed) else { return nil }
            let candidateStem = String(name.dropLast(suffix.count))
            let candidate = listed.deletingLastPathComponent().appendingPathComponent(candidateStem + suffix)
            return try identity(candidate) == own ? candidateStem : nil
        }

        // The reader's first act: `word/document.xml` by that path, or refuse.
        let documentURL = tempDir.appendingPathComponent(documentPart)
        guard let documentIdentity = try identity(documentURL), let documentData = try fileData(documentPart) else {
            throw WordError.invalidDocx("the package has no \(documentPart) (DocxReader refuses it too); no consistency verdict.")
        }

        var declared: [ImageRelationshipRef] = []
        var duplicates: [ImageRelationshipRef] = []
        var referencedByPart: [String: Set<String>] = [:]
        var unparsable: Set<String> = []
        var unparsableContentParts: Set<String> = []
        var consumedRels: Set<FileIdentity> = []

        let documentContent = scanPart(documentData, part: documentPart, countDrawingsAsIn: true)
        if !documentContent.parsed { unparsable.insert(documentPart); unparsableContentParts.insert(documentPart) }
        referencedByPart[documentPart] = documentContent.referenced

        /// Declarations of one part, from the `.rels` the file system serves
        /// at the OPC address `<dir>/_rels/<name>.rels` — looked up by that
        /// constructed path, which is how the reader reaches every rels it
        /// reads (`word/_rels/\(target).rels`). Returns false when there is
        /// no such rels. (`word/charts/chart1.xml` ↔ `word/charts/_rels/chart1.xml.rels`,
        /// not `word/_rels/charts/…` — R3 codex F5 corrected that formula.)
        /// Returns nil when there is no rels; otherwise how many relationships
        /// it declares (0 for an empty one, 0 for one that could not be read —
        /// which is recorded in `unparsableParts`).
        @discardableResult
        func declare(part: String) throws -> Int? {
            let slash = part.lastIndex(of: "/").map { part.index(after: $0) } ?? part.startIndex
            let relsPath = String(part[..<slash] + "_rels/" + part[slash...] + ".rels")
            let relsURL = tempDir.appendingPathComponent(relsPath)
            guard let relsIdentity = try identity(relsURL), let relsData = try fileData(relsPath) else { return nil }
            consumedRels.insert(relsIdentity)
            let rels = scanRels(relsData, part: relsPath)
            if rels.parsed {
                declared.append(contentsOf: rels.imageIds.map { ImageRelationshipRef(part: part, id: $0) })
                duplicates.append(contentsOf: rels.duplicateIds.map { ImageRelationshipRef(part: part, id: $0) })
                return rels.allIds.count
            }
            // Whatever the parser delivered before it failed is not a
            // declaration list, it is a prefix of one. Discard it.
            unparsable.insert(relsPath)
            return 0
        }
        try declare(part: documentPart)

        // Everything under `word/`, as the file system lists it. A listing
        // that fails is not an empty package — it is no verdict (verify R4).
        let wordDir = tempDir.appendingPathComponent("word")
        let listed: [String]
        do { listed = try listSubpaths(wordDir).sorted() }
        catch { throw WordError.invalidDocx("could not list the package's word/ directory after extraction (\(error.localizedDescription)); no consistency verdict.") }

        // Pass 1 — XML parts whose rels declares at least one relationship:
        // take the declarations, then scan the part for references. A part
        // without a rels, or with a rels that declares nothing, has nothing
        // to reconcile and is not read: it may be anything, and reading it
        // would refuse packages the reader opens (verify R4 B2, R5 L4). What
        // IS reconciled is every rels under `word/`, whether or not the reader
        // ever reads it — an unreadable one is "no verdict" (deliberately
        // stricter than the reader; see the class doc).
        for sub in listed {
            let fileURL = wordDir.appendingPathComponent(sub)
            guard let own = try identity(fileURL), own != documentIdentity, try stem(of: fileURL, ifSuffixed: ".xml") != nil else { continue }
            let part = "word/" + sub
            guard let declaredCount = try declare(part: part), declaredCount > 0 else { continue }
            let content = scanPart(try Data(contentsOf: fileURL), part: part)
            if !content.parsed { unparsable.insert(part); unparsableContentParts.insert(part) }
            referencedByPart[part] = content.referenced
        }

        // Pass 2 — a `.rels` whose part is absent: knowable, and the
        // pre-3.7.0 verdict (its relationships are unreferenced → orphans)
        // is right. Recognised by identity alone: the listed file's directory
        // IS `<dir>/_rels` (device + inode) and the file IS `<stem>.xml.rels`
        // to the file system — the `.xml` half is asked of the rels file's own
        // name, so no string is folded for an owner that does not exist.
        for sub in listed {
            let fileURL = wordDir.appendingPathComponent(sub)
            guard let own = try identity(fileURL), !consumedRels.contains(own) else { continue }
            let relsDir = fileURL.deletingLastPathComponent()
            let ownerDir = relsDir.deletingLastPathComponent()
            guard let relsDirIdentity = try identity(relsDir, ofType: .typeDirectory),
                  try identity(ownerDir.appendingPathComponent("_rels"), ofType: .typeDirectory) == relsDirIdentity,
                  let ownerStem = try stem(of: fileURL, ifSuffixed: ".xml.rels") else { continue }
            let ownerName = ownerStem + ".xml"
            guard try identity(ownerDir.appendingPathComponent(ownerName)) == nil else { continue }
            let ownerSub = ownerDir.path.hasPrefix(wordDir.path + "/") ? String(ownerDir.path.dropFirst(wordDir.path.count + 1)) + "/" + ownerName : ownerName
            let part = "word/" + ownerSub
            guard referencedByPart[part] == nil else { continue }
            referencedByPart[part] = []
            try declare(part: part)
        }

        let orphans = declared.filter {
            !unparsableContentParts.contains($0.part) && !(referencedByPart[$0.part] ?? []).contains($0.id)
        }
        let mediaDir = tempDir.appendingPathComponent("word/media")
        var mediaIsDirectory: ObjCBool = false
        var media = 0
        if fm.fileExists(atPath: mediaDir.path, isDirectory: &mediaIsDirectory), mediaIsDirectory.boolValue {
            do { media = try listDirectory(mediaDir).filter { (try? identity(mediaDir.appendingPathComponent($0))) != nil }.count }
            catch { throw WordError.invalidDocx("could not list the package's word/media directory after extraction (\(error.localizedDescription)); no consistency verdict.") }
        }

        return ImageConsistencyReport(
            bodyDrawingCount: documentContent.drawingCount,
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
        if String(data: data, encoding: .utf8) == nil { return "not valid UTF-8" }
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
                else if !(i + 1 < n && b[i + 1] == 0x21) {                      // a start tag (not "<!…")
                    // The reader enters every element, self-closing ones too,
                    // so a self-closing tag at the limit trips it (verify R3).
                    if depth + 1 > maxElementDepth { return "nesting deeper than \(maxElementDepth) elements (the reader's limit)" }
                    if last != 0x2F { depth += 1 }
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
    static func scanRels(_ relsXML: Data, part: String = "") -> (imageIds: [String], duplicateIds: [String], parsed: Bool, allIds: [String], structure: [String]) {
        guard admissible(relsXML, part: part) else { return ([], [], false, [], []) }
        let delegate = RelsScanner()
        let parsed = makeParser(relsXML, delegate: delegate).parse() && !delegate.sawUndeclaredPrefix
        var seen = Set<String>(), flagged = Set<String>(), dupes: [String] = []
        for id in delegate.allIds where !seen.insert(id).inserted && flagged.insert(id).inserted { dupes.append(id) }
        return (delegate.imageIds, dupes, parsed, delegate.allIds, delegate.structure)
    }

    static func scanPart(_ partXML: Data, part: String = "", countDrawingsAsIn countDrawings: Bool = false) -> (referenced: Set<String>, drawingCount: Int, parsed: Bool) {
        guard admissible(partXML, part: part) else { return ([], 0, false) }
        let delegate = PartScanner(countsDrawings: countDrawings)
        let parsed = makeParser(partXML, delegate: delegate).parse() && !delegate.sawUndeclaredPrefix
        return (delegate.referenced, delegate.drawingCount, parsed)
    }

    /// Namespace processing on, prefixes reported: scanners resolve attribute
    /// namespaces from the reported mappings and refuse an undeclared prefix
    /// themselves (see `NamespaceTrackingScanner`); external entities never
    /// resolve.
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
    /// Set when an element or attribute used a prefix no `xmlns:` declares.
    /// libxml2 only *records* that (`nsWellFormed`) and `XMLParser.parse()`
    /// never consults it, while the reader's `XMLDocument` treats it as fatal
    /// (error 201) — so the scanner refuses it itself (verify R3, B7).
    private(set) var sawUndeclaredPrefix = false

    /// Call first in `didStartElement`. False (and the parse aborted) when the
    /// element or any attribute carries a prefix that is not in scope; the
    /// parser reports an element's own `xmlns:` mappings before the element,
    /// so the scope table is complete at this point.
    func namespacesWellFormed(_ parser: XMLParser, qualifiedName: String?, namespaceURI: String?, attributes: [String: String]) -> Bool {
        if let q = qualifiedName, let colon = q.firstIndex(of: ":") {
            let prefix = String(q[..<colon])
            // A prefix the parser reported no mapping for — and only that.
            // That is two of the reader's twelve namespace classes: a prefix
            // no `xmlns:` declares, and one declared with an empty URI
            // (libxml2 reports no mapping for `xmlns:zz=""`). An empty prefix
            // (`:b`), a trailing colon (`w:`), two colons, an invalid URI,
            // `xml` rebound, an expanded duplicate attribute are ill-formed to
            // the reader and deliberately not judged here (see the class doc).
            if !prefix.isEmpty, prefix != "xml", scopes[prefix]?.last == nil { return refuse(parser) }
        }
        for key in attributes.keys {
            guard let colon = key.firstIndex(of: ":") else { continue }
            let prefix = String(key[..<colon])
            // `xml:` is bound implicitly and never reported as a mapping, so it
            // must be exempt; `xmlns:` never reaches this dictionary when
            // namespace processing is on (it arrives as a mapping event) — kept
            // only so the rule reads as the XML rule.
            if prefix == "xml" || prefix == "xmlns" { continue }
            if scopes[prefix]?.last == nil { return refuse(parser) }
        }
        return true
    }
    private func refuse(_ parser: XMLParser) -> Bool {
        sawUndeclaredPrefix = true
        parser.abortParsing()
        return false
    }

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
    /// Lexical structure the parser reported that a text scan of the same
    /// bytes cannot see the way the parser does (`DocxWriter` refuses to
    /// merge into a rels carrying any of it, #142): four kinds, each noted
    /// once in the order first seen (a fixed-size record — verify R5 caught a
    /// per-element list that grew with every distinct prefixed name and made
    /// the dedup quadratic, the very shape #138 exists to prevent). Real Word
    /// output has none of these.
    var structure: [String] { Kind.allCases.compactMap { noted[$0] } }
    private enum Kind: CaseIterable { case comment, cdata, processingInstruction, prefixedElement }
    private var noted: [Kind: String] = [:]
    private func note(_ kind: Kind, _ what: @autoclosure () -> String) { if noted[kind] == nil { noted[kind] = what() } }

    func parser(_ parser: XMLParser, foundComment comment: String) { note(.comment, "an XML comment") }
    func parser(_ parser: XMLParser, foundCDATA CDATABlock: Data) { note(.cdata, "a CDATA section") }
    func parser(_ parser: XMLParser, foundProcessingInstructionWithTarget target: String, data: String?) { note(.processingInstruction, "a processing instruction") }

    func parser(_ parser: XMLParser, didStartElement elementName: String, namespaceURI: String?,
                qualifiedName: String?, attributes: [String: String]) {
        guard namespacesWellFormed(parser, qualifiedName: qualifiedName, namespaceURI: namespaceURI, attributes: attributes) else { return }
        if let q = qualifiedName, q.contains(":") { note(.prefixedElement, "a namespace-prefixed <\(q)> element") }
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
        guard namespacesWellFormed(parser, qualifiedName: qualifiedName, namespaceURI: namespaceURI, attributes: attributes) else { return }
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
