// ReverseExtractor.swift
// format-alignment-engine Phase B tasks 2.2/2.3 — typed reverse extraction
// with the byte-equal upgrade rule (`ooxml-script-transcode`, «Reverse
// extraction covers the five format layers»; Decision 3).
// word-canonical-forms Phase 1 task 1.1 — form-gap measurement: every bail
// records the XML path to the first offending node/attribute so the report
// is the work queue for vocabulary extension (Decision 1).
//
// The upgrade rule: candidate typed ops are extracted from a part, applied
// to an empty authoring document, and re-serialized; the part leaves the raw
// channel ONLY when the trial rebuild reproduces the source bytes exactly.
// No canonical-form exemptions — a form the serializer cannot reproduce
// keeps the part on `carryPart` and the coverage metric reflects it
// honestly (design Risks: slower coverage growth, binary Stage A/B).
//
// Extraction is deliberately narrow: it recognizes exactly the shapes the
// OperationReducer stamps (makeParagraph / setRuns / setSectionProperties).
// Anything else — unknown elements, extra attributes, foreign whitespace —
// bails to raw, recording a FormGap. The final authority is always the byte
// comparison.

import Foundation

public enum ReverseExtractor {

    /// A form the extraction/rebuild could not represent, located for the
    /// measurement report (`ooxml-script-transcode`, «Form-gap measurement»;
    /// Decision 1). `xmlPath` is the breadcrumb to the first offending node
    /// or attribute for an extraction bail, or `byte@<offset> src=… out=…`
    /// for a trial-rebuild byte-mismatch. `contentClass` is the bail tag.
    public struct FormGap: Equatable {
        public let partPath: String
        public let xmlPath: String
        public let contentClass: String

        public init(partPath: String, xmlPath: String, contentClass: String) {
            self.partPath = partPath
            self.xmlPath = xmlPath
            self.contentClass = contentClass
        }
    }

    /// Result of a full-package reverse with the upgrade rule applied.
    public struct Result {
        /// Rebuild script ops: carryPart for raw-channel parts + typed ops
        /// for upgraded parts.
        public let log: OperationLog
        /// Parts rebuilt through the typed DSL channel (byte-equal proven).
        public let dslParts: Set<String>
        /// Why raw-channel parts stayed raw, keyed by part path — the
        /// content-class attribution the coverage report surfaces
        /// (`ooxml-script-transcode`, «DSL-form coverage measurement»:
        /// "records which content classes remain on the raw channel").
        /// Values are class tags like "table", "hyperlink", "byte-mismatch".
        public let rawReasons: [String: String]
        /// Located form-gaps (empty for parts that upgraded). The measurement
        /// report / work queue for vocabulary extension.
        public let formGaps: [FormGap]
    }

    /// Full-package reverse: attempts the document.xml DSL upgrade, carries
    /// every other part (and any failed upgrade) verbatim on the raw channel.
    /// Stage B byte equality holds by construction.
    public static func reverse(parts: [String: Data]) throws -> Result {
        var log = OperationLog()
        var dslParts: Set<String> = []
        var rawReasons: [String: String] = [:]
        var formGaps: [FormGap] = []

        var documentOps: [Operation]? = nil
        if let documentXML = parts["word/document.xml"] {
            switch documentUpgrade(documentXML: documentXML) {
            case .upgraded(let ops):
                documentOps = ops
            case .raw(let reason, let xmlPath):
                rawReasons["word/document.xml"] = reason
                formGaps.append(FormGap(partPath: "word/document.xml",
                                        xmlPath: xmlPath, contentClass: reason))
            }
        }

        // Raw channel first (sorted for deterministic scripts), skipping any
        // part that upgraded.
        for (path, bytes) in parts.sorted(by: { $0.key < $1.key }) {
            if path == "word/document.xml", documentOps != nil { continue }
            guard let xml = String(data: bytes, encoding: .utf8) else { continue }
            log.append(.carryPart(partPath: path, xml: xml), source: .swift)
            if path != "word/document.xml" {
                rawReasons[path] = "sibling-part"
            }
        }
        if let ops = documentOps {
            for op in ops {
                log.append(op, source: .swift)
            }
            dslParts.insert("word/document.xml")
        }
        return Result(log: log, dslParts: dslParts, rawReasons: rawReasons, formGaps: formGaps)
    }

    // MARK: - document.xml typed extraction (trial + byte-compare)

    enum Upgrade {
        case upgraded([Operation])
        case raw(reason: String, xmlPath: String)
    }

    /// Attempts typed extraction of a document.xml. `.upgraded` carries the
    /// ops when the trial rebuild is byte-equal to the source; `.raw` carries
    /// the content-class tag and located XML path that blocked the upgrade.
    static func documentUpgrade(documentXML: Data) -> Upgrade {
        let ops: [Operation]
        do {
            ops = try extractDocumentOps(documentXML: documentXML)
        } catch let bail as Unsupported {
            return .raw(reason: bail.contentClass, xmlPath: bail.xmlPath)
        } catch {
            return .raw(reason: "parse-error", xmlPath: "w:document")
        }
        // Trial rebuild through the same reducer + serializer the script
        // execution path uses — the byte comparison is the upgrade gate.
        var doc = WordDocument.emptyAuthoringDocument()
        guard (try? doc.apply(operations: ops)) != nil,
              let tree = doc.xmlTrees["word/document.xml"],
              let rebuilt = try? XmlTreeWriter.serialize(tree) else {
            return .raw(reason: "rebuild-error", xmlPath: "w:document")
        }
        if rebuilt != documentXML {
            return .raw(reason: "byte-mismatch",
                        xmlPath: byteMismatchLocator(source: documentXML, rebuilt: rebuilt))
        }
        return .upgraded(ops)
    }

    /// Locates the first byte divergence between source and rebuilt, with a
    /// small context window each side so the differing form is identifiable.
    static func byteMismatchLocator(source: Data, rebuilt: Data) -> String {
        let a = [UInt8](source)
        let b = [UInt8](rebuilt)
        let n = min(a.count, b.count)
        var i = 0
        while i < n, a[i] == b[i] { i += 1 }
        func window(_ bytes: [UInt8], _ at: Int) -> String {
            let lo = max(0, at - 24)
            let hi = min(bytes.count, at + 24)
            return String(decoding: bytes[lo..<hi], as: UTF8.self)
        }
        return "byte@\(i) src=…\(window(a, i))… out=…\(window(b, i))…"
    }

    /// The root attribute set `emptyAuthoringDocument` emits — a source root
    /// matching this needs no `setDocumentRoot` op (word-canonical-forms 2.1).
    static let authoringDefaultRootAttributes: [RootAttribute] = [
        RootAttribute(prefix: "xmlns", localName: "w",
                      value: "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
        RootAttribute(prefix: "xmlns", localName: "w14",
                      value: "http://schemas.microsoft.com/office/word/2010/wordml"),
    ]

    /// Extraction bail-out carrying the content class + located XML path
    /// (word-canonical-forms task 1.1). Throw sites thread a `path` breadcrumb
    /// so the report names exactly where the unsupported form sits.
    private struct Unsupported: Error {
        let contentClass: String
        let xmlPath: String
        init(_ contentClass: String = "unsupported-form", _ xmlPath: String = "w:document") {
            self.contentClass = contentClass
            self.xmlPath = xmlPath
        }
    }

    /// Walks the document tree and derives typed ops. Throws `Unsupported`
    /// (with a located path) on any shape outside the reducer's vocabulary.
    private static func extractDocumentOps(documentXML: Data) throws -> [Operation] {
        let tree = try XmlTreeReader.parse(documentXML)
        let root = tree.root
        guard root.localName == "document" else {
            throw Unsupported("root-element", "w:\(root.localName)")
        }
        let elementChildren = try elementsOnly(root, path: "w:document")
        guard elementChildren.count == 1, elementChildren[0].localName == "body" else {
            throw Unsupported("root-shape", "w:document")
        }
        let body = elementChildren[0]

        var ops: [Operation] = []
        // word-canonical-forms task 2.1: when the root's attribute cloud
        // differs from the authoring default (minimal w + w14), emit a
        // setDocumentRoot first so the rebuild reproduces the namespace set.
        let rootAttrs = root.attributes.map {
            RootAttribute(prefix: $0.prefix, localName: $0.localName, value: $0.value)
        }
        if rootAttrs != Self.authoringDefaultRootAttributes {
            ops.append(.setDocumentRoot(attributes: rootAttrs))
        }
        let bodyChildren = try elementsOnly(body, path: "w:document/w:body")
        for (index, child) in bodyChildren.enumerated() {
            let cpath = "w:document/w:body/w:\(child.localName)[\(index)]"
            switch child.localName {
            case "p":
                ops.append(contentsOf: try extractParagraph(child, path: cpath))
            case "tbl":
                ops.append(.appendTable(in: nil, table: try extractTable(child, path: cpath)))
            case "sectPr":
                // Trailing body sectPr only (the reducer stamps it last).
                guard index == bodyChildren.count - 1 else {
                    throw Unsupported("sectPr-position", cpath)
                }
                ops.append(.setSectionProperties(at: nil, section: try extractSectPr(child, path: cpath)))
            default:
                throw Unsupported(child.localName, cpath)
            }
        }
        return ops
    }

    // MARK: Table extraction (task 2.5)

    /// Recognizes exactly `makeTable`'s canonical shape: `<w:tblGrid>` of
    /// bare gridCols, then rows of cells each holding a single plain
    /// paragraph. Any richer table form (tblPr, merged cells, styled cell
    /// paragraphs) bails with the "table" class tag → raw channel.
    private static func extractTable(_ tbl: XmlNode, path: String) throws -> TablePayload {
        guard tbl.attributes.isEmpty else { throw Unsupported("table", "\(path)/@attrs") }
        let children = try elementsOnly(tbl, path: path)
        guard let grid = children.first, grid.localName == "tblGrid" else {
            throw Unsupported("table", "\(path)/w:tblGrid")
        }
        guard grid.attributes.isEmpty else { throw Unsupported("table", "\(path)/w:tblGrid/@attrs") }
        let gridCols = try elementsOnly(grid, path: "\(path)/w:tblGrid")
        let columns = gridCols.count
        guard columns > 0, gridCols.allSatisfy({
            $0.localName == "gridCol" && $0.attributes.isEmpty && $0.children.isEmpty
        }) else {
            throw Unsupported("table", "\(path)/w:tblGrid/w:gridCol")
        }

        var cells: [[String]] = []
        for (ri, tr) in children.dropFirst().enumerated() {
            let rpath = "\(path)/w:tr[\(ri)]"
            guard tr.localName == "tr", tr.attributes.isEmpty else { throw Unsupported("table", rpath) }
            let tcs = try elementsOnly(tr, path: rpath)
            guard tcs.count == columns else { throw Unsupported("table", "\(rpath)/w:tc") }
            var row: [String] = []
            for (ci, tc) in tcs.enumerated() {
                let cpath = "\(rpath)/w:tc[\(ci)]"
                guard tc.localName == "tc", tc.attributes.isEmpty else { throw Unsupported("table", cpath) }
                let tcChildren = try elementsOnly(tc, path: cpath)
                guard tcChildren.count == 1, tcChildren[0].localName == "p" else {
                    throw Unsupported("table", "\(cpath)/w:p")
                }
                row.append(try plainCellText(tcChildren[0], path: "\(cpath)/w:p"))
            }
            cells.append(row)
        }
        guard !cells.isEmpty else { throw Unsupported("table", "\(path)/w:tr") }
        return TablePayload(rows: cells.count, columns: columns, cells: cells)
    }

    /// Cell paragraph in `makeParagraph(text:)`'s exact shape:
    /// `<w:p><w:r><w:t>text</w:t></w:r></w:p>` — no attributes, no pPr, no rPr.
    private static func plainCellText(_ p: XmlNode, path: String) throws -> String {
        guard p.attributes.isEmpty else { throw Unsupported("table", "\(path)/@attrs") }
        let children = try elementsOnly(p, path: path)
        guard children.count == 1, children[0].localName == "r",
              children[0].attributes.isEmpty else {
            throw Unsupported("table", "\(path)/w:r")
        }
        let rChildren = try elementsOnly(children[0], path: "\(path)/w:r")
        guard rChildren.count == 1, rChildren[0].localName == "t",
              rChildren[0].attributes.isEmpty else {
            throw Unsupported("table", "\(path)/w:r/w:t")
        }
        return rChildren[0].children
            .compactMap { $0.kind == .text ? $0.textContent : nil }
            .joined()
    }

    /// Element children of a node; throws on ANY interleaved non-element
    /// content (text, comments, PIs). The reducer's stamping emits compact
    /// element-only structure, so even whitespace-only text between elements
    /// is a foreign form (pretty-printed source) — bail before the trial.
    private static func elementsOnly(_ node: XmlNode, path: String) throws -> [XmlNode] {
        var out: [XmlNode] = []
        for child in node.children {
            guard child.kind == .element else {
                throw Unsupported("non-element-content", "\(path)/#\(child.kind)")
            }
            out.append(child)
        }
        return out
    }

    // MARK: Paragraph extraction (tasks 2.2 + 2.3)

    private static func extractParagraph(_ p: XmlNode, path: String) throws -> [Operation] {
        // Attributes: w14:paraId required (the stable setRuns target across
        // script round-trips) + Word-authored w14:textId / w:rsid* (task 2.2).
        // Any other attribute bails, named for the report.
        var payload = ParagraphPayload(text: "")
        var paraIdValue: String? = nil
        for attr in p.attributes {
            switch (attr.prefix, attr.localName) {
            case ("w14", "paraId"): paraIdValue = attr.value
            case ("w14", "textId"): payload.textId = attr.value
            case ("w", "rsidR"): payload.rsidR = attr.value
            case ("w", "rsidRPr"): payload.rsidRPr = attr.value
            case ("w", "rsidRDefault"): payload.rsidRDefault = attr.value
            case ("w", "rsidP"): payload.rsidP = attr.value
            default:
                throw Unsupported("paragraph-attrs", "\(path)/\(attrToken(attr))")
            }
        }
        guard let paraId = paraIdValue, !paraId.isEmpty else {
            throw Unsupported("paragraph-no-paraId", path)
        }
        payload.paraId = paraId
        var runs: [RunPayload] = []
        var sectionOp: Operation? = nil

        let children = try elementsOnly(p, path: path)
        var idx = 0
        if idx < children.count, children[idx].localName == "pPr" {
            let section = try extractPPr(children[idx], into: &payload, path: "\(path)/w:pPr")
            if let section {
                sectionOp = .setSectionProperties(
                    at: ElementID(rawString: "w14:paraId=\(paraId)"), section: section)
            }
            idx += 1
        }
        while idx < children.count {
            // Non-run inline content (hyperlink, bookmarkStart, …) tags the
            // bail with its element name for class attribution.
            guard children[idx].localName == "r" else {
                throw Unsupported(children[idx].localName, "\(path)/w:\(children[idx].localName)")
            }
            runs.append(try extractRun(children[idx], path: "\(path)/w:r[\(idx)]"))
            idx += 1
        }

        // Uniform emission: appendParagraph stamps pPr + a placeholder run;
        // setRuns then replaces the inline content (keeping pPr). A single
        // plain-text run matches makeParagraph's own shape, so the shorter
        // spelling is used when possible.
        var ops: [Operation] = []
        if runs.count == 1, runs[0] == RunPayload(text: runs[0].text) {
            payload.text = runs[0].text
            ops.append(.appendParagraph(in: nil, paragraph: payload))
        } else {
            payload.text = ""
            ops.append(.appendParagraph(in: nil, paragraph: payload))
            ops.append(.setRuns(target: ElementID(rawString: "w14:paraId=\(paraId)"), runs: runs))
        }
        if let sectionOp {
            ops.append(sectionOp)
        }
        return ops
    }

    /// Extracts pPr fields into the payload; returns a SectionPayload when a
    /// trailing mid-body `<w:sectPr>` is present (task 2.4). Order and
    /// vocabulary must be exactly what `makePPrChildren` stamps.
    private static func extractPPr(_ pPr: XmlNode,
                                   into payload: inout ParagraphPayload,
                                   path: String) throws -> SectionPayload? {
        guard pPr.attributes.isEmpty else { throw Unsupported("pPr-attrs", "\(path)/@attrs") }
        var section: SectionPayload? = nil
        for child in try elementsOnly(pPr, path: path) {
            guard section == nil else {
                throw Unsupported("pPr-after-sectPr", "\(path)/w:\(child.localName)")
            }
            let cpath = "\(path)/w:\(child.localName)"
            switch child.localName {
            case "pStyle":
                payload.styleId = try singleVal(child, path: cpath)
            case "numPr":
                for numChild in try elementsOnly(child, path: cpath) {
                    switch numChild.localName {
                    case "ilvl": payload.numLevel = try intVal(numChild, path: "\(cpath)/w:ilvl")
                    case "numId": payload.numId = try intVal(numChild, path: "\(cpath)/w:numId")
                    default: throw Unsupported("numPr-element", "\(cpath)/w:\(numChild.localName)")
                    }
                }
            case "spacing":
                for attr in child.attributes {
                    guard attr.prefix == "w" else { throw Unsupported("spacing-attr", "\(cpath)/\(attrToken(attr))") }
                    switch attr.localName {
                    case "before": payload.spacingBefore = try int(attr.value, path: cpath)
                    case "after": payload.spacingAfter = try int(attr.value, path: cpath)
                    case "line": payload.spacingLine = try int(attr.value, path: cpath)
                    case "lineRule": payload.spacingLineRule = attr.value
                    default: throw Unsupported("spacing-attr", "\(cpath)/\(attrToken(attr))")
                    }
                }
                guard child.children.isEmpty else { throw Unsupported("spacing-children", cpath) }
            case "ind":
                for attr in child.attributes {
                    guard attr.prefix == "w" else { throw Unsupported("ind-attr", "\(cpath)/\(attrToken(attr))") }
                    switch attr.localName {
                    case "left": payload.indentLeft = try int(attr.value, path: cpath)
                    case "right": payload.indentRight = try int(attr.value, path: cpath)
                    case "firstLine": payload.indentFirstLine = try int(attr.value, path: cpath)
                    case "hanging": payload.indentHanging = try int(attr.value, path: cpath)
                    default: throw Unsupported("ind-attr", "\(cpath)/\(attrToken(attr))")
                    }
                }
                guard child.children.isEmpty else { throw Unsupported("ind-children", cpath) }
            case "jc":
                payload.alignment = try singleVal(child, path: cpath)
            case "sectPr":
                section = try extractSectPr(child, path: cpath)
            default:
                throw Unsupported("pPr-element", cpath)
            }
        }
        return section
    }

    private static func extractRun(_ r: XmlNode, path: String) throws -> RunPayload {
        var run = RunPayload(text: "")
        // Run-level rsids (task 2.2); any other run attribute bails.
        for attr in r.attributes {
            switch (attr.prefix, attr.localName) {
            case ("w", "rsidR"): run.rsidR = attr.value
            case ("w", "rsidRPr"): run.rsidRPr = attr.value
            default: throw Unsupported("run-attrs", "\(path)/\(attrToken(attr))")
            }
        }
        let children = try elementsOnly(r, path: path)
        var idx = 0
        if idx < children.count, children[idx].localName == "rPr" {
            try extractRPr(children[idx], into: &run, path: "\(path)/w:rPr")
            idx += 1
        }
        guard idx < children.count, children[idx].localName == "t",
              idx == children.count - 1 else {
            let where_ = idx < children.count ? "w:\(children[idx].localName)" : "w:t(missing)"
            throw Unsupported("run-body", "\(path)/\(where_)")
        }
        let t = children[idx]
        // xml:space="preserve" (task 2.3); any other <w:t> attribute bails.
        for attr in t.attributes {
            if attr.prefix == "xml", attr.localName == "space", attr.value == "preserve" {
                run.preserveSpace = true
            } else {
                throw Unsupported("t-attrs", "\(path)/w:t/\(attrToken(attr))")
            }
        }
        run.text = t.children
            .compactMap { $0.kind == .text ? $0.textContent : nil }
            .joined()
        return run
    }

    /// rPr vocabulary and order must be exactly what setRuns stamps:
    /// rFonts, b, i, color, sz, u, vertAlign.
    private static func extractRPr(_ rPr: XmlNode, into run: inout RunPayload, path: String) throws {
        guard rPr.attributes.isEmpty else { throw Unsupported("rPr-attrs", "\(path)/@attrs") }
        for child in try elementsOnly(rPr, path: path) {
            let cpath = "\(path)/w:\(child.localName)"
            switch child.localName {
            case "rFonts":
                for attr in child.attributes {
                    guard attr.prefix == "w" else { throw Unsupported("rFonts-attr", "\(cpath)/\(attrToken(attr))") }
                    switch attr.localName {
                    case "ascii": run.fontAscii = attr.value
                    case "eastAsia": run.fontEastAsia = attr.value
                    default: throw Unsupported("rFonts-attr", "\(cpath)/\(attrToken(attr))")
                    }
                }
                guard child.children.isEmpty else { throw Unsupported("rFonts-children", cpath) }
            case "b":
                guard child.attributes.isEmpty, child.children.isEmpty else { throw Unsupported("b-shape", cpath) }
                run.bold = true
            case "i":
                guard child.attributes.isEmpty, child.children.isEmpty else { throw Unsupported("i-shape", cpath) }
                run.italic = true
            case "color":
                run.color = try singleVal(child, path: cpath)
            case "sz":
                run.sizeHalfPoints = try intVal(child, path: cpath)
            case "u":
                run.underline = try singleVal(child, path: cpath)
            case "vertAlign":
                run.vertAlign = try singleVal(child, path: cpath)
            default:
                throw Unsupported("rPr-element", cpath)
            }
        }
    }

    // MARK: Section extraction (task 2.4)

    /// sectPr vocabulary and order must be exactly what makeSectPr stamps:
    /// headerReference*, footerReference*, pgSz, pgMar, cols.
    private static func extractSectPr(_ sectPr: XmlNode, path: String) throws -> SectionPayload {
        var section = SectionPayload()
        // sectPr element rsids (task 2.2); any other attribute bails.
        for attr in sectPr.attributes {
            switch (attr.prefix, attr.localName) {
            case ("w", "rsidR"): section.rsidR = attr.value
            case ("w", "rsidSect"): section.rsidSect = attr.value
            default: throw Unsupported("sectPr-attrs", "\(path)/\(attrToken(attr))")
            }
        }
        for child in try elementsOnly(sectPr, path: path) {
            let cpath = "\(path)/w:\(child.localName)"
            switch child.localName {
            case "headerReference", "footerReference":
                var type: String? = nil
                var rId: String? = nil
                for attr in child.attributes {
                    if attr.prefix == "w", attr.localName == "type" { type = attr.value }
                    else if attr.prefix == "r", attr.localName == "id" { rId = attr.value }
                    else { throw Unsupported("hf-ref-attr", "\(cpath)/\(attrToken(attr))") }
                }
                guard let t = type, let id = rId, child.children.isEmpty else {
                    throw Unsupported("hf-ref-shape", cpath)
                }
                let ref = HeaderFooterReference(type: t, relationshipId: id)
                if child.localName == "headerReference" {
                    section.headerReferences = (section.headerReferences ?? []) + [ref]
                } else {
                    section.footerReferences = (section.footerReferences ?? []) + [ref]
                }
            case "pgSz":
                for attr in child.attributes {
                    guard attr.prefix == "w" else { throw Unsupported("pgSz-attr", "\(cpath)/\(attrToken(attr))") }
                    switch attr.localName {
                    case "w": section.pageWidth = try int(attr.value, path: cpath)
                    case "h": section.pageHeight = try int(attr.value, path: cpath)
                    case "orient": section.orientation = attr.value
                    default: throw Unsupported("pgSz-attr", "\(cpath)/\(attrToken(attr))")
                    }
                }
                guard child.children.isEmpty else { throw Unsupported("pgSz-children", cpath) }
            case "pgMar":
                for attr in child.attributes {
                    guard attr.prefix == "w" else { throw Unsupported("pgMar-attr", "\(cpath)/\(attrToken(attr))") }
                    switch attr.localName {
                    case "top": section.marginTop = try int(attr.value, path: cpath)
                    case "right": section.marginRight = try int(attr.value, path: cpath)
                    case "bottom": section.marginBottom = try int(attr.value, path: cpath)
                    case "left": section.marginLeft = try int(attr.value, path: cpath)
                    case "header": section.marginHeader = try int(attr.value, path: cpath)
                    case "footer": section.marginFooter = try int(attr.value, path: cpath)
                    case "gutter": section.marginGutter = try int(attr.value, path: cpath)
                    default: throw Unsupported("pgMar-attr", "\(cpath)/\(attrToken(attr))")
                    }
                }
                guard child.children.isEmpty else { throw Unsupported("pgMar-children", cpath) }
            case "cols":
                for attr in child.attributes {
                    guard attr.prefix == "w" else { throw Unsupported("cols-attr", "\(cpath)/\(attrToken(attr))") }
                    switch attr.localName {
                    case "num": section.columnCount = try int(attr.value, path: cpath)
                    case "space": section.columnSpace = try int(attr.value, path: cpath)
                    default: throw Unsupported("cols-attr", "\(cpath)/\(attrToken(attr))")
                    }
                }
                guard child.children.isEmpty else { throw Unsupported("cols-children", cpath) }
            case "docGrid":
                for attr in child.attributes {
                    guard attr.prefix == "w" else { throw Unsupported("docGrid-attr", "\(cpath)/\(attrToken(attr))") }
                    switch attr.localName {
                    case "type": section.docGridType = attr.value
                    case "linePitch": section.docGridLinePitch = try int(attr.value, path: cpath)
                    default: throw Unsupported("docGrid-attr", "\(cpath)/\(attrToken(attr))")
                    }
                }
                guard child.children.isEmpty else { throw Unsupported("docGrid-children", cpath) }
            default:
                throw Unsupported("sectPr-element", cpath)
            }
        }
        return section
    }

    // MARK: Small helpers

    /// Renders an attribute as a `@prefix:localName` token for gap paths.
    private static func attrToken(_ attr: XmlAttribute?) -> String {
        guard let attr else { return "@attrs" }
        let prefix = attr.prefix.map { "\($0):" } ?? ""
        return "@\(prefix)\(attr.localName)"
    }

    /// The sole `w:val` attribute of an empty element.
    private static func singleVal(_ node: XmlNode, path: String) throws -> String {
        guard node.children.isEmpty,
              node.attributes.count == 1,
              let attr = node.attributes.first,
              attr.prefix == "w", attr.localName == "val" else {
            throw Unsupported("val-attr-shape", path)
        }
        return attr.value
    }

    private static func intVal(_ node: XmlNode, path: String) throws -> Int {
        try int(try singleVal(node, path: path), path: path)
    }

    private static func int(_ s: String, path: String) throws -> Int {
        guard let v = Int(s) else { throw Unsupported("non-integer", path) }
        return v
    }
}
