// ReverseExtractor.swift
// format-alignment-engine Phase B tasks 2.2/2.3 — typed reverse extraction
// with the byte-equal upgrade rule (`ooxml-script-transcode`, «Reverse
// extraction covers the five format layers»; Decision 3).
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
// bails to raw. The final authority is always the byte comparison.

import Foundation

public enum ReverseExtractor {

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
    }

    /// Full-package reverse: attempts the document.xml DSL upgrade, carries
    /// every other part (and any failed upgrade) verbatim on the raw channel.
    /// Stage B byte equality holds by construction.
    public static func reverse(parts: [String: Data]) throws -> Result {
        var log = OperationLog()
        var dslParts: Set<String> = []
        var rawReasons: [String: String] = [:]

        var documentOps: [Operation]? = nil
        if let documentXML = parts["word/document.xml"] {
            switch documentUpgrade(documentXML: documentXML) {
            case .upgraded(let ops):
                documentOps = ops
            case .raw(let reason):
                rawReasons["word/document.xml"] = reason
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
        return Result(log: log, dslParts: dslParts, rawReasons: rawReasons)
    }

    // MARK: - document.xml typed extraction (trial + byte-compare)

    enum Upgrade {
        case upgraded([Operation])
        case raw(reason: String)
    }

    /// Attempts typed extraction of a document.xml. `.upgraded` carries the
    /// ops when the trial rebuild is byte-equal to the source; `.raw` carries
    /// the content-class tag that blocked the upgrade.
    static func documentUpgrade(documentXML: Data) -> Upgrade {
        let ops: [Operation]
        do {
            ops = try extractDocumentOps(documentXML: documentXML)
        } catch let bail as Unsupported {
            return .raw(reason: bail.contentClass)
        } catch {
            return .raw(reason: "parse-error")
        }
        // Trial rebuild through the same reducer + serializer the script
        // execution path uses — the byte comparison is the upgrade gate.
        var doc = WordDocument.emptyAuthoringDocument()
        guard (try? doc.apply(operations: ops)) != nil,
              let tree = doc.xmlTrees["word/document.xml"],
              let rebuilt = try? XmlTreeWriter.serialize(tree),
              rebuilt == documentXML else {
            return .raw(reason: "byte-mismatch")
        }
        return .upgraded(ops)
    }

    /// Extraction bail-out carrying the content class that cannot ride the
    /// typed channel (task 2.5 "class attribution").
    private struct Unsupported: Error {
        let contentClass: String
        init(_ contentClass: String = "unsupported-form") {
            self.contentClass = contentClass
        }
    }

    /// Walks the document tree and derives typed ops. Throws `Unsupported`
    /// on any shape outside the reducer's stamping vocabulary.
    private static func extractDocumentOps(documentXML: Data) throws -> [Operation] {
        let tree = try XmlTreeReader.parse(documentXML)
        let root = tree.root
        guard root.localName == "document" else { throw Unsupported() }
        let elementChildren = try elementsOnly(root)
        guard elementChildren.count == 1, elementChildren[0].localName == "body" else {
            throw Unsupported()
        }
        let body = elementChildren[0]

        var ops: [Operation] = []
        let bodyChildren = try elementsOnly(body)
        for (index, child) in bodyChildren.enumerated() {
            switch child.localName {
            case "p":
                ops.append(contentsOf: try extractParagraph(child))
            case "tbl":
                ops.append(.appendTable(in: nil, table: try extractTable(child)))
            case "sectPr":
                // Trailing body sectPr only (the reducer stamps it last).
                guard index == bodyChildren.count - 1 else { throw Unsupported("sectPr-position") }
                ops.append(.setSectionProperties(at: nil, section: try extractSectPr(child)))
            default:
                throw Unsupported(child.localName)
            }
        }
        return ops
    }

    // MARK: Table extraction (task 2.5)

    /// Recognizes exactly `makeTable`'s canonical shape: `<w:tblGrid>` of
    /// bare gridCols, then rows of cells each holding a single plain
    /// paragraph. Any richer table form (tblPr, merged cells, styled cell
    /// paragraphs) bails with the "table" class tag → raw channel.
    private static func extractTable(_ tbl: XmlNode) throws -> TablePayload {
        guard tbl.attributes.isEmpty else { throw Unsupported("table") }
        let children = try elementsOnly(tbl)
        guard let grid = children.first, grid.localName == "tblGrid" else {
            throw Unsupported("table")
        }
        guard grid.attributes.isEmpty else { throw Unsupported("table") }
        let gridCols = try elementsOnly(grid)
        let columns = gridCols.count
        guard columns > 0, gridCols.allSatisfy({
            $0.localName == "gridCol" && $0.attributes.isEmpty && $0.children.isEmpty
        }) else {
            throw Unsupported("table")
        }

        var cells: [[String]] = []
        for tr in children.dropFirst() {
            guard tr.localName == "tr", tr.attributes.isEmpty else { throw Unsupported("table") }
            let tcs = try elementsOnly(tr)
            guard tcs.count == columns else { throw Unsupported("table") }
            var row: [String] = []
            for tc in tcs {
                guard tc.localName == "tc", tc.attributes.isEmpty else { throw Unsupported("table") }
                let tcChildren = try elementsOnly(tc)
                guard tcChildren.count == 1, tcChildren[0].localName == "p" else {
                    throw Unsupported("table")
                }
                row.append(try plainCellText(tcChildren[0]))
            }
            cells.append(row)
        }
        guard !cells.isEmpty else { throw Unsupported("table") }
        return TablePayload(rows: cells.count, columns: columns, cells: cells)
    }

    /// Cell paragraph in `makeParagraph(text:)`'s exact shape:
    /// `<w:p><w:r><w:t>text</w:t></w:r></w:p>` — no attributes, no pPr, no rPr.
    private static func plainCellText(_ p: XmlNode) throws -> String {
        guard p.attributes.isEmpty else { throw Unsupported("table") }
        let children = try elementsOnly(p)
        guard children.count == 1, children[0].localName == "r",
              children[0].attributes.isEmpty else {
            throw Unsupported("table")
        }
        let rChildren = try elementsOnly(children[0])
        guard rChildren.count == 1, rChildren[0].localName == "t",
              rChildren[0].attributes.isEmpty else {
            throw Unsupported("table")
        }
        return rChildren[0].children
            .compactMap { $0.kind == .text ? $0.textContent : nil }
            .joined()
    }

    /// Element children of a node; throws on ANY interleaved non-element
    /// content (text, comments, PIs). The reducer's stamping emits compact
    /// element-only structure, so even whitespace-only text between elements
    /// is a foreign form (pretty-printed source) — bail before the trial.
    private static func elementsOnly(_ node: XmlNode) throws -> [XmlNode] {
        var out: [XmlNode] = []
        for child in node.children {
            guard child.kind == .element else { throw Unsupported() }
            out.append(child)
        }
        return out
    }

    // MARK: Paragraph extraction (tasks 2.2 + 2.3)

    private static func extractParagraph(_ p: XmlNode) throws -> [Operation] {
        // Attributes: exactly w14:paraId (needed as the stable setRuns
        // target across script round-trips) and nothing else.
        guard p.attributes.count == 1,
              let paraIdAttr = p.attributes.first,
              paraIdAttr.prefix == "w14", paraIdAttr.localName == "paraId",
              !paraIdAttr.value.isEmpty else {
            throw Unsupported()
        }
        let paraId = paraIdAttr.value

        var payload = ParagraphPayload(text: "", paraId: paraId)
        var runs: [RunPayload] = []
        var sectionOp: Operation? = nil

        let children = try elementsOnly(p)
        var idx = 0
        if idx < children.count, children[idx].localName == "pPr" {
            let section = try extractPPr(children[idx], into: &payload)
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
                throw Unsupported(children[idx].localName)
            }
            runs.append(try extractRun(children[idx]))
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
                                   into payload: inout ParagraphPayload) throws -> SectionPayload? {
        guard pPr.attributes.isEmpty else { throw Unsupported() }
        var section: SectionPayload? = nil
        for child in try elementsOnly(pPr) {
            guard section == nil else { throw Unsupported() }  // sectPr must be last
            switch child.localName {
            case "pStyle":
                payload.styleId = try singleVal(child)
            case "numPr":
                for numChild in try elementsOnly(child) {
                    switch numChild.localName {
                    case "ilvl": payload.numLevel = try intVal(numChild)
                    case "numId": payload.numId = try intVal(numChild)
                    default: throw Unsupported()
                    }
                }
            case "spacing":
                for attr in child.attributes {
                    guard attr.prefix == "w" else { throw Unsupported() }
                    switch attr.localName {
                    case "before": payload.spacingBefore = try int(attr.value)
                    case "after": payload.spacingAfter = try int(attr.value)
                    case "line": payload.spacingLine = try int(attr.value)
                    case "lineRule": payload.spacingLineRule = attr.value
                    default: throw Unsupported()
                    }
                }
                guard child.children.isEmpty else { throw Unsupported() }
            case "ind":
                for attr in child.attributes {
                    guard attr.prefix == "w" else { throw Unsupported() }
                    switch attr.localName {
                    case "left": payload.indentLeft = try int(attr.value)
                    case "right": payload.indentRight = try int(attr.value)
                    case "firstLine": payload.indentFirstLine = try int(attr.value)
                    case "hanging": payload.indentHanging = try int(attr.value)
                    default: throw Unsupported()
                    }
                }
                guard child.children.isEmpty else { throw Unsupported() }
            case "jc":
                payload.alignment = try singleVal(child)
            case "sectPr":
                section = try extractSectPr(child)
            default:
                throw Unsupported()
            }
        }
        return section
    }

    private static func extractRun(_ r: XmlNode) throws -> RunPayload {
        guard r.attributes.isEmpty else { throw Unsupported() }
        var run = RunPayload(text: "")
        let children = try elementsOnly(r)
        var idx = 0
        if idx < children.count, children[idx].localName == "rPr" {
            try extractRPr(children[idx], into: &run)
            idx += 1
        }
        guard idx < children.count, children[idx].localName == "t",
              children[idx].attributes.isEmpty,
              idx == children.count - 1 else {
            throw Unsupported()
        }
        run.text = children[idx].children
            .compactMap { $0.kind == .text ? $0.textContent : nil }
            .joined()
        return run
    }

    /// rPr vocabulary and order must be exactly what setRuns stamps:
    /// rFonts, b, i, color, sz, u, vertAlign.
    private static func extractRPr(_ rPr: XmlNode, into run: inout RunPayload) throws {
        guard rPr.attributes.isEmpty else { throw Unsupported() }
        for child in try elementsOnly(rPr) {
            switch child.localName {
            case "rFonts":
                for attr in child.attributes {
                    guard attr.prefix == "w" else { throw Unsupported() }
                    switch attr.localName {
                    case "ascii": run.fontAscii = attr.value
                    case "eastAsia": run.fontEastAsia = attr.value
                    default: throw Unsupported()
                    }
                }
                guard child.children.isEmpty else { throw Unsupported() }
            case "b":
                guard child.attributes.isEmpty, child.children.isEmpty else { throw Unsupported() }
                run.bold = true
            case "i":
                guard child.attributes.isEmpty, child.children.isEmpty else { throw Unsupported() }
                run.italic = true
            case "color":
                run.color = try singleVal(child)
            case "sz":
                run.sizeHalfPoints = try intVal(child)
            case "u":
                run.underline = try singleVal(child)
            case "vertAlign":
                run.vertAlign = try singleVal(child)
            default:
                throw Unsupported()
            }
        }
    }

    // MARK: Section extraction (task 2.4)

    /// sectPr vocabulary and order must be exactly what makeSectPr stamps:
    /// headerReference*, footerReference*, pgSz, pgMar, cols.
    private static func extractSectPr(_ sectPr: XmlNode) throws -> SectionPayload {
        guard sectPr.attributes.isEmpty else { throw Unsupported() }
        var section = SectionPayload()
        for child in try elementsOnly(sectPr) {
            switch child.localName {
            case "headerReference", "footerReference":
                var type: String? = nil
                var rId: String? = nil
                for attr in child.attributes {
                    if attr.prefix == "w", attr.localName == "type" { type = attr.value }
                    else if attr.prefix == "r", attr.localName == "id" { rId = attr.value }
                    else { throw Unsupported() }
                }
                guard let t = type, let id = rId, child.children.isEmpty else { throw Unsupported() }
                let ref = HeaderFooterReference(type: t, relationshipId: id)
                if child.localName == "headerReference" {
                    section.headerReferences = (section.headerReferences ?? []) + [ref]
                } else {
                    section.footerReferences = (section.footerReferences ?? []) + [ref]
                }
            case "pgSz":
                for attr in child.attributes {
                    guard attr.prefix == "w" else { throw Unsupported() }
                    switch attr.localName {
                    case "w": section.pageWidth = try int(attr.value)
                    case "h": section.pageHeight = try int(attr.value)
                    case "orient": section.orientation = attr.value
                    default: throw Unsupported()
                    }
                }
                guard child.children.isEmpty else { throw Unsupported() }
            case "pgMar":
                for attr in child.attributes {
                    guard attr.prefix == "w" else { throw Unsupported() }
                    switch attr.localName {
                    case "top": section.marginTop = try int(attr.value)
                    case "right": section.marginRight = try int(attr.value)
                    case "bottom": section.marginBottom = try int(attr.value)
                    case "left": section.marginLeft = try int(attr.value)
                    case "header": section.marginHeader = try int(attr.value)
                    case "footer": section.marginFooter = try int(attr.value)
                    case "gutter": section.marginGutter = try int(attr.value)
                    default: throw Unsupported()
                    }
                }
                guard child.children.isEmpty else { throw Unsupported() }
            case "cols":
                for attr in child.attributes {
                    guard attr.prefix == "w" else { throw Unsupported() }
                    switch attr.localName {
                    case "num": section.columnCount = try int(attr.value)
                    case "space": section.columnSpace = try int(attr.value)
                    default: throw Unsupported()
                    }
                }
                guard child.children.isEmpty else { throw Unsupported() }
            default:
                throw Unsupported()
            }
        }
        return section
    }

    // MARK: Small helpers

    /// The sole `w:val` attribute of an empty element.
    private static func singleVal(_ node: XmlNode) throws -> String {
        guard node.children.isEmpty,
              node.attributes.count == 1,
              let attr = node.attributes.first,
              attr.prefix == "w", attr.localName == "val" else {
            throw Unsupported()
        }
        return attr.value
    }

    private static func intVal(_ node: XmlNode) throws -> Int {
        try int(try singleVal(node))
    }

    private static func int(_ s: String) throws -> Int {
        guard let v = Int(s) else { throw Unsupported() }
        return v
    }
}
