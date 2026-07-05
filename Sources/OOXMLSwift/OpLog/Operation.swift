// Operation taxonomy for the word-aligned-state-sync op log (Phase 2a).
//
// Spectra change: operation-log-scaffold-impl, target ooxml-swift v0.31.3.
// Capability: ooxml-operation-log
//
// `Operation` enumerates every mutation kind the Phase 2 op log persists. Phase
// 2b (`operation-reducer-impl`) consumes these cases via exhaustive switch in
// the reducer; Phase 2c (`operation-log-setter-wiring-impl`) emits these cases
// from typed-view setters. The `unknown` fallback case carries any op_type the
// local code does not declare so the JSONL log round-trips byte-equal across
// version skews.

import Foundation

// MARK: - Operation taxonomy

/// A single mutation in the OOXML op log.
///
/// 21 cases total: 16 element-level + 4 tree-node-level fallback + 1 unknown
/// forward-compat fallback. See `openspec/specs/ooxml-operation-log/spec.md`
/// requirement "Operation taxonomy covers full OOXML mutation surface" for the
/// authoritative enumeration and meanings.
public enum Operation: Equatable, Sendable {

    // MARK: Element-level operations (typed)

    case insertParagraphAfter(after: ElementID, paragraph: ParagraphPayload)
    case insertParagraphBefore(before: ElementID, paragraph: ParagraphPayload)
    case removeParagraph(id: ElementID)
    case setText(target: ElementID, text: String)
    case setParagraphStyle(target: ElementID, styleId: String?)
    case insertTable(at: ElementID, table: TablePayload)
    case removeTable(id: ElementID)
    case setCellText(table: ElementID, row: Int, column: Int, text: String)
    case insertRun(in: ElementID, position: Int, run: RunPayload)
    case setRunFormat(target: ElementID, format: RunFormatPayload)
    case insertBookmark(at: ElementID, bookmarkId: Int, name: String)
    case insertComment(anchor: ElementID, commentId: Int, text: String, author: String)
    case undo(targetOpID: UUID)
    case redo(targetOpID: UUID)
    case batchBegin(label: String?)
    case batchEnd

    // MARK: Authoring operations (word-aligned-state-sync §4b, macdoc#128)
    // Additive; names + payload fields mirror ECMA-376 WordprocessingML per
    // the `ooxml-operation-log` requirement "Authoring operations extend the
    // taxonomy additively with OOXML-mirror naming".

    /// Appends a paragraph as the last block-level child of the container
    /// addressed by `in` (`nil` = the document body). Construction-order
    /// authoring anchor: the first paragraph in an empty body has no sibling
    /// to anchor on; subsequent inserts use `insertParagraphAfter`.
    case appendParagraph(in: ElementID?, paragraph: ParagraphPayload)

    /// Replaces the addressed paragraph's inline content with the given runs
    /// (`<w:pPr>` preserved). Formatting fields live on `RunPayload`
    /// (bold ↔ `<w:b>`, italic ↔ `<w:i>`, color ↔ `<w:color w:val>`).
    case setRuns(target: ElementID, runs: [RunPayload])

    /// Registers a style definition in `word/styles.xml` (define-on-first-use;
    /// replay of a duplicate `styleId` is an idempotent no-op).
    case defineStyle(payload: StylePayload)

    /// Component envelope markers — documented exception to OOXML-mirror
    /// naming: op-log metadata only (batch-marker pattern), zero OOXML output.
    case beginComponent(type: String, id: ElementID)
    case endComponent(id: ElementID)

    /// Inline atoms (`<w:tab/>` / `<w:br/>` / `<w:noBreakHyphen/>`), appended
    /// in construction order. `in:` addresses a run; a paragraph target makes
    /// the reducer synthesize a wrapping `<w:r>` (atoms are schema-invalid
    /// outside runs). Bare `<w:br/>` is the text-wrapping break; page/column
    /// variants await a future additive `type:` parameter.
    case insertTab(in: ElementID)
    case insertBreak(in: ElementID)
    case insertNoBreakHyphen(in: ElementID)

    // MARK: Tree-node-level fallback operations (typed)

    case insertNode(parent: ElementID, position: Int, nodeXML: String)
    case removeNode(target: ElementID)
    case updateAttribute(target: ElementID, prefix: String?, localName: String, value: String?)
    case moveNode(source: ElementID, destinationParent: ElementID, destinationIndex: Int)

    /// Inserts a new XML node as the next sibling of `after` (i.e., into
    /// after's parent at after's index + 1). Cleaner primitive than
    /// `insertNode` for sibling-relative insertion — caller doesn't need
    /// to resolve the parent and position; Reducer walks the tree to find
    /// after's parent and computes the index.
    ///
    /// Used by `OOXMLEdit.insertHyperlink` to insert the `<w:hyperlink>`
    /// wrapper after a target Run, and other "after X" sibling insertions
    /// where the caller has the sibling ID but not the parent ID.
    ///
    /// `nodeXML` is parsed at Reducer time (via XmlTreeReader on a
    /// fragment wrapped with namespace declarations).
    case insertSiblingAfter(after: ElementID, nodeXML: String)

    /// Wraps `target` element with a `<w:hyperlink r:id="rId">` wrapper.
    /// The Reducer:
    ///   1. Finds target's parent + index
    ///   2. Deep-clones target
    ///   3. Builds `<w:hyperlink r:id="rId">` element with the clone as child
    ///   4. Replaces parent.children[idx] with the wrapper
    ///
    /// Atomic single-op — avoids the placeholder substitution problem of
    /// `[insertNode + removeNode]` decomposition. Used by
    /// `OOXMLEdit.wrapWithHyperlink` and (downstream) by
    /// `WordEdit.applyLink` for Cmd-K parity.
    ///
    /// `rId` must match the Id in a paired `addRelationship` op (caller's
    /// responsibility) for referential integrity between document.xml
    /// and the rels part.
    case wrapWithHyperlink(target: ElementID, rId: String)

    // MARK: Rels-part operations (typed)

    /// Adds a `<Relationship Id="..." Type="..." Target="..." TargetMode="..."/>`
    /// entry to the specified rels part (e.g., "word/_rels/document.xml.rels").
    ///
    /// Rels-part xml has a rigid, well-known structure: a single
    /// `<Relationships>` root with `<Relationship>` children. Treating it
    /// as arbitrary XML mutation (via insertNode + updateAttribute) invites
    /// round-trip bugs because rels parts use a different namespace,
    /// attribute spelling, and validation rules than document.xml. This
    /// typed operation captures the rels-specific shape.
    ///
    /// Used by OOXMLEdit.insertHyperlink / wrapWithHyperlink composite
    /// emission to register the URL target on the rels side.
    ///
    /// `targetMode`: typically "External" for hyperlinks; nil omits the
    /// attribute. See macdoc#110 §5 design walkthrough Q3.
    case addRelationship(part: String, id: String, type: String, target: String, targetMode: String?)

    // MARK: Forward-compat fallback (preserves unrecognized op_type byte-equal)

    /// Carries any op_type the local code does not declare so the JSONL log
    /// round-trips byte-equal across version skews. Phase 2b reducer treats
    /// `.unknown` as opaque (logs a warning, passes through to next op).
    case unknown(opType: String, payload: JSONValue)
}

// MARK: - Payload value types

/// Minimal data needed by `insertParagraphAfter` / `insertParagraphBefore` to
/// reconstruct a paragraph in the reducer (Phase 2b). Carries text-only +
/// optional style; rich formatting goes through `setRunFormat` follow-up ops.
public struct ParagraphPayload: Equatable, Sendable, Codable {
    public var text: String
    public var styleId: String?
    /// §4b (#128): explicit stable id (↔ `w14:paraId`) carrying the mdocx
    /// DSL's mandatory identifier. When present the reducer stamps it on the
    /// created `<w:p>`; when absent the opID-derived libraryUUID behavior
    /// applies unchanged.
    public var paraId: String?

    public init(text: String, styleId: String? = nil, paraId: String? = nil) {
        self.text = text
        self.styleId = styleId
        self.paraId = paraId
    }
}

/// Minimal data needed by `insertTable` to reconstruct an empty grid in the
/// reducer. Cell content arrives via subsequent `setCellText` ops.
public struct TablePayload: Equatable, Sendable, Codable {
    public var rows: Int
    public var columns: Int

    public init(rows: Int, columns: Int) {
        self.rows = rows
        self.columns = columns
    }
}

/// Minimal data needed by `insertRun` to reconstruct a run inside a paragraph
/// in the reducer. Carries text-only; formatting goes through `setRunFormat`.
public struct RunPayload: Equatable, Sendable, Codable {
    public var text: String
    /// §4b (#128) `setRuns` formatting fields (bold ↔ `<w:b>`,
    /// italic ↔ `<w:i>` — spelled `italic` for cross-payload consistency
    /// with `RunFormatPayload.italic`; ECMA-376 titles the element
    /// "Italics" — color ↔ `<w:color w:val>`). Optional so pre-#128 JSONL
    /// decodes unchanged.
    public var bold: Bool?
    public var italic: Bool?
    public var color: String?

    public init(text: String, bold: Bool? = nil, italic: Bool? = nil, color: String? = nil) {
        self.text = text
        self.bold = bold
        self.italic = italic
        self.color = color
    }
}

/// Style definition carried by `defineStyle` (§4b, #128). `styleId` mirrors
/// `<w:style w:styleId>`; `fontSize` is in points.
public struct StylePayload: Equatable, Sendable, Codable {
    public var styleId: String
    public var name: String?
    public var font: String?
    public var fontSize: Int?
    public var color: String?
    public var bold: Bool?
    public var italic: Bool?

    public init(styleId: String, name: String? = nil, font: String? = nil,
                fontSize: Int? = nil, color: String? = nil,
                bold: Bool? = nil, italic: Bool? = nil) {
        self.styleId = styleId
        self.name = name
        self.font = font
        self.fontSize = fontSize
        self.color = color
        self.bold = bold
        self.italic = italic
    }
}

/// Run-level format descriptor for `setRunFormat`. Phase 2a covers the most
/// common boolean toggles + font-size; richer formatting expands additively
/// in Phase 2b/2c without breaking the JSONL contract (forward-compat).
public struct RunFormatPayload: Equatable, Sendable, Codable {
    public var bold: Bool?
    public var italic: Bool?
    public var underline: Bool?
    public var fontSizeHalfPoints: Int?
    public var color: String?

    public init(
        bold: Bool? = nil,
        italic: Bool? = nil,
        underline: Bool? = nil,
        fontSizeHalfPoints: Int? = nil,
        color: String? = nil
    ) {
        self.bold = bold
        self.italic = italic
        self.underline = underline
        self.fontSizeHalfPoints = fontSizeHalfPoints
        self.color = color
    }
}

// MARK: - JSONValue (free-form forward-compat payload)

/// Sendable, Equatable, Codable representation of an arbitrary JSON value.
///
/// Used to carry the payload of `Operation.unknown(opType:payload:)` —
/// preserves any extra fields beyond the four required JSONL discriminator
/// fields (`op_id`, `ts`, `source`, `op_type`) byte-equal across encode/decode
/// cycles. Indirect enum so deeply-nested structures don't blow the stack.
public indirect enum JSONValue: Equatable, Sendable, Codable {
    case null
    case bool(Bool)
    case int(Int)
    case double(Double)
    case string(String)
    case array([JSONValue])
    case object([String: JSONValue])

    public init(from decoder: Decoder) throws {
        let container = try decoder.singleValueContainer()
        if container.decodeNil() {
            self = .null
        } else if let v = try? container.decode(Bool.self) {
            self = .bool(v)
        } else if let v = try? container.decode(Int.self) {
            self = .int(v)
        } else if let v = try? container.decode(Double.self) {
            self = .double(v)
        } else if let v = try? container.decode(String.self) {
            self = .string(v)
        } else if let v = try? container.decode([JSONValue].self) {
            self = .array(v)
        } else if let v = try? container.decode([String: JSONValue].self) {
            self = .object(v)
        } else {
            throw DecodingError.dataCorruptedError(
                in: container,
                debugDescription: "JSONValue: unrecognized JSON kind"
            )
        }
    }

    public func encode(to encoder: Encoder) throws {
        var container = encoder.singleValueContainer()
        switch self {
        case .null: try container.encodeNil()
        case .bool(let v): try container.encode(v)
        case .int(let v): try container.encode(v)
        case .double(let v): try container.encode(v)
        case .string(let v): try container.encode(v)
        case .array(let v): try container.encode(v)
        case .object(let v): try container.encode(v)
        }
    }
}
