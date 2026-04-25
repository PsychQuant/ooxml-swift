import Foundation

// MARK: - Position-indexed paragraph child markers (PsychQuant/che-word-mcp#56)
//
// Every legal child of `<w:p>` per ECMA-376 Part 1 §17.3.1 (`CT_P`) needs to
// survive a Reader → mutate → Writer round-trip. Some children (Run, Hyperlink,
// FieldSimple, AlternateContent) have full typed models. Others are
// either purely positional (range markers carrying just `w:id`) or have such
// a wide attribute surface that a typed model is impractical (custom XML,
// smart tags, bidi overrides). For those, we store the verbatim source XML in
// a raw-carrier struct alongside a `position: Int` so the Writer can emit
// them in source order via `Paragraph.toXML()`'s sort-by-position pass.
//
// Phase 2 lands `BookmarkRangeMarker` only. Phase 4 (tasks 4.2–4.4) extends
// this file with the remaining six raw-carrier types.

// MARK: - Range marker shared shape

/// Two-state range marker kind shared by every `*RangeStart` / `*RangeEnd`
/// pair in OOXML (bookmarks, comments, permissions). Phase 4 lands all three
/// pair types using this same enum.
public enum RangeMarkerKind: String, Equatable {
    case start
    case end
}

// MARK: - BookmarkRangeMarker

/// Position-indexed marker recording where a `<w:bookmarkStart>` or
/// `<w:bookmarkEnd>` element sits inside its parent `<w:p>`. The `Bookmark`
/// model on `Paragraph.bookmarks` carries the bookmark's name + id pair;
/// this marker carries the *position* information needed to re-emit the
/// bookmark elements at their original relative positions during Writer
/// sort-by-position emit (see Phase 4 task 4.5).
public struct BookmarkRangeMarker: Equatable {
    public typealias Kind = RangeMarkerKind

    public var kind: Kind
    public var id: Int
    public var position: Int

    public init(kind: Kind, id: Int, position: Int) {
        self.kind = kind
        self.id = id
        self.position = position
    }
}

// MARK: - CommentRangeMarker

/// Position-indexed marker for `<w:commentRangeStart w:id="…"/>` and
/// `<w:commentRangeEnd w:id="…"/>`. Comments themselves live in
/// `word/comments.xml`; this marker only records *where* the highlighted
/// range begins / ends inside the paragraph so a no-op round-trip emits the
/// markers at their original source positions.
///
/// Note: legacy `Paragraph.commentIds` (populated by the v0.1.0 parser)
/// only captures `commentRangeStart` ids without position. v0.19.0+ keeps
/// `commentIds` populated for backward compat with 218 MCP tools, while
/// adding `commentRangeMarkers` for round-trip preservation.
public struct CommentRangeMarker: Equatable {
    public typealias Kind = RangeMarkerKind

    public var kind: Kind
    public var id: Int
    public var position: Int

    public init(kind: Kind, id: Int, position: Int) {
        self.kind = kind
        self.id = id
        self.position = position
    }
}

// MARK: - PermissionRangeMarker

/// Position-indexed marker for `<w:permStart>` / `<w:permEnd>`. Permissions
/// gate which authors / groups may edit the wrapped range. Captures the id,
/// editor / group attributes, and source position so non-editable regions
/// survive a no-op round-trip.
public struct PermissionRangeMarker: Equatable {
    public typealias Kind = RangeMarkerKind

    public var kind: Kind
    public var id: String
    /// `w:edGrp` (when present, e.g., "everyone" / "current") for permStart.
    public var editorGroup: String?
    /// `w:ed` (specific editor name) for permStart.
    public var editor: String?
    public var position: Int

    public init(kind: Kind, id: String, editorGroup: String? = nil, editor: String? = nil, position: Int) {
        self.kind = kind
        self.id = id
        self.editorGroup = editorGroup
        self.editor = editor
        self.position = position
    }
}

// MARK: - ProofErrorMarker

/// Position-indexed marker for `<w:proofErr w:type="…"/>`. Proof errors
/// flag spelling / grammar issues for Word's Proofing UI. They carry no
/// id and can appear without a closing partner — `kind` records whether
/// this is a `start` or `end` per the `w:type` attribute (`spellStart`,
/// `spellEnd`, `gramStart`, `gramEnd`).
public struct ProofErrorMarker: Equatable {
    public enum ErrorType: String, Equatable {
        case spellStart
        case spellEnd
        case gramStart
        case gramEnd
    }

    public var type: ErrorType
    public var position: Int

    public init(type: ErrorType, position: Int) {
        self.type = type
        self.position = position
    }
}

// MARK: - SmartTagBlock

/// Position-indexed raw-carrier for `<w:smartTag>`. Smart tags wrap runs
/// with semantic annotations (e.g., dates, addresses, stock tickers) and
/// have a wide attribute surface (`w:uri`, `w:element`, plus per-vendor
/// child `<w:smartTagPr>`). Stored as verbatim `rawXML` so round-trip is
/// byte-equivalent without typed modeling.
public struct SmartTagBlock: Equatable {
    public var rawXML: String
    public var position: Int

    public init(rawXML: String, position: Int) {
        self.rawXML = rawXML
        self.position = position
    }
}

// MARK: - CustomXmlBlock

/// Position-indexed raw-carrier for `<w:customXml>`. Custom XML blocks are
/// extension points for line-of-business apps to embed structured data.
/// Stored verbatim so the binding semantics (which we don't model) survive
/// a no-op round-trip.
public struct CustomXmlBlock: Equatable {
    public var rawXML: String
    public var position: Int

    public init(rawXML: String, position: Int) {
        self.rawXML = rawXML
        self.position = position
    }
}

// MARK: - BidiOverrideBlock

/// Position-indexed raw-carrier for `<w:dir>` (direction) and `<w:bdo>`
/// (bidirectional override). Both wrap runs to flip text direction
/// (RTL / LTR) for mixed-script content. Stored verbatim because the inner
/// `<w:r>` children inherit the wrapper's directional context — pulling
/// them up into `Paragraph.runs` would lose that semantic.
public struct BidiOverrideBlock: Equatable {
    public enum Element: String, Equatable {
        case dir
        case bdo
    }

    public var element: Element
    public var rawXML: String
    public var position: Int

    public init(element: Element, rawXML: String, position: Int) {
        self.element = element
        self.rawXML = rawXML
        self.position = position
    }
}

// MARK: - UnrecognizedChild

/// Fallback raw-carrier for any `<w:p>` direct child that does not match
/// any typed model or registered raw-carrier above. Stores element local
/// name + verbatim XML + source position. Used to surface ECMA-376 spec
/// gaps (any element appearing here triggers an XCTFail in the round-trip
/// test suite, naming the element so we can add a dedicated raw-carrier).
public struct UnrecognizedChild: Equatable {
    public var name: String
    public var rawXML: String
    public var position: Int

    public init(name: String, rawXML: String, position: Int) {
        self.name = name
        self.rawXML = rawXML
        self.position = position
    }
}
