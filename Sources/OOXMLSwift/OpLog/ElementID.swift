// ElementID — stable identifier for OOXML elements addressed by the op log.
//
// Spectra change: operation-log-scaffold-impl, target ooxml-swift v0.31.3.
// Capability: ooxml-operation-log
//
// Byte-aligned with `XmlNode.stableID` (landed v0.30.0 in `Tree/XmlNode.swift`)
// so the Phase 2b reducer can match an op's ElementID payload against the
// XmlTree's nodes via straight string equality. Same priority chain:
//   1. w14:paraId
//   2. w:bookmarkId
//   3. w:id
//   4. r:id
//   5. w14:textId
//   6. libraryUUID (UUID v4 fallback assigned to nodes lacking native ID)
// Returns nil only when the node carries none of the above and has no
// libraryUUID — caller responsibility (per design.md Decision 2 risk note) is
// to assign a libraryUUID before constructing ops that reference such a node.

import Foundation

/// Stable identifier for an OOXML element, used as the addressing primitive
/// in the operation log.
///
/// The `raw` String matches the format produced by `XmlNode.stableID` for
/// nodes that have a native stable-ID attribute (e.g., `"w14:paraId=0ABC1234"`)
/// and `"lib:<UUID>"` for nodes addressed by library-generated UUIDs only.
public struct ElementID: Equatable, Hashable, Sendable, Codable {

    /// The string form of the identifier — same format as `XmlNode.stableID`
    /// for native IDs, or `"lib:<UUID>"` for library-generated fallbacks.
    public let raw: String

    /// Derives an `ElementID` from an `XmlNode` using the priority chain
    /// `w14:paraId` → `w:bookmarkId` → `w:id` → `r:id` → `w14:textId` →
    /// `libraryUUID`. Returns `nil` when none of these are set — caller must
    /// assign a `libraryUUID` to the node before constructing ops that
    /// reference it (Phase 2b reducer concern).
    public init?(node: XmlNode) {
        if let stable = node.stableID {
            self.raw = stable
            return
        }
        if let uuid = node.libraryUUID {
            self.raw = "lib:\(uuid.uuidString)"
            return
        }
        return nil
    }

    /// Constructs an `ElementID` from a library-generated UUID. Used when the
    /// caller knows the UUID already and does not need the priority-chain
    /// lookup against an `XmlNode`.
    public init(libraryUUID: UUID) {
        self.raw = "lib:\(libraryUUID.uuidString)"
    }

    /// Constructs an `ElementID` from a verbatim string. Used by JSONL
    /// decoding to reconstruct `ElementID` values from on-disk bytes — the
    /// raw string is whatever the encoder emitted, no validation here.
    public init(rawString: String) {
        self.raw = rawString
    }
}
