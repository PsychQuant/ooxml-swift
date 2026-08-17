// RawPartChannel.swift
// format-alignment-engine Phase A task 1.3 — reverse + compare utilities that
// wire the all-parts raw channel into a byte-equal acceptance pipeline
// (`format-alignment-pipeline` capability; Decision 1 Stage C exemption,
// Decision 2 dual-track raw floor).

import Foundation
import ZIPFoundation

public enum RawPartChannel {

    /// Reads every file entry of a `.docx` package as raw bytes, keyed by part
    /// path. Container-normalizing by construction: the map is keyed by part
    /// path and holds part CONTENT only — zip entry ordering, compression
    /// parameters, and timestamps (Stage C) never enter it, per the acceptance
    /// contract (`format-alignment-pipeline`, Decision 1). Feed two of these to
    /// `PartFidelity` to evaluate Stage A/B.
    public static func readAllParts(from docxURL: URL) throws -> [String: Data] {
        let archive = try Archive(url: docxURL, accessMode: .read)
        var parts: [String: Data] = [:]
        for entry in archive where entry.type == .file {
            var buffer = Data()
            _ = try archive.extract(entry) { buffer.append($0) }
            parts[entry.path] = buffer
        }
        return parts
    }

    /// Reads every regular file under an already-extracted package root.
    /// Used by sync to compare Word's new package against the preserved
    /// pre-import archive without re-zipping it first.
    public static func readAllParts(fromDirectory root: URL) throws -> [String: Data] {
        let fm = FileManager.default
        let normalizedRoot = root.resolvingSymlinksInPath().path
        guard let walk = fm.enumerator(at: root, includingPropertiesForKeys: [.isRegularFileKey]) else {
            return [:]
        }
        var parts: [String: Data] = [:]
        for case let url as URL in walk {
            let values = try url.resourceValues(forKeys: [.isRegularFileKey])
            guard values.isRegularFile == true else { continue }
            let normalized = url.resolvingSymlinksInPath().path
            guard normalized.hasPrefix(normalizedRoot + "/") else { continue }
            let path = String(normalized.dropFirst(normalizedRoot.count + 1))
            parts[path] = try Data(contentsOf: url)
        }
        return parts
    }

    /// Reverses the raw channel: every XML part becomes a `carryPart` op so a
    /// rebuild script reproduces it verbatim (the honest-copy baseline).
    /// Deterministic order (sorted by path) for reproducible scripts.
    ///
    /// Binary media parts use a base64 raw operation so the text script can
    /// reconstruct every OPC part without byte corruption.
    public static func carriedPartOps(from parts: [String: Data]) -> [Operation] {
        parts.sorted { $0.key < $1.key }.compactMap { path, bytes in
            if let xml = String(data: bytes, encoding: .utf8) {
                return .carryPart(partPath: path, xml: xml)
            }
            return .carryBinaryPart(partPath: path, base64: bytes.base64EncodedString())
        }
    }

    /// Part-level DSL/raw coverage (Phase A granularity): each XML part counts
    /// wholly as DSL or raw depending on membership in `dslParts`. Binary parts
    /// are excluded from the denominator — coverage is measured over XML parts
    /// (`format-alignment-pipeline` Q1 working answer: aggregate over all XML
    /// parts). Later phases refine this to sub-part byte splits as content
    /// classes move from raw to DSL within a single part.
    public static func partLevelCoverage(parts: [String: Data],
                                         dslParts: Set<String>) -> PartFidelity.CoverageReport {
        let coverages = parts.compactMap { path, bytes -> PartFidelity.PartCoverage? in
            // XML parts only in the denominator (binary media can't ride the
            // String-typed DSL/raw channel yet).
            guard String(data: bytes, encoding: .utf8) != nil else { return nil }
            let isDSL = dslParts.contains(path)
            return PartFidelity.PartCoverage(
                partPath: path,
                dslBytes: isDSL ? bytes.count : 0,
                rawBytes: isDSL ? 0 : bytes.count
            )
        }
        return PartFidelity.coverage(coverages)
    }
}
