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
    ///
    /// Every caller of this function takes an arbitrary `.docx` path — a
    /// `--verify-against` target, an `execute_script` / `export_script`
    /// `source_path`, or a package Word itself just wrote — so, unlike
    /// `ZipHelper.unzip`, nothing here refused an entry count or a byte total
    /// before this was written (PsychQuant/ooxml-swift#107). `limits` bounds
    /// three things, each refused on its own evidence and before the whole
    /// package is held in memory: the number of entries, one entry's DECLARED
    /// size, and the ACTUAL bytes a decompressing entry produces (checked
    /// per chunk, so an entry whose declaration understates what it
    /// decompresses to is stopped within one read-chunk of the limit, not
    /// after the fact). The running total across all entries is bounded too.
    public static func readAllParts(from docxURL: URL, limits: ZipHelper.Limits = ZipHelper.defaultLimits) throws -> [String: Data] {
        let archive = try Archive(url: docxURL, accessMode: .read)
        var parts: [String: Data] = [:]
        var entryCount = 0
        var declaredTotal: Int64 = 0
        var actualTotal: Int64 = 0
        for entry in archive where entry.type == .file {
            entryCount += 1
            if entryCount > limits.maximumEntryCount {
                throw WordError.invalidDocx("the package declares more than \(limits.maximumEntryCount) entries; refusing to read it.")
            }
            let declared = Int64(clamping: entry.uncompressedSize)
            if declared > limits.maximumEntryBytes {
                throw WordError.invalidDocx("the package declares an entry of \(ZipHelper.describeBytes(declared)) (\(ZipHelper.displayName(entry.path))), over the \(ZipHelper.describeBytes(limits.maximumEntryBytes)) limit; refusing to read it.")
            }
            let (declaredSum, declaredOverflowed) = declaredTotal.addingReportingOverflow(declared)
            declaredTotal = declaredOverflowed ? Int64.max : declaredSum
            if declaredOverflowed || declaredTotal > limits.maximumTotalBytes {
                throw WordError.invalidDocx("the package declares more than \(ZipHelper.describeBytes(limits.maximumTotalBytes)) of content in total; refusing to read it.")
            }
            var buffer = Data()
            var entryActual: Int64 = 0
            // Checked per chunk, DURING decompression — not after the entry is
            // fully buffered — so an entry whose declared size understates
            // what it actually decompresses to is stopped within one chunk of
            // the limit rather than after the whole thing is already in
            // memory.
            _ = try archive.extract(entry) { chunk in
                buffer.append(chunk)
                entryActual += Int64(chunk.count)
                if entryActual > limits.maximumEntryBytes {
                    throw WordError.invalidDocx("the package decompresses an entry to more than \(ZipHelper.describeBytes(limits.maximumEntryBytes)) (\(ZipHelper.displayName(entry.path))), whatever it declared; refusing to read it.")
                }
            }
            let (actualSum, actualOverflowed) = actualTotal.addingReportingOverflow(entryActual)
            actualTotal = actualOverflowed ? Int64.max : actualSum
            if actualOverflowed || actualTotal > limits.maximumTotalBytes {
                throw WordError.invalidDocx("the package decompresses to more than \(ZipHelper.describeBytes(limits.maximumTotalBytes)) of content in total; refusing to read it.")
            }
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
