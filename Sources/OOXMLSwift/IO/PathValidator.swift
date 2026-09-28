import Foundation

/// Validates that an OOXML relationship target is a safe relative path.
///
/// Background (che-word-mcp#55): `_rels/document.xml.rels` `Target` attributes
/// flow into `URL.appendingPathComponent`, which does NOT normalize `..`
/// traversal. Without validation a malicious .docx can read OR write outside
/// the intended `word/` directory at the user's UID.
///
/// Defense-in-depth: this validator runs at the parse boundary (DocxReader)
/// AND at property setters on `Header.originalFileName` / `Footer.originalFileName`.
///
/// Returns `true` when the path is acceptable for use as a relative OOXML
/// part filename. Rejects:
/// - Empty / oversized paths (DoS guard at 256 chars)
/// - Absolute paths (leading `/`)
/// - Parent-directory traversal (any `..` segment, including URL-encoded)
/// - Control characters (NUL, newline, etc. — would terminate path parsing in C-string sinks)
///
/// Accepts non-ASCII Unicode in the printable range (CJK-named header parts
/// are valid in Word installs with CJK locales).
public func isSafeRelativeOOXMLPath(_ path: String) -> Bool {
    // DoS guard
    guard !path.isEmpty, path.count <= 256 else {
        return false
    }

    // Absolute paths
    if path.hasPrefix("/") {
        return false
    }

    // URL-encoded traversal (decode-then-check: %2e%2e, %2E%2E, mixed case)
    let lowercased = path.lowercased()
    if lowercased.contains("%2e%2e") || lowercased.contains("%2f") || lowercased.contains("%5c") {
        return false
    }

    // Direct `..` segments (split by both `/` and `\`)
    let segments = path.split { $0 == "/" || $0 == "\\" }
    for segment in segments where segment == ".." {
        return false
    }

    // Control characters (NUL, newlines, tabs — anything < 0x20 or 0x7F)
    for scalar in path.unicodeScalars {
        if scalar.value < 0x20 || scalar.value == 0x7F {
            return false
        }
    }

    return true
}

// MARK: - Relationship-target containment (che-word-mcp export_all_images / #ox-image-target-containment)

/// Resolves an OOXML relationship `Target` to a regular file strictly
/// inside `packageRoot` (the directory an already-`ZipHelper.unzip`-ed
/// package was extracted to), or returns `nil` when it cannot.
///
/// This is the SINGLE shared containment gate for relationship-target-driven
/// filesystem reads that legitimately need `..` to work (a Target may climb
/// out of its owning part's directory and back down into a sibling — e.g.
/// `../media/image1.png` from `word/_rels/document.xml.rels`, whose base
/// directory is `word/`, names the package-root part `media/image1.png`).
/// `isSafeRelativeOOXMLPath` above is intentionally stricter (rejects EVERY
/// `..`) and stays the right guard for header/footer Targets, which (a)
/// never legitimately contain `..` and (b) have the raw Target string reused
/// a SECOND time verbatim to build a sibling `_rels/<target>.rels` path — a
/// blanket "no `..` anywhere" invariant protects that reuse for free, while
/// a resolve-then-check-the-final-file function like this one only
/// guarantees safety for the ONE candidate it is asked to resolve. Images
/// have exactly one sink (the read itself), so this function's OPC-correct
/// resolution is both necessary (to keep supporting legitimate `..`) and
/// sufficient there.
///
/// Resolution follows OPC (ECMA-376 Part 2 §9.3 / RFC 3986 §5.2): a Target
/// starting with `/` is resolved against the PACKAGE root; otherwise
/// against `partDirectory` (e.g. `/word/` for `word/_rels/document.xml.rels`,
/// whose relationships are relative to the directory `word/document.xml`
/// itself lives in). This reuses `ProfileXML.normalizedRelationshipTarget`,
/// which already performs this exact lexical merge-and-remove-dot-segments
/// algorithm for styles/theme/fontTable/numbering resolution (macdoc#196) —
/// one definition of "resolve a relationship Target" now backs both.
///
/// That algorithm can never leave a literal `..` in its output (a `..`
/// segment only ever POPS a previously-appended segment; it is never itself
/// appended) and clamps at the virtual package root (a `..` with nothing
/// left to pop is silently dropped, not carried negative) — so the
/// candidate this function builds is always a package-root-relative,
/// dot-segment-free path before it ever touches the real filesystem. What
/// still needs a live filesystem check is a symlink physically sitting
/// inside the extracted tree (`ZipHelper.unzip` itself refuses to extract a
/// symlink *zip entry*, but this is defense-in-depth for any caller that
/// prepared `packageRoot` a different way) — `resolvingSymlinksInPath()`
/// on both the root and the candidate, compared with `hasPrefix`, catches
/// that case the same way `RawPartChannel.readAllParts(fromDirectory:)`
/// already does for its own directory walk.
///
/// Returns `nil` (never throws — an unresolvable relationship is an "orphan,
/// skip it" situation for a general reader, exactly like a Target naming a
/// part that simply does not exist) when:
/// - `targetMode == "External"` — never a package-internal read
/// - `target` is empty, or normalizes to a directory name (trailing `/`)
/// - the resolved, symlink-resolved candidate is not `packageRoot` itself or
///   strictly inside it (the actual containment gate)
/// - nothing exists at the resolved path, or it exists but is not a regular
///   file (directories, device nodes, FIFOs, … are refused)
internal func resolveContainedOOXMLTarget(
    _ target: String,
    targetMode: String?,
    relativeToPartDirectory partDirectory: String,
    packageRoot: URL
) -> URL? {
    guard targetMode != "External" else { return nil }
    guard !target.isEmpty else { return nil }

    let normalizedPartName = ProfileXML.normalizedRelationshipTarget(target, relativeTo: partDirectory)
    // `normalizedPartName` always starts with "/" (see
    // `normalizedRelationshipTarget`); `count > 1` excludes the bare "/"
    // (package root itself, not a part).
    guard normalizedPartName.count > 1, !normalizedPartName.hasSuffix("/") else { return nil }
    let relativePartPath = String(normalizedPartName.dropFirst())

    let normalizedRoot = packageRoot.resolvingSymlinksInPath().path
    let candidate = packageRoot.appendingPathComponent(relativePartPath)
    let resolvedCandidate = candidate.resolvingSymlinksInPath().path

    guard resolvedCandidate == normalizedRoot || resolvedCandidate.hasPrefix(normalizedRoot + "/") else {
        return nil
    }

    var isDirectory: ObjCBool = false
    guard FileManager.default.fileExists(atPath: resolvedCandidate, isDirectory: &isDirectory),
          !isDirectory.boolValue else {
        return nil
    }
    // Regular files only — mirrors `PackageInspector`'s own
    // `identity(_:ofType:)` regular-file check.
    guard let attributes = try? FileManager.default.attributesOfItem(atPath: resolvedCandidate),
          (attributes[.type] as? FileAttributeType) == .typeRegular else {
        return nil
    }
    return URL(fileURLWithPath: resolvedCandidate)
}
