import Foundation

// Cross-document OMath splice — main API implementation.
//
// Spec: openspec/changes/cross-document-omath-splice/specs/omath-splice/spec.md
// Design: openspec/changes/cross-document-omath-splice/design.md
// Issue: PsychQuant/ooxml-swift#57

extension WordDocument {

    /// Copy a verbatim `<m:oMath>` XML block from a source `Paragraph` and splice it
    /// into a target body paragraph at the specified position.
    ///
    /// Source carrier shape is preserved (Decision Q1):
    /// - Inline OMath in source `Run.rawXML` → target gets a new `Run` with `rawXML`
    /// - Direct-child OMath in source `unrecognizedChildren` → target gets a new
    ///   `UnrecognizedChild(name: "oMath", ...)`
    ///
    /// `omathIndex` selects which OMath to splice when source paragraph contains
    /// multiple, in source-document order joint-sorted across both carriers (Q2).
    ///
    /// Mid-paragraph splice (`.afterText` / `.beforeText` resolving inside a run)
    /// triggers run-split: the anchor run is divided into prefix/suffix segments,
    /// rPr is copied to both, OMath Run is inserted between them, all sharing
    /// the original run's `position` value (Q3 — relies on `Paragraph.toXML`'s
    /// stable sort to retain insertion order).
    ///
    /// - Parameters:
    ///   - sourceParagraph: paragraph from which to extract OMath
    ///   - toBodyParagraphIndex: body-children paragraph index in `self`
    ///     (counts only `.paragraph` direct children of body, NOT including tables / SDTs)
    ///   - position: where within the target paragraph to splice the OMath
    ///   - omathIndex: 0-based index into source's joint-ordered OMath list (default 0)
    ///   - rPrMode: how to propagate source Run rPr to the new OMath Run (Q4)
    ///   - namespacePolicy: how to handle prefix / URI mismatch (Q6)
    /// - Returns: number of OMath blocks spliced (always 1 in this single-OMath API)
    /// - Throws: `OMathSpliceError` per the failure taxonomy
    @discardableResult
    public mutating func spliceOMath(
        from sourceParagraph: Paragraph,
        toBodyParagraphIndex: Int,
        position: OMathSplicePosition,
        omathIndex: Int = 0,
        rPrMode: OMathSpliceRpRMode = .full,
        namespacePolicy: OMathSpliceNamespacePolicy = .lenient
    ) throws -> Int {
        // === Validate target paragraph index ===
        let targetParagraphIndices = Self.bodyParagraphIndices(in: body)
        guard toBodyParagraphIndex >= 0 && toBodyParagraphIndex < targetParagraphIndices.count else {
            throw OMathSpliceError.targetParagraphOutOfRange(toBodyParagraphIndex)
        }
        let bodyChildIdx = targetParagraphIndices[toBodyParagraphIndex]

        // === Extract OMath from source ===
        let extracted = OMathExtractor.extract(from: sourceParagraph)
        guard !extracted.isEmpty else {
            throw OMathSpliceError.sourceHasNoOMath
        }
        guard omathIndex >= 0 && omathIndex < extracted.count else {
            throw OMathSpliceError.omathIndexOutOfRange(requested: omathIndex, available: extracted.count)
        }
        let omath = extracted[omathIndex]

        // === Namespace policy check (Q6) ===
        if case .paragraph(let targetPara) = body.children[bodyChildIdx] {
            try Self.checkNamespacePolicy(
                source: omath.xml,
                targetParagraph: targetPara,
                policy: namespacePolicy
            )
        }

        // === Splice into target ===
        guard case .paragraph(var targetPara) = body.children[bodyChildIdx] else {
            throw OMathSpliceError.targetParagraphOutOfRange(toBodyParagraphIndex)
        }

        try Self.performSplice(
            into: &targetPara,
            omath: omath,
            position: position,
            rPrMode: rPrMode
        )

        body.children[bodyChildIdx] = .paragraph(targetPara)
        modifiedParts.insert("word/document.xml")
        return 1
    }

    /// Paragraph-level batch splice — copies all OMath blocks from one source paragraph
    /// to a corresponding target body paragraph in source-document order, auto-deriving
    /// the splice anchor for each OMath from its source-text-context (Q5).
    ///
    /// For each OMath in source order, this method:
    /// 1. Slices `flattenedDisplayText()` ~10 chars before and after the OMath's source position
    /// 2. Routes to `spliceOMath(..., position: .afterText(prefix, instance: 1), ...)` for the
    ///    target side
    /// 3. Throws `.contextAnchorNotFound(omathIndex:, snippet:)` per OMath where prefix lookup fails
    ///
    /// Partial-success semantics: any OMath blocks already spliced before a failure remain
    /// in target — caller can inspect target paragraph state.
    @discardableResult
    public mutating func spliceParagraphOMath(
        from sourceParagraph: Paragraph,
        toBodyParagraphIndex: Int,
        rPrMode: OMathSpliceRpRMode = .full,
        namespacePolicy: OMathSpliceNamespacePolicy = .lenient
    ) throws -> Int {
        let extracted = OMathExtractor.extract(from: sourceParagraph)
        guard !extracted.isEmpty else {
            return 0  // No OMath to splice — graceful no-op for batch driver loops
        }

        // Build a flattened text view of source paragraph WITH OMath visibleText included.
        // For each OMath we need the prefix substring (from prose only, NOT including
        // OMath visible glyphs themselves — we use original Run text positions).
        //
        // Approach: walk the source paragraph collecting per-run text, finding each OMath's
        // anchor as the ~10 chars of plain prose immediately preceding it.
        let prefixContexts = Self.deriveContextAnchors(
            from: sourceParagraph,
            forExtracted: extracted,
            charsBefore: 10
        )

        var spliced = 0
        for (i, omath) in extracted.enumerated() {
            let snippet = prefixContexts[i]
            // Empty prefix → fall back to .atEnd (OMath at start of paragraph).
            let position: OMathSplicePosition = snippet.isEmpty
                ? .atEnd
                : .afterText(snippet, instance: 1, options: AnchorLookupOptions())

            do {
                try self.spliceOMath(
                    from: sourceParagraph,
                    toBodyParagraphIndex: toBodyParagraphIndex,
                    position: position,
                    omathIndex: i,
                    rPrMode: rPrMode,
                    namespacePolicy: namespacePolicy
                )
                spliced += 1
            } catch OMathSpliceError.anchorNotFound(_, _) {
                throw OMathSpliceError.contextAnchorNotFound(
                    omathIndex: i,
                    snippet: snippet
                )
                _ = omath  // keep reference; Swift unused-var hint suppression
            }
        }
        return spliced
    }

    // MARK: - Internal helpers

    /// Indices into `body.children` that are `.paragraph(_)` cases.
    /// Used to translate caller's body-paragraph index into the actual `body.children` index.
    internal static func bodyParagraphIndices(in body: Body) -> [Int] {
        var out: [Int] = []
        for (i, child) in body.children.enumerated() {
            if case .paragraph = child { out.append(i) }
        }
        return out
    }

    /// Validates source XML's namespace against target paragraph's first OMath occurrence
    /// (or, if target has no existing OMath, against the OMML standard URI).
    ///
    /// `.lenient`: accept prefix mismatch with same URI; throw on URI mismatch.
    /// `.strict`: throw on any prefix or URI mismatch.
    internal static func checkNamespacePolicy(
        source: String,
        targetParagraph: Paragraph,
        policy: OMathSpliceNamespacePolicy
    ) throws {
        let standardOMMLURI = "http://schemas.openxmlformats.org/officeDocument/2006/math"
        let sourceURI = OMathNamespace.extractURI(from: source) ?? standardOMMLURI
        let sourcePrefix = OMathNamespace.extractPrefix(from: source) ?? ""

        // Find target OMath URI/prefix from existing OMath in target paragraph (if any).
        var targetURI: String = standardOMMLURI
        var targetPrefix: String = "m"
        for run in targetParagraph.runs {
            if let raw = run.rawXML, raw.contains(":oMath") || raw.contains("<oMath") {
                if let uri = OMathNamespace.extractURI(from: raw) { targetURI = uri }
                if let pre = OMathNamespace.extractPrefix(from: raw) { targetPrefix = pre }
                break
            }
        }
        for child in targetParagraph.unrecognizedChildren where child.name == "oMath" || child.name == "oMathPara" {
            if let uri = OMathNamespace.extractURI(from: child.rawXML) { targetURI = uri }
            if let pre = OMathNamespace.extractPrefix(from: child.rawXML) { targetPrefix = pre }
            break
        }

        if sourceURI != targetURI {
            throw OMathSpliceError.namespaceMismatch(sourceURI: sourceURI, targetURI: targetURI)
        }
        if policy == .strict && sourcePrefix != targetPrefix {
            throw OMathSpliceError.namespaceMismatch(sourceURI: sourceURI, targetURI: targetURI)
        }
    }

    /// Performs the actual splice into the given target paragraph (in-place mutation).
    internal static func performSplice(
        into targetPara: inout Paragraph,
        omath: ExtractedOMath,
        position: OMathSplicePosition,
        rPrMode: OMathSpliceRpRMode
    ) throws {
        switch omath.kind {
        case .inRun:
            try spliceAsRun(into: &targetPara, omath: omath, position: position, rPrMode: rPrMode)
        case .directChild:
            try spliceAsDirectChild(into: &targetPara, omath: omath, position: position)
        }
    }

    /// Splice OMath as a new `Run.rawXML` into target paragraph.
    private static func spliceAsRun(
        into targetPara: inout Paragraph,
        omath: ExtractedOMath,
        position: OMathSplicePosition,
        rPrMode: OMathSpliceRpRMode
    ) throws {
        // Build the new OMath Run.
        var omathRun = Run(text: "")
        omathRun.rawXML = omath.xml
        if let sourceRpR = omath.sourceRunProperties {
            omathRun.properties = sourceRpR.filteredForOMathSplice(mode: rPrMode)
        } else {
            // Source carrier was directChild — no Run rPr. Use empty.
            omathRun.properties = RunProperties()
        }

        switch position {
        case .atStart:
            // Position 0 routes through the post-content legacy path and emits at end.
            // To get "atStart" semantically, give it a position smaller than any existing.
            let minPos = targetPara.runs.compactMap { $0.position }.min() ?? 1
            omathRun.position = max(1, minPos - 1)
            targetPara.runs.insert(omathRun, at: 0)

        case .atEnd:
            // Place after all current content. Use max existing position + 1.
            let maxPos = targetPara.runs.compactMap { $0.position }.max() ?? 0
            omathRun.position = maxPos + 1
            targetPara.runs.append(omathRun)

        case .afterText(let anchor, let instance, let options):
            try insertAtAnchor(
                anchor: anchor,
                instance: instance,
                options: options,
                position: .after,
                omathRun: omathRun,
                in: &targetPara
            )

        case .beforeText(let anchor, let instance, let options):
            try insertAtAnchor(
                anchor: anchor,
                instance: instance,
                options: options,
                position: .before,
                omathRun: omathRun,
                in: &targetPara
            )
        }
    }

    private enum AnchorSide { case before, after }

    /// Resolve anchor + perform run-split + insert OMath Run at split point.
    /// Q3 decision: shared position with original run, stable sort retains order.
    ///
    /// Handles three cases:
    /// 1. Anchor falls entirely within one run → split that run.
    /// 2. Anchor spans multiple runs and `.after` is requested → split only the END run
    ///    at the position where the anchor ends.
    /// 3. Anchor spans multiple runs and `.before` is requested → split only the START
    ///    run at the position where the anchor begins.
    private static func insertAtAnchor(
        anchor: String,
        instance: Int,
        options: AnchorLookupOptions,
        position side: AnchorSide,
        omathRun: Run,
        in para: inout Paragraph
    ) throws {
        guard let resolved = resolveRunAnchor(
            anchor: anchor,
            instance: instance,
            in: para
        ) else {
            throw OMathSpliceError.anchorNotFound(anchor, instance: instance)
        }

        let splitRunIdx = side == .after ? resolved.endRunIdx : resolved.startRunIdx
        let charOffset = side == .after ? resolved.endOffsetInEndRun : resolved.startOffsetInStartRun
        let originalRun = para.runs[splitRunIdx]

        // Split run into prefix [0..<charOffset] and suffix [charOffset..<end].
        let (prefix, suffix) = splitRun(originalRun, atCharOffset: charOffset)

        var newOmath = omathRun
        newOmath.position = originalRun.position

        var newRuns: [Run] = []
        if !prefix.text.isEmpty || prefix.rawXML != nil {
            newRuns.append(prefix)
        }
        newRuns.append(newOmath)
        if !suffix.text.isEmpty || suffix.rawXML != nil {
            newRuns.append(suffix)
        }

        para.runs.replaceSubrange(splitRunIdx...splitRunIdx, with: newRuns)
    }

    /// Splits a run's `text` at the given UTF-16 character offset, returning prefix and suffix.
    /// `properties`, `position`, `rawXML`, etc. are deep-copied to both sides (rawXML stays
    /// only on whichever segment carries the underlying content — but for plain-text runs,
    /// rawXML is typically nil, so both sides get nil).
    internal static func splitRun(_ original: Run, atCharOffset offset: Int) -> (prefix: Run, suffix: Run) {
        let text = original.text
        // Use UTF-16 offset semantics consistent with the rest of OOXMLSwift's text math.
        let utf16 = text.utf16
        let safeOffset = max(0, min(offset, utf16.count))
        let splitIdx = utf16.index(utf16.startIndex, offsetBy: safeOffset)
        let prefixText = String(String.UnicodeScalarView(text.unicodeScalars.prefix(safeOffset)))
            // Fallback to UTF-16 substring mapped back to String when scalar count diverges.
        let prefixStr: String
        let suffixStr: String
        if let pStr = String(utf16.prefix(safeOffset)),
           let sStr = String(utf16.suffix(utf16.count - safeOffset)) {
            prefixStr = pStr
            suffixStr = sStr
        } else {
            // Fallback: use Character-level slicing (safe for ASCII / BMP).
            let chars = Array(text)
            let cap = min(safeOffset, chars.count)
            prefixStr = String(chars[0..<cap])
            suffixStr = String(chars[cap..<chars.count])
        }
        _ = prefixText  // silence unused
        _ = splitIdx

        var prefixRun = original
        prefixRun.text = prefixStr
        // Prefix keeps drawing/rawXML/etc. only if it's non-text content (rare). For
        // plain text runs, drawing/rawXML are typically nil — they pass through.
        // For OMath-bearing runs, we shouldn't be splitting them in the first place
        // (anchor resolution skips them). So this is safe.

        var suffixRun = original
        suffixRun.text = suffixStr

        return (prefixRun, suffixRun)
    }

    /// Result of resolving an anchor to runs.
    /// `startRunIdx` and `startOffsetInStartRun` describe where the anchor begins;
    /// `endRunIdx` and `endOffsetInEndRun` describe where it ends. For single-run
    /// anchors, `startRunIdx == endRunIdx`. For cross-run anchors, they differ.
    internal struct RunAnchorResolution {
        let startRunIdx: Int
        let startOffsetInStartRun: Int
        let endRunIdx: Int
        let endOffsetInEndRun: Int
    }

    /// Walks paragraph runs in array order, accumulating text-only flatten, and locates
    /// the Nth occurrence of the anchor string. Supports cross-run anchors — returns
    /// the start and end run indices + offsets so callers can split only the relevant run.
    ///
    /// Skips runs whose rawXML contains OMath — those are pure-OMath runs whose visibleText
    /// would mislead the splitter (we cannot split inside an OMath run).
    internal static func resolveRunAnchor(
        anchor: String,
        instance: Int,
        in para: Paragraph
    ) -> RunAnchorResolution? {
        guard !anchor.isEmpty, instance >= 1 else { return nil }
        let anchorUtf16 = Array(anchor.utf16)

        // Build (runIdx, runText, startGlobal) excluding OMath-bearing runs.
        var runSpans: [(runIdx: Int, text: String, startGlobal: Int)] = []
        var globalOffset = 0
        for (i, run) in para.runs.enumerated() {
            if let raw = run.rawXML, raw.contains(":oMath") || raw.contains("<oMath")
                || raw.contains(":oMathPara") || raw.contains("<oMathPara") {
                continue
            }
            runSpans.append((i, run.text, globalOffset))
            globalOffset += run.text.utf16.count
        }

        let combined = runSpans.map { $0.text }.joined()
        let combinedUtf16 = Array(combined.utf16)
        guard combinedUtf16.count >= anchorUtf16.count else { return nil }

        var occurrencesFound = 0
        var searchStart = 0
        while searchStart + anchorUtf16.count <= combinedUtf16.count {
            var match = true
            for j in 0..<anchorUtf16.count {
                if combinedUtf16[searchStart + j] != anchorUtf16[j] {
                    match = false
                    break
                }
            }
            if match {
                occurrencesFound += 1
                if occurrencesFound == instance {
                    let globalStart = searchStart
                    let globalEnd = searchStart + anchorUtf16.count

                    // Find run containing globalStart.
                    var startSpan: (runIdx: Int, text: String, startGlobal: Int)? = nil
                    for span in runSpans {
                        let runEnd = span.startGlobal + span.text.utf16.count
                        if globalStart >= span.startGlobal && globalStart < runEnd {
                            startSpan = span
                            break
                        }
                    }
                    // Find run containing globalEnd-1 (last anchor char).
                    var endSpan: (runIdx: Int, text: String, startGlobal: Int)? = nil
                    let lastCharGlobal = globalEnd - 1
                    for span in runSpans {
                        let runEnd = span.startGlobal + span.text.utf16.count
                        if lastCharGlobal >= span.startGlobal && lastCharGlobal < runEnd {
                            endSpan = span
                            break
                        }
                    }
                    guard let start = startSpan, let end = endSpan else { return nil }

                    return RunAnchorResolution(
                        startRunIdx: start.runIdx,
                        startOffsetInStartRun: globalStart - start.startGlobal,
                        endRunIdx: end.runIdx,
                        endOffsetInEndRun: globalEnd - end.startGlobal
                    )
                }
                searchStart += 1
            } else {
                searchStart += 1
            }
        }
        return nil
    }

    /// Splice OMath as a direct child of `<w:p>` via `unrecognizedChildren`.
    private static func spliceAsDirectChild(
        into targetPara: inout Paragraph,
        omath: ExtractedOMath,
        position: OMathSplicePosition
    ) throws {
        var newPos: Int
        switch position {
        case .atStart:
            let allPositions = paragraphAllPositions(targetPara)
            newPos = max(1, (allPositions.min() ?? 1) - 1)

        case .atEnd:
            let allPositions = paragraphAllPositions(targetPara)
            newPos = (allPositions.max() ?? 0) + 1

        case .afterText, .beforeText:
            // For direct-child OMath splice with anchor, insert at end as a v0.1 simplification.
            // (Direct-child OMath is rarely mid-paragraph in practice — typically Pandoc
            // display equations stand alone in their paragraph.)
            let allPositions = paragraphAllPositions(targetPara)
            newPos = (allPositions.max() ?? 0) + 1
        }

        targetPara.unrecognizedChildren.append(
            UnrecognizedChild(name: "oMath", rawXML: omath.xml, position: newPos)
        )
    }

    private static func paragraphAllPositions(_ para: Paragraph) -> [Int] {
        var positions: [Int] = []
        positions += para.runs.compactMap { $0.position }
        positions += para.unrecognizedChildren.compactMap { $0.position }
        positions += para.bookmarkMarkers.compactMap { $0.position }
        positions += para.commentRangeMarkers.compactMap { $0.position }
        positions += para.permissionRangeMarkers.compactMap { $0.position }
        positions += para.proofErrorMarkers.compactMap { $0.position }
        positions += para.smartTags.compactMap { $0.position }
        positions += para.customXmlBlocks.compactMap { $0.position }
        positions += para.bidiOverrides.compactMap { $0.position }
        positions += para.contentControls.compactMap { $0.position }
        return positions
    }

    /// For each extracted OMath in source order, derive a context anchor (~N chars of
    /// plain prose immediately preceding the OMath) for batch splice.
    ///
    /// Walks runs in array order, accumulating text. When we hit a Run whose rawXML
    /// matches `extracted[i].xml`, we capture the `charsBefore`-trailing chars of the
    /// concatenated prose so far.
    internal static func deriveContextAnchors(
        from sourceParagraph: Paragraph,
        forExtracted extracted: [ExtractedOMath],
        charsBefore: Int
    ) -> [String] {
        var anchors = [String](repeating: "", count: extracted.count)
        var matchedExtractedIndices = Set<Int>()
        var proseAccumulator = ""

        // Walk runs in array order; whenever a run's rawXML matches an extracted OMath,
        // capture the trailing N chars of proseAccumulator.
        for run in sourceParagraph.runs {
            if let raw = run.rawXML,
               (raw.contains(":oMath") || raw.contains("<oMath") || raw.contains(":oMathPara") || raw.contains("<oMathPara")) {
                // Find which extracted index this run corresponds to (by xml byte equality).
                for (i, ex) in extracted.enumerated() where !matchedExtractedIndices.contains(i) {
                    if ex.kind == .inRun && ex.xml == raw {
                        let trailing = String(proseAccumulator.suffix(charsBefore)).trimmingCharacters(in: .whitespaces)
                        anchors[i] = trailing
                        matchedExtractedIndices.insert(i)
                        break
                    }
                }
                continue
            }
            proseAccumulator += run.text
        }

        // Direct-child OMath in unrecognizedChildren — derive from its source position
        // by walking runs up to that position.
        for (i, ex) in extracted.enumerated() where ex.kind == .directChild && !matchedExtractedIndices.contains(i) {
            // For direct-child OMath, use the prose accumulated up to the OMath's position.
            // (Best-effort — direct-child OMath typically doesn't have meaningful surrounding prose.)
            guard let omathPos = ex.sourcePosition else { continue }
            var prose = ""
            for run in sourceParagraph.runs {
                if let runPos = run.position, runPos < omathPos {
                    if run.rawXML?.contains("oMath") != true {
                        prose += run.text
                    }
                }
            }
            anchors[i] = String(prose.suffix(charsBefore)).trimmingCharacters(in: .whitespaces)
        }

        return anchors
    }
}

// MARK: - String UTF-16 helpers

private extension String {
    /// Returns String from a UTF-16 view subsequence, or nil if invalid sequence.
    init?<S: Sequence>(_ utf16: S) where S.Element == UInt16 {
        let array = Array(utf16)
        let scalars = String.UnicodeScalarView()
        let view: String.UTF16View? = nil
        _ = scalars
        _ = view
        // Use Foundation's NSString bridge for reliable round-trip.
        let ns = NSString(characters: array, length: array.count)
        self = ns as String
    }
}
