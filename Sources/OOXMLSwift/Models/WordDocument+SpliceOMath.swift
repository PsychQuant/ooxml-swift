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
    /// - Throws: `OMathSpliceError` for splice/anchor/namespace failures, or
    ///   `OMathSpliceMalformedXMLError` when a source/target OMath fragment is invalid
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
        markTypedDirty("word/document.xml")
        return 1
    }

    /// Paragraph-level batch splice — copies all OMath blocks from one source paragraph
    /// to a corresponding target body paragraph in source-document order, auto-deriving
    /// the splice anchor for each OMath from its source-text-context (Q5).
    ///
    /// For each OMath in source order, this method:
    /// 1. Derives the trailing ~10-character prose prefix and its source occurrence instance
    /// 2. Routes to `.atStart` for a matched leading OMath or `.afterText(prefix, instance:)`
    /// 3. Globally preflights every XML fragment before anchor derivation/mutation
    /// 4. Preflights each shared-anchor group and throws `.contextAnchorNotFound` on failure
    ///
    /// Partial-success semantics: earlier source anchor groups remain after a later group fails;
    /// a failing shared-anchor group does not partially mutate the target.
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

        let targetParagraphIndices = Self.bodyParagraphIndices(in: body)
        guard toBodyParagraphIndex >= 0 && toBodyParagraphIndex < targetParagraphIndices.count,
              case .paragraph(let initialTarget) = body.children[targetParagraphIndices[toBodyParagraphIndex]] else {
            throw OMathSpliceError.targetParagraphOutOfRange(toBodyParagraphIndex)
        }

        // Phase 1 — XML validity only. Malformed XML has precedence over
        // namespace policy regardless of source order, so scan every source
        // and the relevant initial-target fragment before comparing any URI.
        for omath in extracted {
            try Self.validateOMathFragment(omath.xml)
        }
        if let existingTarget = OMathExtractor.extract(from: initialTarget).first {
            try Self.validateOMathFragment(existingTarget.xml)
        }

        // Phase 2 — namespace policy. Only a fully valid batch reaches this
        // phase; policy errors therefore cannot mask a later malformed item.
        for omath in extracted {
            try Self.checkNamespacePolicy(
                source: omath.xml,
                targetParagraph: initialTarget,
                policy: namespacePolicy
            )
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

        struct AnchorGroup {
            let anchor: DerivedContextAnchor
            var indices: [Int]
        }

        var groups: [AnchorGroup] = []
        for i in extracted.indices {
            guard let anchor = prefixContexts[i] else {
                throw OMathSpliceError.contextAnchorNotFound(
                    omathIndex: i,
                    snippet: ""
                )
            }

            if let groupIndex = groups.firstIndex(where: { $0.anchor == anchor }) {
                groups[groupIndex].indices.append(i)
            } else {
                groups.append(AnchorGroup(anchor: anchor, indices: [i]))
            }
        }

        var spliced = 0
        // Process anchor groups in source order so a later failure preserves
        // only earlier source groups. Within one shared boundary, apply from
        // right to left so the final OMath order remains source order.
        for group in groups {
            if group.anchor.boundaryPlacement == nil,
               !group.anchor.snippet.isEmpty,
               Self.resolveRunAnchor(
                   anchor: group.anchor.snippet,
                   instance: group.anchor.instance,
                   options: AnchorLookupOptions(),
                   in: initialTarget
               ) == nil {
                throw OMathSpliceError.contextAnchorNotFound(
                    omathIndex: group.indices[0],
                    snippet: group.anchor.snippet
                )
            }

            let position: OMathSplicePosition
            switch group.anchor.boundaryPlacement {
            case .start?:
                position = .atStart
            case .end?:
                position = .atEnd
            case nil:
                position = group.anchor.snippet.isEmpty
                    ? .atStart
                    : .afterText(
                        group.anchor.snippet,
                        instance: group.anchor.instance,
                        options: AnchorLookupOptions()
                    )
            }

            // Start/text insertion prepends at a shared boundary and therefore
            // applies right-to-left. End-lane insertion appends, so it must apply
            // left-to-right to retain source order.
            let applicationIndices = group.anchor.boundaryPlacement == .end
                ? group.indices
                : Array(group.indices.reversed())
            for i in applicationIndices {
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
                        snippet: group.anchor.snippet
                    )
                }
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
        try validateOMathFragment(source)
        guard let sourceURI = OMathNamespace.extractURI(from: source) else {
            throw OMathSpliceMalformedXMLError()
        }
        let sourcePrefix = OMathNamespace.extractPrefix(from: source) ?? ""

        // Find the first target OMath in the same joint serializer order used
        // by omathIndex. Only a paragraph with no OMath uses the standard m:
        // baseline; an explicit default namespace has prefix "".
        var targetURI: String = standardOMMLURI
        var targetPrefix: String = "m"
        if let existing = OMathExtractor.extract(from: targetParagraph).first {
            try validateOMathFragment(existing.xml)
            guard let uri = OMathNamespace.extractURI(from: existing.xml) else {
                throw OMathSpliceMalformedXMLError()
            }
            targetURI = uri
            targetPrefix = OMathNamespace.extractPrefix(from: existing.xml) ?? ""
        }

        if sourceURI != targetURI {
            throw OMathSpliceError.namespaceMismatch(sourceURI: sourceURI, targetURI: targetURI)
        }
        if policy == .strict && sourcePrefix != targetPrefix {
            throw OMathSpliceError.namespaceMismatch(sourceURI: sourceURI, targetURI: targetURI)
        }
    }

    /// XML admission primitive deliberately separated from namespace policy.
    /// Batch mode runs this for the entire source set first so a malformed item
    /// always wins over URI/prefix mismatches in earlier items.
    internal static func validateOMathFragment(_ xml: String) throws {
        guard OMathNamespace.isWellFormed(xml),
              OMathNamespace.extractURI(from: xml) != nil else {
            throw OMathSpliceMalformedXMLError()
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
            omathRun.paragraphBoundaryPlacement = .start
            omathRun.paragraphBoundaryOrder = prepareBoundaryInsertion(
                placement: .start,
                prepend: true,
                in: &targetPara
            )
            targetPara.runs.insert(omathRun, at: 0)

        case .atEnd:
            omathRun.paragraphBoundaryPlacement = .end
            omathRun.paragraphBoundaryOrder = prepareBoundaryInsertion(
                placement: .end,
                prepend: false,
                in: &targetPara
            )
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
            options: options,
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
        let utf16 = text.utf16
        let safeOffset = max(0, min(offset, utf16.count))
        let prefixStr: String
        let suffixStr: String
        if let pStr = String(utf16.prefix(safeOffset)),
           let sStr = String(utf16.suffix(utf16.count - safeOffset)) {
            prefixStr = pStr
            suffixStr = sStr
        } else {
            // Fallback: Character-level slicing (safe for ASCII / BMP).
            let chars = Array(text)
            let cap = min(safeOffset, chars.count)
            prefixStr = String(chars[0..<cap])
            suffixStr = String(chars[cap..<chars.count])
        }

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
        options: AnchorLookupOptions = AnchorLookupOptions(),
        in para: Paragraph
    ) -> RunAnchorResolution? {
        guard !anchor.isEmpty, instance >= 1 else { return nil }

        // Build (runIdx, runText, startGlobal) excluding OMath-bearing runs.
        var runSpans: [(runIdx: Int, text: String, startGlobal: Int)] = []
        var globalOffset = 0
        for (i, run) in para.runs.enumerated() {
            guard let visibleText = serializedTypedText(in: run) else { continue }
            runSpans.append((i, visibleText, globalOffset))
            globalOffset += visibleText.utf16.count
        }

        let combined = runSpans.map { $0.text }.joined()
        let anchorText = options.mathScriptInsensitive
            ? AnchorLookupOptions.canonicalizeMathScriptVariants(anchor)
            : anchor
        let anchorUtf16 = Array(anchorText.utf16)
        guard !anchorUtf16.isEmpty else { return nil }

        let combinedUtf16: [UInt16]
        let originalStarts: [Int]
        let originalEnds: [Int]
        if options.mathScriptInsensitive {
            let normalized = normalizedUTF16WithOriginalOffsets(combined)
            combinedUtf16 = normalized.units
            originalStarts = normalized.originalStarts
            originalEnds = normalized.originalEnds
        } else {
            combinedUtf16 = Array(combined.utf16)
            originalStarts = Array(combinedUtf16.indices)
            originalEnds = combinedUtf16.indices.map { $0 + 1 }
        }
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
                    let globalStart = originalStarts[searchStart]
                    let globalEnd = originalEnds[searchStart + anchorUtf16.count - 1]

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

    /// Typed `Run.text` participates in anchors only when `Run.toXML()` would
    /// actually serialize it. Raw run overrides, raw property overrides, and
    /// drawings all replace the typed text branch; matching their hidden text
    /// would split and duplicate opaque content.
    private static func serializedTypedText(in run: Run) -> String? {
        guard run.rawXML == nil,
              run.properties.rawXML == nil,
              run.drawing == nil else {
            return nil
        }
        return run.text
    }

    /// Normalize one extended grapheme at a time and retain a map from every
    /// normalized UTF-16 unit back to the original UTF-16 span. This lets
    /// math-script-insensitive lookup still split the original Run at a valid
    /// boundary after folding subscripts, superscripts, and combining marks.
    private static func normalizedUTF16WithOriginalOffsets(
        _ text: String
    ) -> (units: [UInt16], originalStarts: [Int], originalEnds: [Int]) {
        var units: [UInt16] = []
        var starts: [Int] = []
        var ends: [Int] = []
        var originalOffset = 0

        for character in text {
            let original = String(character)
            let originalLength = original.utf16.count
            let normalized = AnchorLookupOptions.canonicalizeMathScriptVariants(original)
            for unit in normalized.utf16 {
                units.append(unit)
                starts.append(originalOffset)
                ends.append(originalOffset + originalLength)
            }
            originalOffset += originalLength
        }

        return (units, starts, ends)
    }

    /// Splice OMath as a direct child of `<w:p>` via `unrecognizedChildren`.
    private static func spliceAsDirectChild(
        into targetPara: inout Paragraph,
        omath: ExtractedOMath,
        position: OMathSplicePosition
    ) throws {
        switch position {
        case .atStart:
            var child = UnrecognizedChild(
                name: omath.directChildName ?? "oMath",
                rawXML: omath.xml,
                position: nil
            )
            child.paragraphBoundaryPlacement = .start
            child.paragraphBoundaryOrder = prepareBoundaryInsertion(
                placement: .start,
                prepend: true,
                in: &targetPara
            )
            targetPara.unrecognizedChildren.append(child)
            return

        case .atEnd:
            var child = UnrecognizedChild(
                name: omath.directChildName ?? "oMath",
                rawXML: omath.xml,
                position: nil
            )
            child.paragraphBoundaryPlacement = .end
            child.paragraphBoundaryOrder = prepareBoundaryInsertion(
                placement: .end,
                prepend: false,
                in: &targetPara
            )
            targetPara.unrecognizedChildren.append(child)
            return

        case .afterText(let anchor, let instance, let options):
            try insertDirectChildAtAnchor(
                anchor: anchor,
                instance: instance,
                options: options,
                side: .after,
                omath: omath,
                in: &targetPara
            )
            return

        case .beforeText(let anchor, let instance, let options):
            try insertDirectChildAtAnchor(
                anchor: anchor,
                instance: instance,
                options: options,
                side: .before,
                omath: omath,
                in: &targetPara
            )
            return
        }
    }

    private static func insertDirectChildAtAnchor(
        anchor: String,
        instance: Int,
        options: AnchorLookupOptions,
        side: AnchorSide,
        omath: ExtractedOMath,
        in para: inout Paragraph
    ) throws {
        _ = canonicalizeParagraphCarriers(in: &para, startingAt: 1)
        guard let resolved = resolveRunAnchor(
            anchor: anchor,
            instance: instance,
            options: options,
            in: para
        ) else {
            throw OMathSpliceError.anchorNotFound(anchor, instance: instance)
        }

        let splitRunIndex = side == .after ? resolved.endRunIdx : resolved.startRunIdx
        let offset = side == .after ? resolved.endOffsetInEndRun : resolved.startOffsetInStartRun
        let original = para.runs[splitRunIndex]
        let originalPosition = original.position ?? 1
        let (prefix, suffix) = splitRun(original, atCharOffset: offset)
        let hasPrefix = !prefix.text.isEmpty || prefix.rawXML != nil
        let hasSuffix = !suffix.text.isEmpty || suffix.rawXML != nil
        let componentCount = (hasPrefix ? 1 : 0) + 1 + (hasSuffix ? 1 : 0)
        shiftCarrierPositions(
            after: originalPosition,
            by: componentCount - 1,
            in: &para
        )

        var cursor = originalPosition
        var replacementRuns: [Run] = []
        if hasPrefix {
            var positionedPrefix = prefix
            positionedPrefix.position = cursor
            replacementRuns.append(positionedPrefix)
            cursor += 1
        }

        let childPosition = cursor
        cursor += 1

        if hasSuffix {
            var positionedSuffix = suffix
            positionedSuffix.position = cursor
            replacementRuns.append(positionedSuffix)
        }

        var runs = para.runs
        runs.replaceSubrange(splitRunIndex...splitRunIndex, with: replacementRuns)
        para.runs = runs
        para.unrecognizedChildren.append(
            UnrecognizedChild(
                name: omath.directChildName ?? "oMath",
                rawXML: omath.xml,
                position: childPosition
            )
        )
    }

    private enum ParagraphCarrierReference: Hashable {
        case run(Int)
        case hyperlink(Int)
        case fieldSimple(Int)
        case alternateContent(Int)
        case bookmarkMarker(Int)
        case commentMarker(Int)
        case permissionMarker(Int)
        case proofError(Int)
        case smartTag(Int)
        case customXml(Int)
        case bidiOverride(Int)
        case unrecognized(Int)
        case contentControl(Int)
    }

    /// Normalize only existing absolute-boundary OMath carriers and reserve
    /// an order for the new carrier. The normal paragraph arrays remain the
    /// typed source of truth; no page-break/note/bookmark state is rewritten.
    private static func prepareBoundaryInsertion(
        placement: ParagraphBoundaryPlacement,
        prepend: Bool,
        in para: inout Paragraph
    ) -> Int {
        var refs: [(order: Int, stable: Int, ref: ParagraphCarrierReference)] = []
        var stable = 0
        for index in para.runs.indices where para.runs[index].paragraphBoundaryPlacement == placement {
            refs.append((para.runs[index].paragraphBoundaryOrder ?? 0, stable, .run(index)))
            stable += 1
        }
        for index in para.unrecognizedChildren.indices
            where para.unrecognizedChildren[index].paragraphBoundaryPlacement == placement {
            refs.append((
                para.unrecognizedChildren[index].paragraphBoundaryOrder ?? 0,
                stable,
                .unrecognized(index)
            ))
            stable += 1
        }
        refs.sort {
            $0.order == $1.order ? $0.stable < $1.stable : $0.order < $1.order
        }

        for (offset, entry) in refs.enumerated() {
            setBoundaryOrder(entry.ref, to: prepend ? offset + 1 : offset, in: &para)
        }
        return prepend ? 0 : refs.count
    }

    private static func setBoundaryOrder(
        _ ref: ParagraphCarrierReference,
        to order: Int,
        in para: inout Paragraph
    ) {
        switch ref {
        case .run(let index):
            var runs = para.runs
            runs[index].paragraphBoundaryOrder = order
            para.runs = runs
        case .unrecognized(let index):
            para.unrecognizedChildren[index].paragraphBoundaryOrder = order
        default:
            break
        }
    }

    /// Compact the 13 position-indexed collections in their current serialized
    /// order. Used only when a direct-child mid-text splice needs a unique
    /// cross-collection slot; absolute boundaries use the separate serializer
    /// lane and never rewrite legacy typed state or existing positions.
    @discardableResult
    private static func canonicalizeParagraphCarriers(
        in para: inout Paragraph,
        startingAt firstPosition: Int
    ) -> Int {
        var positive: [(position: Int, stableOrder: Int, ref: ParagraphCarrierReference)] = []
        var postContent: [ParagraphCarrierReference] = []
        var stableOrder = 0

        func collect(_ refs: [ParagraphCarrierReference], postContentOrder: Bool = false) {
            for ref in refs {
                let position = carrierPosition(ref, in: para) ?? 0
                if postContentOrder {
                    if position <= 0 { postContent.append(ref) }
                } else if position > 0 {
                    positive.append((position, stableOrder, ref))
                    stableOrder += 1
                }
            }
        }

        // Positive-position collection order mirrors Paragraph's sorted-list
        // builder and supplies an explicit stable tie-breaker.
        collect(para.runs.indices.map(ParagraphCarrierReference.run))
        collect(para.hyperlinks.indices.map(ParagraphCarrierReference.hyperlink))
        collect(para.fieldSimples.indices.map(ParagraphCarrierReference.fieldSimple))
        collect(para.alternateContents.indices.map(ParagraphCarrierReference.alternateContent))
        collect(para.bookmarkMarkers.indices.map(ParagraphCarrierReference.bookmarkMarker))
        collect(para.commentRangeMarkers.indices.map(ParagraphCarrierReference.commentMarker))
        collect(para.permissionRangeMarkers.indices.map(ParagraphCarrierReference.permissionMarker))
        collect(para.proofErrorMarkers.indices.map(ParagraphCarrierReference.proofError))
        collect(para.smartTags.indices.map(ParagraphCarrierReference.smartTag))
        collect(para.customXmlBlocks.indices.map(ParagraphCarrierReference.customXml))
        collect(para.bidiOverrides.indices.map(ParagraphCarrierReference.bidiOverride))
        collect(para.unrecognizedChildren.indices.map(ParagraphCarrierReference.unrecognized))
        collect(para.contentControls.indices.map(ParagraphCarrierReference.contentControl))

        // Nil/zero/negative entries follow Paragraph's legacy post-content
        // bucket order. Negative positions are normalized instead of silently
        // disappearing from both serializer predicates.
        collect(para.contentControls.indices.map(ParagraphCarrierReference.contentControl), postContentOrder: true)
        collect(para.runs.indices.map(ParagraphCarrierReference.run), postContentOrder: true)
        collect(para.hyperlinks.indices.map(ParagraphCarrierReference.hyperlink), postContentOrder: true)
        collect(para.fieldSimples.indices.map(ParagraphCarrierReference.fieldSimple), postContentOrder: true)
        collect(para.alternateContents.indices.map(ParagraphCarrierReference.alternateContent), postContentOrder: true)
        collect(para.bookmarkMarkers.indices.map(ParagraphCarrierReference.bookmarkMarker), postContentOrder: true)
        collect(para.commentRangeMarkers.indices.map(ParagraphCarrierReference.commentMarker), postContentOrder: true)
        collect(para.permissionRangeMarkers.indices.map(ParagraphCarrierReference.permissionMarker), postContentOrder: true)
        collect(para.proofErrorMarkers.indices.map(ParagraphCarrierReference.proofError), postContentOrder: true)
        collect(para.smartTags.indices.map(ParagraphCarrierReference.smartTag), postContentOrder: true)
        collect(para.customXmlBlocks.indices.map(ParagraphCarrierReference.customXml), postContentOrder: true)
        collect(para.bidiOverrides.indices.map(ParagraphCarrierReference.bidiOverride), postContentOrder: true)
        collect(para.unrecognizedChildren.indices.map(ParagraphCarrierReference.unrecognized), postContentOrder: true)

        positive.sort {
            $0.position == $1.position
                ? $0.stableOrder < $1.stableOrder
                : $0.position < $1.position
        }
        let ordered = positive.map(\.ref) + postContent

        var next = firstPosition
        for ref in ordered {
            setCarrierPosition(ref, to: next, in: &para)
            if next < Int.max { next += 1 }
        }
        return next
    }

    private static func carrierPosition(
        _ ref: ParagraphCarrierReference,
        in para: Paragraph
    ) -> Int? {
        switch ref {
        case .run(let i): return para.runs[i].position
        case .hyperlink(let i): return para.hyperlinks[i].position
        case .fieldSimple(let i): return para.fieldSimples[i].position
        case .alternateContent(let i): return para.alternateContents[i].position
        case .bookmarkMarker(let i): return para.bookmarkMarkers[i].position
        case .commentMarker(let i): return para.commentRangeMarkers[i].position
        case .permissionMarker(let i): return para.permissionRangeMarkers[i].position
        case .proofError(let i): return para.proofErrorMarkers[i].position
        case .smartTag(let i): return para.smartTags[i].position
        case .customXml(let i): return para.customXmlBlocks[i].position
        case .bidiOverride(let i): return para.bidiOverrides[i].position
        case .unrecognized(let i): return para.unrecognizedChildren[i].position
        case .contentControl(let i): return para.contentControls[i].position
        }
    }

    private static func setCarrierPosition(
        _ ref: ParagraphCarrierReference,
        to position: Int,
        in para: inout Paragraph
    ) {
        switch ref {
        case .run(let i):
            var runs = para.runs
            runs[i].position = position
            para.runs = runs
        case .hyperlink(let i): para.hyperlinks[i].position = position
        case .fieldSimple(let i): para.fieldSimples[i].position = position
        case .alternateContent(let i): para.alternateContents[i].position = position
        case .bookmarkMarker(let i): para.bookmarkMarkers[i].position = position
        case .commentMarker(let i): para.commentRangeMarkers[i].position = position
        case .permissionMarker(let i): para.permissionRangeMarkers[i].position = position
        case .proofError(let i): para.proofErrorMarkers[i].position = position
        case .smartTag(let i): para.smartTags[i].position = position
        case .customXml(let i): para.customXmlBlocks[i].position = position
        case .bidiOverride(let i): para.bidiOverrides[i].position = position
        case .unrecognized(let i): para.unrecognizedChildren[i].position = position
        case .contentControl(let i): para.contentControls[i].position = position
        }
    }

    private static func shiftCarrierPositions(
        after boundary: Int,
        by delta: Int,
        in para: inout Paragraph
    ) {
        guard delta > 0 else { return }
        let refs = allCarrierReferences(in: para)
        for ref in refs {
            guard let position = carrierPosition(ref, in: para), position > boundary else { continue }
            let shifted = position.addingReportingOverflow(delta)
            setCarrierPosition(ref, to: shifted.overflow ? Int.max : shifted.partialValue, in: &para)
        }
    }

    private static func allCarrierReferences(in para: Paragraph) -> [ParagraphCarrierReference] {
        para.runs.indices.map(ParagraphCarrierReference.run)
            + para.hyperlinks.indices.map(ParagraphCarrierReference.hyperlink)
            + para.fieldSimples.indices.map(ParagraphCarrierReference.fieldSimple)
            + para.alternateContents.indices.map(ParagraphCarrierReference.alternateContent)
            + para.bookmarkMarkers.indices.map(ParagraphCarrierReference.bookmarkMarker)
            + para.commentRangeMarkers.indices.map(ParagraphCarrierReference.commentMarker)
            + para.permissionRangeMarkers.indices.map(ParagraphCarrierReference.permissionMarker)
            + para.proofErrorMarkers.indices.map(ParagraphCarrierReference.proofError)
            + para.smartTags.indices.map(ParagraphCarrierReference.smartTag)
            + para.customXmlBlocks.indices.map(ParagraphCarrierReference.customXml)
            + para.bidiOverrides.indices.map(ParagraphCarrierReference.bidiOverride)
            + para.unrecognizedChildren.indices.map(ParagraphCarrierReference.unrecognized)
            + para.contentControls.indices.map(ParagraphCarrierReference.contentControl)
    }

    internal struct DerivedContextAnchor: Equatable {
        let snippet: String
        let instance: Int
        let boundaryPlacement: ParagraphBoundaryPlacement?

        init(
            snippet: String,
            instance: Int,
            boundaryPlacement: ParagraphBoundaryPlacement? = nil
        ) {
            self.snippet = snippet
            self.instance = instance
            self.boundaryPlacement = boundaryPlacement
        }
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
    ) -> [DerivedContextAnchor?] {
        var anchors = [DerivedContextAnchor?](repeating: nil, count: extracted.count)
        var matchedExtractedIndices = Set<Int>()
        var proseAccumulator = ""

        // Boundary lanes are authoritative metadata, not prose-derived hints.
        // Preserve both start and end so an empty-prose end source cannot be
        // reclassified as a leading equation during batch reuse before reload.
        for (i, ex) in extracted.enumerated() {
            if let placement = ex.boundaryPlacement {
                anchors[i] = DerivedContextAnchor(
                    snippet: "",
                    instance: 1,
                    boundaryPlacement: placement
                )
                matchedExtractedIndices.insert(i)
            }
        }

        // Walk runs in array order; whenever a run's rawXML matches an extracted OMath,
        // capture the trailing N chars of proseAccumulator.
        for run in sourceParagraph.runs {
            if let raw = run.rawXML, OMathNamespace.hasOMathRoot(in: raw) {
                // Find which extracted index this run corresponds to (by xml byte equality).
                for (i, ex) in extracted.enumerated() where !matchedExtractedIndices.contains(i) {
                    if ex.kind == .inRun
                        && ex.xml == OMathExtractor.ensureXmlnsDeclared(in: raw) {
                        let trailing = String(proseAccumulator.suffix(charsBefore)).trimmingCharacters(in: .whitespaces)
                        anchors[i] = DerivedContextAnchor(
                            snippet: trailing,
                            instance: trailing.isEmpty
                                ? 1
                                : occurrenceCount(of: trailing, in: proseAccumulator)
                        )
                        matchedExtractedIndices.insert(i)
                        break
                    }
                }
                continue
            }
            if let visibleText = serializedTypedText(in: run) {
                proseAccumulator += visibleText
            }
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
                    if let visibleText = serializedTypedText(in: run) {
                        prose += visibleText
                    }
                }
            }
            let trailing = String(prose.suffix(charsBefore)).trimmingCharacters(in: .whitespaces)
            anchors[i] = DerivedContextAnchor(
                snippet: trailing,
                instance: trailing.isEmpty ? 1 : occurrenceCount(of: trailing, in: prose)
            )
        }

        return anchors
    }

    private static func occurrenceCount(of needle: String, in haystack: String) -> Int {
        guard !needle.isEmpty else { return 1 }
        var count = 0
        var searchStart = haystack.startIndex
        while searchStart < haystack.endIndex,
              let range = haystack.range(of: needle, range: searchStart..<haystack.endIndex) {
            count += 1
            searchStart = haystack.index(after: range.lowerBound)
        }
        return max(count, 1)
    }
}

// MARK: - String UTF-16 helpers

private extension String {
    /// Construct a String from a UTF-16 unit sequence (e.g. `String.UTF16View.SubSequence`)
    /// via NSString bridge. Returns nil for empty sequences only when explicitly empty;
    /// the bridge always succeeds for finite input.
    init?<S: Sequence>(_ utf16: S) where S.Element == UInt16 {
        let array = Array(utf16)
        let ns = NSString(characters: array, length: array.count)
        self = ns as String
    }
}
