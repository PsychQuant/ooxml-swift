// ScriptTranscoder.swift
// word-aligned-state-sync Phase 4 tasks 5.1 / 5.2 / 5.4 — the bidirectional
// codec between `OperationLog` and canonical `.mdocx` Swift source
// (`ooxml-script-transcode` capability; surface per `mdocx-grammar`,
// op wire format per `ooxml-operation-log` — the single normative home).
//
// Projection contract (v0.34 vertical slice):
//
// - DSL form covers the bijective authoring subset: a swift-sourced
//   `appendParagraph(in: nil, payload-with-paraId)` becomes
//   `Paragraph(id: "...", style: .x) { "text" }` inside a synthesized
//   `Section(id: "main")` envelope. The Section wrapper is pure syntax —
//   it emits no operation, so log → script → log stays equivalent.
// - EVERY other op (mutations, word-sourced ops, component envelopes,
//   inline atoms, tables, unknown future ops, …) round-trips via the raw
//   escape line `// @op {canonical JSON fields}` — the forward-compat
//   mechanism the spec mandates for unknown op_types, extended to all
//   DSL-unrepresentable ops so round-trip holds for ARBITRARY logs.
//   Richer DSL-form coverage (setRuns bodies, components, tables) grows
//   in later Phase 4 iterations without breaking this contract.
// - Deterministic formatting: one op ↔ one block/line in log order,
//   4-space indent, no timestamps — adding one op to the log produces
//   exactly one inserted hunk in the exported source (task 5.4).
//
// The importer is a hand-written line-based parser over the CANONICAL
// grammar subset only. It is NOT a Swift compiler: any construct outside
// the canonical form (arbitrary expressions, side effects, unknown
// elements) throws `TranscodeError.unsupportedSyntax(line:column:reason:)`
// with a precise 1-based location.

import Foundation

// MARK: - TranscodeError

public enum TranscodeError: Error, Equatable {
    /// The source contains a construct outside the canonical exporter
    /// grammar. `line`/`column` are 1-based.
    case unsupportedSyntax(line: Int, column: Int, reason: String)
    /// A `// @op` raw line failed to parse back into an Operation.
    case malformedRawOp(line: Int, reason: String)
    /// A slot designation could not be honored (format-alignment-engine
    /// Phase D task 4.1). Strict mode never guesses: an unusable designation
    /// fails loudly instead of silently degrading to verbatim content.
    case slotDesignationFailure(name: String, reason: String)
}

// MARK: - SlotDesignation (format-alignment-engine Phase D, task 4.1)

/// A named content slot: `name` becomes a Swift function parameter of the
/// emitted script; `paraId` designates the paragraph whose text the slot
/// replaces (`template-content-slots`, «Slot designation is explicit in
/// strict mode» — no inference).
public struct SlotDesignation: Equatable, Sendable {
    public let name: String
    public let paraId: String

    public init(name: String, paraId: String) {
        self.name = name
        self.paraId = paraId
    }
}

// MARK: - WordStyleMap

/// Bidirectional mapping between OOXML `styleId` strings and `WordStyle`
/// member names (`mdocx-grammar`: "Style references via typed enum").
/// Predefined pairs follow the spec's own example (`Heading1` ↔ `.heading1`);
/// everything else maps verbatim (member name == styleId) when the styleId
/// is a valid Swift identifier that doesn't collide with a predefined member.
public enum WordStyleMap {

    public static let predefined: [String: String] = {
        var map: [String: String] = [:]   // member -> styleId
        for i in 1...9 { map["heading\(i)"] = "Heading\(i)" }
        map["quote"] = "Quote"
        map["listItem"] = "ListItem"
        map["title"] = "Title"
        map["subtitle"] = "Subtitle"
        return map
    }()

    private static let identifierPattern = try! NSRegularExpression(
        pattern: "^[A-Za-z_][A-Za-z0-9_]*$")

    public static func isIdentifier(_ s: String) -> Bool {
        identifierPattern.firstMatch(
            in: s, range: NSRange(s.startIndex..., in: s)) != nil
    }

    /// styleId → `.member` name, or nil when the styleId has no lossless
    /// DSL spelling (caller falls back to the raw `// @op` escape).
    public static func member(forStyleId styleId: String) -> String? {
        if let hit = predefined.first(where: { $0.value == styleId }) {
            return hit.key
        }
        guard isIdentifier(styleId) else { return nil }
        // A verbatim member must not collide with a predefined member,
        // or the reverse mapping would rewrite it (e.g. styleId "heading1"
        // verbatim would read back as "Heading1").
        guard predefined[styleId] == nil else { return nil }
        return styleId
    }

    /// `.member` name → verbatim styleId.
    public static func styleId(forMember member: String) -> String {
        predefined[member] ?? member
    }
}

// MARK: - ScriptExporter (task 5.1)

public enum ScriptExporter {

    /// Emits canonical `.mdocx` Swift source for the log. Deterministic:
    /// the same log always produces byte-identical output.
    /// Slot-parameterized export (format-alignment-engine Phase D task 4.1,
    /// design Q2 working answer: Swift function parameters). The designated
    /// paragraphs' text becomes function parameters — `func makeDocument(…)`
    /// with the extracted content as the call-site arguments — while every
    /// other op is emitted exactly as the canonical form. Empty `slots`
    /// delegates to the canonical exporter (no-designation invariant, 4.2).
    ///
    /// Strict-mode failures (unknown paraId, non-DSL-spellable paragraph,
    /// invalid slot name, duplicates) throw `slotDesignationFailure`.
    public static func exportSwift(log: OperationLog,
                                   slots: [SlotDesignation]) throws -> String {
        guard !slots.isEmpty else { return exportSwift(log: log) }

        // Validate designations up front — strict mode fails loudly.
        var slotByParaId: [String: SlotDesignation] = [:]
        var seenNames: Set<String> = []
        for slot in slots {
            guard WordStyleMap.isIdentifier(slot.name),
                  slot.name.first.map({ $0.isLowercase || $0 == "_" }) == true else {
                throw TranscodeError.slotDesignationFailure(
                    name: slot.name, reason: "slot name must be a lowercase Swift identifier")
            }
            guard seenNames.insert(slot.name).inserted else {
                throw TranscodeError.slotDesignationFailure(
                    name: slot.name, reason: "duplicate slot name")
            }
            guard slotByParaId.updateValue(slot, forKey: slot.paraId) == nil else {
                throw TranscodeError.slotDesignationFailure(
                    name: slot.name, reason: "paragraph \(slot.paraId) designated twice")
            }
            // The designated paragraph must exist and be DSL-spellable —
            // a slot inside a raw `// @op` line would not be substitutable.
            let target = log.entries.first { entry in
                if case .appendParagraph(let c, let payload) = entry.op, c == nil,
                   payload.paraId == slot.paraId { return true }
                return false
            }
            guard let targetEntry = target,
                  case .appendParagraph(_, let payload) = targetEntry.op else {
                throw TranscodeError.slotDesignationFailure(
                    name: slot.name,
                    reason: "no body paragraph with id \(slot.paraId) in the log")
            }
            guard paragraphBlock(payload: payload, paraId: slot.paraId, indent: 8) != nil else {
                throw TranscodeError.slotDesignationFailure(
                    name: slot.name,
                    reason: "paragraph \(slot.paraId) has no DSL spelling (extended formatting rides the raw escape)")
            }
        }

        let body = emitBody(entries: log.entries, slotByParaId: slotByParaId)

        // Defaults = the extracted content (call site executes the verbatim
        // rebuild; callers substitute new values for new content).
        var defaults: [String: String] = [:]
        for entry in log.entries {
            if case .appendParagraph(let c, let payload) = entry.op, c == nil,
               let paraId = payload.paraId, let slot = slotByParaId[paraId] {
                defaults[slot.name] = payload.text
            }
        }

        var out: [String] = []
        out.append("// Generated by ScriptExporter (word-aligned-state-sync Phase 4).")
        out.append("// Canonical .mdocx form — round-trips via ScriptImporter.parse(source:).")
        out.append("// Slotted template (format-alignment-engine Phase D): named content")
        out.append("// slots are function parameters; call-site arguments carry the content.")
        out.append("import WordDSLSwift")
        out.append("")
        out.append("func makeDocument(")
        for (idx, slot) in slots.enumerated() {
            let comma = idx == slots.count - 1 ? "" : ","
            out.append("    \(slot.name): String\(comma)")
        }
        out.append(") -> WordDocument {")
        out.append("    WordDocument {")
        out.append("    Section(id: \"main\") {")
        out.append(contentsOf: body)
        out.append("    }")
        out.append("    }")
        out.append("}")
        out.append("")
        out.append("let document = makeDocument(")
        for (idx, slot) in slots.enumerated() {
            let comma = idx == slots.count - 1 ? "" : ","
            out.append("    \(slot.name): \(quote(defaults[slot.name] ?? ""))\(comma)")
        }
        out.append(")")
        out.append("")
        return out.joined(separator: "\n")
    }

    public static func exportSwift(log: OperationLog) -> String {
        let body = emitBody(entries: log.entries, slotByParaId: [:])
        var out: [String] = []
        out.append("// Generated by ScriptExporter (word-aligned-state-sync Phase 4).")
        out.append("// Canonical .mdocx form — round-trips via ScriptImporter.parse(source:).")
        out.append("import WordDSLSwift")
        out.append("")
        out.append("let document = WordDocument {")
        out.append("    Section(id: \"main\") {")
        out.append(contentsOf: body)
        out.append("    }")
        out.append("}")
        out.append("")
        return out.joined(separator: "\n")
    }

    /// Shared body emission for the canonical and slotted forms. A paragraph
    /// whose paraId is designated emits its body as the slot's bare
    /// identifier instead of a quoted string literal.
    private static func emitBody(entries: [LogEntry],
                                 slotByParaId: [String: SlotDesignation]) -> [String] {
        var body: [String] = []

        var i = 0
        while i < entries.count {
            let entry = entries[i]

            // Component envelope projection: beginComponent … endComponent
            // becomes `<Type>(id: "…") { … }` (mdocx-grammar "reverse
            // direction reconstructs component invocation"). Nested
            // envelopes fall back to raw lines (post-v0.34 refinement).
            if case .beginComponent(let type, let id) = entry.op,
               entry.source == .swift,
               WordStyleMap.isIdentifier(type),
               let endIdx = entries[(i + 1)...].firstIndex(where: {
                   if case .endComponent(let eid) = $0.op { return eid == id }
                   return false
               }),
               !entries[(i + 1)..<endIdx].contains(where: {
                   if case .beginComponent = $0.op { return true }
                   return false
               }) {
                body.append("        \(type)(id: \(quote(id.raw))) {")
                for inner in entries[(i + 1)..<endIdx] {
                    if case .appendParagraph(let c, let payload) = inner.op,
                       c == nil, inner.source == .swift,
                       let paraId = payload.paraId, !paraId.isEmpty,
                       let block = paragraphBlock(payload: payload, paraId: paraId, indent: 12,
                                                  slotName: slotByParaId[paraId]?.name) {
                        body.append(contentsOf: block)
                    } else {
                        body.append("            " + rawOpLine(entry: inner))
                    }
                }
                body.append("        }")
                i = endIdx + 1
                continue
            }

            if case .appendParagraph(let container, let payload) = entry.op,
               container == nil,
               entry.source == .swift,
               let paraId = payload.paraId, !paraId.isEmpty,
               let block = paragraphBlock(payload: payload, paraId: paraId, indent: 8,
                                          slotName: slotByParaId[paraId]?.name) {
                body.append(contentsOf: block)
            } else {
                body.append("        " + rawOpLine(entry: entry))
            }
            i += 1
        }

        return body
    }

    /// DSL-form paragraph block, or nil when the payload has no lossless
    /// DSL spelling (non-identifier styleId etc.). A non-nil `slotName`
    /// spells the body as the slot's bare identifier (a function parameter
    /// reference) instead of the quoted text literal (Phase D task 4.1).
    private static func paragraphBlock(payload: ParagraphPayload, paraId: String,
                                       indent: Int, slotName: String? = nil) -> [String]? {
        // format-alignment-engine Phase B: the DSL block spells only
        // id/style/text. A payload carrying any extended pPr field
        // (alignment/spacing/indent/numPr) has no lossless DSL spelling yet —
        // fall back to the `// @op` escape so the fields survive round-trip.
        guard payload.alignment == nil,
              payload.spacingBefore == nil, payload.spacingAfter == nil,
              payload.spacingLine == nil, payload.spacingLineRule == nil,
              payload.indentLeft == nil, payload.indentRight == nil,
              payload.indentFirstLine == nil, payload.indentHanging == nil,
              payload.numId == nil, payload.numLevel == nil,
              // word-canonical-forms task 2.2/3.1: Word-authored paragraph
              // attrs + pPr/rPr have no lossless DSL spelling — fall back to
              // the // @op escape.
              payload.textId == nil, payload.rsidR == nil, payload.rsidRPr == nil,
              payload.rsidRDefault == nil, payload.rsidP == nil,
              payload.indentFirstLineChars == nil, payload.indentHangingChars == nil,
              payload.paragraphMarkRun == nil else { return nil }
        let pad = String(repeating: " ", count: indent)
        var head = "\(pad)Paragraph(id: \(quote(paraId))"
        if let styleId = payload.styleId {
            guard let member = WordStyleMap.member(forStyleId: styleId) else { return nil }
            head += ", style: .\(member)"
        }
        head += ") {"
        return [head,
                "\(pad)    \(slotName ?? quote(payload.text))",
                "\(pad)}"]
    }

    /// `// @op {"op_type":...,"source":...,<op fields>}` — canonical raw
    /// escape reusing the JSONL codec's field encoding (single source of
    /// truth for op shapes; op_id/timestamp regenerate on import per the
    /// round-trip contract).
    private static func rawOpLine(entry: LogEntry) -> String {
        let (opType, fields) = JSONLLineCoder.encodeOp(entry.op)
        var parts: [String] = []
        parts.append("\"op_type\":\(quote(opType))")
        parts.append("\"source\":\(quote(entry.source.rawValue))")
        for (key, value) in fields {
            parts.append("\"\(key)\":\(value)")
        }
        return "// @op {\(parts.joined(separator: ","))}"
    }

    static func quote(_ s: String) -> String {
        var escaped = ""
        for ch in s {
            switch ch {
            case "\\": escaped += "\\\\"
            case "\"": escaped += "\\\""
            case "\n": escaped += "\\n"
            case "\r": escaped += "\\r"
            case "\t": escaped += "\\t"
            default: escaped.append(ch)
            }
        }
        return "\"\(escaped)\""
    }
}

// MARK: - ScriptImporter (task 5.2)

public enum ScriptImporter {

    private enum Scope: Equatable {
        case top, document, section, component(id: String), paragraph
        /// Slotted-template form (format-alignment-engine Phase D task 4.1):
        /// inside `func makeDocument(` parameter list.
        case slotSignature
        /// Inside the makeDocument function body (wraps a `WordDocument {`).
        case functionBody
        /// Inside the `let document = makeDocument(` argument list.
        case callSite
    }

    private struct ParagraphState {
        var paraId: String
        var styleId: String?
        var textParts: [String] = []
        /// Inline-atom ops collected from the body; flushed AFTER the
        /// paragraph's own appendParagraph op so the atoms' target exists
        /// at replay time.
        var pendingAtoms: [Operation] = []
        var line: Int
    }

    /// Parses canonical `.mdocx` source back into an `OperationLog`.
    /// Accepts exactly the exporter's grammar subset plus hand-written
    /// scripts of the same shape; anything else throws
    /// `TranscodeError.unsupportedSyntax` with a 1-based location.
    public static func parse(source: String) throws -> OperationLog {
        var log = OperationLog()
        var scopeStack: [Scope] = [.top]
        var paragraph: ParagraphState?

        // Slotted-template state (Phase D task 4.1): parameter names declared
        // by `func makeDocument(…)` and their call-site argument values.
        var slotNames: Set<String> = []
        let slotBindings = collectSlotBindings(source: source)

        let lines = source.components(separatedBy: "\n")
        for (idx, rawLine) in lines.enumerated() {
            let lineNo = idx + 1
            let line = rawLine.trimmingCharacters(in: .whitespaces)
            let column = (rawLine.count - rawLine.drop(while: { $0 == " " }).count) + 1

            if line.isEmpty { continue }

            // Raw op escape — decode via the JSONL codec (canonical shapes).
            if line.hasPrefix("// @op ") {
                let json = String(line.dropFirst("// @op ".count))
                let entry = try decodeRawOp(json: json, line: lineNo)
                log.append(entry.op, source: entry.source)
                continue
            }
            if line.hasPrefix("//") { continue }               // ordinary comment
            if line == "import WordDSLSwift" { continue }

            // Slotted-template form (Phase D task 4.1).
            if line == "func makeDocument(" {
                guard scopeStack.last == .top else {
                    throw TranscodeError.unsupportedSyntax(
                        line: lineNo, column: column, reason: "unexpected function declaration here")
                }
                scopeStack.append(.slotSignature)
                continue
            }
            if scopeStack.last == .slotSignature {
                if line == ") -> WordDocument {" {
                    scopeStack.removeLast()
                    scopeStack.append(.functionBody)
                    continue
                }
                if let m = match(line, #"^([a-z_][A-Za-z0-9_]*): String,?$"#) {
                    slotNames.insert(m[0])
                    continue
                }
                throw TranscodeError.unsupportedSyntax(
                    line: lineNo, column: column,
                    reason: "slot signature accepts only `name: String` parameters")
            }
            if line == "WordDocument {", scopeStack.last == .functionBody {
                scopeStack.append(.document)
                continue
            }
            if line == "let document = makeDocument(", scopeStack.last == .top {
                scopeStack.append(.callSite)
                continue
            }
            if scopeStack.last == .callSite {
                if line == ")" {
                    scopeStack.removeLast()
                    continue
                }
                if match(line, #"^[a-z_][A-Za-z0-9_]*: "(?:[^"\\]|\\.)*",?$"#) != nil {
                    continue  // bindings pre-collected by collectSlotBindings
                }
                throw TranscodeError.unsupportedSyntax(
                    line: lineNo, column: column,
                    reason: "call site accepts only `name: \"value\"` arguments")
            }

            switch line {
            case "let document = WordDocument {":
                guard scopeStack.last == .top else {
                    throw TranscodeError.unsupportedSyntax(
                        line: lineNo, column: column, reason: "unexpected WordDocument declaration here")
                }
                scopeStack.append(.document)
                continue
            case "}":
                if paragraph != nil {
                    // close of a Paragraph body
                    let p = paragraph!
                    log.append(.appendParagraph(in: nil, paragraph: ParagraphPayload(
                        text: p.textParts.joined(),
                        styleId: p.styleId,
                        paraId: p.paraId)), source: .swift)
                    for atom in p.pendingAtoms {
                        log.append(atom, source: .swift)
                    }
                    paragraph = nil
                    if scopeStack.last == .paragraph { scopeStack.removeLast() }
                } else if case .component(let cid)? = scopeStack.last {
                    log.append(.endComponent(id: ElementID(rawString: cid)), source: .swift)
                    scopeStack.removeLast()
                } else if scopeStack.count > 1 {
                    scopeStack.removeLast()
                } else {
                    throw TranscodeError.unsupportedSyntax(
                        line: lineNo, column: column, reason: "unbalanced closing brace")
                }
                continue
            default:
                break
            }

            if let m = match(line, #"^Section\(id: "((?:[^"\\]|\\.)*)"\) \{$"#) {
                guard scopeStack.last == .document else {
                    throw TranscodeError.unsupportedSyntax(
                        line: lineNo, column: column, reason: "Section must be a direct child of WordDocument")
                }
                _ = m   // section id is envelope-only in this slice (no op emitted)
                scopeStack.append(.section)
                continue
            }

            // Component invocation: `<Type>(id: "…") {` with a non-reserved
            // capitalized identifier inside a Section (mdocx-grammar
            // component reconstruction).
            if scopeStack.last == .section,
               let m = match(line, #"^([A-Z][A-Za-z0-9_]*)\(id: "((?:[^"\\]|\\.)*)"\) \{$"#),
               !["Paragraph", "Section", "Table", "TableRow", "TableCell",
                 "Hyperlink", "Bookmark", "WordDocument", "Run"].contains(m[0]) {
                let componentID = unescape(m[1])
                log.append(.beginComponent(type: m[0],
                                           id: ElementID(rawString: componentID)),
                           source: .swift)
                scopeStack.append(.component(id: componentID))
                continue
            }

            if line.hasPrefix("Paragraph") {
                guard scopeStack.last == .section || isComponentScope(scopeStack.last) else {
                    throw TranscodeError.unsupportedSyntax(
                        line: lineNo, column: column, reason: "Paragraph must appear inside a Section")
                }
                if let m = match(line, #"^Paragraph\(id: "((?:[^"\\]|\\.)*)"\) \{$"#) {
                    paragraph = ParagraphState(paraId: unescape(m[0]), styleId: nil, line: lineNo)
                    scopeStack.append(.paragraph)
                    continue
                }
                if let m = match(line, #"^Paragraph\(id: "((?:[^"\\]|\\.)*)", style: \.([A-Za-z_][A-Za-z0-9_]*)\) \{$"#) {
                    paragraph = ParagraphState(
                        paraId: unescape(m[0]),
                        styleId: WordStyleMap.styleId(forMember: m[1]),
                        line: lineNo)
                    scopeStack.append(.paragraph)
                    continue
                }
                if match(line, #"^Paragraph\(id: "(?:[^"\\]|\\.)*", style: "(?:[^"\\]|\\.)*"\) \{$"#) != nil {
                    throw TranscodeError.unsupportedSyntax(
                        line: lineNo, column: column,
                        reason: "raw-string style is rejected — style references require a typed WordStyle value (`style: .heading1`)")
                }
                throw TranscodeError.unsupportedSyntax(
                    line: lineNo, column: column,
                    reason: "Paragraph requires an explicit `id:` argument (mandatory explicit identifiers)")
            }

            if scopeStack.last == .paragraph, paragraph != nil {
                if let m = match(line, #"^"((?:[^"\\]|\\.)*)"$"#) {
                    paragraph!.textParts.append(unescape(m[0]))
                    continue
                }
                // Inline atoms as standalone children (`mdocx-grammar`
                // "Special-character inline atoms"). The paragraph itself is
                // the op target — the reducer synthesizes the wrapping
                // `<w:r>` (rule pinned in `ooxml-operation-log`, §4b).
                // NOTE: the paragraph's appendParagraph op is emitted at the
                // closing brace, AFTER these atom ops — the atoms queue on
                // the pending-atom list and flush after the paragraph op so
                // replay order stays applicable.
                let atomOps: [String: (ElementID) -> Operation] = [
                    "Tab()": { .insertTab(in: $0) },
                    "Break()": { .insertBreak(in: $0) },
                    "NoBreakHyphen()": { .insertNoBreakHyphen(in: $0) },
                ]
                if let makeOp = atomOps[line] {
                    let target = ElementID(rawString: "w14:paraId=\(paragraph!.paraId)")
                    paragraph!.pendingAtoms.append(makeOp(target))
                    continue
                }
                // Slot reference (Phase D task 4.1): a bare identifier that
                // names a declared function parameter resolves to its
                // call-site argument value.
                if match(line, #"^[a-z_][A-Za-z0-9_]*$"#) != nil, slotNames.contains(line) {
                    guard let value = slotBindings[line] else {
                        throw TranscodeError.unsupportedSyntax(
                            line: lineNo, column: column,
                            reason: "slot `\(line)` has no call-site argument value")
                    }
                    paragraph!.textParts.append(value)
                    continue
                }
                throw TranscodeError.unsupportedSyntax(
                    line: lineNo, column: column,
                    reason: "only String literals and inline atoms (Tab()/Break()/NoBreakHyphen()) are accepted in this paragraph body (canonical subset)")
            }

            throw TranscodeError.unsupportedSyntax(
                line: lineNo, column: column,
                reason: "construct outside the canonical .mdocx transcoder grammar: \(line.prefix(60))")
        }

        guard scopeStack == [.top] else {
            throw TranscodeError.unsupportedSyntax(
                line: lines.count, column: 1, reason: "unterminated block at end of source")
        }
        return log
    }

    // MARK: helpers

    /// Pre-pass over the source collecting `let document = makeDocument(…)`
    /// call-site arguments: `name: "value"` lines between the call open and
    /// its closing `)`. Returns an empty map for non-slotted scripts.
    private static func collectSlotBindings(source: String) -> [String: String] {
        var bindings: [String: String] = [:]
        var inCallSite = false
        for rawLine in source.components(separatedBy: "\n") {
            let line = rawLine.trimmingCharacters(in: .whitespaces)
            if line == "let document = makeDocument(" {
                inCallSite = true
                continue
            }
            guard inCallSite else { continue }
            if line == ")" {
                inCallSite = false
                continue
            }
            if let m = match(line, #"^([a-z_][A-Za-z0-9_]*): "((?:[^"\\]|\\.)*)",?$"#) {
                bindings[m[0]] = unescape(m[1])
            }
        }
        return bindings
    }

    private static func decodeRawOp(json: String, line: Int) throws -> (op: Operation, source: OpSource) {
        guard let data = json.data(using: .utf8),
              let obj = (try? JSONSerialization.jsonObject(with: data)) as? [String: Any],
              let opType = obj["op_type"] as? String else {
            throw TranscodeError.malformedRawOp(line: line, reason: "raw op line is not a JSON object with op_type")
        }
        let source = OpSource(rawValue: (obj["source"] as? String) ?? "swift") ?? .swift
        do {
            let op = try JSONLLineCoder.decodeOp(opType: opType, fullObject: obj, lineIndex: line)
            return (op, source)
        } catch {
            throw TranscodeError.malformedRawOp(line: line, reason: "cannot decode op '\(opType)': \(error)")
        }
    }

    private static func isComponentScope(_ scope: Scope?) -> Bool {
        if case .component = scope { return true }
        return false
    }

    /// First-match capture groups (regex anchored by caller's pattern).
    private static func match(_ s: String, _ pattern: String) -> [String]? {
        guard let re = try? NSRegularExpression(pattern: pattern) else { return nil }
        let range = NSRange(s.startIndex..., in: s)
        guard let m = re.firstMatch(in: s, range: range) else { return nil }
        var groups: [String] = []
        for i in 1..<m.numberOfRanges {
            guard let r = Range(m.range(at: i), in: s) else { return nil }
            groups.append(String(s[r]))
        }
        return groups
    }

    private static func unescape(_ s: String) -> String {
        var out = ""
        var iterator = s.makeIterator()
        while let ch = iterator.next() {
            if ch == "\\", let next = iterator.next() {
                switch next {
                case "n": out.append("\n")
                case "r": out.append("\r")
                case "t": out.append("\t")
                default: out.append(next)   // \" \\ and anything else literal
                }
            } else {
                out.append(ch)
            }
        }
        return out
    }
}
