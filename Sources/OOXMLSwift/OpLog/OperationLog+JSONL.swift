// JSONL encode/decode for OperationLog.
//
// Spectra change: operation-log-scaffold-impl, target ooxml-swift v0.31.3.
// Capability: ooxml-operation-log
//
// On-disk format: one JSON object per line, separated by Unix LF (0x0A).
// Each line carries 4 required discriminator fields followed by op-specific
// fields:
//   {"op_id":"...","ts":"2026-05-07T01:30:00Z","source":"swift",
//    "op_type":"setText","target":"w14:paraId=X","text":"Hello"}
//
// Field order (deterministic for byte-equal round-trip):
//   1. op_id   — UUID string, uppercase hex with dashes
//   2. ts      — ISO-8601 UTC, second precision
//   3. source  — "swift" or "word"
//   4. op_type — case discriminator string (e.g., "setText", "insertParagraphAfter")
//   5. op-specific fields in case associated-value declaration order
//   6. (unknown ops only) payload object keys sorted lexicographically
//
// Auto-synth Codable cannot apply because `Operation`'s associated-value cases
// require manual encode/decode dispatch on the op_type discriminator. Custom
// Codable conformances for `Operation` and `LogEntry` live here, co-located
// with the JSONL string-building logic that drives them.

import Foundation

// MARK: - Errors

public enum OperationLogJSONLError: Error, Equatable {
    case malformedLine(lineIndex: Int)
}

// MARK: - OperationLog encode/decode entry points

extension OperationLog {

    /// Serializes the log to UTF-8 bytes containing one JSON object per
    /// `LogEntry`, separated by Unix LF (`0x0A`). Each line is a complete,
    /// self-contained JSON object — no JSON array wrapper, no leading/trailing
    /// brackets at the file level.
    public func encodeJSONL() -> Data {
        var result = Data()
        for entry in entries {
            let line = JSONLLineCoder.encodeLine(entry: entry)
            result.append(line)
            result.append(0x0A) // \n
        }
        return result
    }

    /// Parses newline-delimited JSON objects back into a log. Each line MUST
    /// have the four required discriminator fields (`op_id`, `ts`, `source`,
    /// `op_type`); otherwise throws `.malformedLine(lineIndex:)`.
    /// Unknown `op_type` strings decode to `.unknown(opType:payload:)`
    /// carrying the entire JSON object minus the required fields as the
    /// payload.
    public static func decodeJSONL(_ data: Data) throws -> OperationLog {
        var log = OperationLog()
        let lines = data.split(separator: 0x0A, omittingEmptySubsequences: true)
        for (index, lineSlice) in lines.enumerated() {
            let lineData = Data(lineSlice)
            let entry = try JSONLLineCoder.decodeLine(data: lineData, lineIndex: index)
            log.appendEntry(entry)
        }
        return log
    }

    /// Internal escape hatch: append a fully-formed `LogEntry` to `entries`
    /// without re-running through the public `append(_:source:)`. Used only
    /// by `decodeJSONL` to reconstruct a log with the original opIDs and
    /// timestamps preserved verbatim.
    fileprivate mutating func appendEntry(_ entry: LogEntry) {
        // Cannot simply use `entries.append(entry)` from this extension because
        // `entries` is `private(set)`. Workaround: route through a typed
        // `append(_:source:opID:at:)` call — the public signature exposes the
        // exact knobs we need to reconstruct the entry losslessly.
        append(entry.op, source: entry.source, opID: entry.opID, at: entry.timestamp)
    }
}

// MARK: - JSONL line coder (the real work)

/// Internal helper for line-level encode/decode. Kept as an enum-namespace
/// rather than methods on OperationLog so the file's public surface is just
/// `encodeJSONL` / `decodeJSONL`.
internal enum JSONLLineCoder {

    static func encodeLine(entry: LogEntry) -> Data {
        var fields: [(key: String, value: String)] = []

        // Required discriminator fields in fixed order.
        fields.append(("op_id", jsonString(entry.opID.uuidString)))
        fields.append(("ts", jsonString(iso8601(entry.timestamp))))
        fields.append(("source", jsonString(entry.source.rawValue)))

        let (opType, opFields) = encodeOp(entry.op)
        fields.append(("op_type", jsonString(opType)))
        fields.append(contentsOf: opFields)

        // Manual JSON object construction to control field order. The output
        // must be byte-exact for round-trip — JSONEncoder does not give the
        // required ordering control.
        var line = Data("{".utf8)
        for (i, field) in fields.enumerated() {
            if i > 0 { line.append(0x2C) } // ,
            line.append(0x22) // "
            line.append(Data(field.key.utf8))
            line.append(0x22) // "
            line.append(0x3A) // :
            line.append(Data(field.value.utf8))
        }
        line.append(0x7D) // }
        return line
    }

    static func decodeLine(data: Data, lineIndex: Int) throws -> LogEntry {
        // Parse the line as a generic JSON object first.
        let raw = try? JSONSerialization.jsonObject(with: data, options: [])
        guard let obj = raw as? [String: Any] else {
            throw OperationLogJSONLError.malformedLine(lineIndex: lineIndex)
        }

        guard let opIDStr = obj["op_id"] as? String,
              let opID = UUID(uuidString: opIDStr),
              let tsStr = obj["ts"] as? String,
              let ts = parseISO8601(tsStr),
              let sourceStr = obj["source"] as? String,
              let source = OpSource(rawValue: sourceStr),
              let opType = obj["op_type"] as? String else {
            throw OperationLogJSONLError.malformedLine(lineIndex: lineIndex)
        }

        let op = try decodeOp(opType: opType, fullObject: obj, lineIndex: lineIndex)
        return LogEntry(opID: opID, op: op, source: source, timestamp: ts)
    }

    // MARK: Op encode dispatch

    /// Returns `(op_type discriminator, [(key, JSON-encoded value)])`.
    static func encodeOp(_ op: Operation) -> (String, [(key: String, value: String)]) {
        switch op {
        case .insertParagraphAfter(let after, let paragraph):
            return ("insertParagraphAfter", [
                ("after", jsonString(after.raw)),
                ("paragraph", encodeCodable(paragraph))
            ])
        case .insertParagraphBefore(let before, let paragraph):
            return ("insertParagraphBefore", [
                ("before", jsonString(before.raw)),
                ("paragraph", encodeCodable(paragraph))
            ])
        case .removeParagraph(let id):
            return ("removeParagraph", [("id", jsonString(id.raw))])
        case .setText(let target, let text):
            return ("setText", [
                ("target", jsonString(target.raw)),
                ("text", jsonString(text))
            ])
        case .setParagraphStyle(let target, let styleId):
            return ("setParagraphStyle", [
                ("target", jsonString(target.raw)),
                ("styleId", styleId.map(jsonString) ?? "null")
            ])
        case .insertTable(let at, let table):
            return ("insertTable", [
                ("at", jsonString(at.raw)),
                ("table", encodeCodable(table))
            ])
        case .removeTable(let id):
            return ("removeTable", [("id", jsonString(id.raw))])
        case .setCellText(let table, let row, let column, let text):
            return ("setCellText", [
                ("table", jsonString(table.raw)),
                ("row", String(row)),
                ("column", String(column)),
                ("text", jsonString(text))
            ])
        case .insertRun(let inID, let position, let run):
            return ("insertRun", [
                ("in", jsonString(inID.raw)),
                ("position", String(position)),
                ("run", encodeCodable(run))
            ])
        case .setRunFormat(let target, let format):
            return ("setRunFormat", [
                ("target", jsonString(target.raw)),
                ("format", encodeCodable(format))
            ])
        case .insertBookmark(let at, let bookmarkId, let name):
            return ("insertBookmark", [
                ("at", jsonString(at.raw)),
                ("bookmarkId", String(bookmarkId)),
                ("name", jsonString(name))
            ])
        case .insertComment(let anchor, let commentId, let text, let author):
            return ("insertComment", [
                ("anchor", jsonString(anchor.raw)),
                ("commentId", String(commentId)),
                ("text", jsonString(text)),
                ("author", jsonString(author))
            ])
        case .undo(let targetOpID):
            return ("undo", [("targetOpID", jsonString(targetOpID.uuidString))])
        case .redo(let targetOpID):
            return ("redo", [("targetOpID", jsonString(targetOpID.uuidString))])
        case .batchBegin(let label):
            return ("batchBegin", [
                ("label", label.map(jsonString) ?? "null")
            ])
        case .batchEnd:
            return ("batchEnd", [])
        case .insertNode(let parent, let position, let nodeXML):
            return ("insertNode", [
                ("parent", jsonString(parent.raw)),
                ("position", String(position)),
                ("nodeXML", jsonString(nodeXML))
            ])
        case .removeNode(let target):
            return ("removeNode", [("target", jsonString(target.raw))])
        case .updateAttribute(let target, let prefix, let localName, let value):
            return ("updateAttribute", [
                ("target", jsonString(target.raw)),
                ("prefix", prefix.map(jsonString) ?? "null"),
                ("localName", jsonString(localName)),
                ("value", value.map(jsonString) ?? "null")
            ])
        case .moveNode(let source, let destinationParent, let destinationIndex):
            return ("moveNode", [
                ("source", jsonString(source.raw)),
                ("destinationParent", jsonString(destinationParent.raw)),
                ("destinationIndex", String(destinationIndex))
            ])
        case .insertSiblingAfter(let after, let nodeXML):
            return ("insertSiblingAfter", [
                ("after", jsonString(after.raw)),
                ("nodeXML", jsonString(nodeXML))
            ])
        case .wrapWithHyperlink(let target, let rId):
            return ("wrapWithHyperlink", [
                ("target", jsonString(target.raw)),
                ("rId", jsonString(rId))
            ])
        case .addRelationship(let part, let id, let type, let target, let targetMode):
            return ("addRelationship", [
                ("part", jsonString(part)),
                ("id", jsonString(id)),
                ("type", jsonString(type)),
                ("target", jsonString(target)),
                ("targetMode", targetMode.map(jsonString) ?? "null")
            ])
        case .appendParagraph(let container, let paragraph):
            return ("appendParagraph", [
                ("in", container.map { jsonString($0.raw) } ?? "null"),
                ("paragraph", encodeCodable(paragraph))
            ])
        case .setRuns(let target, let runs):
            return ("setRuns", [
                ("target", jsonString(target.raw)),
                ("runs", encodeCodable(runs))
            ])
        case .defineStyle(let payload):
            return ("defineStyle", [("payload", encodeCodable(payload))])
        case .beginComponent(let type, let id):
            return ("beginComponent", [
                ("type", jsonString(type)),
                ("id", jsonString(id.raw))
            ])
        case .endComponent(let id):
            return ("endComponent", [("id", jsonString(id.raw))])
        case .insertTab(let c):
            return ("insertTab", [("in", jsonString(c.raw))])
        case .insertBreak(let c):
            return ("insertBreak", [("in", jsonString(c.raw))])
        case .insertNoBreakHyphen(let c):
            return ("insertNoBreakHyphen", [("in", jsonString(c.raw))])
        case .unknown(let opType, let payload):
            // Merge payload keys sorted lexicographically. Required fields
            // (op_id, ts, source, op_type) are already emitted upstream — the
            // unknown payload is everything else.
            guard case .object(let dict) = payload else {
                return (opType, [])
            }
            let sortedKeys = dict.keys.sorted()
            let fields = sortedKeys.map { key -> (key: String, value: String) in
                return (key, encodeJSONValue(dict[key]!))
            }
            return (opType, fields)
        }
    }

    // MARK: Op decode dispatch

    static func decodeOp(opType: String, fullObject: [String: Any], lineIndex: Int) throws -> Operation {
        // Helper to read ElementID from a string field.
        func eid(_ key: String) throws -> ElementID {
            guard let s = fullObject[key] as? String else {
                throw OperationLogJSONLError.malformedLine(lineIndex: lineIndex)
            }
            return ElementID(rawString: s)
        }
        func str(_ key: String) throws -> String {
            guard let s = fullObject[key] as? String else {
                throw OperationLogJSONLError.malformedLine(lineIndex: lineIndex)
            }
            return s
        }
        func optStr(_ key: String) -> String? {
            // Distinguishes JSON null (NSNull) from missing key from non-string.
            let v = fullObject[key]
            if v == nil || v is NSNull { return nil }
            return v as? String
        }
        func int(_ key: String) throws -> Int {
            if let n = fullObject[key] as? Int { return n }
            if let n = fullObject[key] as? NSNumber { return n.intValue }
            throw OperationLogJSONLError.malformedLine(lineIndex: lineIndex)
        }
        func uuid(_ key: String) throws -> UUID {
            guard let s = fullObject[key] as? String, let u = UUID(uuidString: s) else {
                throw OperationLogJSONLError.malformedLine(lineIndex: lineIndex)
            }
            return u
        }
        func payload<T: Decodable>(_ key: String, _ type: T.Type) throws -> T {
            guard let nested = fullObject[key] else {
                throw OperationLogJSONLError.malformedLine(lineIndex: lineIndex)
            }
            let nestedData = try JSONSerialization.data(withJSONObject: nested, options: [])
            return try JSONDecoder().decode(T.self, from: nestedData)
        }

        switch opType {
        case "insertParagraphAfter":
            return .insertParagraphAfter(after: try eid("after"), paragraph: try payload("paragraph", ParagraphPayload.self))
        case "insertParagraphBefore":
            return .insertParagraphBefore(before: try eid("before"), paragraph: try payload("paragraph", ParagraphPayload.self))
        case "removeParagraph":
            return .removeParagraph(id: try eid("id"))
        case "setText":
            return .setText(target: try eid("target"), text: try str("text"))
        case "setParagraphStyle":
            return .setParagraphStyle(target: try eid("target"), styleId: optStr("styleId"))
        case "insertTable":
            return .insertTable(at: try eid("at"), table: try payload("table", TablePayload.self))
        case "removeTable":
            return .removeTable(id: try eid("id"))
        case "setCellText":
            return .setCellText(table: try eid("table"), row: try int("row"), column: try int("column"), text: try str("text"))
        case "insertRun":
            return .insertRun(in: try eid("in"), position: try int("position"), run: try payload("run", RunPayload.self))
        case "setRunFormat":
            return .setRunFormat(target: try eid("target"), format: try payload("format", RunFormatPayload.self))
        case "insertBookmark":
            return .insertBookmark(at: try eid("at"), bookmarkId: try int("bookmarkId"), name: try str("name"))
        case "insertComment":
            return .insertComment(anchor: try eid("anchor"), commentId: try int("commentId"), text: try str("text"), author: try str("author"))
        case "undo":
            return .undo(targetOpID: try uuid("targetOpID"))
        case "redo":
            return .redo(targetOpID: try uuid("targetOpID"))
        case "batchBegin":
            return .batchBegin(label: optStr("label"))
        case "batchEnd":
            return .batchEnd
        case "insertNode":
            return .insertNode(parent: try eid("parent"), position: try int("position"), nodeXML: try str("nodeXML"))
        case "removeNode":
            return .removeNode(target: try eid("target"))
        case "updateAttribute":
            return .updateAttribute(
                target: try eid("target"),
                prefix: optStr("prefix"),
                localName: try str("localName"),
                value: optStr("value")
            )
        case "moveNode":
            return .moveNode(source: try eid("source"), destinationParent: try eid("destinationParent"), destinationIndex: try int("destinationIndex"))
        case "insertSiblingAfter":
            return .insertSiblingAfter(after: try eid("after"), nodeXML: try str("nodeXML"))
        case "wrapWithHyperlink":
            return .wrapWithHyperlink(target: try eid("target"), rId: try str("rId"))
        case "appendParagraph":
            let container = optStr("in").map { ElementID(rawString: $0) }
            return .appendParagraph(in: container, paragraph: try payload("paragraph", ParagraphPayload.self))
        case "setRuns":
            return .setRuns(target: try eid("target"), runs: try payload("runs", [RunPayload].self))
        case "defineStyle":
            return .defineStyle(payload: try payload("payload", StylePayload.self))
        case "beginComponent":
            return .beginComponent(type: try str("type"), id: try eid("id"))
        case "endComponent":
            return .endComponent(id: try eid("id"))
        case "insertTab":
            return .insertTab(in: try eid("in"))
        case "insertBreak":
            return .insertBreak(in: try eid("in"))
        case "insertNoBreakHyphen":
            return .insertNoBreakHyphen(in: try eid("in"))
        case "addRelationship":
            return .addRelationship(
                part: try str("part"),
                id: try str("id"),
                type: try str("type"),
                target: try str("target"),
                targetMode: optStr("targetMode")
            )
        default:
            // Unknown op_type — preserve the full payload (everything except
            // the four required fields) byte-equal in a JSONValue object.
            var payloadDict: [String: JSONValue] = [:]
            for (key, value) in fullObject {
                if key == "op_id" || key == "ts" || key == "source" || key == "op_type" {
                    continue
                }
                payloadDict[key] = jsonValueFromAny(value)
            }
            return .unknown(opType: opType, payload: .object(payloadDict))
        }
    }

    // MARK: Helpers

    /// JSON-encodes a String (escapes quotes, backslashes, control chars).
    static func jsonString(_ s: String) -> String {
        let arrayData = try! JSONSerialization.data(withJSONObject: [s], options: [.fragmentsAllowed])
        let arrayStr = String(data: arrayData, encoding: .utf8)!
        // arrayStr is like ["foo"]; strip opening [ and closing ]
        return String(arrayStr.dropFirst().dropLast())
    }

    /// JSON-encodes any `Codable` value as a single JSON string.
    static func encodeCodable<T: Encodable>(_ value: T) -> String {
        let encoder = JSONEncoder()
        encoder.outputFormatting = [.sortedKeys]
        let data = try! encoder.encode(value)
        return String(data: data, encoding: .utf8)!
    }

    /// JSON-encodes a `JSONValue` with sorted nested keys for byte-equal
    /// round-trip on unknown payloads.
    static func encodeJSONValue(_ v: JSONValue) -> String {
        let encoder = JSONEncoder()
        encoder.outputFormatting = [.sortedKeys]
        let data = try! encoder.encode(v)
        return String(data: data, encoding: .utf8)!
    }

    /// Converts a `JSONSerialization`-decoded `Any` to a typed `JSONValue`.
    /// Used by the `unknown` op decoder to wrap the residual payload.
    static func jsonValueFromAny(_ v: Any) -> JSONValue {
        if v is NSNull { return .null }
        if let b = v as? Bool, type(of: v) == type(of: NSNumber(value: true)) {
            // NSNumber wraps Bool — must distinguish from Int/Double.
            // (Foundation's JSON returns NSNumber for booleans, which `as? Bool` matches.)
            return .bool(b)
        }
        if let n = v as? NSNumber {
            // Distinguish int vs double via the underlying CFNumber type.
            // For simplicity here: prefer int when fractional part is zero.
            let d = n.doubleValue
            if d.truncatingRemainder(dividingBy: 1) == 0,
               d >= Double(Int.min), d <= Double(Int.max) {
                return .int(Int(d))
            }
            return .double(d)
        }
        if let s = v as? String { return .string(s) }
        if let arr = v as? [Any] { return .array(arr.map(jsonValueFromAny)) }
        if let dict = v as? [String: Any] {
            var result: [String: JSONValue] = [:]
            for (k, val) in dict {
                result[k] = jsonValueFromAny(val)
            }
            return .object(result)
        }
        return .null
    }

    static func iso8601(_ date: Date) -> String {
        let f = ISO8601DateFormatter()
        f.formatOptions = [.withInternetDateTime]
        return f.string(from: date)
    }

    static func parseISO8601(_ str: String) -> Date? {
        let f = ISO8601DateFormatter()
        f.formatOptions = [.withInternetDateTime]
        if let d = f.date(from: str) { return d }
        // Fallback for fractional-second timestamps in older logs.
        f.formatOptions = [.withInternetDateTime, .withFractionalSeconds]
        return f.date(from: str)
    }
}
