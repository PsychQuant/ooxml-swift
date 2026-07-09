import XCTest
@testable import OOXMLSwift

/// word-aligned-state-sync Phase 4 tasks 5.1 / 5.2 / 5.4 / 5.6 —
/// `ooxml-script-transcode` capability: the bidirectional codec between
/// `OperationLog` and canonical `.mdocx` Swift source.
///
/// Projection contract (v0.34 vertical slice):
/// - DSL form covers the bijective authoring subset:
///   `appendParagraph(in: nil, payload with paraId)` ↔
///   `Paragraph(id: "...", style: .x) { "text" }` inside a synthesized
///   `Section(id: "main")` envelope (the Section wrapper emits no op).
/// - EVERY other op round-trips via the raw escape line
///   `// @op {canonical JSONL op fields}` — the same forward-compat
///   mechanism the spec mandates for unknown op_types, extended to all
///   DSL-unrepresentable ops so round-trip holds for ARBITRARY logs.
/// - Style references use the `WordStyleMap` predefined table
///   (`.heading1` ↔ "Heading1") falling back to verbatim member names.
final class ScriptTranscodeTests: XCTestCase {

    private func log(_ ops: [(OOXMLSwift.Operation, OpSource)]) -> OperationLog {
        var l = OperationLog()
        for (op, src) in ops { l.append(op, source: src) }
        return l
    }

    private func opsEquivalent(_ a: OperationLog, _ b: OperationLog,
                               file: StaticString = #filePath, line: UInt = #line) {
        XCTAssertEqual(a.entries.count, b.entries.count,
                       "entry count must match", file: file, line: line)
        for (i, (ea, eb)) in zip(a.entries, b.entries).enumerated() {
            XCTAssertEqual(ea.op, eb.op, "op \(i) must be equivalent", file: file, line: line)
            XCTAssertEqual(ea.source, eb.source, "op \(i) source must survive", file: file, line: line)
        }
    }

    // MARK: - 5.1 exporter shape (mdocx-grammar conformance on the DSL subset)

    func testExportEmitsCanonicalMdocxShape() throws {
        let l = log([
            (.appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "本章探討", styleId: nil, paraId: "P1")), .swift),
            (.appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "Title", styleId: "Heading1", paraId: "H1")), .swift),
        ])
        let src = ScriptExporter.exportSwift(log: l)

        XCTAssertTrue(src.contains("import WordDSLSwift"),
                      "canonical source imports WordDSLSwift")
        XCTAssertTrue(src.contains("let document = WordDocument {"),
                      "top-level WordDocument result-builder declaration")
        XCTAssertTrue(src.contains("Section(id: \"main\") {"),
                      "synthesized Section envelope")
        XCTAssertTrue(src.contains("Paragraph(id: \"P1\") {"),
                      "OOXML-mirror element naming: Paragraph, explicit id")
        XCTAssertTrue(src.contains("\"本章探討\""),
                      "unstyled text is a plain String literal (flat-Run implicit String)")
        XCTAssertTrue(src.contains("Paragraph(id: \"H1\", style: .heading1) {"),
                      "styleId maps through WordStyleMap: Heading1 -> .heading1")
        XCTAssertFalse(src.contains("Heading1(\""),
                       "no semantic-shortcut wrapper components (grammar prohibition)")
    }

    func testExportEscapesNonRepresentableOpsAsRawLines() throws {
        let pid = ElementID(rawString: "w14:paraId=ABC")
        let l = log([
            (.setText(target: pid, text: "mutated"), .word),
            (.unknown(opType: "future_op_v2",
                      payload: JSONValue.object(["k": JSONValue.int(1)])), .swift),
        ])
        let src = ScriptExporter.exportSwift(log: l)

        let rawLines = src.split(separator: "\n").filter {
            $0.trimmingCharacters(in: .whitespaces).hasPrefix("// @op ")
        }
        XCTAssertEqual(rawLines.count, 2,
                       "non-representable ops must each become one // @op raw line")
        XCTAssertTrue(src.contains("\"op_type\":\"setText\"") || src.contains("setText"),
                      "raw line carries the canonical op form")
    }

    // MARK: - 5.2 importer

    func testImportHandWrittenCanonicalScript() throws {
        let src = """
        import WordDSLSwift

        let document = WordDocument {
            Section(id: "main") {
                Paragraph(id: "p-title", style: .heading1) {
                    "Title"
                }
                Paragraph(id: "p-intro") {
                    "Body intro"
                }
            }
        }
        """
        let l = try ScriptImporter.parse(source: src)

        XCTAssertEqual(l.entries.count, 2)
        guard case .appendParagraph(let c0, let p0) = l.entries[0].op else {
            return XCTFail("expected appendParagraph, got \(l.entries[0].op)")
        }
        XCTAssertNil(c0)
        XCTAssertEqual(p0.paraId, "p-title")
        XCTAssertEqual(p0.styleId, "Heading1",
                       ".heading1 maps back to verbatim styleId Heading1")
        XCTAssertEqual(p0.text, "Title")
        guard case .appendParagraph(_, let p1) = l.entries[1].op else {
            return XCTFail("expected appendParagraph")
        }
        XCTAssertEqual(p1.styleId, nil)
        XCTAssertEqual(p1.text, "Body intro")
        XCTAssertEqual(l.entries.map(\.source), [.swift, .swift],
                       "hand-written script ops carry source swift")
    }

    func testImportRejectsArbitrarySwiftWithPreciseLocation() throws {
        let src = """
        import WordDSLSwift

        let document = WordDocument {
            Section(id: "main") {
                FileManager.default.removeItem(atPath: "/")
            }
        }
        """
        XCTAssertThrowsError(try ScriptImporter.parse(source: src)) { error in
            guard case TranscodeError.unsupportedSyntax(let line, _, let reason) = error else {
                return XCTFail("expected unsupportedSyntax, got \(error)")
            }
            XCTAssertEqual(line, 5, "error must carry the offending line number")
            XCTAssertFalse(reason.isEmpty)
        }
    }

    func testImportRejectsRawStringStyle() throws {
        // mdocx-grammar: raw-string style MUST be rejected.
        let src = """
        import WordDSLSwift

        let document = WordDocument {
            Section(id: "main") {
                Paragraph(id: "p1", style: "heading1") {
                    "Title"
                }
            }
        }
        """
        XCTAssertThrowsError(try ScriptImporter.parse(source: src)) { error in
            guard case TranscodeError.unsupportedSyntax(_, _, let reason) = error else {
                return XCTFail("expected unsupportedSyntax, got \(error)")
            }
            XCTAssertTrue(reason.contains("WordStyle") || reason.contains("style"),
                          "reason should name the typed-enum requirement, got: \(reason)")
        }
    }

    func testImportRejectsParagraphWithoutExplicitID() throws {
        // mdocx-grammar: mandatory explicit identifiers.
        let src = """
        import WordDSLSwift

        let document = WordDocument {
            Section(id: "main") {
                Paragraph {
                    "text"
                }
            }
        }
        """
        XCTAssertThrowsError(try ScriptImporter.parse(source: src)) { error in
            guard case TranscodeError.unsupportedSyntax(_, _, let reason) = error else {
                return XCTFail("expected unsupportedSyntax, got \(error)")
            }
            XCTAssertTrue(reason.contains("id"), "reason should name the id requirement")
        }
    }

    // MARK: - 5.6 round-trip: log → script → log

    func testRoundTripLogScriptLogPreservesOperations() throws {
        let pid = ElementID(rawString: "w14:paraId=ABC")
        let rid = ElementID(rawString: "lib:11111111-2222-3333-4444-555555555555")
        let l = log([
            // DSL-representable
            (.appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "first", styleId: nil, paraId: "P1")), .swift),
            (.appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "heading", styleId: "Heading2", paraId: "P2")), .swift),
            // raw-escape territory: mutation, word-sourced, components, atoms, unknown
            (.setText(target: pid, text: "edited in Word"), .word),
            (.beginComponent(type: "Summary", id: ElementID(rawString: "c1")), .swift),
            (.setRuns(target: pid, runs: [RunPayload(text: "styled", bold: true)]), .swift),
            (.endComponent(id: ElementID(rawString: "c1")), .swift),
            (.insertTab(in: rid), .swift),
            (.defineStyle(payload: StylePayload(styleId: "titleBrown", font: "Noto Serif TC",
                                                fontSize: 36, color: "663300", bold: true)), .swift),
            (.removeParagraph(id: pid), .word),
            (.unknown(opType: "future_op_v2",
                      payload: JSONValue.object(["k": JSONValue.int(1)])), .swift),
        ])

        let script = ScriptExporter.exportSwift(log: l)
        let reconstructed = try ScriptImporter.parse(source: script)
        opsEquivalent(l, reconstructed)
    }

    // MARK: - 5.6 idempotency: script → log → script

    func testScriptLogScriptIsIdempotentOnCanonicalForm() throws {
        let l = log([
            (.appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "alpha", styleId: nil, paraId: "A")), .swift),
            (.setText(target: ElementID(rawString: "w14:paraId=A"), text: "beta"), .word),
        ])
        let script1 = ScriptExporter.exportSwift(log: l)
        let log2 = try ScriptImporter.parse(source: script1)
        let script2 = ScriptExporter.exportSwift(log: log2)
        XCTAssertEqual(script1, script2,
                       "canonical form must be a fixed point of the transcoder")
    }

    // MARK: - 5.4 stable formatting: one added op → one localized hunk

    func testAddingOneOpProducesOneLocalizedInsertion() throws {
        let base: [(OOXMLSwift.Operation, OpSource)] = [
            (.appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "one", styleId: nil, paraId: "P1")), .swift),
            (.appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "two", styleId: nil, paraId: "P2")), .swift),
        ]
        var extended = base
        extended.append((.appendParagraph(in: nil, paragraph: ParagraphPayload(
            text: "three", styleId: nil, paraId: "P3")), .swift))

        let src1 = ScriptExporter.exportSwift(log: log(base)).split(
            separator: "\n", omittingEmptySubsequences: false).map(String.init)
        let src2 = ScriptExporter.exportSwift(log: log(extended)).split(
            separator: "\n", omittingEmptySubsequences: false).map(String.init)

        // src2 must equal src1 with one contiguous block inserted.
        XCTAssertGreaterThan(src2.count, src1.count)
        var i = 0
        while i < src1.count && src1[i] == src2[i] { i += 1 }          // common prefix
        var j1 = src1.count - 1, j2 = src2.count - 1
        while j1 >= i && src1[j1] == src2[j2] { j1 -= 1; j2 -= 1 }     // common suffix
        XCTAssertGreaterThan(j1 + 1, i - 1)
        XCTAssertTrue(j1 < i,
            "all of src1 must be prefix+suffix (no changed lines) — insertion only; " +
            "divergent middle src1[\(i)...\(j1)]: \(src1[max(0,i)...max(0,j1)].joined(separator: " | "))")
        let inserted = src2[i...j2]
        XCTAssertTrue(inserted.contains { $0.contains("three") || $0.contains("P3") },
                      "the inserted block corresponds to the new op")
    }

    // MARK: - determinism

    func testExportIsDeterministic() throws {
        let l = log([
            (.appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "x", styleId: "Quote", paraId: "Q")), .swift),
            (.defineStyle(payload: StylePayload(styleId: "s", bold: true)), .swift),
        ])
        XCTAssertEqual(ScriptExporter.exportSwift(log: l),
                       ScriptExporter.exportSwift(log: l),
                       "same log must always export byte-identical source")
    }
}

extension ScriptTranscodeTests {

    /// mdocx-grammar "Special-character inline atoms as standalone children" —
    /// importer accepts Tab()/Break()/NoBreakHyphen() in a paragraph body,
    /// emitting paragraph-targeted atom ops (the reducer synthesizes the
    /// wrapping <w:r> per the §4b rule).
    func testImportAcceptsInlineAtomsAsStandaloneChildren() throws {
        let src = """
        import WordDSLSwift

        let document = WordDocument {
            Section(id: "main") {
                Paragraph(id: "p1") {
                    "Header"
                    Tab()
                    "Right-aligned"
                    Break()
                }
            }
        }
        """
        let l = try ScriptImporter.parse(source: src)

        XCTAssertEqual(l.entries.count, 3, "appendParagraph + insertTab + insertBreak")
        guard case .appendParagraph(_, let p) = l.entries[0].op else {
            return XCTFail("first op must be the paragraph itself")
        }
        XCTAssertEqual(p.text, "HeaderRight-aligned")
        guard case .insertTab(let t1) = l.entries[1].op else {
            return XCTFail("expected insertTab, got \(l.entries[1].op)")
        }
        XCTAssertEqual(t1.raw, "w14:paraId=p1",
                       "atom targets the containing paragraph (reducer wraps in <w:r>)")
        guard case .insertBreak = l.entries[2].op else {
            return XCTFail("expected insertBreak, got \(l.entries[2].op)")
        }
    }
}

extension ScriptTranscodeTests {

    /// 5.1 "compiles as a Swift source file" — the exported source is piped
    /// through `swiftc -parse` (syntax-level compile; imports not resolved).
    /// The full semantic compile of this grammar is exercised by
    /// WordDSLRuntimeTests, whose inline DSL bodies use identical syntax.
    func testExportedSourceParsesUnderSwiftc() throws {
        let swiftc = "/usr/bin/swiftc"
        guard FileManager.default.fileExists(atPath: swiftc) else {
            throw XCTSkip("swiftc not available")
        }
        let l: OperationLog = {
            var log = OperationLog()
            log.append(.appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "本章探討", styleId: "Heading1", paraId: "P1")), source: .swift)
            log.append(.setText(target: ElementID(rawString: "w14:paraId=P1"),
                                text: "raw escape line"), source: .word)
            return log
        }()
        let source = ScriptExporter.exportSwift(log: l)

        let dir = FileManager.default.temporaryDirectory
            .appendingPathComponent("mdocx-parse-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: dir) }
        let file = dir.appendingPathComponent("exported.mdocx.swift")
        try source.write(to: file, atomically: true, encoding: .utf8)

        let process = Process()
        process.executableURL = URL(fileURLWithPath: swiftc)
        process.arguments = ["-parse", file.path]
        let pipe = Pipe(); process.standardError = pipe
        try process.run(); process.waitUntilExit()
        let diag = String(decoding: pipe.fileHandleForReading.readDataToEndOfFile(),
                          as: UTF8.self)
        XCTAssertEqual(process.terminationStatus, 0,
                       "exported source must be syntactically valid Swift: \(diag)")
    }
}

extension ScriptTranscodeTests {

    /// 5.5 "Script export covers all operation types in the log" — EVERY
    /// Operation case (all 36, same construction list as the enum pin in
    /// OperationLogTests) must have a Swift representation that survives
    /// export → import with op equivalence. DSL-form for the authoring
    /// subset, `// @op` raw escape (the spec's comment-marker mechanism)
    /// for everything else.
    func testEveryOperationTypeRoundTripsThroughExport() throws {
        let id = ElementID(rawString: "w14:paraId=A")
        let id2 = ElementID(rawString: "w14:paraId=B")
        let uuid = UUID(uuidString: "11111111-1111-4111-8111-111111111111")!

        let cases: [OOXMLSwift.Operation] = [
            .insertParagraphAfter(after: id, paragraph: ParagraphPayload(text: "p")),
            .insertParagraphBefore(before: id, paragraph: ParagraphPayload(text: "p")),
            .removeParagraph(id: id),
            .setText(target: id, text: "Hello"),
            .setParagraphStyle(target: id, styleId: "Heading1"),
            .insertTable(at: id, table: TablePayload(rows: 2, columns: 3)),
            .removeTable(id: id),
            .setCellText(table: id, row: 0, column: 1, text: "cell"),
            .insertRun(in: id, position: 0, run: RunPayload(text: "r")),
            .setRunFormat(target: id, format: RunFormatPayload(bold: true)),
            .insertBookmark(at: id, bookmarkId: 7, name: "anchor"),
            .insertComment(anchor: id, commentId: 3, text: "ct", author: "auth"),
            .undo(targetOpID: uuid),
            .redo(targetOpID: uuid),
            .batchBegin(label: "rename"),
            .batchEnd,
            .insertNode(parent: id, position: 0, nodeXML: "<w:p/>"),
            .removeNode(target: id),
            .updateAttribute(target: id, prefix: "w", localName: "id", value: "5"),
            .moveNode(source: id, destinationParent: id2, destinationIndex: 0),
            .insertSiblingAfter(after: id, nodeXML: "<w:t>x</w:t>"),
            .wrapWithHyperlink(target: id, rId: "rId99"),
            .addRelationship(
                part: "word/_rels/document.xml.rels",
                id: "rId99",
                type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                target: "https://example.com",
                targetMode: "External"
            ),
            .unknown(opType: "future", payload: JSONValue.object(["k": JSONValue.int(1)])),
            .appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "x", styleId: nil, paraId: "P1")),
            .setRuns(target: id, runs: [RunPayload(text: "x", bold: true)]),
            .defineStyle(payload: StylePayload(styleId: "s1")),
            .beginComponent(type: "Summary", id: id),
            .endComponent(id: id),
            .insertTab(in: id),
            .insertBreak(in: id),
            .insertNoBreakHyphen(in: id),
            .carryPart(partPath: "word/styles.xml", xml: "<w:styles/>"),
            .setSectionProperties(at: nil, section: SectionPayload(pageWidth: 11906, columnCount: 2)),
            .appendTable(in: nil, table: TablePayload(rows: 1, columns: 2, cells: [["左", "右"]])),
            .setDocumentRoot(attributes: [RootAttribute(prefix: "xmlns", localName: "w", value: "NS")]),
        ]
        XCTAssertEqual(cases.count, 36,
                       "update this list when the Operation enum grows — every case must round-trip")

        var l = OperationLog()
        for op in cases { l.append(op, source: .swift) }

        let script = ScriptExporter.exportSwift(log: l)
        let reconstructed = try ScriptImporter.parse(source: script)
        opsEquivalent(l, reconstructed)
    }
}
