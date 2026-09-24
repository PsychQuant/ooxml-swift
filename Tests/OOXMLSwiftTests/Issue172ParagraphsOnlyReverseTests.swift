import Foundation
import XCTest
import ZIPFoundation
@testable import OOXMLSwift

/// Regression / parity tests for [PsychQuant/ooxml-swift#172](https://github.com/PsychQuant/ooxml-swift/issues/172).
///
/// `ReverseExtractor.paragraphsOnly(url:slots:)` is the single library
/// implementation replacing two independently-maintained copies:
/// `MacDoc.Word.Reverse.reverseEngineer(from:)` (macdoc CLI,
/// `Sources/MacDocCLI/MacDoc+Word.swift`) and che-word-mcp's
/// `paragraphsOnlyReverse(from:)` (`Sources/CheWordMCP/ScriptPipelineTools.swift`,
/// explicitly documented there as "PORT, not a shared call", #227).
///
/// Section A checks the closed `OmittedBodyBlockReason` enum in isolation.
/// Section B is the byte-equal oracle comparison: the macdoc 0.11.0 release
/// binary's `word reverse --paragraphs-only` output, against this library
/// function's output, on fixtures covering a table, an empty paragraph, and
/// a raw (unrecognized) body-level block — the representative shapes named
/// in the issue.
final class Issue172ParagraphsOnlyReverseTests: XCTestCase {

    // MARK: - Fixture

    /// body:
    ///   0: "Intro"                    — plain paragraph
    ///   1: (empty)                    — empty paragraph, no <w:pPr>, no runs
    ///   2: <w:commentRangeStart>      — raw block element (unrecognized at body level)
    ///   3: <w:tbl>                    — one row, two cells
    ///   4: "Styled" (Heading1)        — plain paragraph with a pStyle
    private static let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\
        <w:p><w:r><w:t>Intro</w:t></w:r></w:p>\
        <w:p/>\
        <w:commentRangeStart w:id="0"/>\
        <w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>\
        <w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Cell A</w:t></w:r></w:p></w:tc>\
        <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Cell B</w:t></w:r></w:p></w:tc>\
        </w:tr></w:tbl>\
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Styled</w:t></w:r></w:p>\
        <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>\
        </w:body></w:document>
        """

    private func buildFixture() throws -> URL {
        let staging = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue172-staging-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: staging, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: staging) }

        func write(_ content: String, to relativePath: String) throws {
            let url = staging.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: url.deletingLastPathComponent(), withIntermediateDirectories: true)
            try content.write(to: url, atomically: true, encoding: .utf8)
        }

        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="xml" ContentType="application/xml"/>
                <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>
            """, to: "[Content_Types].xml")
        try write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """, to: "_rels/.rels")
        try write(Self.documentXML, to: "word/document.xml")

        let docxURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue172-\(UUID().uuidString).docx")
        let archive = try Archive(url: docxURL, accessMode: .create)
        let base = staging.resolvingSymlinksInPath().path
        let enumerator = FileManager.default.enumerator(
            at: staging, includingPropertiesForKeys: [.isDirectoryKey])!
        for case let fileURL as URL in enumerator {
            let isDir = (try? fileURL.resourceValues(forKeys: [.isDirectoryKey]))?.isDirectory ?? false
            if isDir { continue }
            let entry = String(fileURL.resolvingSymlinksInPath().path.dropFirst(base.count + 1))
            try archive.addEntry(with: entry, fileURL: fileURL, compressionMethod: .deflate)
        }
        return docxURL
    }

    // MARK: - Section A: OmittedBodyBlockReason (closed enum)

    func testOmittedBlocksLabelEachNonParagraphBodyChildKind() throws {
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }

        let result = try ReverseExtractor.paragraphsOnly(url: fixture)
        XCTAssertEqual(result.omittedBlocks.map(\.reason), [
            .rawBlockElement(name: "commentRangeStart"),
            .table,
        ], "omitted blocks must appear in document order with the right reason each")
        // Positional addressing matches body.children indices: Intro(0),
        // empty paragraph(1), commentRangeStart(2), table(3), Styled(4).
        XCTAssertEqual(result.omittedBlocks.map(\.index), [2, 3])
    }

    func testOmittedBodyBlockReasonIsExhaustiveOverNonParagraphBodyChildCases() {
        // Compile-time guard, not a runtime assertion: this switch has no
        // `default:`. If `BodyChild` (Document.swift) ever grows a case
        // beyond paragraph/table/contentControl/bookmarkMarker/rawBlockElement,
        // ReverseExtractor.swift's paragraphsOnly(url:) switch fails to build
        // — the loud signal #172 asked for instead of a silent gap.
        let sample = BodyChild.table(Table(rows: []))
        switch sample {
        case .paragraph: XCTFail("not exercised by this sample")
        case .table: break
        case .contentControl: XCTFail("not exercised by this sample")
        case .bookmarkMarker: XCTFail("not exercised by this sample")
        case .rawBlockElement: XCTFail("not exercised by this sample")
        }
    }

    // MARK: - Section B: byte-equal oracle vs. macdoc 0.11.0 CLI

    /// `word reverse --paragraphs-only` behavior is env-gated the same way
    /// `TemplateFixtureGate` gates real-doc fixtures: set
    /// `MACDOC_RELEASE_BINARY` to the path of a `macdoc` CLI build ≥0.11.0
    /// (the release that shipped `--paragraphs-only`) to run this oracle
    /// comparison locally. Skips loudly (not silently) when unset so CI,
    /// which has no macdoc binary, stays green.
    ///
    ///   MACDOC_RELEASE_BINARY=/path/to/macdoc swift test --filter Issue172ParagraphsOnlyReverseTests
    private func requireReleaseBinary() throws -> URL {
        guard let path = ProcessInfo.processInfo.environment["MACDOC_RELEASE_BINARY"] else {
            throw XCTSkip("set MACDOC_RELEASE_BINARY to a macdoc CLI build (>=0.11.0) to run the byte-equal oracle comparison")
        }
        guard FileManager.default.isExecutableFile(atPath: path) else {
            throw XCTSkip("MACDOC_RELEASE_BINARY (\(path)) is not an executable file")
        }
        return URL(fileURLWithPath: path)
    }

    private func runCLIReverse(_ binary: URL, input: URL, output: URL, extraArgs: [String] = []) throws {
        let process = Process()
        process.executableURL = binary
        process.arguments = ["word", "reverse", input.path,
                             "--paragraphs-only", "--to-mdocx", output.path] + extraArgs
        let stderrPipe = Pipe()
        process.standardError = stderrPipe
        try process.run()
        process.waitUntilExit()
        if process.terminationStatus != 0 {
            let stderrData = stderrPipe.fileHandleForReading.readDataToEndOfFile()
            XCTFail("macdoc word reverse exited \(process.terminationStatus): "
                + String(decoding: stderrData, as: UTF8.self))
        }
    }

    func testScriptIsByteEqualToReleaseCLI() throws {
        let binary = try requireReleaseBinary()
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }

        let cliOutput = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue172-cli-\(UUID().uuidString).mdocx.swift")
        defer { try? FileManager.default.removeItem(at: cliOutput) }
        try runCLIReverse(binary, input: fixture, output: cliOutput)
        let cliScript = try String(contentsOf: cliOutput, encoding: .utf8)

        let result = try ReverseExtractor.paragraphsOnly(url: fixture)
        XCTAssertEqual(result.script, cliScript,
                       "ReverseExtractor.paragraphsOnly must byte-match macdoc 0.11.0's word reverse --paragraphs-only")
    }

    func testScriptWithSlotIsByteEqualToReleaseCLI() throws {
        let binary = try requireReleaseBinary()
        let fixture = try buildFixture()
        defer { try? FileManager.default.removeItem(at: fixture) }

        let cliOutput = FileManager.default.temporaryDirectory
            .appendingPathComponent("issue172-cli-slot-\(UUID().uuidString).mdocx.swift")
        defer { try? FileManager.default.removeItem(at: cliOutput) }
        try runCLIReverse(binary, input: fixture, output: cliOutput, extraArgs: ["--slot", "intro=p1"])
        let cliScript = try String(contentsOf: cliOutput, encoding: .utf8)

        let result = try ReverseExtractor.paragraphsOnly(
            url: fixture, slots: [SlotDesignation(name: "intro", paraId: "p1")])
        XCTAssertEqual(result.script, cliScript,
                       "slot-designated export must byte-match the CLI's --slot output too")
    }
}
