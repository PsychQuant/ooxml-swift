import XCTest
@testable import OOXMLSwift

/// word-aligned-state-sync Phase 5 tasks 6.6 / 6.7 — v1.0 risk benchmarks.
/// Gated behind RUN_BENCHMARKS=1 (numbers are recorded in
/// docs/benchmarks-word-aligned-state-sync.md; CI never runs these).
final class V1BenchmarkTests: XCTestCase {

    private func requireBenchmarks() throws {
        guard ProcessInfo.processInfo.environment["RUN_BENCHMARKS"] == "1" else {
            throw XCTSkip("benchmarks gated behind RUN_BENCHMARKS=1")
        }
    }

    private var residentMB: Double {
        var info = mach_task_basic_info()
        var count = mach_msg_type_number_t(
            MemoryLayout<mach_task_basic_info>.size / MemoryLayout<natural_t>.size)
        let kr = withUnsafeMutablePointer(to: &info) {
            $0.withMemoryRebound(to: integer_t.self, capacity: Int(count)) {
                task_info(mach_task_self_, task_flavor_t(MACH_TASK_BASIC_INFO), $0, &count)
            }
        }
        guard kr == KERN_SUCCESS else { return -1 }
        return Double(info.resident_size) / 1_048_576
    }

    /// 6.6 — tree memory cost on a large real-world docx. The design's risk
    /// bar: xmlTrees residency under 50 MB for a thesis-scale document, else
    /// document mitigation.
    func testTreeMemoryCostOnLargeDocument() throws {
        try requireBenchmarks()
        let candidates = [
            "/Users/che/Downloads/20260505v.docx",   // 1.6 MB real-world
            "/Users/che/Downloads/ETST.docx",        // 0.9 MB fallback
        ]
        guard let path = candidates.first(where: { FileManager.default.fileExists(atPath: $0) }) else {
            throw XCTSkip("no large local fixture available")
        }
        let url = URL(fileURLWithPath: path)

        let before = residentMB
        var doc = try DocxReader.read(from: url, wireTreeBackedViews: true)
        let after = residentMB
        let delta = after - before

        let partCount = doc.xmlTrees.count
        let paragraphCount = doc.body.children.filter {
            if case .paragraph = $0 { return true } else { return false }
        }.count
        print("BENCH[6.6] fixture=\(url.lastPathComponent) size=\((try? Data(contentsOf: url).count) ?? 0)B parts=\(partCount) paragraphs=\(paragraphCount) rss_before=\(String(format: "%.1f", before))MB rss_after=\(String(format: "%.1f", after))MB delta=\(String(format: "%.1f", delta))MB")
        doc.close()

        // The 50 MB risk bar applies to the TREE cost (attribution test
        // below: ~38 MB for the 3.9 MB document.xml). The full read() RSS
        // delta additionally contains the typed model and one-shot parse
        // peaks — recorded here for the benchmark doc, sanity-capped only.
        XCTAssertLessThan(delta, 500.0, "read() RSS delta sanity cap")
    }

    /// 6.7 — typed-view performance: read with tree wiring OFF (pre) vs ON
    /// (post) on a 200-paragraph fixture, N iterations each. `get_paragraphs`
    /// in che-word-mcp is read + body enumeration, so read() dominates.
    func testTypedViewReadPerformance200Paragraphs() throws {
        try requireBenchmarks()

        // Build the 200-paragraph fixture from the DSL (self-contained).
        let dir = FileManager.default.temporaryDirectory
            .appendingPathComponent("bench-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: dir) }
        let url = dir.appendingPathComponent("p200.docx")

        var log = OperationLog()
        for i in 1...200 {
            log.append(.appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "第 \(i) 段：benchmark paragraph with 中英 mixed content for realistic sizing.",
                styleId: nil, paraId: "p\(i)")), source: .swift)
        }
        var builder = WordDocument.emptyAuthoringDocument()
        try builder.apply(operations: log.entries.map(\.op), source: .swift)
        try builder.writeAuthoringPackage(to: url)

        let iterations = 30
        func measure(wire: Bool) throws -> Double {
            let t0 = DispatchTime.now()
            for _ in 0..<iterations {
                var d = try DocxReader.read(from: url, wireTreeBackedViews: wire)
                let n = d.body.children.count            // enumeration = get_paragraphs shape
                precondition(n == 200)
                d.close()
            }
            return Double(DispatchTime.now().uptimeNanoseconds - t0.uptimeNanoseconds)
                / Double(iterations) / 1_000_000
        }

        let pre = try measure(wire: false)
        let post = try measure(wire: true)
        let overheadPct = (post - pre) / pre * 100
        print("BENCH[6.7] iterations=\(iterations) read_pre=\(String(format: "%.2f", pre))ms read_post=\(String(format: "%.2f", post))ms overhead=\(String(format: "%.1f", overheadPct))%")

        XCTAssertLessThan(post, pre * 2.0,
                          "tree wiring must not double the typed read path")
    }
}

extension V1BenchmarkTests {

    /// 6.6 attribution — how much of the read() RSS delta is the XmlTree
    /// alone vs typed model + ZIP buffers. Parses word/document.xml through
    /// XmlTreeParser in isolation.
    func testTreeMemoryAttribution() throws {
        try requireBenchmarks()
        let path = "/Users/che/Downloads/20260505v.docx"
        guard FileManager.default.fileExists(atPath: path) else {
            throw XCTSkip("fixture missing")
        }
        // Extract document.xml bytes only.
        var doc = try DocxReader.read(from: URL(fileURLWithPath: path))
        guard let tree = doc.xmlTrees["word/document.xml"] else {
            doc.close(); throw XCTSkip("no document tree")
        }
        let xml = try XmlTreeWriter.serialize(tree)
        doc.close()

        let before = residentMB
        let standalone = try XmlTreeReader.parse(xml)
        let after = residentMB
        var nodeCount = 0
        func walk(_ n: XmlNode) { nodeCount += 1; n.children.forEach(walk) }
        walk(standalone.root)
        print("BENCH[6.6-attr] documentXML=\(xml.count)B nodes=\(nodeCount) tree_alone_delta=\(String(format: "%.1f", after - before))MB")
        XCTAssertLessThan(after - before, 50.0,
                          "the design risk bar: tree cost for a thesis-scale main part under 50 MB")
    }
}
