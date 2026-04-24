import XCTest
@testable import OOXMLSwift

/// Phase C regression test for che-word-mcp#52.
///
/// Spec: `openspec/changes/che-word-mcp-header-footer-raw-element-preservation/specs/che-word-mcp-field-equation-crud/spec.md`
/// MODIFIED requirement: "update_all_fields MCP tool recomputes SEQ counters across the document"
///
/// 2 spec scenarios:
/// 1. Default behavior shares counters across containers (regression check)
/// 2. Isolation flag resets counters per container family
final class UpdateAllFieldsCounterIsolationTests: XCTestCase {

    private func captionParagraph(identifier: String, initialCached: String = "0") -> Paragraph {
        let field = SequenceField(identifier: identifier, cachedResult: initialCached)
        var run = Run(text: "")
        run.rawXML = field.toFieldXML()
        var para = Paragraph()
        para.runs = [Run(text: "\(identifier) "), run]
        para.properties.style = "Caption"
        return para
    }

    /// Build doc with 3 Figure SEQ in body + 1 Figure SEQ in header.
    private func makeFixture() -> WordDocument {
        var doc = WordDocument()
        for _ in 0..<3 {
            doc.appendParagraph(captionParagraph(identifier: "Figure"))
        }
        let header = Header(id: "rId10",
                            paragraphs: [captionParagraph(identifier: "Figure")])
        doc.headers.append(header)
        return doc
    }

    // MARK: - Scenario 1: default behavior shares counters globally

    func testDefaultBehaviorSharesCountersAcrossContainers() {
        var doc = makeFixture()
        let counters = doc.updateAllFields()  // default isolatePerContainer: false

        XCTAssertEqual(counters["Figure"], 4,
                       "Default global mode SHALL produce Figure: 4 (3 body + 1 header sees count 4)")
    }

    // MARK: - Scenario 2: isolation flag resets counters per container family

    func testIsolationFlagResetsCountersPerContainerFamily() {
        var doc = makeFixture()
        let counters = doc.updateAllFields(isolatePerContainer: true)

        // Per spec scenario: body's three Figure counters become 1, 2, 3 and
        // header's Figure counter becomes 1. Per-container breakdown — return
        // value semantic SHALL preserve the maximum per-identifier across
        // all containers (for backward compat with single-counter callers),
        // OR document the breakdown shape if changed.
        //
        // Simplest contract: returned dict is the FINAL counter state of the
        // body container family (since body is processed first). Per-container
        // state is ephemeral and not exposed in the return value (caller can
        // inspect rawXML of each container's SEQ runs to see actual cached
        // values per container).
        XCTAssertEqual(counters["Figure"], 3,
                       "Isolation mode SHALL produce Figure: 3 in body container; header has its own Figure: 1")

        // Verify header's SEQ cached value is "1" (isolated, not "4")
        let header = doc.headers[0]
        let headerCachedValue = extractCachedValue(from: header.paragraphs[0])
        XCTAssertEqual(headerCachedValue, "1",
                       "Header's SEQ Figure cached value SHALL be '1' under isolation mode (not '4')")
    }

    /// Extract the cached <w:t> value from a SEQ field run's rawXML.
    private func extractCachedValue(from para: Paragraph) -> String? {
        for run in para.runs {
            guard let raw = run.rawXML else { continue }
            // Match the cached <w:t>X</w:t> after fldCharType="separate"
            let pattern = #"<w:fldChar[^/]*fldCharType="separate"[^/]*/>\s*</w:r>\s*<w:r[^>]*>\s*<w:t[^>]*>([^<]+)</w:t>"#
            if let regex = try? NSRegularExpression(pattern: pattern, options: [.dotMatchesLineSeparators]),
               let match = regex.firstMatch(in: raw, options: [], range: NSRange(raw.startIndex..., in: raw)),
               match.numberOfRanges >= 2,
               let captureRange = Range(match.range(at: 1), in: raw) {
                return String(raw[captureRange])
            }
        }
        return nil
    }
}
