// MdocxFixtureNormalizerTests
//
// Unit coverage for `MdocxFixtureNormalizer` — the Phase A helper that strips
// identity-noise from a docx so byte-diff comparisons against hand-crafted
// goldens stay stable.
//
// Each strip rule has at least one explicit input/output test. Plus:
//   - testIdempotence: normalize(normalize(x)) == normalize(x)
//   - testDeterminism: same logical content with different ordering / extra
//     stripped fields normalizes to the same bytes
//   - testCrossReferencePreservation: hyperlink anchor + bookmark id pair
//     remain consistent after re-numbering
//   - testNonDefaultThemePreserved / testDefaultThemeStripped: theme drop
//     gating
//
// Inputs are built in-memory as ZIP byte buffers (no disk fixture files) so
// the tests are self-contained.

import XCTest
import ZIPFoundation
@testable import OOXMLSwift

final class MdocxFixtureNormalizerTests: XCTestCase {

    // MARK: - Strip rule 1: RSID attributes

    func testStrips_RSIDAttributesOnEveryElement() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
        <w:p w:rsidR="00ABC123" w:rsidRDefault="00DEF456" w14:paraId="0AAA0001"><w:r w:rsidRPr="00111111"><w:t>hello</w:t></w:r></w:p>
        <w:tbl><w:tr w:rsidTr="00222222"><w:tc><w:p w:rsidP="00333333" w14:paraId="0AAA0002"><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
        </w:body>
        </w:document>
        """
        let docx = try buildDocx(parts: ["word/document.xml": Data(documentXML.utf8)])
        let normalized = try MdocxFixtureNormalizer.normalize(
            docxBytes: docx,
            defaultThemeBytes: Data()
        )
        let outDocXML = try readPart("word/document.xml", from: normalized)
        let outString = String(decoding: outDocXML, as: UTF8.self)
        // No w:rsid* attribute survives.
        for name in ["rsidR", "rsidRDefault", "rsidP", "rsidRPr", "rsidTr", "rsidSect"] {
            XCTAssertFalse(
                outString.contains("w:\(name)="),
                "expected w:\(name) to be stripped from\n\(outString)"
            )
        }
        // Non-rsid content is preserved (text, structure).
        XCTAssertTrue(outString.contains("<w:t>hello</w:t>"))
        XCTAssertTrue(outString.contains("<w:t>cell</w:t>"))
    }

    // MARK: - Strip rule 2: <w:rsids> element from settings

    func testStrips_RsidsElementFromSettings() throws {
        let settingsXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:characterSpacingControl w:val="compressPunctuation"/>
        <w:rsids><w:rsidRoot w:val="00ABC123"/><w:rsid w:val="00ABC123"/><w:rsid w:val="00DEF456"/></w:rsids>
        <w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
        </w:settings>
        """
        let docx = try buildDocx(parts: [
            "word/document.xml": Data(emptyBody.utf8),
            "word/settings.xml": Data(settingsXML.utf8),
        ])
        let normalized = try MdocxFixtureNormalizer.normalize(
            docxBytes: docx,
            defaultThemeBytes: Data()
        )
        let out = String(decoding: try readPart("word/settings.xml", from: normalized), as: UTF8.self)
        XCTAssertFalse(out.contains("<w:rsids"), "rsids element survived: \(out)")
        XCTAssertFalse(out.contains("w:rsidRoot"))
        // Non-default keys preserved.
        XCTAssertTrue(out.contains("characterSpacingControl"))
        XCTAssertTrue(out.contains("<w:compat>"))
    }

    // MARK: - Strip rule 4: structural strip of default settings keys

    func testStrips_DefaultSettingsKeys_ProofStateAndZoom() throws {
        let settingsXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:zoom w:percent="100"/>
        <w:proofState w:spelling="clean" w:grammar="clean"/>
        <w:defaultTabStop w:val="720"/>
        <w:characterSpacingControl w:val="compressPunctuation"/>
        </w:settings>
        """
        let docx = try buildDocx(parts: [
            "word/document.xml": Data(emptyBody.utf8),
            "word/settings.xml": Data(settingsXML.utf8),
        ])
        let normalized = try MdocxFixtureNormalizer.normalize(
            docxBytes: docx,
            defaultThemeBytes: Data()
        )
        let out = String(decoding: try readPart("word/settings.xml", from: normalized), as: UTF8.self)
        XCTAssertFalse(out.contains("<w:zoom"), "zoom survived: \(out)")
        XCTAssertFalse(out.contains("<w:proofState"), "proofState survived: \(out)")
        XCTAssertFalse(out.contains("<w:defaultTabStop"), "defaultTabStop survived: \(out)")
        XCTAssertTrue(out.contains("characterSpacingControl"), "non-default key dropped: \(out)")
    }

    // MARK: - Strip rule 3: theme drop iff bytes equal vendored default

    func testDefaultThemeStripped() throws {
        let theme = "<theme>canonical default theme bytes</theme>"
        let docx = try buildDocx(parts: [
            "word/document.xml": Data(emptyBody.utf8),
            "word/theme/theme1.xml": Data(theme.utf8),
        ])
        let normalized = try MdocxFixtureNormalizer.normalize(
            docxBytes: docx,
            defaultThemeBytes: Data(theme.utf8)
        )
        XCTAssertFalse(
            entryNames(of: normalized).contains("word/theme/theme1.xml"),
            "default theme should have been stripped from output"
        )
    }

    func testNonDefaultThemePreserved() throws {
        let canonical = "<theme>canonical default theme bytes</theme>"
        // Modify by one byte → no longer "default", must be preserved.
        let custom = "<theme>canonical default theme bytes!</theme>"
        let docx = try buildDocx(parts: [
            "word/document.xml": Data(emptyBody.utf8),
            "word/theme/theme1.xml": Data(custom.utf8),
        ])
        let normalized = try MdocxFixtureNormalizer.normalize(
            docxBytes: docx,
            defaultThemeBytes: Data(canonical.utf8)
        )
        XCTAssertTrue(
            entryNames(of: normalized).contains("word/theme/theme1.xml"),
            "non-default theme must be preserved"
        )
        let outBytes = try readPart("word/theme/theme1.xml", from: normalized)
        XCTAssertEqual(outBytes, Data(custom.utf8))
    }

    // MARK: - Strip rule 5: stable-ID renumber + cross-reference preservation

    func testStrips_RenumbersParaIdsInDocumentOrder() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
        <w:p w14:paraId="0AAA1111"><w:r><w:t>first</w:t></w:r></w:p>
        <w:p w14:paraId="0BBB2222"><w:r><w:t>second</w:t></w:r></w:p>
        <w:p w14:paraId="0CCC3333"><w:r><w:t>third</w:t></w:r></w:p>
        </w:body>
        </w:document>
        """
        let docx = try buildDocx(parts: ["word/document.xml": Data(documentXML.utf8)])
        let normalized = try MdocxFixtureNormalizer.normalize(
            docxBytes: docx,
            defaultThemeBytes: Data()
        )
        let out = String(decoding: try readPart("word/document.xml", from: normalized), as: UTF8.self)
        XCTAssertTrue(out.contains(#"w14:paraId="00000001""#))
        XCTAssertTrue(out.contains(#"w14:paraId="00000002""#))
        XCTAssertTrue(out.contains(#"w14:paraId="00000003""#))
        XCTAssertFalse(out.contains("0AAA1111"))
        XCTAssertFalse(out.contains("0BBB2222"))
        XCTAssertFalse(out.contains("0CCC3333"))
    }

    func testCrossReferencePreservation_HyperlinkAnchorAndBookmarkPair() throws {
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
        <w:p w14:paraId="0ABC1234"><w:bookmarkStart w:id="42" w:name="targetPara"/><w:r><w:t>destination</w:t></w:r><w:bookmarkEnd w:id="42"/></w:p>
        <w:p w14:paraId="0XYZ7890"><w:hyperlink w:anchor="0XYZ7890"><w:r><w:t>self link</w:t></w:r></w:hyperlink></w:p>
        </w:body>
        </w:document>
        """
        let docx = try buildDocx(parts: ["word/document.xml": Data(documentXML.utf8)])
        let normalized = try MdocxFixtureNormalizer.normalize(
            docxBytes: docx,
            defaultThemeBytes: Data()
        )
        let out = String(decoding: try readPart("word/document.xml", from: normalized), as: UTF8.self)
        // First paragraph paraId becomes 00000001, second becomes 00000002.
        XCTAssertTrue(out.contains(#"w14:paraId="00000001""#))
        XCTAssertTrue(out.contains(#"w14:paraId="00000002""#))
        // bookmarkStart/bookmarkEnd pair both renumbered to 00000001 (first
        // bookmark in document order). They MUST stay paired.
        XCTAssertTrue(out.contains(#"<w:bookmarkStart w:id="00000001""#),
                      "bookmarkStart not renumbered: \(out)")
        XCTAssertTrue(out.contains(#"<w:bookmarkEnd w:id="00000001""#),
                      "bookmarkEnd not paired-renumbered: \(out)")
        // hyperlink anchor was 0XYZ7890 → second paragraph's new paraId
        // 00000002. Referential integrity preserved.
        XCTAssertTrue(out.contains(#"w:anchor="00000002""#),
                      "hyperlink anchor not rewritten: \(out)")
        // Old IDs gone.
        XCTAssertFalse(out.contains("0ABC1234"))
        XCTAssertFalse(out.contains("0XYZ7890"))
        XCTAssertFalse(out.contains(#"w:id="42""#))
    }

    // MARK: - Idempotence and determinism

    func testIdempotence() throws {
        // Build a docx with all five flavors of identity-noise present.
        let documentXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
        <w:p w:rsidR="00111111" w14:paraId="0AAA0001"><w:r><w:t>alpha</w:t></w:r></w:p>
        <w:p w:rsidR="00222222" w14:paraId="0BBB0002"><w:r><w:t>beta</w:t></w:r></w:p>
        </w:body>
        </w:document>
        """
        let settingsXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:zoom w:percent="100"/>
        <w:rsids><w:rsidRoot w:val="00111111"/></w:rsids>
        <w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
        </w:settings>
        """
        let docx = try buildDocx(parts: [
            "word/document.xml": Data(documentXML.utf8),
            "word/settings.xml": Data(settingsXML.utf8),
        ])
        let once = try MdocxFixtureNormalizer.normalize(docxBytes: docx, defaultThemeBytes: Data())
        let twice = try MdocxFixtureNormalizer.normalize(docxBytes: once, defaultThemeBytes: Data())
        XCTAssertEqual(once, twice, "normalize is not idempotent")
    }

    func testDeterminism_TwoInputsDifferingOnlyInStrippedFields() throws {
        // Input A: rsids + zoom present, paraIds in style "0AAA0001".
        let docA = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
        <w:p w:rsidR="00ABC123" w14:paraId="0AAA0001"><w:r><w:t>X</w:t></w:r></w:p>
        <w:p w:rsidR="00DEF456" w14:paraId="0BBB0002"><w:r><w:t>Y</w:t></w:r></w:p>
        </w:body>
        </w:document>
        """
        let settA = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:zoom w:percent="100"/>
        <w:rsids><w:rsidRoot w:val="00ABC123"/></w:rsids>
        <w:compat/>
        </w:settings>
        """
        // Input B: same logical content but no rsids, no zoom, different paraIds.
        let docB = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
        <w:p w14:paraId="0FFF9999"><w:r><w:t>X</w:t></w:r></w:p>
        <w:p w14:paraId="0EEE8888"><w:r><w:t>Y</w:t></w:r></w:p>
        </w:body>
        </w:document>
        """
        let settB = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:compat/>
        </w:settings>
        """
        let dA = try buildDocx(parts: [
            "word/document.xml": Data(docA.utf8),
            "word/settings.xml": Data(settA.utf8),
        ])
        let dB = try buildDocx(parts: [
            "word/document.xml": Data(docB.utf8),
            "word/settings.xml": Data(settB.utf8),
        ])
        let nA = try MdocxFixtureNormalizer.normalize(docxBytes: dA, defaultThemeBytes: Data())
        let nB = try MdocxFixtureNormalizer.normalize(docxBytes: dB, defaultThemeBytes: Data())
        // Compare the document.xml part bytes. ZIP framing can carry
        // implementation-detail metadata (timestamps if any), but the per-part
        // payload after normalization must be identical.
        let outA = try readPart("word/document.xml", from: nA)
        let outB = try readPart("word/document.xml", from: nB)
        XCTAssertEqual(outA, outB,
                       "document.xml bytes should be identical post-normalize\nA=\(String(decoding: outA, as: UTF8.self))\nB=\(String(decoding: outB, as: UTF8.self))")
        let setA = try readPart("word/settings.xml", from: nA)
        let setB = try readPart("word/settings.xml", from: nB)
        XCTAssertEqual(setA, setB,
                       "settings.xml bytes should be identical post-normalize")
    }

    // MARK: - Helpers

    /// Minimal `<w:body>` for parts where we only need a syntactically valid
    /// document.xml present in the ZIP.
    private let emptyBody = """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>
    """

    /// Build an in-memory docx (ZIP) from a map of part path → bytes. No
    /// [Content_Types].xml or _rels/.rels is added — the normalizer doesn't
    /// inspect them, and tests only assert on word/* parts.
    private func buildDocx(parts: [String: Data]) throws -> Data {
        let archive = try Archive(accessMode: .create)
        // Sort to make the ZIP layout itself deterministic across test runs.
        for name in parts.keys.sorted() {
            let data = parts[name]!
            try archive.addEntry(
                with: name,
                type: .file,
                uncompressedSize: Int64(data.count),
                compressionMethod: .deflate,
                provider: { position, size in
                    let start = data.startIndex.advanced(by: Int(position))
                    let end = start.advanced(by: size)
                    return data.subdata(in: start..<end)
                }
            )
        }
        guard let bytes = archive.data else {
            XCTFail("failed to obtain in-memory archive bytes")
            return Data()
        }
        return bytes
    }

    /// Extract a single part from a docx as raw bytes.
    private func readPart(_ path: String, from docxBytes: Data) throws -> Data {
        let archive = try Archive(data: docxBytes, accessMode: .read)
        guard let entry = archive[path] else {
            XCTFail("part \(path) not found in archive (entries: \(entryNames(of: docxBytes)))")
            return Data()
        }
        var buffer = Data()
        _ = try archive.extract(entry) { buffer.append($0) }
        return buffer
    }

    /// List all entry names in a docx (for assertions about presence/absence).
    private func entryNames(of docxBytes: Data) -> [String] {
        do {
            let archive = try Archive(data: docxBytes, accessMode: .read)
            return archive.map { $0.path }
        } catch {
            return []
        }
    }
}
