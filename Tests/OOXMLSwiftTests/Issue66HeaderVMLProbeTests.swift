import XCTest
@testable import OOXMLSwift

final class Issue66HeaderVMLProbeTests: XCTestCase {
    func testHeaderVMLPictParsedIntoRawElements() throws {
        // Local-only real-world fixture (path via env — a third party's document
        // name does not belong in the repo; see the repo privacy discipline).
        guard let fixturePath = ProcessInfo.processInfo.environment["OOXML_LOCAL_THESIS_FIXTURE"],
              FileManager.default.fileExists(atPath: fixturePath) else {
            throw XCTSkip("set OOXML_LOCAL_THESIS_FIXTURE to a local thesis docx to run this probe")
        }
        let raw = URL(fileURLWithPath: fixturePath)
        let doc = try DocxReader.read(from: raw)
        XCTAssertGreaterThan(doc.headers.count, 0, "expected at least one header")
        var foundPict = false
        for h in doc.headers {
            for child in h.bodyChildren {
                if case .paragraph(let p) = child {
                    for r in p.runs {
                        if let raws = r.rawElements {
                            for el in raws {
                                if el.name == "pict" {
                                    foundPict = true
                                    print("✓ found <w:pict> in header \(h.id) — xml prefix: \(el.xml.prefix(120))")
                                }
                            }
                        }
                    }
                }
            }
        }
        XCTAssertTrue(foundPict, "<w:pict> not captured into rawElements during DocxReader.read")
    }

    func testHeaderRoundTripPreservesVML() throws {
        // Local-only real-world fixture (path via env — a third party's document
        // name does not belong in the repo; see the repo privacy discipline).
        guard let fixturePath = ProcessInfo.processInfo.environment["OOXML_LOCAL_THESIS_FIXTURE"],
              FileManager.default.fileExists(atPath: fixturePath) else {
            throw XCTSkip("set OOXML_LOCAL_THESIS_FIXTURE to a local thesis docx to run this probe")
        }
        let raw = URL(fileURLWithPath: fixturePath)
        let out = URL(fileURLWithPath: NSTemporaryDirectory()).appendingPathComponent("issue66_repro.docx")
        try? FileManager.default.removeItem(at: out)

        let doc = try DocxReader.read(from: raw)
        try DocxWriter.write(doc, to: out)

        // Re-read the output, count <w:pict> in any header
        // Simpler: read raw bytes via Foundation + unzip helper
        let p = Process()
        p.executableURL = URL(fileURLWithPath: "/usr/bin/unzip")
        p.arguments = ["-p", out.path, "word/header1.xml"]
        let pipe = Pipe()
        p.standardOutput = pipe
        try p.run()
        p.waitUntilExit()
        let data = pipe.fileHandleForReading.readDataToEndOfFile()
        let header1 = String(data: data, encoding: .utf8) ?? ""
        print("output header1.xml chars: \(header1.count)")
        print("contains <w:pict: \(header1.contains("<w:pict"))")
        print("contains <v:shape: \(header1.contains("<v:shape"))")
        XCTAssertTrue(header1.contains("<w:pict"), "round-trip output dropped <w:pict>")
    }
}
