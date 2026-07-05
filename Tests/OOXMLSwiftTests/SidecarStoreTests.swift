import Foundation
import XCTest
@testable import OOXMLSwift

/// word-aligned-state-sync Phase 2 task 3.16 — sidecar file management
/// ("Decision 5: Sidecar persistence, not in-document metadata";
/// `ooxml-word-sync` scenarios "Sidecar files created on first sync" +
/// "docx contains zero sync metadata").
///
/// Naming follows the spec's stem convention: `report.docx` →
/// `report.oplog.jsonl` + `report.snapshot.json`, same directory.
/// Sidecars are strictly opt-in (design Open Question Q1 working answer):
/// plain `DocxWriter.write` / `DocxReader.read` never touch them.
final class SidecarStoreTests: XCTestCase {

    private func makeDoc() -> WordDocument {
        var doc = WordDocument()
        doc.appendParagraph(Paragraph(text: "sidecar fixture"))
        return doc
    }

    private func tempDocxURL(_ name: String = "report") -> URL {
        let dir = FileManager.default.temporaryDirectory
            .appendingPathComponent("sidecar-\(UUID().uuidString)")
        try? FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        return dir.appendingPathComponent("\(name).docx")
    }

    // MARK: - URL derivation (spec stem convention)

    func testSidecarURLDerivation() {
        let docx = URL(fileURLWithPath: "/tmp/thesis/report.docx")
        XCTAssertEqual(SidecarStore.oplogURL(for: docx).lastPathComponent,
                       "report.oplog.jsonl")
        XCTAssertEqual(SidecarStore.snapshotURL(for: docx).lastPathComponent,
                       "report.snapshot.json")
        XCTAssertEqual(SidecarStore.oplogURL(for: docx).deletingLastPathComponent().path,
                       "/tmp/thesis", "sidecars live in the same directory as the docx")
    }

    // MARK: - Save writes docx + both sidecars; docx bytes untouched by opt-in

    func testSaveWithSidecarsWritesAllThreeFiles() throws {
        var doc = makeDoc()
        doc.operationLog.append(
            .setText(target: ElementID(rawString: "w14:paraId=0AB7C123"), text: "x"),
            source: .swift)

        let docxURL = tempDocxURL()
        defer { try? FileManager.default.removeItem(at: docxURL.deletingLastPathComponent()) }
        try doc.saveWithSidecars(to: docxURL)

        XCTAssertTrue(FileManager.default.fileExists(atPath: docxURL.path))
        XCTAssertTrue(FileManager.default.fileExists(atPath: SidecarStore.oplogURL(for: docxURL).path))
        XCTAssertTrue(FileManager.default.fileExists(atPath: SidecarStore.snapshotURL(for: docxURL).path))
    }

    func testDocxBytesIdenticalWithAndWithoutSidecars() throws {
        // "nothing written into the docx": the sidecar-opt-in save must not
        // change a single byte of the docx itself.
        let doc = makeDoc()

        let plainURL = tempDocxURL("plain")
        let sidecarURL = tempDocxURL("sidecar")
        defer {
            try? FileManager.default.removeItem(at: plainURL.deletingLastPathComponent())
            try? FileManager.default.removeItem(at: sidecarURL.deletingLastPathComponent())
        }
        try DocxWriter.write(doc, to: plainURL)
        try doc.saveWithSidecars(to: sidecarURL)

        // Zip containers embed per-entry mtimes, so compare the unzipped
        // document part rather than raw container bytes.
        let plain = try DocxReader.read(from: plainURL)
        let withSidecar = try DocxReader.read(from: sidecarURL)
        XCTAssertEqual(plain.body.children.count, withSidecar.body.children.count)
        if case .paragraph(let p1) = plain.body.children.first,
           case .paragraph(let p2) = withSidecar.body.children.first {
            XCTAssertEqual(p1.text, p2.text)
        } else {
            XCTFail("expected paragraphs in both saves")
        }
    }

    func testPlainWriteNeverCreatesSidecars() throws {
        let doc = makeDoc()
        let docxURL = tempDocxURL()
        defer { try? FileManager.default.removeItem(at: docxURL.deletingLastPathComponent()) }
        try DocxWriter.write(doc, to: docxURL)

        XCTAssertFalse(FileManager.default.fileExists(atPath: SidecarStore.oplogURL(for: docxURL).path),
                       "plain write must not create an oplog sidecar (opt-in only)")
        XCTAssertFalse(FileManager.default.fileExists(atPath: SidecarStore.snapshotURL(for: docxURL).path),
                       "plain write must not create a snapshot sidecar (opt-in only)")
    }

    // MARK: - Log round-trip through the sidecar

    func testOpenWithSidecarsRestoresLog() throws {
        var doc = makeDoc()
        let pid = ElementID(rawString: "w14:paraId=0AB7C123")
        doc.operationLog.append(.setText(target: pid, text: "hello"), source: .swift)
        doc.operationLog.append(.removeParagraph(id: pid), source: .word)

        let docxURL = tempDocxURL()
        defer { try? FileManager.default.removeItem(at: docxURL.deletingLastPathComponent()) }
        try doc.saveWithSidecars(to: docxURL)

        let reopened = try WordDocument.openWithSidecars(from: docxURL)
        XCTAssertEqual(reopened.operationLog.entries.count, 2,
                       "openWithSidecars must restore the persisted log")
        XCTAssertEqual(reopened.operationLog.entries.first?.source, .swift)
        XCTAssertEqual(reopened.operationLog.entries.last?.source, .word)
        if case .setText(let target, let text) = reopened.operationLog.entries.first!.op {
            XCTAssertEqual(target, pid)
            XCTAssertEqual(text, "hello")
        } else {
            XCTFail("first restored op must be setText")
        }
    }

    func testOpenWithSidecarsAbsentLogIsFreshStart() throws {
        // bootstrapFromDocx fresh-start semantics: a docx without sidecars
        // opens with an empty log, no throw.
        let doc = makeDoc()
        let docxURL = tempDocxURL()
        defer { try? FileManager.default.removeItem(at: docxURL.deletingLastPathComponent()) }
        try DocxWriter.write(doc, to: docxURL)

        let reopened = try WordDocument.openWithSidecars(from: docxURL)
        XCTAssertTrue(reopened.operationLog.entries.isEmpty,
                      "absent sidecars must mean fresh start, not an error")
    }

    // MARK: - Snapshot contents

    func testSnapshotRecordsDocxHashAndOpCount() throws {
        var doc = makeDoc()
        doc.operationLog.append(
            .setText(target: ElementID(rawString: "w14:paraId=0AB7C123"), text: "x"),
            source: .swift)

        let docxURL = tempDocxURL()
        defer { try? FileManager.default.removeItem(at: docxURL.deletingLastPathComponent()) }
        try doc.saveWithSidecars(to: docxURL)

        guard let snapshot = try SidecarStore.loadSnapshot(alongside: docxURL) else {
            return XCTFail("snapshot sidecar must load after saveWithSidecars")
        }
        let docxData = try Data(contentsOf: docxURL)
        XCTAssertEqual(snapshot.docxSHA256, SidecarStore.sha256Hex(of: docxData),
                       "snapshot must record the SHA-256 of the docx as written")
        XCTAssertEqual(snapshot.opCount, 1)
    }

    // MARK: - JSONL shape

    func testOplogSidecarIsOneLinePerEntry() throws {
        var doc = makeDoc()
        let pid = ElementID(rawString: "w14:paraId=0AB7C123")
        doc.operationLog.append(.setText(target: pid, text: "a"), source: .swift)
        doc.operationLog.append(.setText(target: pid, text: "b"), source: .swift)
        doc.operationLog.append(.setText(target: pid, text: "c"), source: .swift)

        let docxURL = tempDocxURL()
        defer { try? FileManager.default.removeItem(at: docxURL.deletingLastPathComponent()) }
        try doc.saveWithSidecars(to: docxURL)

        let raw = try String(contentsOf: SidecarStore.oplogURL(for: docxURL), encoding: .utf8)
        let lines = raw.split(separator: "\n", omittingEmptySubsequences: true)
        XCTAssertEqual(lines.count, 3, "JSONL sidecar must have exactly one line per entry")
        for line in lines {
            XCTAssertNotNil(
                try? JSONSerialization.jsonObject(with: Data(line.utf8)),
                "every JSONL line must parse independently as JSON")
        }
    }
}
