// VisualDiffTests.swift
// format-alignment-engine Phase D task 4.3 — gated visual regression harness
// (`docx-visual-diff-testing`): docx → PDF via live Microsoft Word
// (AppleScript, the WordLiveRoundTripTests driving pattern), pages rendered
// with PDFKit/CoreGraphics (native channel only — no LibreOffice), per-page
// pixel-difference ratio against a threshold. Skips loudly without
// RUN_WORD_INTEGRATION=1 or without Word.
//
//     RUN_WORD_INTEGRATION=1 swift test --filter VisualDiffTests

import XCTest
import PDFKit
import CoreGraphics
@testable import OOXMLSwift

/// The harness proper — reusable by the acceptance run (task 4.4).
enum VisualDiffHarness {

    /// Pixel-difference threshold: a page fails when more than this share of
    /// its pixels differ. Heuristic by design (#130 Residue: aesthetic
    /// judgment is approximated, not captured).
    static let threshold = 0.005

    static func requireGate() throws {
        guard ProcessInfo.processInfo.environment["RUN_WORD_INTEGRATION"] == "1" else {
            throw XCTSkip("visual diff gated behind RUN_WORD_INTEGRATION=1")
        }
        guard FileManager.default.fileExists(atPath: "/Applications/Microsoft Word.app") else {
            throw XCTSkip("Microsoft Word not installed — the docx→PDF channel is Word only")
        }
    }

    /// docx → PDF via live Word. Throws XCTSkip when osascript cannot drive
    /// Word (automation permission) — environment problems skip, not fail.
    static func exportPDF(docx: URL, to pdf: URL) throws {
        let script = """
        tell application "Microsoft Word"
            open POSIX file "\(docx.path)"
            set theDoc to active document
            save as theDoc file name "\(pdf.path)" file format format PDF
            close theDoc saving no
        end tell
        """
        let process = Process()
        process.executableURL = URL(fileURLWithPath: "/usr/bin/osascript")
        process.arguments = ["-e", script]
        let pipe = Pipe()
        process.standardOutput = pipe
        process.standardError = pipe
        try process.run()
        process.waitUntilExit()
        let output = String(decoding: pipe.fileHandleForReading.readDataToEndOfFile(), as: UTF8.self)
        guard process.terminationStatus == 0,
              FileManager.default.fileExists(atPath: pdf.path) else {
            throw XCTSkip("osascript could not drive Word (automation permission?): \(output)")
        }
    }

    /// Per-page pixel-difference ratios between two PDFs. Page-count
    /// mismatches contribute ratio 1.0 for each unpaired page.
    static func pageDiffRatios(_ a: URL, _ b: URL) throws -> [Double] {
        guard let docA = PDFDocument(url: a), let docB = PDFDocument(url: b) else {
            throw NSError(domain: "VisualDiff", code: 1,
                          userInfo: [NSLocalizedDescriptionKey: "cannot open PDFs"])
        }
        let pages = max(docA.pageCount, docB.pageCount)
        var ratios: [Double] = []
        for i in 0..<pages {
            guard let pageA = docA.page(at: i), let pageB = docB.page(at: i) else {
                ratios.append(1.0)  // unpaired page = fully different
                continue
            }
            let pixelsA = render(pageA)
            let pixelsB = render(pageB)
            guard pixelsA.count == pixelsB.count, !pixelsA.isEmpty else {
                ratios.append(1.0)
                continue
            }
            var differing = 0
            // Compare per-pixel (4 bytes RGBA); count a pixel as differing
            // when any channel deviates by more than 8/255 (antialiasing
            // tolerance).
            for px in stride(from: 0, to: pixelsA.count, by: 4) {
                if abs(Int(pixelsA[px]) - Int(pixelsB[px])) > 8
                    || abs(Int(pixelsA[px + 1]) - Int(pixelsB[px + 1])) > 8
                    || abs(Int(pixelsA[px + 2]) - Int(pixelsB[px + 2])) > 8 {
                    differing += 1
                }
            }
            ratios.append(Double(differing) / Double(pixelsA.count / 4))
        }
        return ratios
    }

    /// Renders a PDF page to RGBA bytes at 150 dpi equivalent.
    private static func render(_ page: PDFPage) -> [UInt8] {
        let bounds = page.bounds(for: .mediaBox)
        let scale: CGFloat = 150.0 / 72.0
        let width = Int(bounds.width * scale)
        let height = Int(bounds.height * scale)
        guard width > 0, height > 0 else { return [] }
        var data = [UInt8](repeating: 255, count: width * height * 4)
        let colorSpace = CGColorSpaceCreateDeviceRGB()
        let renderResult: Bool = data.withUnsafeMutableBytes { buffer in
            guard let ctx = CGContext(
                data: buffer.baseAddress, width: width, height: height,
                bitsPerComponent: 8, bytesPerRow: width * 4, space: colorSpace,
                bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue) else { return false }
            ctx.setFillColor(CGColor(red: 1, green: 1, blue: 1, alpha: 1))
            ctx.fill(CGRect(x: 0, y: 0, width: CGFloat(width), height: CGFloat(height)))
            ctx.scaleBy(x: scale, y: scale)
            ctx.translateBy(x: -bounds.minX, y: -bounds.minY)
            page.draw(with: .mediaBox, to: ctx)
            return true
        }
        return renderResult ? data : []
    }
}

final class VisualDiffTests: XCTestCase {

    /// Multi-paragraph two-column source so a layout change is visually
    /// large.
    ///
    /// FIXED scratch path, deliberately NOT per-run UUID: sandboxed Word's
    /// `save as` into a never-authorized folder raises a BLOCKING
    /// "Grant Access" sheet until the 120s AppleEvent timeout, and a fresh
    /// UUID dir re-triggers it on every export of every run (diagnosed
    /// 2026-07-08 — the harness skipped on -1712/-1708 despite TCC being
    /// granted). With a stable path one manual grant persists across runs.
    /// Teardown clears the files but KEEPS the directory — deleting it could
    /// invalidate Word's path-based grant.
    private func makeScratchDir() throws -> URL {
        let dir = FileManager.default.homeDirectoryForCurrentUser
            .appendingPathComponent(".cache/ooxml-swift-visual-diff")
        try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
        // Clear stale artifacts from previous runs up front; keep the dir.
        for file in (try? FileManager.default.contentsOfDirectory(
            at: dir, includingPropertiesForKeys: nil)) ?? [] {
            try? FileManager.default.removeItem(at: file)
        }
        addTeardownBlock {
            for file in (try? FileManager.default.contentsOfDirectory(
                at: dir, includingPropertiesForKeys: nil)) ?? [] {
                try? FileManager.default.removeItem(at: file)
            }
        }
        return dir
    }

    private func makeTwoColumnDocx(at url: URL, twoColumn: Bool) throws {
        var ops: [OOXMLSwift.Operation] = []
        for i in 1...12 {
            ops.append(.appendParagraph(in: nil, paragraph: ParagraphPayload(
                text: "第\(i)段落。二段組の版面検証のための本文テキストがここに続きます。"
                    + "視覚回帰ハーネスは列レイアウトの変化を検出しなければなりません。",
                paraId: "P\(i)")))
        }
        ops.append(.setSectionProperties(at: nil, section: SectionPayload(
            pageWidth: 11906, pageHeight: 16838,
            marginTop: 1985, marginRight: 1701, marginBottom: 1701, marginLeft: 1701,
            columnCount: twoColumn ? 2 : nil, columnSpace: twoColumn ? 425 : nil)))
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: ops)
        try doc.writeAuthoringPackage(to: url)
    }

    /// Spec scenario: identical documents pass — a reference and its
    /// byte-equal rebuild render pixel-identical pages.
    func testIdenticalDocumentsPass() throws {
        try VisualDiffHarness.requireGate()
        let dir = try makeScratchDir()

        let reference = dir.appendingPathComponent("reference.docx")
        try makeTwoColumnDocx(at: reference, twoColumn: true)

        // Byte-equal rebuild through the format-alignment pipeline.
        let parts = try RawPartChannel.readAllParts(from: reference)
        let result = try ReverseExtractor.reverse(parts: parts)
        let parsed = try ScriptImporter.parse(source: ScriptExporter.exportSwift(log: result.log))
        var doc = WordDocument.emptyAuthoringDocument()
        try doc.apply(operations: parsed.entries.map(\.op))
        let rebuilt = dir.appendingPathComponent("rebuilt.docx")
        try doc.writeAuthoringPackage(to: rebuilt)

        let pdfA = dir.appendingPathComponent("reference.pdf")
        let pdfB = dir.appendingPathComponent("rebuilt.pdf")
        try VisualDiffHarness.exportPDF(docx: reference, to: pdfA)
        try VisualDiffHarness.exportPDF(docx: rebuilt, to: pdfB)

        let ratios = try VisualDiffHarness.pageDiffRatios(pdfA, pdfB)
        XCTAssertFalse(ratios.isEmpty)
        for (page, ratio) in ratios.enumerated() {
            XCTAssertEqual(ratio, 0.0, accuracy: 1e-9,
                           "page \(page + 1) must be pixel-identical (ratio \(ratio))")
        }
    }

    /// Spec scenario: layout drift is caught — dropping the second section's
    /// two-column layout pushes a page's ratio past the threshold, and the
    /// failure names the page.
    func testLayoutDriftIsCaught() throws {
        try VisualDiffHarness.requireGate()
        let dir = try makeScratchDir()

        let twoCol = dir.appendingPathComponent("two-column.docx")
        let oneCol = dir.appendingPathComponent("one-column.docx")
        try makeTwoColumnDocx(at: twoCol, twoColumn: true)
        try makeTwoColumnDocx(at: oneCol, twoColumn: false)

        let pdfA = dir.appendingPathComponent("two-column.pdf")
        let pdfB = dir.appendingPathComponent("one-column.pdf")
        try VisualDiffHarness.exportPDF(docx: twoCol, to: pdfA)
        try VisualDiffHarness.exportPDF(docx: oneCol, to: pdfB)

        let ratios = try VisualDiffHarness.pageDiffRatios(pdfA, pdfB)
        let failingPages = ratios.enumerated()
            .filter { $0.element > VisualDiffHarness.threshold }
            .map { $0.offset + 1 }
        XCTAssertFalse(failingPages.isEmpty,
                       "column-layout drift must exceed the threshold on at least one page "
                       + "(ratios: \(ratios))")
        // The comparison failure names the affected page(s).
        print("[visual-diff] layout drift caught on page(s) \(failingPages), ratios \(ratios)")
    }
}
