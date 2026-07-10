// RenderGeometry.swift
// render-effect-semantics task 1.1 — geometry measurement over rendered PDFs
// (`docx-visual-diff-testing`, «Geometry measurement extends the harness
// beyond pixel ratios»; design Decisions 2 + 6). Native frameworks only
// (PDFKit/CoreGraphics). Lives in the TEST target on purpose: no public API
// surface, no release obligation — promotion to library API waits for a
// non-test consumer (design Decision 6).

import Foundation
import PDFKit

enum RenderGeometry {

    static func pageCount(pdf: PDFDocument) -> Int {
        pdf.pageCount
    }

    /// Media box of the page, or nil for an out-of-range index.
    static func pageBox(pdf: PDFDocument, page index: Int) -> CGRect? {
        guard let page = pdf.page(at: index) else { return nil }
        return page.bounds(for: .mediaBox)
    }

    /// Bounding boxes of the page's text lines (page space, bottom-left
    /// origin), ordered top-to-bottom. Empty for out-of-range pages or
    /// pages without extractable text.
    static func lineBoxes(pdf: PDFDocument, page index: Int) -> [CGRect] {
        guard let page = pdf.page(at: index) else { return [] }
        guard let whole = page.selection(for: page.bounds(for: .mediaBox)) else { return [] }
        return whole.selectionsByLine()
            .filter { !($0.string ?? "").trimmingCharacters(in: .whitespacesAndNewlines).isEmpty }
            .map { $0.bounds(for: page) }
            .filter { $0.height > 0 }
            .sorted { $0.midY > $1.midY }
    }

    /// Median baseline-to-baseline distance (pt) between consecutive text
    /// lines — the measured line pitch. Nil when the page has fewer than two
    /// text lines («Sparse page yields nil line pitch»: consumers that
    /// expected text on the page treat nil as a failure, not a skip).
    static func medianLinePitch(pdf: PDFDocument, page index: Int) -> CGFloat? {
        let boxes = lineBoxes(pdf: pdf, page: index)
        guard boxes.count >= 2 else { return nil }
        let deltas = zip(boxes, boxes.dropFirst()).map { $0.midY - $1.midY }.sorted()
        let mid = deltas.count / 2
        return deltas.count % 2 == 1
            ? deltas[mid]
            : (deltas[mid - 1] + deltas[mid]) / 2
    }
}
