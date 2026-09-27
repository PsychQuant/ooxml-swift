import Foundation
import XCTest
@testable import OOXMLSwift

/// Regression tests for [PsychQuant/ooxml-swift#189](https://github.com/PsychQuant/ooxml-swift/issues/189).
///
/// `OperationReducer.applyRunFormat`'s underline-off branch called
/// `removeRunProperty("u", from: run)`, deleting `<w:u>` outright — the same
/// defect class #161 fixed for bold/italic (`setBooleanRunProperty` now
/// writes an explicit `w:val="0"` instead of removing the element). Because
/// `<w:u>` is *valued* (not a bare on/off toggle), the fix isn't reusing
/// `setBooleanRunProperty`, but writing the explicit off value the schema
/// defines for underline: `w:val="none"`. Absence lets a run inside an
/// underlined paragraph/character style re-inherit the underline — so
/// `setUnderline(false)` was a no-op exactly where it mattered, mirroring
/// #161's bold finding exactly.
final class Issue189UnderlineOffWritesNoneTests: XCTestCase {

    private func makeBodyWithSingleRun(text: String) -> (XmlTree, ElementID) {
        let runUUID = UUID()
        let textNode = XmlNode.text(text)
        let wt = XmlNode.element(prefix: "w", localName: "t", children: [textNode])
        let wr = XmlNode.element(prefix: "w", localName: "r", children: [wt])
        wr.libraryUUID = runUUID
        let wp = XmlNode.element(prefix: "w", localName: "p", children: [wr])
        let body = XmlNode.element(prefix: "w", localName: "body", children: [wp])
        let doc = XmlNode.element(prefix: "w", localName: "document", children: [body])
        return (XmlTree.synthesized(root: doc), ElementID(libraryUUID: runUUID))
    }

    /// `<w:u>`'s `w:val` on the given run, or `nil` if the run has no `<w:u>`
    /// at all (element absent — the pre-fix bug's observable symptom).
    private func underlineVal(_ tree: XmlTree, runID: ElementID) -> String?? {
        guard let run = OperationReducer.findNode(elementID: runID, in: tree) else {
            return nil
        }
        guard let rPr = run.children.first(where: { $0.kind == .element && $0.localName == "rPr" }) else {
            return nil
        }
        guard let u = rPr.children.first(where: { $0.kind == .element && $0.localName == "u" }) else {
            return nil
        }
        return .some(u.attributes.first(where: { $0.prefix == "w" && $0.localName == "val" })?.value)
    }

    func testSetRunFormatUnderlineTrueWritesSingle() throws {
        let (base, runID) = makeBodyWithSingleRun(text: "hello")

        var log = OperationLog()
        log.append(.setRunFormat(target: runID, format: RunFormatPayload(underline: true)), source: .swift)

        let result = try OperationReducer.materialize(log: log, base: base)
        XCTAssertEqual(underlineVal(result, runID: runID), .some("single"))
    }

    /// The issue's exact repro: underline on, then off. `<w:u>` must SURVIVE
    /// as an explicit `w:val="none"`, not be deleted.
    func testSetRunFormatUnderlineFalseWritesExplicitNoneNotRemoval() throws {
        let (base, runID) = makeBodyWithSingleRun(text: "hello")

        var log = OperationLog()
        log.append(.setRunFormat(target: runID, format: RunFormatPayload(underline: true)), source: .swift)
        log.append(.setRunFormat(target: runID, format: RunFormatPayload(underline: false)), source: .swift)

        let result = try OperationReducer.materialize(log: log, base: base)
        guard let val = underlineVal(result, runID: runID) else {
            return XCTFail("<w:u> was deleted — underline would silently re-inherit from an enclosing style")
        }
        XCTAssertEqual(val, "none", "explicit off must be w:val=\"none\", not element absence")
    }

    /// Turning underline back on after an explicit off must overwrite the
    /// stale `w:val="none"`, not leave both values contradicting each other.
    func testSetRunFormatUnderlineTrueAfterFalseOverwritesNone() throws {
        let (base, runID) = makeBodyWithSingleRun(text: "hello")

        var log = OperationLog()
        log.append(.setRunFormat(target: runID, format: RunFormatPayload(underline: false)), source: .swift)
        log.append(.setRunFormat(target: runID, format: RunFormatPayload(underline: true)), source: .swift)

        let result = try OperationReducer.materialize(log: log, base: base)
        XCTAssertEqual(underlineVal(result, runID: runID), .some("single"))
    }
}
