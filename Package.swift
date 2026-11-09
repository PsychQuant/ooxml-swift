// swift-tools-version: 5.9
import PackageDescription

let package = Package(
    name: "OOXMLSwift",
    platforms: [.macOS(.v13)],
    products: [
        .library(name: "OOXMLSwift", targets: ["OOXMLSwift"]),
        .library(name: "WordDSLSwift", targets: ["WordDSLSwift"]),
        // Form normalization by string-level edits of word/document.xml
        // (#238, moved from macdoc so che-word-mcp can use it too).
        .library(name: "DocxForm", targets: ["DocxForm"]),
        .executable(name: "WordDSLSwiftScriptRunner",
                    targets: ["WordDSLSwiftScriptRunner"])
    ],
    dependencies: [
        // 0.9.20: the throwing Archive(url:accessMode:) both OOXMLSwift and
        // DocxForm use first shipped in 0.9.17, so 0.9.0–0.9.16 never
        // compiled; 0.9.20 is the version resolved and tested (#238).
        .package(url: "https://github.com/weichsel/ZIPFoundation.git", from: "0.9.20")
    ],
    targets: [
        .target(
            name: "OOXMLSwift",
            dependencies: ["ZIPFoundation"]
        ),
        .target(
            name: "WordDSLSwift",
            dependencies: ["OOXMLSwift"]
        ),
        // Deliberately no OOXMLSwift dependency: DocxForm edits the named
        // elements of word/document.xml as strings and copies every other
        // part byte for byte, so a form is never re-serialized through the
        // typed model (#238; macdoc#249).
        .target(
            name: "DocxForm",
            dependencies: ["ZIPFoundation"]
        ),
        .executableTarget(
            name: "WordDSLSwiftScriptRunner"
        ),
        .testTarget(
            name: "OOXMLSwiftTests",
            dependencies: ["OOXMLSwift"],
            // The mdocx fixture corpus contains `.mdocx.swift` design-frozen
            // surface specs that reference the future `WordDSLSwift` DSL
            // (full result-builder body, not yet implemented). They are NOT
            // compiled as part of the test target — the runner
            // `MdocxFixtureCorpusTests` validates them as Phase A
            // tokenization-only well-formed Swift. Excluded entirely so
            // SwiftPM does not auto-include them as test sources, and listed
            // as resources so test code can locate them via Bundle.module
            // when Phase B activates. See `mdocx-fixture-corpus` change
            // design.md Decision 5.
            exclude: [
                "Fixtures/mdocx",
                "Fixtures/golden-corpus",
                // Committed PDF fixtures for RenderGeometryTests
                // (render-effect-semantics task 1.1) — located via #filePath,
                // not Bundle resources.
                "Fixtures/render"
            ]
        ),
        .testTarget(
            name: "WordDSLSwiftTests",
            dependencies: ["WordDSLSwift"]
        ),
        .testTarget(
            name: "DocxFormTests",
            dependencies: ["DocxForm", "ZIPFoundation"]
        ),
        // Imports DocxForm without @testable, in a target of its own: a
        // @testable import in any file of a test target can make internal
        // names visible to the other files, so only a separate target proves
        // that what client packages call is public (#238).
        .testTarget(
            name: "DocxFormPublicAPITests",
            dependencies: ["DocxForm"]
        )
    ]
)
