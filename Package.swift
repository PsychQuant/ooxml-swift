// swift-tools-version: 5.9
import PackageDescription

let package = Package(
    name: "OOXMLSwift",
    platforms: [.macOS(.v13)],
    products: [
        .library(name: "OOXMLSwift", targets: ["OOXMLSwift"]),
        .library(name: "WordDSLSwift", targets: ["WordDSLSwift"])
    ],
    dependencies: [
        .package(url: "https://github.com/weichsel/ZIPFoundation.git", from: "0.9.0")
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
                "Fixtures/mdocx"
            ]
        ),
        .testTarget(
            name: "WordDSLSwiftTests",
            dependencies: ["WordDSLSwift"]
        )
    ]
)
