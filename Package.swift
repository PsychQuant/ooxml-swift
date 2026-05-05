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
            dependencies: ["OOXMLSwift"]
        ),
        .testTarget(
            name: "WordDSLSwiftTests",
            dependencies: ["WordDSLSwift"]
        )
    ]
)
