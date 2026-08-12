import Foundation

enum RunnerError: Error, CustomStringConvertible {
    case usage
    case processFailed(Int32, String)

    var description: String {
        switch self {
        case .usage:
            return "usage: WordDSLSwiftScriptRunner <script.mdocx.swift> <work-dir>"
        case .processFailed(let status, let output):
            return "fixture compilation/execution failed (status \(status)):\n\(output)"
        }
    }
}

func quotedSwiftString(_ value: String) -> String {
    var result = "\""
    for scalar in value.unicodeScalars {
        switch scalar.value {
        case 0x22: result += "\\\""
        case 0x5C: result += "\\\\"
        case 0x0A: result += "\\n"
        case 0x0D: result += "\\r"
        case 0x09: result += "\\t"
        default: result.unicodeScalars.append(scalar)
        }
    }
    return result + "\""
}

func canonicalizeSidecars(in workDir: URL) throws {
    let logURL = workDir.appendingPathComponent("out.docx.oplog.jsonl")
    if FileManager.default.fileExists(atPath: logURL.path) {
        let source = try String(contentsOf: logURL, encoding: .utf8)
        let lines = source.split(whereSeparator: \.isNewline)
        var canonical: [String] = []
        for (index, line) in lines.enumerated() {
            let data = Data(line.utf8)
            guard var object = try JSONSerialization.jsonObject(with: data) as? [String: Any] else {
                continue
            }
            object["op_id"] = String(format: "00000000-0000-4000-8000-%012d", index + 1)
            object["ts"] = "1970-01-01T00:00:00Z"
            let encoded = try JSONSerialization.data(withJSONObject: object, options: [.sortedKeys])
            canonical.append(String(decoding: encoded, as: UTF8.self))
        }
        try Data((canonical.joined(separator: "\n") + "\n").utf8).write(to: logURL)
    }

    let snapshotURL = workDir.appendingPathComponent("out.docx.snapshot.json")
    if FileManager.default.fileExists(atPath: snapshotURL.path) {
        let data = try Data(contentsOf: snapshotURL)
        guard var object = try JSONSerialization.jsonObject(with: data) as? [String: Any] else {
            return
        }
        object["savedAt"] = "1970-01-01T00:00:00Z"
        object["docxSHA256"] = "canonicalized-by-mdocx-fixture-runner"
        let encoded = try JSONSerialization.data(
            withJSONObject: object, options: [.prettyPrinted, .sortedKeys])
        try encoded.write(to: snapshotURL)
    }
}

func run() throws {
    let arguments = CommandLine.arguments
    guard arguments.count == 3 else { throw RunnerError.usage }

    let scriptURL = URL(fileURLWithPath: arguments[1]).standardizedFileURL
    let workDir = URL(fileURLWithPath: arguments[2]).standardizedFileURL
    let packageRoot = URL(fileURLWithPath: #filePath)
        .deletingLastPathComponent()
        .deletingLastPathComponent()
        .deletingLastPathComponent()
    let fixturePackage = workDir.appendingPathComponent("fixture-package")
    let sourceDir = fixturePackage.appendingPathComponent("Sources/Fixture")
    let fm = FileManager.default
    try fm.createDirectory(at: sourceDir, withIntermediateDirectories: true)

    let manifest = """
    // swift-tools-version: 5.9
    import PackageDescription
    let package = Package(
        name: "MdocxFixtureExecution",
        platforms: [.macOS(.v13)],
        dependencies: [
            .package(name: "OOXMLSwiftLocal", path: \(quotedSwiftString(packageRoot.path)))
        ],
        targets: [
            .executableTarget(
                name: "Fixture",
                dependencies: [
                    .product(name: "WordDSLSwift", package: "OOXMLSwiftLocal")
                ])
        ])
    """
    try Data(manifest.utf8).write(to: fixturePackage.appendingPathComponent("Package.swift"))

    var source = try String(contentsOf: scriptURL, encoding: .utf8)
    source += "\nimport Foundation\n"
    source += "try document.save(to: URL(fileURLWithPath: "
    source += quotedSwiftString(workDir.appendingPathComponent("out.docx").path)
    source += "))\n"
    try Data(source.utf8).write(to: sourceDir.appendingPathComponent("main.swift"))

    let process = Process()
    process.executableURL = URL(fileURLWithPath: "/usr/bin/env")
    process.arguments = [
        "swift", "run",
        "--package-path", fixturePackage.path,
        "--scratch-path", packageRoot.appendingPathComponent(
            ".build/mdocx-fixture-runner").path,
        "Fixture",
    ]
    let logURL = workDir.appendingPathComponent("fixture-build.log")
    try Data().write(to: logURL)
    let output = try FileHandle(forWritingTo: logURL)
    defer { try? output.close() }
    process.standardOutput = output
    process.standardError = output
    try process.run()
    process.waitUntilExit()
    try output.synchronize()
    let message = String(decoding: try Data(contentsOf: logURL), as: UTF8.self)
    guard process.terminationStatus == 0 else {
        throw RunnerError.processFailed(process.terminationStatus, message)
    }
    try canonicalizeSidecars(in: workDir)
}

do {
    try run()
} catch {
    FileHandle.standardError.write(Data("\(error)\n".utf8))
    exit(1)
}
