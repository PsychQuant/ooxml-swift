import Foundation
import ZIPFoundation

/// ZIP 壓縮/解壓縮工具
public struct ZipHelper {
    /// 解壓縮 ZIP 檔案到臨時目錄
    public static func unzip(_ url: URL) throws -> URL {
        let tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("che-word-mcp")
            .appendingPathComponent(UUID().uuidString)

        try FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
        try FileManager.default.unzipItem(at: url, to: tempDir)

        return tempDir
    }

    /// 壓縮目錄內容為 ZIP 檔案（不包含目錄本身的路徑）
    public static func zip(_ directory: URL, to destination: URL) throws {
        let data = try zipToData(directory)

        if FileManager.default.fileExists(atPath: destination.path) {
            try FileManager.default.removeItem(at: destination)
        }
        try data.write(to: destination)
    }

    /// 壓縮目錄內容為 in-memory ZIP bytes（不寫入磁碟）
    public static func zipToData(_ directory: URL) throws -> Data {
        let archive: Archive
        do {
            archive = try Archive(accessMode: .create)
        } catch {
            throw WordError.zipError("無法建立 in-memory ZIP archive: \(error)")
        }

        let files = try getAllFiles(in: directory)

        for (relativePath, fileURL) in files {
            let fileData = try Data(contentsOf: fileURL)
            try archive.addEntry(
                with: relativePath,
                type: .file,
                uncompressedSize: Int64(fileData.count),
                compressionMethod: .deflate,
                provider: { position, size in
                    let startIndex = fileData.startIndex.advanced(by: Int(position))
                    let endIndex = startIndex.advanced(by: size)
                    return fileData.subdata(in: startIndex..<endIndex)
                }
            )
        }

        guard let data = archive.data else {
            throw WordError.zipError("in-memory ZIP archive 無 data 可讀")
        }
        return data
    }

    /// 取得目錄內所有檔案（回傳相對路徑和完整 URL 的配對）
    private static func getAllFiles(in directory: URL) throws -> [(String, URL)] {
        var result: [(String, URL)] = []
        let fileManager = FileManager.default

        // 使用 subpathsOfDirectory 取得所有子路徑
        let directoryPath = directory.path
        let subpaths = try fileManager.subpathsOfDirectory(atPath: directoryPath)

        for subpath in subpaths {
            let fullURL = directory.appendingPathComponent(subpath)
            var isDirectory: ObjCBool = false
            if fileManager.fileExists(atPath: fullURL.path, isDirectory: &isDirectory) {
                if !isDirectory.boolValue {
                    result.append((subpath, fullURL))
                }
            }
        }

        return result
    }

    /// 清理臨時目錄
    public static func cleanup(_ directory: URL) {
        try? FileManager.default.removeItem(at: directory)
    }
}
