import Foundation

/// Native pixel dimensions of an image read from its on-disk header.
public struct ImageDimensions: Equatable {
    public let widthPx: Int
    public let heightPx: Int

    public var aspectRatio: Double {
        guard heightPx > 0 else { return 0 }
        return Double(widthPx) / Double(heightPx)
    }

    public init(widthPx: Int, heightPx: Int) {
        self.widthPx = widthPx
        self.heightPx = heightPx
    }
}

public enum ImageDimensionsError: Error, Equatable {
    case unsupportedFormat(ext: String)
    case fileNotFound(path: String)
    case invalidPNG
    case invalidJPEG
}

extension ImageDimensions {

    /// Read native pixel dimensions from an image file by parsing its format
    /// header. Supported: PNG (IHDR), JPEG (SOF0/SOF2 scan).
    ///
    /// - Throws: `ImageDimensionsError` on missing file, unsupported extension,
    ///   or malformed header.
    public static func detect(path: String) throws -> ImageDimensions {
        guard FileManager.default.fileExists(atPath: path) else {
            throw ImageDimensionsError.fileNotFound(path: path)
        }
        let url = URL(fileURLWithPath: path)
        let ext = url.pathExtension.lowercased()
        let data = try Data(contentsOf: url)
        switch ext {
        case "png":
            return try detectPNG(data: data)
        case "jpg", "jpeg":
            return try detectJPEG(data: data)
        default:
            throw ImageDimensionsError.unsupportedFormat(ext: ext)
        }
    }

    // MARK: PNG

    /// PNG layout: 8-byte signature + chunks. The first chunk is always IHDR,
    /// whose 13-byte payload starts with big-endian uint32 width + height.
    /// Signature: 89 50 4E 47 0D 0A 1A 0A (bytes 0-7)
    /// IHDR payload begins at byte 16 → width at 16-19, height at 20-23.
    private static func detectPNG(data: Data) throws -> ImageDimensions {
        guard data.count >= 24 else { throw ImageDimensionsError.invalidPNG }
        let signature: [UInt8] = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]
        for (i, expected) in signature.enumerated() where data[i] != expected {
            throw ImageDimensionsError.invalidPNG
        }
        let width = readBigEndianUInt32(data: data, at: 16)
        let height = readBigEndianUInt32(data: data, at: 20)
        return ImageDimensions(widthPx: Int(width), heightPx: Int(height))
    }

    // MARK: JPEG

    /// JPEG layout: SOI (FFD8), then a sequence of segments. Each segment is
    /// `FF <marker> <length-big-endian-2-bytes including length bytes> <data>`,
    /// except SOI/EOI/RST which have no length. Frame-start markers (SOFn =
    /// C0-C3, C5-C7, C9-CB, CD-CF) carry the height (2 BE) and width (2 BE).
    private static func detectJPEG(data: Data) throws -> ImageDimensions {
        guard data.count >= 4, data[0] == 0xFF, data[1] == 0xD8 else {
            throw ImageDimensionsError.invalidJPEG
        }
        let sofMarkers: Set<UInt8> = [
            0xC0, 0xC1, 0xC2, 0xC3,
            0xC5, 0xC6, 0xC7,
            0xC9, 0xCA, 0xCB,
            0xCD, 0xCE, 0xCF
        ]
        let standaloneMarkers: Set<UInt8> = [0x01, 0xD0, 0xD1, 0xD2, 0xD3, 0xD4, 0xD5, 0xD6, 0xD7, 0xD8, 0xD9]

        var i = 2
        while i < data.count - 1 {
            // Scan to next marker (0xFF followed by non-0x00 non-0xFF)
            guard data[i] == 0xFF else {
                i += 1
                continue
            }
            // Skip fill bytes
            while i < data.count - 1 && data[i] == 0xFF { i += 1 }
            let marker = data[i]
            i += 1

            if sofMarkers.contains(marker) {
                // SOF payload: 2-byte length + 1-byte precision + 2-byte height + 2-byte width
                guard data.count >= i + 7 else { throw ImageDimensionsError.invalidJPEG }
                let height = (Int(data[i + 3]) << 8) | Int(data[i + 4])
                let width = (Int(data[i + 5]) << 8) | Int(data[i + 6])
                return ImageDimensions(widthPx: width, heightPx: height)
            }
            if standaloneMarkers.contains(marker) {
                continue
            }
            // Other segments: read 2-byte length, skip the rest
            guard data.count >= i + 2 else { throw ImageDimensionsError.invalidJPEG }
            let length = (Int(data[i]) << 8) | Int(data[i + 1])
            i += length
        }
        throw ImageDimensionsError.invalidJPEG
    }

    private static func readBigEndianUInt32(data: Data, at offset: Int) -> UInt32 {
        return (UInt32(data[offset]) << 24)
            | (UInt32(data[offset + 1]) << 16)
            | (UInt32(data[offset + 2]) << 8)
            | UInt32(data[offset + 3])
    }
}
