import Foundation

// MARK: - Image Reference

/// 圖片參照（儲存在 word/media/ 目錄中的圖片）
public struct ImageReference: Equatable {
    public var id: String           // 關係 ID (rId)
    public var fileName: String     // 檔案名稱 (image1.png)
    public var contentType: String  // MIME 類型 (image/png)
    public var data: Data          // 圖片二進位資料

    public init(id: String, fileName: String, contentType: String, data: Data) {
        self.id = id
        self.fileName = fileName
        self.contentType = contentType
        self.data = data
    }

    /// 從檔案路徑建立圖片參照
    public static func from(path: String, id: String) throws -> ImageReference {
        let url = URL(fileURLWithPath: path)
        let data = try Data(contentsOf: url)

        let fileName = url.lastPathComponent
        let ext = url.pathExtension.lowercased()
        let contentType = mimeType(for: ext)

        return ImageReference(
            id: id,
            fileName: fileName,
            contentType: contentType,
            data: data
        )
    }

    /// 從 Base64 字串建立圖片參照
    public static func from(base64: String, fileName: String, id: String) throws -> ImageReference {
        guard let data = Data(base64Encoded: base64) else {
            throw ImageError.invalidBase64
        }

        let ext = (fileName as NSString).pathExtension.lowercased()
        let contentType = mimeType(for: ext)

        return ImageReference(
            id: id,
            fileName: fileName,
            contentType: contentType,
            data: data
        )
    }

    /// 取得副檔名對應的 MIME 類型
    private static func mimeType(for ext: String) -> String {
        switch ext {
        case "png": return "image/png"
        case "jpg", "jpeg": return "image/jpeg"
        case "gif": return "image/gif"
        case "bmp": return "image/bmp"
        case "tiff", "tif": return "image/tiff"
        case "webp": return "image/webp"
        case "emf": return "image/x-emf"
        case "wmf": return "image/x-wmf"
        default: return "image/png"
        }
    }
}

// MARK: - Image Error

public enum ImageError: Error, LocalizedError {
    case invalidBase64
    case fileNotFound(String)
    case unsupportedFormat(String)
    case dimensionRequired

    /// PsychQuant/ooxml-swift#178: a pixel dimension that either converts to
    /// an EMU value outside `ST_PositiveCoordinate`'s valid range (`0` to
    /// `Drawing.maxPositiveCoordinateEMU`) or is non-positive.
    case dimensionOutOfRange(String)

    public var errorDescription: String? {
        switch self {
        case .invalidBase64:
            return "Invalid Base64 encoded image data"
        case .fileNotFound(let path):
            return "Image file not found: \(path)"
        case .unsupportedFormat(let format):
            return "Unsupported image format: \(format)"
        case .dimensionRequired:
            return "Image dimensions (width/height) are required"
        case .dimensionOutOfRange(let reason):
            return "Image dimension out of range: \(reason)"
        }
    }
}

// MARK: - Drawing

/// 繪圖元素（用於將圖片嵌入文件）
public struct Drawing: Equatable {
    public var type: DrawingType       // inline（行內）或 anchor（浮動）
    public var width: Int              // 寬度（EMU）
    public var height: Int             // 高度（EMU）
    public var imageId: String         // 圖片關係 ID (rId)
    public var name: String            // 圖片名稱
    public var description: String     // 圖片描述（alt text）

    // 樣式屬性
    public var hasBorder: Bool = false
    public var borderColor: String = "000000"
    public var borderWidth: Int = 9525  // EMU (約 0.75pt)
    public var hasShadow: Bool = false

    // 浮動定位屬性（僅 anchor 類型使用）
    public var anchorPosition: AnchorPosition = AnchorPosition()

    public init(type: DrawingType = .inline,
         width: Int,
         height: Int,
         imageId: String,
         name: String = "Picture",
         description: String = "") {
        self.type = type
        self.width = width
        self.height = height
        self.imageId = imageId
        self.name = name
        self.description = description
    }

    /// 建立浮動圖片
    public static func anchor(width: Int, height: Int, imageId: String,
                       position: AnchorPosition = AnchorPosition(),
                       name: String = "Picture") -> Drawing {
        var drawing = Drawing(type: .anchor, width: width, height: height, imageId: imageId, name: name)
        drawing.anchorPosition = position
        return drawing
    }

    /// PsychQuant/ooxml-swift#178: 1 pixel = 9525 EMU @ 96 DPI.
    static let emuPerPixel = 9525

    /// PsychQuant/ooxml-swift#178: `wp:extent/@cx`/`@cy`'s XSD type
    /// (`ST_PositiveCoordinate`) only allows `0` through this value
    /// (`27273042316900`). A pixel value that converts to more than this —
    /// or that overflows `Int` before the range check can even run,
    /// starting around `Int.max / emuPerPixel` ≈ 9.7×10¹⁴ — is invalid
    /// OOXML regardless of whether the multiplication traps.
    static let maxPositiveCoordinateEMU = 27_273_042_316_900

    /// PsychQuant/ooxml-swift#178: pixel → EMU conversion that validates
    /// BEFORE converting, throwing `ImageError.dimensionOutOfRange` for any
    /// `px` that is non-positive, or whose EMU equivalent would overflow
    /// `Int` or exceed `ST_PositiveCoordinate`'s range — instead of the
    /// bare `px * emuPerPixel` multiplication `from`/`anchorFromPixels`
    /// used to do unconditionally, which TRAPS (crashing the whole host
    /// process, not just failing this one call) for `px` beyond roughly
    /// `Int.max / emuPerPixel`. Any caller that forwards external input
    /// (an MCP server, a CLI argument, a value read from an untrusted file)
    /// to a pixel-dimension API was one crafted parameter away from taking
    /// down the entire process — including every other open document's
    /// unsaved edits in a session that holds more than one.
    static func validatedEMU(fromPixels px: Int) throws -> Int {
        guard px > 0 else {
            throw ImageError.dimensionOutOfRange(
                "pixel value \(px) must be positive")
        }
        let (emu, overflow) = px.multipliedReportingOverflow(by: emuPerPixel)
        guard !overflow, emu <= maxPositiveCoordinateEMU else {
            throw ImageError.dimensionOutOfRange(
                "pixel value \(px) converts to an EMU value outside ST_PositiveCoordinate's range (0...\(maxPositiveCoordinateEMU))")
        }
        return emu
    }

    /// PsychQuant/ooxml-swift#178: same conversion as `validatedEMU(fromPixels:)`
    /// but NEVER throws and NEVER traps — clamps to `maxPositiveCoordinateEMU`
    /// (or `0` for non-positive input) instead. Backs the deprecated
    /// non-throwing `from`/`anchorFromPixels` overloads below, which cannot
    /// throw without breaking every existing call site (a source-breaking,
    /// major-version change this fix deliberately avoids — see those
    /// overloads' deprecation messages). Prefer `validatedEMU(fromPixels:)`
    /// (or the `validated`/`anchorValidated` constructors) in new code: this
    /// silently produces a DIFFERENT size than requested for out-of-range
    /// input rather than surfacing the problem to the caller.
    private static func clampedEMU(fromPixels px: Int) -> Int {
        guard px > 0 else { return 0 }
        let (emu, overflow) = px.multipliedReportingOverflow(by: emuPerPixel)
        return (overflow || emu > maxPositiveCoordinateEMU) ? maxPositiveCoordinateEMU : emu
    }

    /// 從像素建立浮動圖片（驗證版本；超出 `ST_PositiveCoordinate` 範圍會拋錯，不當機）
    public static func anchorValidated(widthPx: Int, heightPx: Int, imageId: String,
                                 position: AnchorPosition = AnchorPosition(),
                                 name: String = "Picture") throws -> Drawing {
        return anchor(
            width: try validatedEMU(fromPixels: widthPx),
            height: try validatedEMU(fromPixels: heightPx),
            imageId: imageId, position: position, name: name)
    }

    /// 從像素建立（驗證版本；1 像素 = 9525 EMU @ 96 DPI；超出 `ST_PositiveCoordinate`
    /// 範圍會拋錯，不當機）
    public static func validated(widthPx: Int, heightPx: Int, imageId: String, name: String = "Picture") throws -> Drawing {
        return Drawing(
            width: try validatedEMU(fromPixels: widthPx),
            height: try validatedEMU(fromPixels: heightPx),
            imageId: imageId,
            name: name
        )
    }

    /// 從像素建立浮動圖片
    ///
    /// PsychQuant/ooxml-swift#178: kept for source compatibility (changing
    /// this to `throws` would break every existing call site). Extreme
    /// `widthPx`/`heightPx` values no longer TRAP — they clamp to
    /// `Drawing.maxPositiveCoordinateEMU` instead, silently — but still do
    /// not validate against real OOXML bounds the way a caller checking for
    /// an error would expect. Use `anchorValidated(widthPx:heightPx:imageId:position:name:)`
    /// in new code, which throws `ImageError.dimensionOutOfRange` instead.
    @available(*, deprecated, message: "Silently clamps out-of-range pixel values instead of reporting them. Use anchorValidated(widthPx:heightPx:imageId:position:name:) to get a thrown ImageError.dimensionOutOfRange instead.")
    public static func anchorFromPixels(widthPx: Int, heightPx: Int, imageId: String,
                                 position: AnchorPosition = AnchorPosition(),
                                 name: String = "Picture") -> Drawing {
        return anchor(width: clampedEMU(fromPixels: widthPx), height: clampedEMU(fromPixels: heightPx), imageId: imageId, position: position, name: name)
    }

    /// 從像素建立（1 像素 = 9525 EMU @ 96 DPI）
    ///
    /// PsychQuant/ooxml-swift#178: kept for source compatibility — see
    /// `anchorFromPixels`'s doc comment for the full rationale. Use
    /// `validated(widthPx:heightPx:imageId:name:)` in new code.
    @available(*, deprecated, message: "Silently clamps out-of-range pixel values instead of reporting them. Use validated(widthPx:heightPx:imageId:name:) to get a thrown ImageError.dimensionOutOfRange instead.")
    public static func from(widthPx: Int, heightPx: Int, imageId: String, name: String = "Picture") -> Drawing {
        return Drawing(
            width: clampedEMU(fromPixels: widthPx),
            height: clampedEMU(fromPixels: heightPx),
            imageId: imageId,
            name: name
        )
    }

    /// 從英寸建立（1 英寸 = 914400 EMU）
    public static func from(widthInches: Double, heightInches: Double, imageId: String, name: String = "Picture") -> Drawing {
        return Drawing(
            width: Int(widthInches * 914400),
            height: Int(heightInches * 914400),
            imageId: imageId,
            name: name
        )
    }

    /// 從公分建立（1 公分 = 360000 EMU）
    public static func from(widthCm: Double, heightCm: Double, imageId: String, name: String = "Picture") -> Drawing {
        return Drawing(
            width: Int(widthCm * 360000),
            height: Int(heightCm * 360000),
            imageId: imageId,
            name: name
        )
    }

    /// 取得寬度（像素）
    public var widthInPixels: Int {
        width / Drawing.emuPerPixel
    }

    /// 取得高度（像素）
    public var heightInPixels: Int {
        height / Drawing.emuPerPixel
    }
}

/// 繪圖類型
public enum DrawingType: Equatable {
    case inline    // 行內（隨文字流動）
    case anchor    // 浮動（絕對或相對定位）
}

// MARK: - Anchor Positioning (浮動圖片定位)

/// 水平參照點
public enum HorizontalRelativeFrom: String {
    case margin = "margin"          // 頁邊界
    case page = "page"              // 頁面
    case column = "column"          // 欄（預設）
    case character = "character"    // 字元
    case leftMargin = "leftMargin"  // 左邊界
    case rightMargin = "rightMargin" // 右邊界
    case insideMargin = "insideMargin"   // 內側邊界
    case outsideMargin = "outsideMargin" // 外側邊界
}

/// 垂直參照點
public enum VerticalRelativeFrom: String {
    case margin = "margin"           // 頁邊界
    case page = "page"               // 頁面
    case paragraph = "paragraph"     // 段落（預設）
    case line = "line"               // 行
    case topMargin = "topMargin"     // 上邊界
    case bottomMargin = "bottomMargin" // 下邊界
    case insideMargin = "insideMargin"   // 內側邊界
    case outsideMargin = "outsideMargin" // 外側邊界
}

/// 水平對齊方式
public enum HorizontalAlignment: String {
    case left = "left"
    case center = "center"
    case right = "right"
    case inside = "inside"
    case outside = "outside"
}

/// 垂直對齊方式
public enum VerticalAlignment: String {
    case top = "top"
    case center = "center"
    case bottom = "bottom"
    case inside = "inside"
    case outside = "outside"
}

/// 文繞圖方式
public enum WrapType: Equatable {
    case none           // 無文繞圖（圖片在文字上方或下方）
    case square         // 方形文繞圖
    case tight          // 緊密文繞圖
    case through        // 穿透文繞圖
    case topAndBottom   // 上下文繞圖（文字在圖片上下）
    case behindText     // 浮於文字下方
    case inFrontOfText  // 浮於文字上方

    public var xmlElement: String {
        switch self {
        case .none: return ""
        case .square: return "<wp:wrapSquare wrapText=\"bothSides\"/>"
        case .tight: return "<wp:wrapTight wrapText=\"bothSides\"/>"
        case .through: return "<wp:wrapThrough wrapText=\"bothSides\"/>"
        case .topAndBottom: return "<wp:wrapTopAndBottom/>"
        case .behindText: return ""  // 由 behindDoc 屬性控制
        case .inFrontOfText: return ""  // 由 behindDoc 屬性控制
        }
    }

    public var behindDoc: Bool {
        return self == .behindText
    }
}

/// 浮動圖片定位設定
public struct AnchorPosition: Equatable {
    // 水平定位
    public var horizontalRelativeFrom: HorizontalRelativeFrom = .column
    public var horizontalOffset: Int? = nil  // EMU 偏移量
    public var horizontalAlignment: HorizontalAlignment? = nil  // 或使用對齊

    // 垂直定位
    public var verticalRelativeFrom: VerticalRelativeFrom = .paragraph
    public var verticalOffset: Int? = nil    // EMU 偏移量
    public var verticalAlignment: VerticalAlignment? = nil      // 或使用對齊

    // 文繞圖
    public var wrapType: WrapType = .square

    // 其他選項
    public var allowOverlap: Bool = true     // 允許重疊
    public var layoutInCell: Bool = true     // 在表格儲存格內配置
    public var locked: Bool = false          // 鎖定位置
    public var distanceTop: Int = 0          // 上方距離 (EMU)
    public var distanceBottom: Int = 0       // 下方距離 (EMU)
    public var distanceLeft: Int = 114300    // 左方距離 (EMU, 預設約 0.125")
    public var distanceRight: Int = 114300   // 右方距離 (EMU)

    public init() {}

    /// 建立置中於頁面的定位
    public static func centeredOnPage() -> AnchorPosition {
        var pos = AnchorPosition()
        pos.horizontalRelativeFrom = .page
        pos.horizontalAlignment = .center
        pos.verticalRelativeFrom = .page
        pos.verticalAlignment = .center
        pos.wrapType = .square
        return pos
    }

    /// 建立靠右定位
    public static func alignRight() -> AnchorPosition {
        var pos = AnchorPosition()
        pos.horizontalRelativeFrom = .margin
        pos.horizontalAlignment = .right
        pos.verticalRelativeFrom = .paragraph
        pos.verticalOffset = 0
        pos.wrapType = .square
        return pos
    }

    /// 建立靠左定位
    public static func alignLeft() -> AnchorPosition {
        var pos = AnchorPosition()
        pos.horizontalRelativeFrom = .margin
        pos.horizontalAlignment = .left
        pos.verticalRelativeFrom = .paragraph
        pos.verticalOffset = 0
        pos.wrapType = .square
        return pos
    }

    /// 建立絕對定位（使用 EMU）
    public static func absolute(x: Int, y: Int, relativeTo: (HorizontalRelativeFrom, VerticalRelativeFrom) = (.page, .page)) -> AnchorPosition {
        var pos = AnchorPosition()
        pos.horizontalRelativeFrom = relativeTo.0
        pos.horizontalOffset = x
        pos.verticalRelativeFrom = relativeTo.1
        pos.verticalOffset = y
        pos.wrapType = .square
        return pos
    }

    /// 建立浮於文字上方的定位
    public static func floatAboveText() -> AnchorPosition {
        var pos = AnchorPosition()
        pos.wrapType = .inFrontOfText
        return pos
    }

    /// 建立浮於文字下方的定位
    public static func floatBehindText() -> AnchorPosition {
        var pos = AnchorPosition()
        pos.wrapType = .behindText
        return pos
    }
}

// MARK: - XML Generation

extension Drawing {
    /// 轉換為 OOXML XML（完整的 drawing 元素，放在 Run 內）
    func toXML() -> String {
        switch type {
        case .inline:
            return toInlineXML()
        case .anchor:
            return toAnchorXML()
        }
    }

    /// 行內繪圖 XML
    private func toInlineXML() -> String {
        return """
        <w:drawing>
            <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                       distT="0" distB="0" distL="0" distR="0">
                <wp:extent cx="\(width)" cy="\(height)"/>
                <wp:effectExtent l="0" t="0" r="0" b="0"/>
                <wp:docPr id="1" name="\(escapeXML(name))" descr="\(escapeXML(description))"/>
                <wp:cNvGraphicFramePr>
                    <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                </wp:cNvGraphicFramePr>
                <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                            <pic:nvPicPr>
                                <pic:cNvPr id="0" name="\(escapeXML(name))"/>
                                <pic:cNvPicPr/>
                            </pic:nvPicPr>
                            <pic:blipFill>
                                <a:blip r:embed="\(escapeXMLAttribute(imageId))" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                                <a:stretch>
                                    <a:fillRect/>
                                </a:stretch>
                            </pic:blipFill>
                            <pic:spPr>
                                <a:xfrm>
                                    <a:off x="0" y="0"/>
                                    <a:ext cx="\(width)" cy="\(height)"/>
                                </a:xfrm>
                                <a:prstGeom prst="rect">
                                    <a:avLst/>
                                </a:prstGeom>
                                \(borderXML())
                            </pic:spPr>
                        </pic:pic>
                    </a:graphicData>
                </a:graphic>
            </wp:inline>
        </w:drawing>
        """
    }

    /// 浮動繪圖 XML（完整版，支援完整定位選項）
    private func toAnchorXML() -> String {
        let pos = anchorPosition

        // 水平定位 XML
        let horizontalXML: String
        if let alignment = pos.horizontalAlignment {
            horizontalXML = """
            <wp:positionH relativeFrom="\(pos.horizontalRelativeFrom.rawValue)">
                <wp:align>\(alignment.rawValue)</wp:align>
            </wp:positionH>
            """
        } else {
            horizontalXML = """
            <wp:positionH relativeFrom="\(pos.horizontalRelativeFrom.rawValue)">
                <wp:posOffset>\(pos.horizontalOffset ?? 0)</wp:posOffset>
            </wp:positionH>
            """
        }

        // 垂直定位 XML
        let verticalXML: String
        if let alignment = pos.verticalAlignment {
            verticalXML = """
            <wp:positionV relativeFrom="\(pos.verticalRelativeFrom.rawValue)">
                <wp:align>\(alignment.rawValue)</wp:align>
            </wp:positionV>
            """
        } else {
            verticalXML = """
            <wp:positionV relativeFrom="\(pos.verticalRelativeFrom.rawValue)">
                <wp:posOffset>\(pos.verticalOffset ?? 0)</wp:posOffset>
            </wp:positionV>
            """
        }

        // 文繞圖 XML
        let wrapXML = pos.wrapType.xmlElement
        let behindDocValue = pos.wrapType.behindDoc ? "1" : "0"

        return """
        <w:drawing>
            <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                       distT="\(pos.distanceTop)" distB="\(pos.distanceBottom)"
                       distL="\(pos.distanceLeft)" distR="\(pos.distanceRight)"
                       simplePos="0" relativeHeight="0" behindDoc="\(behindDocValue)"
                       locked="\(pos.locked ? "1" : "0")"
                       layoutInCell="\(pos.layoutInCell ? "1" : "0")"
                       allowOverlap="\(pos.allowOverlap ? "1" : "0")">
                <wp:simplePos x="0" y="0"/>
                \(horizontalXML)
                \(verticalXML)
                <wp:extent cx="\(width)" cy="\(height)"/>
                <wp:effectExtent l="0" t="0" r="0" b="0"/>
                \(wrapXML)
                <wp:docPr id="1" name="\(escapeXML(name))" descr="\(escapeXML(description))"/>
                <wp:cNvGraphicFramePr>
                    <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                </wp:cNvGraphicFramePr>
                <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                            <pic:nvPicPr>
                                <pic:cNvPr id="0" name="\(escapeXML(name))"/>
                                <pic:cNvPicPr/>
                            </pic:nvPicPr>
                            <pic:blipFill>
                                <a:blip r:embed="\(escapeXMLAttribute(imageId))" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                                <a:stretch>
                                    <a:fillRect/>
                                </a:stretch>
                            </pic:blipFill>
                            <pic:spPr>
                                <a:xfrm>
                                    <a:off x="0" y="0"/>
                                    <a:ext cx="\(width)" cy="\(height)"/>
                                </a:xfrm>
                                <a:prstGeom prst="rect">
                                    <a:avLst/>
                                </a:prstGeom>
                                \(borderXML())
                            </pic:spPr>
                        </pic:pic>
                    </a:graphicData>
                </a:graphic>
            </wp:anchor>
        </w:drawing>
        """
    }

    /// 邊框 XML
    private func borderXML() -> String {
        guard hasBorder else { return "" }
        return """
        <a:ln w="\(borderWidth)">
            <a:solidFill>
                <a:srgbClr val="\(escapeXMLAttribute(borderColor))"/>
            </a:solidFill>
        </a:ln>
        """
    }

    /// XML 跳脫（僅用於 `wp:docPr` 的 `name`／`descr` 屬性值——PsychQuant/ooxml-swift#165：
    /// 換行／CR／tab 若寫成字面字元，下一次解析會被 XML 正規化成空白，故一併轉成字元參照）
    private func escapeXML(_ string: String) -> String {
        return string
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
            .replacingOccurrences(of: "'", with: "&apos;")
            .replacingOccurrences(of: "\n", with: "&#10;")
            .replacingOccurrences(of: "\r", with: "&#13;")
            .replacingOccurrences(of: "\t", with: "&#9;")
    }
}

// MARK: - Run with Drawing

extension Run {
    /// 建立含圖片的 Run
    public static func withDrawing(_ drawing: Drawing) -> Run {
        var run = Run(text: "")
        run.drawing = drawing
        return run
    }
}
