import Foundation

/// Word MCP 錯誤類型
public enum WordError: Error, LocalizedError {
    // 文件管理錯誤
    case documentNotOpen(String)
    case documentAlreadyOpen(String)
    case documentNotFound(String)

    // 檔案錯誤
    case fileNotFound(String)
    case invalidDocx(String)

    // 操作錯誤
    case invalidIndex(Int)
    case invalidFormat(String)

    // 解析/寫入錯誤
    case parseError(String)
    case writeError(String)
    case zipError(String)

    // MCP 錯誤
    case unknownTool(String)
    case missingParameter(String)
    case invalidParameter(String, String)  // (參數名, 原因)

    // SDT / Content Control 錯誤 (Change A: che-word-mcp-content-controls-read-write)
    case contentControlNotFound(Int)
    case unsupportedSDTType(SDTType)
    case disallowedElement(String)
    case repeatingSectionItemOutOfBounds(index: Int, count: Int)

    // Style / Numbering / Section 錯誤 (#44: che-word-mcp-styles-sections-numbering-foundations)
    case styleNotFound(String)
    case typeMismatch(expected: String, actual: String)
    case numIdNotFound(Int)
    case abstractNumIdNotFound(Int)

    // Tables / Hyperlinks / Headers 錯誤 (#44: che-word-mcp-tables-hyperlinks-headers-builtin)
    case nestedTooDeep(depth: Int, max: Int)
    case hyperlinkNotFound(String)

    // Track Changes 錯誤 (#45: che-word-mcp-track-changes-programmatic-generation)
    case trackChangesNotEnabled

    // 其他
    case unknownError(String)

    public var errorDescription: String? {
        switch self {
        case .documentNotOpen(let id):
            return "Document not open: \(id)"
        case .documentAlreadyOpen(let id):
            return "Document already open: \(id)"
        case .documentNotFound(let id):
            return "Document not found: \(id)"
        case .fileNotFound(let path):
            return "File not found: \(path)"
        case .invalidDocx(let reason):
            return "Invalid .docx file: \(reason)"
        case .invalidIndex(let index):
            return "Invalid index: \(index)"
        case .invalidFormat(let message):
            return "Invalid format: \(message)"
        case .parseError(let message):
            return "Parse error: \(message)"
        case .writeError(let message):
            return "Write error: \(message)"
        case .zipError(let message):
            return "ZIP error: \(message)"
        case .unknownTool(let name):
            return "Unknown tool: \(name)"
        case .missingParameter(let param):
            return "Missing required parameter: \(param)"
        case .invalidParameter(let param, let reason):
            return "Invalid parameter '\(param)': \(reason)"
        case .contentControlNotFound(let id):
            return "Content control not found: \(id)"
        case .unsupportedSDTType(let type):
            return "SDT type does not support this operation: \(type.rawValue)"
        case .disallowedElement(let name):
            return "Disallowed element in content XML: \(name)"
        case .repeatingSectionItemOutOfBounds(let index, let count):
            return "Repeating section item index \(index) out of bounds (count=\(count))"
        case .styleNotFound(let id):
            return "Style not found: \(id)"
        case .typeMismatch(let expected, let actual):
            return "Type mismatch: expected \(expected), got \(actual)"
        case .numIdNotFound(let id):
            return "Numbering numId not found: \(id)"
        case .abstractNumIdNotFound(let id):
            return "Numbering abstractNumId not found: \(id)"
        case .nestedTooDeep(let depth, let max):
            return "Nested table depth \(depth) exceeds maximum \(max)"
        case .hyperlinkNotFound(let id):
            return "Hyperlink not found: \(id)"
        case .trackChangesNotEnabled:
            return "track_changes_not_enabled — call enable_track_changes first"
        case .unknownError(let message):
            return "Unknown error: \(message)"
        }
    }
}
