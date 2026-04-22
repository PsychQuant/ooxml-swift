// MARK: - MathComponent protocol (OMML AST)

/// A node in an Office Math Markup Language (OMML) AST. Concrete types emit
/// ECMA-376 Part 1 §22.1 XML fragments via `toOMML()`.
///
/// Composition model: `[MathComponent]` arrays represent a sequence of math
/// content (e.g. the numerator of a fraction). Nested structures (fraction
/// inside a radical inside a summation base) compose naturally.
public protocol MathComponent {
    func toOMML() -> String
}

/// XML escape for math text (narrower than the Field helper — only the three
/// metacharacters that MUST NOT appear raw inside `<m:t>`).
private func escapeMathXML(_ string: String) -> String {
    return string
        .replacingOccurrences(of: "&", with: "&amp;")
        .replacingOccurrences(of: "<", with: "&lt;")
        .replacingOccurrences(of: ">", with: "&gt;")
}

private func combine(_ components: [MathComponent]) -> String {
    return components.map { $0.toOMML() }.joined()
}

// MARK: - MathStyle

/// `<m:sty>` style for math runs. ECMA-376 §22.1.2.93 `ST_Style`.
public enum MathStyle: String {
    case plain = "p"
    case bold = "b"
    case italic = "i"
    case boldItalic = "bi"
}

// MARK: - MathRun

/// A single run of math text.
public struct MathRun: MathComponent {
    public var text: String
    public var style: MathStyle?

    public init(text: String, style: MathStyle? = nil) {
        self.text = text
        self.style = style
    }

    public func toOMML() -> String {
        let escaped = escapeMathXML(text)
        if let style = style {
            return "<m:r><m:rPr><m:sty m:val=\"\(style.rawValue)\"/></m:rPr><m:t>\(escaped)</m:t></m:r>"
        }
        return "<m:r><m:t>\(escaped)</m:t></m:r>"
    }
}

// MARK: - MathFraction

/// Fraction with numerator / denominator.
public struct MathFraction: MathComponent {
    public var numerator: [MathComponent]
    public var denominator: [MathComponent]
    public var barStyle: FractionBar

    /// Fraction bar style. ECMA-376 §22.1.2.28.
    public enum FractionBar: String {
        case horizontal = "bar"    // default, standard horizontal bar
        case skewed = "skw"        // diagonal bar (a/b)
        case linear = "lin"        // inline (a/b written flat)
        case noBar = "noBar"       // stacked without bar (binomial-like)
    }

    public init(
        numerator: [MathComponent],
        denominator: [MathComponent],
        barStyle: FractionBar = .horizontal
    ) {
        self.numerator = numerator
        self.denominator = denominator
        self.barStyle = barStyle
    }

    public func toOMML() -> String {
        var xml = "<m:f>"
        if barStyle != .horizontal {
            xml += "<m:fPr><m:type m:val=\"\(barStyle.rawValue)\"/></m:fPr>"
        }
        xml += "<m:num>\(combine(numerator))</m:num>"
        xml += "<m:den>\(combine(denominator))</m:den>"
        xml += "</m:f>"
        return xml
    }
}

// MARK: - MathSubSuperScript

/// Subscript, superscript, or both on a base. Emits `<m:sSub>`, `<m:sSup>`,
/// or `<m:sSubSup>` depending on which slots are populated.
public struct MathSubSuperScript: MathComponent {
    public var base: [MathComponent]
    public var sub: [MathComponent]?
    public var sup: [MathComponent]?

    public init(base: [MathComponent], sub: [MathComponent]? = nil, sup: [MathComponent]? = nil) {
        self.base = base
        self.sub = sub
        self.sup = sup
    }

    public func toOMML() -> String {
        let baseXML = combine(base)
        switch (sub, sup) {
        case let (.some(s), .some(p)):
            return "<m:sSubSup><m:e>\(baseXML)</m:e><m:sub>\(combine(s))</m:sub><m:sup>\(combine(p))</m:sup></m:sSubSup>"
        case let (.some(s), nil):
            return "<m:sSub><m:e>\(baseXML)</m:e><m:sub>\(combine(s))</m:sub></m:sSub>"
        case let (nil, .some(p)):
            return "<m:sSup><m:e>\(baseXML)</m:e><m:sup>\(combine(p))</m:sup></m:sSup>"
        case (nil, nil):
            return baseXML
        }
    }
}

// MARK: - MathAccent

/// Accent decorator over a math base: `\hat{x}`, `\bar{x}`, `\tilde{x}`, etc.
/// Emits `<m:acc>` per ECMA-376 §22.1.2.1.
///
/// `accentChar` MUST be a Unicode combining diacritic (e.g. `"\u{0302}"` for
/// hat / circumflex, `"\u{0304}"` for macron / bar, `"\u{0303}"` for tilde,
/// `"\u{0307}"` for dot above). Spacing variants like `"^"` will render but
/// won't compose visually with the base — pick the combining form.
public struct MathAccent: MathComponent {
    public var base: [MathComponent]
    public var accentChar: String

    public init(base: [MathComponent], accentChar: String) {
        self.base = base
        self.accentChar = accentChar
    }

    public func toOMML() -> String {
        let chrEscaped = escapeMathXML(accentChar)
        return "<m:acc><m:accPr><m:chr m:val=\"\(chrEscaped)\"/></m:accPr><m:e>\(combine(base))</m:e></m:acc>"
    }
}

// MARK: - MathRadical

/// Square root or n-th root. `degree == nil` produces `<m:degHide>` for
/// a bare square root.
public struct MathRadical: MathComponent {
    public var radicand: [MathComponent]
    public var degree: [MathComponent]?

    public init(radicand: [MathComponent], degree: [MathComponent]? = nil) {
        self.radicand = radicand
        self.degree = degree
    }

    public func toOMML() -> String {
        var xml = "<m:rad>"
        if degree == nil {
            xml += "<m:radPr><m:degHide m:val=\"1\"/></m:radPr>"
        }
        xml += "<m:deg>\(degree.map(combine) ?? "")</m:deg>"
        xml += "<m:e>\(combine(radicand))</m:e>"
        xml += "</m:rad>"
        return xml
    }
}

// MARK: - MathNary (∑, ∫, ∏)

public struct MathNary: MathComponent {
    public enum NaryOperator: String {
        case sum = "∑"
        case integral = "∫"
        case product = "∏"
        case doubleIntegral = "∬"
        case contourIntegral = "∮"
        case union = "⋃"
        case intersection = "⋂"
    }

    public var op: NaryOperator
    public var sub: [MathComponent]?        // lower bound (under operator)
    public var sup: [MathComponent]?        // upper bound (above operator)
    public var base: [MathComponent]        // integrand / summand

    public init(
        op: NaryOperator,
        sub: [MathComponent]? = nil,
        sup: [MathComponent]? = nil,
        base: [MathComponent]
    ) {
        self.op = op
        self.sub = sub
        self.sup = sup
        self.base = base
    }

    public func toOMML() -> String {
        var xml = "<m:nary>"
        xml += "<m:naryPr><m:chr m:val=\"\(op.rawValue)\"/></m:naryPr>"
        xml += "<m:sub>\(sub.map(combine) ?? "")</m:sub>"
        xml += "<m:sup>\(sup.map(combine) ?? "")</m:sup>"
        xml += "<m:e>\(combine(base))</m:e>"
        xml += "</m:nary>"
        return xml
    }
}

// MARK: - MathDelimiter

/// Delimited expression: `(a, b)`, `{a|b}`, `[x]`, `|x|`.
public struct MathDelimiter: MathComponent {
    public var open: String
    public var close: String
    public var elements: [[MathComponent]]   // one or more elements between delimiters
    public var separator: String             // e.g. "," or "|"; empty means single-element

    public init(
        open: String,
        close: String,
        elements: [[MathComponent]],
        separator: String = ""
    ) {
        self.open = open
        self.close = close
        self.elements = elements
        self.separator = separator
    }

    public func toOMML() -> String {
        var xml = "<m:d>"
        xml += "<m:dPr>"
        xml += "<m:begChr m:val=\"\(open)\"/>"
        xml += "<m:endChr m:val=\"\(close)\"/>"
        if !separator.isEmpty {
            xml += "<m:sepChr m:val=\"\(separator)\"/>"
        }
        xml += "</m:dPr>"
        for element in elements {
            xml += "<m:e>\(combine(element))</m:e>"
        }
        xml += "</m:d>"
        return xml
    }
}

// MARK: - MathFunction

/// Function application like `sin(x)`, `log(y)`. The function name is itself
/// math content so custom identifiers work.
public struct MathFunction: MathComponent {
    public var functionName: [MathComponent]
    public var argument: [MathComponent]

    public init(functionName: [MathComponent], argument: [MathComponent]) {
        self.functionName = functionName
        self.argument = argument
    }

    public func toOMML() -> String {
        var xml = "<m:func>"
        xml += "<m:fName>\(combine(functionName))</m:fName>"
        xml += "<m:e>\(combine(argument))</m:e>"
        xml += "</m:func>"
        return xml
    }
}

// MARK: - MathLimit

/// A base with a limit underneath or above (used for `lim`, `max`, etc.).
public struct MathLimit: MathComponent {
    public enum Position {
        case lower       // `<m:limLow>` — limit below base
        case upper       // `<m:limUpp>` — limit above base
    }

    public var position: Position
    public var base: [MathComponent]
    public var limit: [MathComponent]

    public init(position: Position, base: [MathComponent], limit: [MathComponent]) {
        self.position = position
        self.base = base
        self.limit = limit
    }

    public func toOMML() -> String {
        let tag = position == .upper ? "limUpp" : "limLow"
        return "<m:\(tag)><m:e>\(combine(base))</m:e><m:lim>\(combine(limit))</m:lim></m:\(tag)>"
    }
}

// MARK: - UnknownMath (v0.10.0 opaque fallback)

/// Preserves an `<m:...>` subtree that `OMMLParser` doesn't recognize so
/// round-trip `write → read → write` never loses data.
///
/// Added in ooxml-swift 0.10.0. Callers iterating `[MathComponent]` arrays
/// may encounter this struct — handle it by pattern-matching or `as?` cast.
public struct UnknownMath: MathComponent {
    public let rawXML: String

    public init(rawXML: String) {
        self.rawXML = rawXML
    }

    public func toOMML() -> String {
        return rawXML
    }
}

// MARK: - MathMatrix

/// A matrix. `rows[r][c]` is the content of the cell at row r, column c.
public struct MathMatrix: MathComponent {
    public var rows: [[[MathComponent]]]

    public init(rows: [[[MathComponent]]]) {
        self.rows = rows
    }

    public func toOMML() -> String {
        var xml = "<m:m>"
        for row in rows {
            xml += "<m:mr>"
            for cell in row {
                xml += "<m:e>\(combine(cell))</m:e>"
            }
            xml += "</m:mr>"
        }
        xml += "</m:m>"
        return xml
    }
}
