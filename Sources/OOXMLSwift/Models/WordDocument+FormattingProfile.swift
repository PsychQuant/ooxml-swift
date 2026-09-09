import Foundation

/// Durable value-owned XML, separate from transient tree freshness/carry flags.
internal struct DocumentFormattingState {
    var defaultsXML: String?
    var originalStylesXML: String
    var baselineStyles: [Style]
    var themeData: Data?
    var fontsData: Data?
    var explicitlyApplied = true
}

extension WordDocument {
    // 標楷體的 Office family identifier；localized 名稱在 Mac Word 會被替代字型接管。
    private static let officialEastAsianFont = "DFKai-SB"

    /// The theme currently selected by explicit part operations, live trees,
    /// formatting profiles, or the source archive, in writer precedence order.
    /// This does not import or sanitize caller-owned theme XML.
    public func effectiveThemeData() throws -> Data? {
        try effectiveFormattingPartData("word/theme/theme1.xml", fallback: formattingState?.themeData)
    }

    private func effectiveFormattingPartData(_ path: String, fallback: Data?) throws -> Data? {
        if let carried = carriedParts[path] { return carried }
        if treeFreshParts.contains(path), let tree = xmlTrees[path] { return try XmlTreeWriter.serialize(tree) }
        func archived() throws -> Data? {
            guard let archive = archiveTempDir else { return nil }
            let url = archive.appendingPathComponent(path)
            guard FileManager.default.fileExists(atPath: url.path) else { return nil }
            return try Data(contentsOf: url)
        }
        // Legacy archive edits remain valid for unprofiled documents. For an
        // applied profile, the old archive is not evidence of a newer edit.
        if formattingState?.explicitlyApplied != true, modifiedParts.contains(path), let data = try archived() { return data }
        if let fallback { return fallback }
        return try archived()
    }
    /// Capture completed authoritative operations before committing the op
    /// transaction. Later typed edits use this baseline instead of the import.
    internal func refreshedFormattingState(trees: [String: XmlTree], carried: [String: Data],
                                           freshParts: Set<String>, carriedPaths: Set<String>) throws
        -> (state: DocumentFormattingState, styles: [Style]?)? {
        guard var state = formattingState else { return nil }
        func bytes(_ path: String) throws -> Data? {
            if state.explicitlyApplied, carriedPaths.contains(path), let data = carried[path] { return data }
            if freshParts.contains(path), let tree = trees[path] { return try XmlTreeWriter.serialize(tree) }
            return nil
        }
        var updatedStyles: [Style]?
        if let data = try bytes("word/styles.xml") {
            let root = ProfileXML.canonicalWordTree(try ProfileXML.parse(data))
            let xml = try ProfileXML.string(root)
            let parsed = try DocxReader.parseStyles(from: XMLDocument(xmlString: xml, options: []))
            state.defaultsXML = try ProfileXML.child(root, "docDefaults").map(ProfileXML.string)
            state.originalStylesXML = xml
            state.baselineStyles = parsed
            updatedStyles = parsed
        }
        if let data = try bytes("word/theme/theme1.xml") { state.themeData = data }
        if let data = try bytes("word/fontTable.xml") { state.fontsData = data }
        return (state, updatedStyles)
    }

    /// Applies an explicit profile atomically to this value. No app settings
    /// or Normal template are read here. Existing-document inherit is a no-op.
    public mutating func applyFormattingProfile(_ profile: DocumentFormattingProfile,
                                                context: DocumentFormattingContext) throws {
        try profile.validate()
        if profile.kind == .inherit, context == .existingDocument { return }
        var next = self
        next.xmlTrees = xmlTrees.mapValues { $0.deepCopy() }
        next.operationReplayBase = nil
        if profile.kind == .inherit {
            for index in next.styles.indices {
                guard next.styles[index].runProperties?.fontOrigin.generated == true else { continue }
                next.styles[index].runProperties?.fontName = nil
            }
            let styles = try ProfileXML.parse(next.styles.toStylesXML())
            let defaults = ProfileXML.child(styles, "docDefaults")!
            for run in ProfileXML.walk(defaults) where run.localName == "rPr" {
                run.children.removeAll { $0.localName == "rFonts" }
            }
            if var owned = next.formattingState {
                // Read-back XML has no generator provenance. Retain its
                // defaults and raw styling, merging only known typed edits.
                owned.explicitlyApplied = true
                next.formattingState = owned
            } else {
                next.formattingState = DocumentFormattingState(
                    defaultsXML: try ProfileXML.string(defaults.withWordNamespace()),
                    originalStylesXML: try ProfileXML.string(styles), baselineStyles: next.styles,
                    fontsData: Data("<w:fonts xmlns:w=\"\(ProfileXML.w)\"/>".utf8))
            }
            next.markTypedDirty("word/styles.xml")
            next.markTypedDirty("word/fontTable.xml")
            self = next
            return
        }

        let imported = try ProfileXML.checked(profile.stylesXML!, root: "styles")
        let defaultStyle = imported.children.first { $0.localName == "style" && ProfileXML.value($0, "type") == "paragraph" && ["1", "true", "on"].contains(ProfileXML.value($0, "default") ?? "") }!
        let sourceDefault = ProfileXML.value(defaultStyle, "styleId")!
        let rawTarget: XmlNode
        if let carried = next.carriedParts["word/styles.xml"] { rawTarget = try ProfileXML.parse(carried) }
        else if let tree = next.xmlTrees["word/styles.xml"], !next.modifiedParts.contains("word/styles.xml") || next.treeFreshParts.contains("word/styles.xml") { rawTarget = tree.root }
        else { rawTarget = try ProfileXML.parse(next.styles.toStylesXML()) }
        let targetXML = ProfileXML.canonicalWordTree(rawTarget)
        let targetDefault = targetXML.children.first { $0.namespaceURI == ProfileXML.w && $0.localName == "style" && ProfileXML.value($0, "type") == "paragraph" && ["1", "true", "on"].contains(ProfileXML.value($0, "default") ?? "") }
            .flatMap { ProfileXML.value($0, "styleId") } ?? sourceDefault
        // A localized default ID maps to the target's default, preserving body
        // references and every retained style's existing basedOn chain.
        if sourceDefault != targetDefault {
            guard !imported.children.contains(where: { $0.localName == "style" && ProfileXML.value($0, "styleId") == targetDefault }) else {
                throw DocumentFormattingProfileError.invalidSnapshot("default style mapping collision")
            }
            for node in ProfileXML.walk(imported) {
                if node === defaultStyle { node.setWordAttribute("styleId", targetDefault) }
                if ["basedOn", "next", "link"].contains(node.localName), ProfileXML.value(node, "val") == sourceDefault {
                    node.setWordAttribute("val", targetDefault)
                }
            }
        }
        defaultStyle.setWordAttribute("default", "1")
        let importedIDs = Set(imported.children.compactMap { ProfileXML.value($0, "styleId") })
        imported.copyMissingNamespaces(from: targetXML)
        for style in targetXML.children where style.kind == .element && style.namespaceURI == ProfileXML.w && style.localName == "style" {
            guard let id = ProfileXML.value(style, "styleId"), !importedIDs.contains(id) else { continue }
            let retained = style.deepClone()
            if ProfileXML.value(retained, "type") == "paragraph" { retained.attributes.removeAll { $0.localName == "default" } }
            imported.children.append(retained)
        }
        for run in ProfileXML.walk(imported) where run.localName == "rPr" && run.namespaceURI == ProfileXML.w {
            if ProfileXML.child(run, "rFonts") == nil {
                run.children.insert(.element(prefix: "w", localName: "rFonts", namespaceURI: ProfileXML.w), at: 0)
            }
            let fonts = ProfileXML.child(run, "rFonts")!
            fonts.attributes.removeAll { $0.localName == "eastAsiaTheme" }
            fonts.setWordAttribute("eastAsia", Self.officialEastAsianFont)
        }
        try ProfileXML.validateStyles(imported)
        let stylesXML = try ProfileXML.string(imported)
        next.styles = try DocxReader.parseStyles(from: XMLDocument(xmlString: stylesXML, options: []))
        let defaults = ProfileXML.child(imported, "docDefaults")!

        var themeXML: String?
        if let source = profile.themeXML {
            let theme = try ProfileXML.checked(source, root: "theme")
            for font in ProfileXML.walk(theme) where font.namespaceURI == ProfileXML.a {
                if font.localName == "ea" || (font.localName == "font" && font.attributeValue(prefix: nil, localName: "script") == "Hant") {
                    font.attributes.removeAll { $0.localName == "typeface" }
                    font.attributes.append(XmlAttribute(localName: "typeface", value: Self.officialEastAsianFont))
                }
            }
            themeXML = try ProfileXML.string(theme)
        }
        let fonts = try profile.fontsXML.map { try ProfileXML.checked($0, root: "fonts") }
            ?? ProfileXML.parse("<w:fonts xmlns:w=\"\(ProfileXML.w)\"/>")
        if !fonts.children.contains(where: { ProfileXML.value($0, "name") == Self.officialEastAsianFont }) {
            fonts.children.append(.element(prefix: "w", localName: "font", namespaceURI: ProfileXML.w,
                                           attributes: [XmlAttribute(prefix: "w", localName: "name", value: Self.officialEastAsianFont)]))
        }
        next.formattingState = DocumentFormattingState(defaultsXML: try ProfileXML.string(defaults.withWordNamespace()),
            originalStylesXML: stylesXML, baselineStyles: next.styles, themeData: themeXML.map { Data($0.utf8) }, fontsData: try Data(ProfileXML.string(fonts).utf8))
        next.markTypedDirty("word/styles.xml")
        next.xmlTrees["word/styles.xml"] = try XmlTreeReader.parse(Data(stylesXML.utf8))
        next.treeFreshParts.insert("word/styles.xml")

        if let carried = next.carriedParts["word/document.xml"] {
            next.xmlTrees["word/document.xml"] = try XmlTreeReader.parse(carried)
            next.treeFreshParts.insert("word/document.xml")
        }
        let section = try ProfileXML.checked(profile.sectionXML!, root: "sectPr")
        let parsedSection = DocxReader.parseSectionProperties(try XMLElement(xmlString: ProfileXML.string(section)))
        next.sectionProperties.pageSize = parsedSection.pageSize
        next.sectionProperties.pageMargins = parsedSection.pageMargins
        next.sectionProperties.orientation = parsedSection.orientation
        if ProfileXML.child(section, "cols") != nil {
            next.sectionProperties.columns = parsedSection.columns
            next.sectionProperties.columnSpacing = parsedSection.columnSpacing
        }
        if ProfileXML.child(section, "docGrid") != nil { next.sectionProperties.docGrid = parsedSection.docGrid }
        next.sectionProperties.xmlNode = nil
        // Materialize a typed-only document before patching the final body
        // section. Existing trees preserve non-final sections and references.
        if next.xmlTrees["word/document.xml"] == nil {
            let trees = try DocxWriter.materializeTypedTrees(next, parts: ["word/document.xml"])
            next.xmlTrees["word/document.xml"] = trees["word/document.xml"]
        } else if next.modifiedParts.contains("word/document.xml"), !next.treeFreshParts.contains("word/document.xml") {
            let trees = try DocxWriter.materializeTypedTrees(next, parts: ["word/document.xml"])
            next.xmlTrees["word/document.xml"] = trees["word/document.xml"]
        }
        guard let tree = next.xmlTrees["word/document.xml"], let body = ProfileXML.child(tree.root, "body") else {
            throw DocumentFormattingProfileError.invalidSnapshot("target document body")
        }
        let final = body.children.last(where: { $0.namespaceURI == ProfileXML.w && $0.localName == "sectPr" })
            ?? XmlNode.element(prefix: "w", localName: "sectPr", namespaceURI: ProfileXML.w)
        if !body.children.contains(where: { $0 === final }) { body.children.append(final) }
        for safe in section.children where safe.kind == .element {
            final.children.removeAll { $0.namespaceURI == ProfileXML.w && $0.localName == safe.localName }
            final.children.append(safe.deepClone())
        }
        let sectionOrder = ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"]
        final.children = final.children.enumerated().sorted {
            let left = sectionOrder.firstIndex(of: $0.element.localName) ?? sectionOrder.count
            let right = sectionOrder.firstIndex(of: $1.element.localName) ?? sectionOrder.count
            return left == right ? $0.offset < $1.offset : left < right
        }.map(\.element)
        // The tree's nodes do not maintain parent pointers. Dirty the spine
        // explicitly so a clean ancestor cannot blob-copy the old section.
        body.markDirty()
        tree.root.markDirty()
        let sectionForParsing = final.deepClone()
        sectionForParsing.attributes.append(XmlAttribute(prefix: "xmlns", localName: "w", value: ProfileXML.w))
        sectionForParsing.attributes.append(XmlAttribute(prefix: "xmlns", localName: "r", value: "http://schemas.openxmlformats.org/officeDocument/2006/relationships"))
        next.sectionProperties = DocxReader.parseSectionProperties(try XMLElement(xmlString: ProfileXML.string(sectionForParsing)))
        next.markTypedDirty("word/document.xml")
        next.treeFreshParts.insert("word/document.xml")
        // Metadata remains authoritative until the finalizer merges the
        // formatting registrations. Invalidating it here loses raw replay's
        // only source and needlessly invokes the ordinary typed rels writer.
        next.markTypedDirty("word/fontTable.xml")
        if let themeXML {
            next.markTypedDirty("word/theme/theme1.xml")
            next.xmlTrees["word/theme/theme1.xml"] = try XmlTreeReader.parse(Data(themeXML.utf8))
            next.treeFreshParts.insert("word/theme/theme1.xml")
        }
        self = next
    }

    /// Shared writer finalization. Rebuild defaults from durable state while
    /// allowing later typed or reducer style edits to take precedence.
    internal func writeFormattingParts(to directory: URL) throws {
        let state = formattingState
        // Generic carried styles retain their pre-profile writer semantics.
        // An independently carried theme still needs publication and metadata.
        let preservesCarriedStyles = state?.explicitlyApplied == false && carriedParts["word/styles.xml"] != nil
        let writesStyles = state.map { $0.explicitlyApplied || modifiedParts.contains("word/styles.xml") } == true && !preservesCarriedStyles
        let freshAncillary = treeFreshParts.intersection(modifiedParts).intersection(["word/theme/theme1.xml", "word/fontTable.xml"])
        // Bare carry operations also power byte-equal replay. Only explicitly
        // dirty carried parts need publication through the ordinary writer;
        // untouched replay metadata must remain byte-preserved.
        let carriedAncillary = Set(carriedParts.keys).intersection(modifiedParts).intersection(["word/theme/theme1.xml", "word/fontTable.xml"])
        guard writesStyles || !freshAncillary.isEmpty || !carriedAncillary.isEmpty else { return }
        func write(_ bytes: Data, _ path: String) throws {
            let url = directory.appendingPathComponent(path)
            try FileManager.default.createDirectory(at: url.deletingLastPathComponent(), withIntermediateDirectories: true)
            try bytes.write(to: url)
        }
        var parts: [(path: String, type: String, rel: String, target: String)] = []
        if writesStyles, let state {
            let root: XmlNode
            if state.explicitlyApplied, let carried = carriedParts["word/styles.xml"] {
                root = try ProfileXML.parse(carried)
            } else if treeFreshParts.contains("word/styles.xml"), let tree = xmlTrees["word/styles.xml"] {
                root = tree.root.deepClone()
            } else {
                root = try ProfileXML.parse(DocxWriter.stylesXML(styles, latentStyles: latentStyles))
                let original = try ProfileXML.parse(state.originalStylesXML)
                let baseline = try ProfileXML.parse(state.baselineStyles.toStylesXML())
                root.copyMissingNamespaces(from: original)
                for index in root.children.indices {
                    let node = root.children[index]
                    guard let id = ProfileXML.value(node, "styleId"),
                          let oldTyped = baseline.children.first(where: { ProfileXML.value($0, "styleId") == id }),
                          let preserved = original.children.first(where: { ProfileXML.value($0, "styleId") == id }) else { continue }
                    root.children[index] = try Self.mergeFormattingChanges(original: preserved, baseline: oldTyped, current: node)
                }
                root.children.removeAll { $0.namespaceURI == ProfileXML.w && $0.localName == "docDefaults" }
                if let defaults = state.defaultsXML { root.children.insert(try ProfileXML.parse(defaults), at: 0) }
            }
            try write(Data(ProfileXML.string(root).utf8), "word/styles.xml")
            parts.append(("word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", "styles", "styles.xml"))
        }
        func currentAncillary(_ path: String, fallback: Data?) throws -> Data? {
            guard writesStyles || freshAncillary.contains(path) || carriedAncillary.contains(path) else { return nil }
            return try effectiveFormattingPartData(path, fallback: fallback)
        }
        if let fonts = try currentAncillary("word/fontTable.xml", fallback: state?.fontsData) {
            try write(fonts, "word/fontTable.xml")
            parts.append(("word/fontTable.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml", "fontTable", "fontTable.xml"))
        }
        if let theme = try currentAncillary("word/theme/theme1.xml", fallback: state?.themeData) {
            try write(theme, "word/theme/theme1.xml")
            parts.append(("word/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml", "theme", "theme/theme1.xml"))
        }
        let typesURL = directory.appendingPathComponent("[Content_Types].xml")
        let relsURL = directory.appendingPathComponent("word/_rels/document.xml.rels")
        let types = try ProfileXML.parse(Data(contentsOf: typesURL))
        let rels = try ProfileXML.parse(Data(contentsOf: relsURL))
        let relNS = "http://schemas.openxmlformats.org/package/2006/relationships"
        let typeNS = "http://schemas.openxmlformats.org/package/2006/content-types"
        var ids = Set(rels.children.compactMap { $0.attributeValue(prefix: nil, localName: "Id") })
        var typesChanged = false, relsChanged = false
        func setAttribute(_ node: XmlNode, _ name: String, _ value: String) {
            if let index = node.attributes.firstIndex(where: { $0.prefix == nil && $0.localName == name }) {
                node.attributes[index] = XmlAttribute(localName: name, value: value)
            } else { node.attributes.append(XmlAttribute(localName: name, value: value)) }
        }
        for part in parts {
            if let existing = types.children.first(where: { $0.namespaceURI == typeNS && $0.localName == "Override" && $0.attributeValue(prefix: nil, localName: "PartName") == "/" + part.path }) {
                if existing.attributeValue(prefix: nil, localName: "ContentType") != part.type {
                    setAttribute(existing, "ContentType", part.type)
                    typesChanged = true
                }
            } else {
                types.children.append(.element(prefix: types.prefix, localName: "Override", namespaceURI: typeNS, attributes: [
                    XmlAttribute(localName: "PartName", value: "/" + part.path), XmlAttribute(localName: "ContentType", value: part.type)]))
                typesChanged = true
            }
            let relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + part.rel
            let previous = rels.children.first { $0.namespaceURI == relNS && $0.attributeValue(prefix: nil, localName: "Type") == relationshipType }
            if let previous {
                let target = previous.attributeValue(prefix: nil, localName: "Target") ?? ""
                let resolvedTarget = NSString(string: target.hasPrefix("/") ? target : "/word/" + target).standardizingPath
                if resolvedTarget != "/" + part.path {
                    setAttribute(previous, "Target", part.target)
                    relsChanged = true
                }
                if previous.attributeValue(prefix: nil, localName: "TargetMode") == "External" {
                    previous.attributes.removeAll { $0.prefix == nil && $0.localName == "TargetMode" }
                    relsChanged = true
                }
            } else {
                var n = 1
                while ids.contains("rId\(n)") { n += 1 }
                let id = "rId\(n)"
                ids.insert(id)
                rels.children.append(.element(prefix: rels.prefix, localName: "Relationship", namespaceURI: relNS, attributes: [
                    XmlAttribute(localName: "Id", value: id), XmlAttribute(localName: "Type", value: relationshipType), XmlAttribute(localName: "Target", value: part.target)]))
                relsChanged = true
            }
        }
        if typesChanged { try write(Data(ProfileXML.string(types).utf8), "[Content_Types].xml") }
        if relsChanged { try write(Data(ProfileXML.string(rels).utf8), "word/_rels/document.xml.rels") }
    }

    /// Apply only typed changes to the preserved style. A rename must not
    /// erase table properties, ligatures or theme bindings absent from Style.
    private static func mergeFormattingChanges(original: XmlNode, baseline: XmlNode, current: XmlNode) throws -> XmlNode {
        let merged = original.deepClone()
        let names = Set(baseline.attributes.map(\.qualifiedName) + current.attributes.map(\.qualifiedName))
        for name in names {
            let old = baseline.attributes.first { $0.qualifiedName == name }
            let new = current.attributes.first { $0.qualifiedName == name }
            if old != new {
                merged.attributes.removeAll { $0.qualifiedName == name }
                if let new { merged.attributes.append(new) }
            }
        }
        let children = Set(baseline.children.filter { $0.kind == .element }.map(\.localName)
            + current.children.filter { $0.kind == .element }.map(\.localName))
        for name in children.sorted() {
            let old = baseline.children.filter { $0.kind == .element && $0.localName == name }
            let new = current.children.filter { $0.kind == .element && $0.localName == name }
            guard try old.map(ProfileXML.string) != new.map(ProfileXML.string) else { continue }
            if ["rPr", "pPr"].contains(name), old.count == 1, new.count == 1,
               let index = merged.children.firstIndex(where: { $0.localName == name }) {
                merged.children[index] = try mergeFormattingChanges(original: merged.children[index], baseline: old[0], current: new[0])
            } else {
                let index = merged.children.firstIndex { $0.localName == name } ?? merged.children.count
                merged.children.removeAll { $0.kind == .element && $0.localName == name }
                merged.children.insert(contentsOf: new.map { $0.deepClone() }, at: min(index, merged.children.count))
            }
        }
        return merged
    }
}

private extension XmlNode {
    func copyMissingNamespaces(from other: XmlNode) {
        for attr in other.attributes where attr.isNamespaceDeclaration {
            if !attributes.contains(where: { $0.qualifiedName == attr.qualifiedName }) { attributes.append(attr) }
        }
    }
    func setWordAttribute(_ name: String, _ value: String) {
        attributes.removeAll { $0.prefix == "w" && $0.localName == name }
        attributes.append(XmlAttribute(prefix: "w", localName: name, value: value))
    }
    func withWordNamespace() -> XmlNode {
        let copy = deepClone()
        copy.attributes.removeAll { $0.isNamespaceDeclaration }
        copy.attributes.insert(XmlAttribute(prefix: "xmlns", localName: "w", value: ProfileXML.w), at: 0)
        return copy
    }
}
