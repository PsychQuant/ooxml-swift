// WordDocument+Authoring.swift
// word-aligned-state-sync Phase 4 task 5.3 — public authoring entry points
// consumed by the WordDSLSwift module ("Build a docx end-to-end from a
// Swift script": no prior docx as input).
//
// WordDSLSwift lives in a separate module, so it cannot reach the internal
// `appendAndMaterialize` pipeline or `xmlTrees` setter directly. These
// wrappers expose exactly the authoring surface the DSL runtime needs:
// an empty tree-backed document, op application through the canonical
// log+reducer path, and a minimal-package serializer that writes the tree
// (preserving `w14:paraId` — the typed scratch writer would drop it).

import Foundation

extension WordDocument {

    /// An empty document whose `word/document.xml` tree exists and is ready
    /// for authoring ops (`appendParagraph(in: nil)` targets its `<w:body>`).
    public static func emptyAuthoringDocument() -> WordDocument {
        var doc = WordDocument()
        let body = XmlNode.element(
            prefix: "w", localName: "body",
            namespaceURI: "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        let root = XmlNode.element(
            prefix: "w", localName: "document",
            namespaceURI: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            attributes: [
                XmlAttribute(prefix: "xmlns", localName: "w",
                             value: "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
                XmlAttribute(prefix: "xmlns", localName: "w14",
                             value: "http://schemas.microsoft.com/office/word/2010/wordml"),
            ],
            children: [body])
        doc.xmlTrees["word/document.xml"] = XmlTree.synthesized(root: root)
        return doc
    }

    /// Applies authoring operations through the canonical log + reducer
    /// pipeline (same path as `apply(_ edit:)` and the typed setters).
    /// Op IDs are generated here; callers relying on round-trip semantics
    /// may regenerate IDs per the `ooxml-script-transcode` contract.
    public mutating func apply(operations: [Operation], source: OpSource = .swift) throws {
        try appendAndMaterialize(operations, source: source)
        resyncBodyFromDocumentTree()
    }

    /// Serializes a minimal, Word-valid package from the current trees:
    /// `[Content_Types].xml`, `_rels/.rels`, `word/_rels/document.xml.rels`,
    /// `word/document.xml` (tree-serialized — `w14:paraId` preserved), and
    /// `word/styles.xml` when a styles tree exists. Atomic single-file write
    /// (temp + rename).
    public func writeAuthoringPackage(to url: URL) throws {
        let tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("mdocx-authoring-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: tempDir) }

        func write(_ content: Data, _ relativePath: String) throws {
            let fileURL = tempDir.appendingPathComponent(relativePath)
            try FileManager.default.createDirectory(
                at: fileURL.deletingLastPathComponent(), withIntermediateDirectories: true)
            try content.write(to: fileURL)
        }

        guard let docTree = xmlTrees["word/document.xml"] else {
            throw WordError.parseError("authoring package requires a word/document.xml tree")
        }
        try write(try XmlTreeWriter.serialize(docTree), "word/document.xml")

        let hasStyles = xmlTrees["word/styles.xml"] != nil
        if let stylesTree = xmlTrees["word/styles.xml"] {
            try write(try XmlTreeWriter.serialize(stylesTree), "word/styles.xml")
        }

        let stylesOverride = hasStyles
            ? "\n    <Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>"
            : ""
        try write(Data("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="xml" ContentType="application/xml"/>
                <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\(stylesOverride)
            </Types>
            """.utf8), "[Content_Types].xml")
        try write(Data("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """.utf8), "_rels/.rels")
        let stylesRel = hasStyles
            ? "\n    <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
            : ""
        try write(Data("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\(stylesRel)
            </Relationships>
            """.utf8), "word/_rels/document.xml.rels")

        let data = try ZipHelper.zipToData(tempDir)
        let tempFile = url.appendingPathExtension("tmp.\(UUID().uuidString)")
        defer { try? FileManager.default.removeItem(at: tempFile) }
        try data.write(to: tempFile)
        _ = try FileManager.default.replaceItemAt(url, withItemAt: tempFile,
                                                  backupItemName: nil, options: [])
    }
}
