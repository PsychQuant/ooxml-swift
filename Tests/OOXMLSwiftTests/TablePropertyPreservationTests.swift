import Foundation
import XCTest
@testable import OOXMLSwift

final class TablePropertyPreservationTests: XCTestCase {
    func testCellMutationPreservesUnmodeledTableProperties() throws {
        let directory = FileManager.default.temporaryDirectory
            .appendingPathComponent("table-property-preservation-\(UUID().uuidString)")
        try FileManager.default.createDirectory(
            at: directory, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: directory) }

        let seedURL = directory.appendingPathComponent("seed.docx")
        var seed = WordDocument()
        var tableProperties = TableProperties()
        tableProperties.width = 0
        tableProperties.widthType = .auto
        seed.appendTable(Table(
            rows: [
                TableRow(cells: [TableCell(text: "A"), TableCell(text: "B")]),
                TableRow(cells: [TableCell(text: "C"), TableCell(text: "")]),
            ],
            properties: tableProperties))
        try DocxWriter.write(seed, to: seedURL)

        let extracted = try ZipHelper.unzip(seedURL)
        defer { ZipHelper.cleanup(extracted) }
        let documentURL = extracted.appendingPathComponent("word/document.xml")
        var xml = try String(contentsOf: documentURL, encoding: .utf8)
        let richProperties = """
        <w:tblStyle w:val="TableGrid"/>
        <w:jc w:val="center" w15:keep="alignment-extension"/>
        <mc:AlternateContent><mc:Choice Requires="w15"><w:tblCellSpacing w:w="20" w:type="dxa"/></mc:Choice><mc:Fallback><w:tblCellSpacing w:w="10" w:type="dxa"/></mc:Fallback></mc:AlternateContent>
        <w:tblInd w:w="720" w:type="dxa" w15:keep="indent-extension"/>
        <w:tblLayout w:type="fixed" w15:keep="layout-extension"/>
        <w:tblLook w:firstColumn="1" w:firstRow="1" w:lastColumn="0" w:lastRow="0" w:noHBand="0" w:noVBand="1" w:val="04A0"/>
        <w:tblBorders><w:top w:val="single" w:sz="12" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="12" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="12" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="12" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="6" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="6" w:space="0" w:color="auto"/></w:tblBorders>
        <w:tblCaption w:val="Government form"/>
        <w:tblDescription w:val="Untouched metadata"/>
        <w:tblStylePr w:type="firstRow" w15:keep="style-extension"><w:rPr><w:b/></w:rPr></w:tblStylePr>
        <w:tblStylePr w:type="futureType"><w:rPr><w:i/></w:rPr></w:tblStylePr>
        <v:future v:flag="1"/>
        <v:tblW v:flag="foreign-same-name"/>
        <w:tblPrChange w:id="7" w:author="Reviewer"><w:tblPr><w:tblW w:w="1234" w:type="dxa"/></w:tblPr></w:tblPrChange>
        """.replacingOccurrences(of: "\n", with: "")
        xml = xml.replacingOccurrences(
            of: "<w:tblPr>",
            with: "<w:tblPr xmlns:v=\"urn:table-property-test\" mc:Ignorable=\"v\">\(richProperties)")
        xml = xml.replacingOccurrences(
            of: "<w:tblW w:w=\"0\"",
            with: "<w:tblW w15:keep=\"width-extension\" w:w=\"0\"")
        try xml.write(to: documentURL, atomically: true, encoding: .utf8)

        let inputURL = directory.appendingPathComponent("input.docx")
        try ZipHelper.zip(extracted, to: inputURL)
        var document = try DocxReader.read(from: inputURL)
        defer { document.close() }
        try document.updateCell(tableIndex: 0, row: 1, col: 1, text: "D")

        let outputURL = directory.appendingPathComponent("output.docx")
        try DocxWriter.write(document, to: outputURL)
        let outputPackage = try ZipHelper.unzip(outputURL)
        defer { ZipHelper.cleanup(outputPackage) }
        let outputXML = try String(
            contentsOf: outputPackage.appendingPathComponent("word/document.xml"),
            encoding: .utf8)

        for token in [
            "<w:tblStyle", "w:val=\"TableGrid\"",
            "w15:keep=\"width-extension\"",
            "w15:keep=\"alignment-extension\"",
            "<mc:AlternateContent", "<w:tblCellSpacing w:w=\"20\"",
            "w15:keep=\"indent-extension\"",
            "w15:keep=\"layout-extension\"",
            "w15:keep=\"style-extension\"",
            "<w:tblLook", "w:val=\"04A0\"", "w:noVBand=\"1\"",
            "<w:tblBorders", "<w:insideV", "w:sz=\"6\"", "w:space=\"0\"",
            "<w:tblCaption", "w:val=\"Government form\"",
            "<w:tblDescription", "w:val=\"Untouched metadata\"",
            "xmlns:v=\"urn:table-property-test\"",
            "mc:Ignorable=\"v\"",
            "<v:future", "v:flag=\"1\"",
            "<v:tblW", "v:flag=\"foreign-same-name\"",
            "w:type=\"futureType\"",
            "<w:tblPrChange",
        ] {
            XCTAssertTrue(outputXML.contains(token), "missing table property: \(token)")
        }
        XCTAssertEqual(
            outputXML.components(separatedBy: "<w:tblLayout").count - 1,
            1,
            "source-loaded typed table layout must not be emitted twice")
        let outerIndent = try XCTUnwrap(outputXML.range(of: "<w:tblInd"))
        let outerAlignment = try XCTUnwrap(outputXML.range(of: "<w:jc"))
        let alternateContent = try XCTUnwrap(outputXML.range(of: "<mc:AlternateContent"))
        let outerBorders = try XCTUnwrap(outputXML.range(of: "<w:tblBorders"))
        let outerLayout = try XCTUnwrap(outputXML.range(of: "<w:tblLayout"))
        let outerLook = try XCTUnwrap(outputXML.range(of: "<w:tblLook"))
        let revision = try XCTUnwrap(outputXML.range(of: "<w:tblPrChange"))
        XCTAssertLessThan(outerAlignment.lowerBound, alternateContent.lowerBound)
        XCTAssertLessThan(alternateContent.lowerBound, outerIndent.lowerBound)
        XCTAssertLessThan(outerIndent.lowerBound, outerBorders.lowerBound)
        XCTAssertLessThan(outerBorders.lowerBound, outerLayout.lowerBound)
        XCTAssertLessThan(outerLayout.lowerBound, outerLook.lowerBound)
        XCTAssertLessThan(outerLook.lowerBound, revision.lowerBound)
        XCTAssertNoThrow(try XMLDocument(
            data: Data(outputXML.utf8), options: [.nodePreserveAll]))
        let tablePropertiesStart = try XCTUnwrap(outputXML.range(of: "<w:tblPr"))
        let tablePropertiesTagEnd = try XCTUnwrap(
            outputXML[tablePropertiesStart.lowerBound...].firstIndex(of: ">"))
        let tablePropertiesOpeningTag = String(
            outputXML[tablePropertiesStart.lowerBound...tablePropertiesTagEnd])
        XCTAssertFalse(
            tablePropertiesOpeningTag.contains("xmlns:w15"),
            "document-root namespaces must stay on the part root instead of being copied to every table")
        XCTAssertEqual(
            try directTablePropertyChildren(in: outputXML).sorted(),
            try directTablePropertyChildren(in: xml).sorted(),
            "an unrelated cell mutation must preserve the complete tblPr child multiset")
        XCTAssertTrue(outputXML.contains(">D</w:t>"))
    }

    func testTypedTableBorderReplacementDoesNotDuplicateRawBorder() {
        var properties = TableProperties()
        properties.rawChildren = [PreservedTableProperty(
            qualifiedName: "w:tblBorders",
            localName: "tblBorders",
            namespaceURI: wordprocessingMLNamespace,
            styleProjectionIndex: nil,
            representedStyleTypes: [],
            slotPosition: 110,
            sourceOrder: 0,
            representedWMLNames: ["tblBorders"],
            namespaceBindings: ["w": wordprocessingMLNamespace],
            xml: #"<w:tblBorders><w:top w:val="dashed"/></w:tblBorders>"#)]
        properties.borders = .all(Border(
            style: .single, size: 4, color: "000000"))

        let xml = properties.toXML()
        XCTAssertEqual(xml.components(separatedBy: "<w:tblBorders>").count - 1, 1)
        XCTAssertFalse(xml.contains("dashed"))
        XCTAssertTrue(xml.contains(#"<w:top w:val="single""#))
    }

    func testConditionalStyleMutationPreservesFalseAndDuplicateOccurrence() throws {
        let directory = FileManager.default.temporaryDirectory
            .appendingPathComponent("conditional-table-style-\(UUID().uuidString)")
        try FileManager.default.createDirectory(
            at: directory, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: directory) }

        let seedURL = directory.appendingPathComponent("seed.docx")
        var seed = WordDocument()
        seed.appendTable(Table(rows: [TableRow(cells: [TableCell(text: "A")])]))
        try DocxWriter.write(seed, to: seedURL)
        let extracted = try ZipHelper.unzip(seedURL)
        defer { ZipHelper.cleanup(extracted) }
        let documentURL = extracted.appendingPathComponent("word/document.xml")
        var xml = try String(contentsOf: documentURL, encoding: .utf8)
        let sourceStyles = """
        <q:tblStylePr q:type="firstRow"><q:rPr><q:b q:val="0"/></q:rPr></q:tblStylePr>
        <w:tblStylePr w:type="firstRow"><w:rPr><w:i/></w:rPr></w:tblStylePr>
        """.replacingOccurrences(of: "\n", with: "")
        xml = xml.replacingOccurrences(
            of: "<w:tblPr>",
            with: "<w:tblPr xmlns:q=\"\(wordprocessingMLNamespace)\">\(sourceStyles)")
        try xml.write(to: documentURL, atomically: true, encoding: .utf8)
        let inputURL = directory.appendingPathComponent("input.docx")
        try ZipHelper.zip(extracted, to: inputURL)

        var document = try DocxReader.read(from: inputURL)
        defer { document.close() }
        var updatedProperties = try XCTUnwrap(
            document.getTables().first?.conditionalStyles.first?.properties)
        XCTAssertEqual(updatedProperties.bold, false)
        updatedProperties.backgroundColor = "FF0000"
        try document.setTableConditionalStyle(
            tableIndex: 0,
            type: .firstRow,
            properties: updatedProperties)
        let outputURL = directory.appendingPathComponent("output.docx")
        try DocxWriter.write(document, to: outputURL)
        let outputPackage = try ZipHelper.unzip(outputURL)
        defer { ZipHelper.cleanup(outputPackage) }
        let output = try String(
            contentsOf: outputPackage.appendingPathComponent("word/document.xml"),
            encoding: .utf8)

        XCTAssertEqual(
            output.components(separatedBy: ":tblStylePr ").count - 1,
            2)
        XCTAssertTrue(output.contains(#"<w:b w:val="0"/>"#))
        XCTAssertTrue(output.contains("<w:i"))
        XCTAssertTrue(output.contains(#"w:fill="FF0000""#))
    }

    func testTypedLayoutMutationReplacesMarkupCompatibilityCarrier() {
        var properties = TableProperties()
        properties.rawChildren = [PreservedTableProperty(
            qualifiedName: "mc:AlternateContent",
            localName: "AlternateContent",
            namespaceURI: "http://schemas.openxmlformats.org/markup-compatibility/2006",
            styleProjectionIndex: nil,
            representedStyleTypes: [],
            slotPosition: 130,
            sourceOrder: 0,
            representedWMLNames: ["tblLayout"],
            namespaceBindings: [
                "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
                "w": wordprocessingMLNamespace,
            ],
            xml: """
            <mc:AlternateContent><mc:Choice Requires="w15"><w:tblLayout w:type="fixed"/></mc:Choice><mc:Fallback><w:tblLayout w:type="fixed"/></mc:Fallback></mc:AlternateContent>
            """.replacingOccurrences(of: "\n", with: ""))]
        properties.sourceProjection = TablePropertiesSourceProjection(
            width: nil,
            widthType: nil,
            alignment: nil,
            borders: nil,
            cellMargins: nil,
            layout: .fixed)
        properties.layout = .autofit

        let output = properties.toXML()
        XCTAssertFalse(output.contains("mc:AlternateContent"))
        XCTAssertEqual(
            output.components(separatedBy: "<w:tblLayout").count - 1,
            1)
        XCTAssertTrue(output.contains(#"w:type="autofit""#))
    }

    func testTypedLayoutMutationKeepsOtherPropertiesInMarkupCompatibilityCarrier() {
        var properties = TableProperties()
        let carrier = """
        <mc:AlternateContent><mc:Choice Requires="w15"><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0"/></mc:Choice><mc:Fallback><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0"/></mc:Fallback></mc:AlternateContent>
        """.replacingOccurrences(of: "\n", with: "")
        properties.rawChildren = [PreservedTableProperty(
            qualifiedName: "mc:AlternateContent",
            localName: "AlternateContent",
            namespaceURI: "http://schemas.openxmlformats.org/markup-compatibility/2006",
            styleProjectionIndex: nil,
            representedStyleTypes: [],
            slotPosition: 130,
            sourceOrder: 0,
            representedWMLNames: ["tblLayout", "tblLook"],
            namespaceBindings: [
                "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
                "w": wordprocessingMLNamespace,
            ],
            xml: carrier)]
        properties.inScopeNamespaces = [
            "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
            "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
        ]
        properties.sourceProjection = TablePropertiesSourceProjection(
            width: nil,
            widthType: nil,
            alignment: nil,
            borders: nil,
            cellMargins: nil,
            layout: .fixed)
        properties.layout = .autofit
        properties.width = 6000
        properties.widthType = .dxa

        let output = properties.toXML()
        XCTAssertTrue(output.contains("mc:AlternateContent"))
        XCTAssertEqual(
            output.components(separatedBy: "<w:tblLayout").count - 1,
            1)
        XCTAssertEqual(
            output.components(separatedBy: "<w:tblLook").count - 1,
            2)
        XCTAssertTrue(output.contains(#"w:type="autofit""#))
        let width = try? XCTUnwrap(output.range(of: "<w:tblW"))
        let layout = try? XCTUnwrap(output.range(of: "<w:tblLayout"))
        let carrierRange = try? XCTUnwrap(output.range(of: "<mc:AlternateContent"))
        XCTAssertNotNil(width)
        XCTAssertNotNil(layout)
        XCTAssertNotNil(carrierRange)
        if let width, let layout, let carrierRange {
            XCTAssertLessThan(width.lowerBound, layout.lowerBound)
            XCTAssertLessThan(layout.lowerBound, carrierRange.lowerBound)
        }
    }

    func testProcessContentCarrierParticipatesInTypedReplacement() throws {
        let directory = FileManager.default.temporaryDirectory
            .appendingPathComponent("process-content-table-property-\(UUID().uuidString)")
        try FileManager.default.createDirectory(
            at: directory, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: directory) }

        var seed = WordDocument()
        seed.appendTable(Table(rows: [TableRow(cells: [TableCell(text: "A")])]))
        let seedURL = directory.appendingPathComponent("seed.docx")
        try DocxWriter.write(seed, to: seedURL)
        let extracted = try ZipHelper.unzip(seedURL)
        defer { ZipHelper.cleanup(extracted) }
        let documentURL = extracted.appendingPathComponent("word/document.xml")
        var xml = try String(contentsOf: documentURL, encoding: .utf8)
        let carrier = """
        <q:layoutWrap xmlns:q="urn:process-content-test"><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0"/></q:layoutWrap>
        """.replacingOccurrences(of: "\n", with: "")
        xml = xml.replacingOccurrences(
            of: "<w:tblPr>",
            with: "<w:tblPr xmlns:v=\"urn:process-content-test\" mc:Ignorable=\"v\" mc:ProcessContent=\"v:layoutWrap\">\(carrier)")
        try xml.write(to: documentURL, atomically: true, encoding: .utf8)
        let inputURL = directory.appendingPathComponent("input.docx")
        try ZipHelper.zip(extracted, to: inputURL)

        var document = try DocxReader.read(from: inputURL)
        defer { document.close() }
        try document.setTableLayout(tableIndex: 0, type: .autofit)
        let outputURL = directory.appendingPathComponent("output.docx")
        try DocxWriter.write(document, to: outputURL)
        let outputPackage = try ZipHelper.unzip(outputURL)
        defer { ZipHelper.cleanup(outputPackage) }
        let output = try String(
            contentsOf: outputPackage.appendingPathComponent("word/document.xml"),
            encoding: .utf8)

        XCTAssertEqual(output.components(separatedBy: "<w:tblLayout").count - 1, 1)
        XCTAssertTrue(output.contains(#"w:type="autofit""#))
        XCTAssertTrue(output.contains("<q:layoutWrap"))
        XCTAssertTrue(output.contains(#"<w:tblLook w:val="04A0""#))
        let layout = try XCTUnwrap(output.range(of: "<w:tblLayout"))
        let wrapper = try XCTUnwrap(output.range(of: "<q:layoutWrap"))
        XCTAssertLessThan(layout.lowerBound, wrapper.lowerBound)
    }

    func testConditionalStyleMutationOnlyReplacesMatchingStyleInCarrier() {
        let firstRowSource = TableConditionalStyle(
            type: .firstRow,
            properties: TableConditionalStyleProperties(bold: false))
        let lastRowSource = TableConditionalStyle(
            type: .lastRow,
            properties: TableConditionalStyleProperties(italic: true))
        let carrier = """
        <mc:AlternateContent><mc:Choice Requires="w15"><w:tblStylePr w:type="firstRow"><w:rPr><w:b w:val="0"/></w:rPr></w:tblStylePr><w:tblStylePr w:type="lastRow"><w:rPr><w:i/></w:rPr></w:tblStylePr></mc:Choice><mc:Fallback><w:tblStylePr w:type="firstRow"><w:rPr><w:b w:val="0"/></w:rPr></w:tblStylePr><w:tblStylePr w:type="lastRow"><w:rPr><w:i/></w:rPr></w:tblStylePr></mc:Fallback></mc:AlternateContent>
        """.replacingOccurrences(of: "\n", with: "")
        var properties = TableProperties()
        properties.rawChildren = [PreservedTableProperty(
            qualifiedName: "mc:AlternateContent",
            localName: "AlternateContent",
            namespaceURI: "http://schemas.openxmlformats.org/markup-compatibility/2006",
            styleProjectionIndex: nil,
            representedStyleTypes: ["firstRow", "lastRow"],
            slotPosition: 800,
            sourceOrder: 0,
            representedWMLNames: ["tblStylePr"],
            namespaceBindings: [
                "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
                "w": wordprocessingMLNamespace,
            ],
            xml: carrier)]
        var table = Table(rows: [TableRow(cells: [TableCell(text: "A")])], properties: properties)
        table.conditionalStyles = [
            TableConditionalStyle(
                type: .firstRow,
                properties: TableConditionalStyleProperties(bold: true)),
            lastRowSource,
        ]
        table.sourcePropertyProjection = TableSourcePropertyProjection(
            tableIndent: nil,
            explicitLayout: nil,
            conditionalStyles: [firstRowSource, lastRowSource])

        let output = table.extendedTablePropertiesXML()

        XCTAssertTrue(output.contains("mc:AlternateContent"))
        XCTAssertEqual(output.components(separatedBy: #"w:type="firstRow""#).count - 1, 1)
        XCTAssertEqual(output.components(separatedBy: #"w:type="lastRow""#).count - 1, 2)
        XCTAssertTrue(output.contains("<w:b/>"))
        XCTAssertTrue(output.contains("<w:i/>"))
    }

    func testForeignWrapperRemainsOpaqueWithoutProcessContent() throws {
        let directory = FileManager.default.temporaryDirectory
            .appendingPathComponent("opaque-table-property-\(UUID().uuidString)")
        try FileManager.default.createDirectory(
            at: directory, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: directory) }

        var seed = WordDocument()
        seed.appendTable(Table(rows: [TableRow(cells: [TableCell(text: "A")])]))
        let seedURL = directory.appendingPathComponent("seed.docx")
        try DocxWriter.write(seed, to: seedURL)
        let extracted = try ZipHelper.unzip(seedURL)
        defer { ZipHelper.cleanup(extracted) }
        let documentURL = extracted.appendingPathComponent("word/document.xml")
        var xml = try String(contentsOf: documentURL, encoding: .utf8)
        let carrier = """
        <mc:AlternateContent><mc:Choice Requires="v"><v:metadata><w:tblLayout w:type="fixed"/></v:metadata><w:tblLook w:val="04A0"/></mc:Choice><mc:Fallback><v:metadata><w:tblLayout w:type="fixed"/></v:metadata><w:tblLook w:val="04A0"/></mc:Fallback></mc:AlternateContent>
        """.replacingOccurrences(of: "\n", with: "")
        xml = xml.replacingOccurrences(
            of: "<w:tblPr>",
            with: "<w:tblPr xmlns:v=\"urn:opaque-table-property\" mc:Ignorable=\"v\">\(carrier)")
        try xml.write(to: documentURL, atomically: true, encoding: .utf8)
        let inputURL = directory.appendingPathComponent("input.docx")
        try ZipHelper.zip(extracted, to: inputURL)

        var document = try DocxReader.read(from: inputURL)
        defer { document.close() }
        try document.setTableLayout(tableIndex: 0, type: .autofit)
        let outputURL = directory.appendingPathComponent("output.docx")
        try DocxWriter.write(document, to: outputURL)
        let outputPackage = try ZipHelper.unzip(outputURL)
        defer { ZipHelper.cleanup(outputPackage) }
        let output = try String(
            contentsOf: outputPackage.appendingPathComponent("word/document.xml"),
            encoding: .utf8)

        XCTAssertEqual(output.components(separatedBy: "<v:metadata").count - 1, 2)
        XCTAssertEqual(output.components(separatedBy: #"w:type="fixed""#).count - 1, 2)
        XCTAssertEqual(output.components(separatedBy: #"w:type="autofit""#).count - 1, 1)
        XCTAssertEqual(output.components(separatedBy: "<w:tblLook").count - 1, 2)
    }

    func testLocallyScopedProcessContentParticipatesInTypedReplacement() throws {
        let carrier = """
        <mc:AlternateContent><mc:Choice Requires="v" mc:ProcessContent="v:layoutWrap"><q:layoutWrap><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0"/></q:layoutWrap></mc:Choice><mc:Fallback mc:ProcessContent="v:layoutWrap"><q:layoutWrap><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0"/></q:layoutWrap></mc:Fallback></mc:AlternateContent>
        """.replacingOccurrences(of: "\n", with: "")
        let output = try mutateLayoutInInjectedProperties(
            carrier,
            tablePropertiesAttributes: "xmlns:v=\"urn:local-process-content\" xmlns:q=\"urn:local-process-content\" mc:Ignorable=\"v\"")

        XCTAssertEqual(output.components(separatedBy: "<w:tblLayout").count - 1, 1)
        XCTAssertTrue(output.contains(#"w:type="autofit""#))
        XCTAssertEqual(output.components(separatedBy: "<w:tblLook").count - 1, 2)
        XCTAssertTrue(output.contains("<q:layoutWrap"))
    }

    func testInheritedProcessContentBindingWinsOverEarlierLocalPrefixRebinding() throws {
        let carrier = """
        <mc:AlternateContent><mc:Choice Requires="x" xmlns:v="urn:inner" mc:ProcessContent="v:layoutWrap"><x:layoutWrap><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0"/></x:layoutWrap></mc:Choice><mc:Fallback mc:ProcessContent="v:layoutWrap"><q:layoutWrap><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0"/></q:layoutWrap></mc:Fallback></mc:AlternateContent>
        """.replacingOccurrences(of: "\n", with: "")
        let output = try mutateLayoutInInjectedProperties(
            carrier,
            tablePropertiesAttributes: "xmlns:v=\"urn:outer\" xmlns:q=\"urn:outer\" xmlns:x=\"urn:inner\" mc:Ignorable=\"v x\"")

        XCTAssertEqual(output.components(separatedBy: "<w:tblLayout").count - 1, 1)
        XCTAssertTrue(output.contains(#"w:type="autofit""#))
        XCTAssertEqual(output.components(separatedBy: "<w:tblLook").count - 1, 2)
        XCTAssertTrue(output.contains(#"xmlns:v="urn:inner""#))
        XCTAssertTrue(output.contains("<x:layoutWrap"))
        XCTAssertTrue(output.contains("<q:layoutWrap"))
    }

    func testUnicodeNamespacePrefixSurvivesCarrierRewrite() throws {
        let carrier = """
        <mc:AlternateContent><mc:Choice Requires="w15"><w:tblLayout w:type="fixed"/><擴:meta 擴:flag="1"/></mc:Choice><mc:Fallback><w:tblLayout w:type="fixed"/><擴:meta 擴:flag="1"/></mc:Fallback></mc:AlternateContent>
        """.replacingOccurrences(of: "\n", with: "")
        var properties = TableProperties()
        properties.rawChildren = [PreservedTableProperty(
            qualifiedName: "mc:AlternateContent",
            localName: "AlternateContent",
            namespaceURI: "http://schemas.openxmlformats.org/markup-compatibility/2006",
            styleProjectionIndex: nil,
            representedStyleTypes: [],
            slotPosition: 130,
            sourceOrder: 0,
            representedWMLNames: ["tblLayout"],
            namespaceBindings: [
                "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
                "w": wordprocessingMLNamespace,
                "擴": "urn:unicode-prefix",
            ],
            xml: carrier)]
        properties.inScopeNamespaces = ["擴": "urn:unicode-prefix"]
        properties.sourceProjection = TablePropertiesSourceProjection(
            width: nil,
            widthType: nil,
            alignment: nil,
            borders: nil,
            cellMargins: nil,
            layout: .fixed)
        properties.layout = .autofit

        let output = properties.toXML()

        XCTAssertEqual(output.components(separatedBy: "<w:tblLayout").count - 1, 1)
        XCTAssertEqual(output.components(separatedBy: "<擴:meta").count - 1, 2)
        XCTAssertTrue(output.contains(#"擴:flag="1""#))
        let wrapped = """
        <root xmlns:w="\(wordprocessingMLNamespace)" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">\(output)</root>
        """
        XCTAssertNoThrow(try XMLDocument(data: Data(wrapped.utf8)), output)
    }

    private func mutateLayoutInInjectedProperties(
        _ propertiesXML: String,
        tablePropertiesAttributes: String
    ) throws -> String {
        let directory = FileManager.default.temporaryDirectory
            .appendingPathComponent("carrier-table-property-\(UUID().uuidString)")
        try FileManager.default.createDirectory(
            at: directory, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: directory) }

        var seed = WordDocument()
        seed.appendTable(Table(rows: [TableRow(cells: [TableCell(text: "A")])]))
        let seedURL = directory.appendingPathComponent("seed.docx")
        try DocxWriter.write(seed, to: seedURL)
        let extracted = try ZipHelper.unzip(seedURL)
        defer { ZipHelper.cleanup(extracted) }
        let documentURL = extracted.appendingPathComponent("word/document.xml")
        var xml = try String(contentsOf: documentURL, encoding: .utf8)
        xml = xml.replacingOccurrences(
            of: "<w:tblPr>",
            with: "<w:tblPr \(tablePropertiesAttributes)>\(propertiesXML)")
        try xml.write(to: documentURL, atomically: true, encoding: .utf8)
        let inputURL = directory.appendingPathComponent("input.docx")
        try ZipHelper.zip(extracted, to: inputURL)

        var document = try DocxReader.read(from: inputURL)
        defer { document.close() }
        try document.setTableLayout(tableIndex: 0, type: .autofit)
        let outputURL = directory.appendingPathComponent("output.docx")
        try DocxWriter.write(document, to: outputURL)
        let outputPackage = try ZipHelper.unzip(outputURL)
        defer { ZipHelper.cleanup(outputPackage) }
        return try String(
            contentsOf: outputPackage.appendingPathComponent("word/document.xml"),
            encoding: .utf8)
    }

    private func directTablePropertyChildren(in xml: String) throws -> [String] {
        let document = try XMLDocument(
            data: Data(xml.utf8), options: [.nodePreserveAll])
        let tableProperties = try XCTUnwrap(
            document.nodes(forXPath: "//*[local-name()='tblPr']").first as? XMLElement)
        return (tableProperties.children ?? []).compactMap { node in
            (node as? XMLElement)?.xmlString(options: [.nodeCompactEmptyElement])
        }
    }
}
