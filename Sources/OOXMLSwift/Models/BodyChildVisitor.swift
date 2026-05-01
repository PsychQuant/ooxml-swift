/// Shared recursive traversal for `BodyChild` trees.
///
/// The visitor makes recursion policy explicit at the call site: table cells,
/// nested tables, and block-level content controls can be enabled or skipped
/// independently. This is intentionally small so existing ad-hoc walkers can
/// migrate one at a time without changing their mutation semantics (#28).
protocol BodyChildVisitor {
    associatedtype State

    var initialState: State { get }
    var recursesIntoTableCells: Bool { get }
    var recursesIntoNestedTables: Bool { get }
    var recursesIntoContentControls: Bool { get }

    mutating func visitParagraph(_ paragraph: inout Paragraph, state: inout State)
    mutating func visitSkippedTable(_ table: inout Table, state: inout State)
    mutating func visitSkippedBodyChild(_ child: inout BodyChild, state: inout State)
}

extension BodyChildVisitor {
    var recursesIntoTableCells: Bool { true }
    var recursesIntoNestedTables: Bool { true }
    var recursesIntoContentControls: Bool { true }

    mutating func visitSkippedTable(_ table: inout Table, state: inout State) {}
    mutating func visitSkippedBodyChild(_ child: inout BodyChild, state: inout State) {}
}

enum BodyChildWalker {
    @discardableResult
    static func walk<V: BodyChildVisitor>(
        _ children: [BodyChild],
        visitor: inout V
    ) -> V.State {
        var mutableChildren = children
        return walk(&mutableChildren, visitor: &visitor)
    }

    @discardableResult
    static func walk<V: BodyChildVisitor>(
        _ children: inout [BodyChild],
        visitor: inout V
    ) -> V.State {
        var state = visitor.initialState
        walk(&children, visitor: &visitor, state: &state)
        return state
    }

    static func walk<V: BodyChildVisitor>(
        _ children: inout [BodyChild],
        visitor: inout V,
        state: inout V.State
    ) {
        for index in children.indices {
            walk(&children[index], visitor: &visitor, state: &state)
        }
    }

    private static func walk<V: BodyChildVisitor>(
        _ child: inout BodyChild,
        visitor: inout V,
        state: inout V.State
    ) {
        switch child {
        case .paragraph(var paragraph):
            visitor.visitParagraph(&paragraph, state: &state)
            child = .paragraph(paragraph)

        case .table(var table):
            if visitor.recursesIntoTableCells {
                walkTable(&table, visitor: &visitor, state: &state)
                child = .table(table)
            } else {
                visitor.visitSkippedTable(&table, state: &state)
                child = .table(table)
            }

        case .contentControl(let control, var children):
            if visitor.recursesIntoContentControls {
                walk(&children, visitor: &visitor, state: &state)
            }
            child = .contentControl(control, children: children)

        case .bookmarkMarker, .rawBlockElement:
            visitor.visitSkippedBodyChild(&child, state: &state)
        }
    }

    private static func walkTable<V: BodyChildVisitor>(
        _ table: inout Table,
        visitor: inout V,
        state: inout V.State
    ) {
        for rowIndex in table.rows.indices {
            for cellIndex in table.rows[rowIndex].cells.indices {
                for paragraphIndex in table.rows[rowIndex].cells[cellIndex].paragraphs.indices {
                    visitor.visitParagraph(
                        &table.rows[rowIndex].cells[cellIndex].paragraphs[paragraphIndex],
                        state: &state
                    )
                }

                for nestedIndex in table.rows[rowIndex].cells[cellIndex].nestedTables.indices {
                    if visitor.recursesIntoNestedTables {
                        walkTable(
                            &table.rows[rowIndex].cells[cellIndex].nestedTables[nestedIndex],
                            visitor: &visitor,
                            state: &state
                        )
                    } else {
                        visitor.visitSkippedTable(
                            &table.rows[rowIndex].cells[cellIndex].nestedTables[nestedIndex],
                            state: &state
                        )
                    }
                }
            }
        }
    }
}
