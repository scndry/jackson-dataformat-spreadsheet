package io.github.scndry.jackson.dataformat.spreadsheet.schema;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.Incubating;

/**
 * Visitor invoked by {@link SpreadsheetSchema#forEachHeaderCell(HeaderLayoutVisitor)}
 * for each cell in the header region of a multi-row header layout. All
 * methods default to no-op so callers override only what they need.
 */
@Incubating
public interface HeaderLayoutVisitor {

    /** Leaf column header cell — written at row = origin + hierarchy depth. */
    default void visitColumnHeader(final int row, final int col, final Column column) {}

    /** Group header cell — possibly merged horizontally across firstCol..lastCol. */
    default void visitGroupCell(final int row, final int firstCol, final int lastCol,
                                final DataColumnGroup.Value group) {}

    /** Vertical merge region for shallow-hierarchy columns (firstRow..lastRow at col). */
    default void visitVerticalMerge(final int firstRow, final int lastRow, final int col) {}
}
