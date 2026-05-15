package io.github.scndry.jackson.dataformat.spreadsheet.schema;

import org.apache.poi.ss.usermodel.CellStyle;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;

/**
 * Resolves Apache POI {@link CellStyle} instances for data and
 * header cells. Built per-workbook by
 * {@link io.github.scndry.jackson.dataformat.spreadsheet.schema.style.StylesBuilder}.
 *
 * @see Column
 * @see SpreadsheetSchema
 */
public interface Styles {

    CellStyle getStyle(Column column);

    default CellStyle getHeaderStyle(final Column column) {
        return getStyle(column);
    }

    CellStyle getStyle(String name);

    /** Resolves the cell style for a {@code @DataColumnGroup} header cell.
     *  Returns {@code null} when no style is registered for the group. */
    default CellStyle getGroupHeaderStyle(final DataColumnGroup.Value group) {
        return null;
    }
}
