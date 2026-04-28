package io.github.scndry.jackson.dataformat.spreadsheet.schema;

import org.apache.poi.ss.usermodel.CellStyle;

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
}
