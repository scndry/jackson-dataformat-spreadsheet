package io.github.scndry.jackson.dataformat.spreadsheet.schema;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Resolves Apache POI {@link CellStyle} instances for data and
 * header cells. Implementations are created per-workbook via
 * the {@link Builder} interface.
 *
 * @see Column
 * @see SpreadsheetSchema
 */
public interface Styles {

    CellStyle getStyle(Column column);

    default CellStyle getHeaderStyle(final Column column) {
        return getStyle(column);
    }

    interface Builder {
        Styles build(Workbook workbook);
    }
}
