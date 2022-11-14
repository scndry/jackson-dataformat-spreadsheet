package io.github.scndry.jackson.dataformat.spreadsheet.schema;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

public interface Styles {

    CellStyle getStyle(Column column);

    default CellStyle getHeaderStyle(final Column column) {
        return getStyle(column);
    }

    interface Builder {
        Styles build(Workbook workbook);
    }
}
