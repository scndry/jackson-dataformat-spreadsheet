package io.github.scndry.jackson.dataformat.spreadsheet.deser;

/**
 * Low-level token enum representing structural events produced
 * by a {@link SheetReader} during sheet traversal.
 *
 * @see SheetReader
 */
public enum SheetToken {
    SHEET_DATA_START, ROW_START, CELL_VALUE, ROW_END, SHEET_DATA_END
}
