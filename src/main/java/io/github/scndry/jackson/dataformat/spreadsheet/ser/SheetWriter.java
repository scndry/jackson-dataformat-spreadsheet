package io.github.scndry.jackson.dataformat.spreadsheet.ser;

import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.CellAddress;

import io.github.scndry.jackson.dataformat.spreadsheet.schema.ColumnPointer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;

/**
 * Low-level interface for writing values into spreadsheet cells.
 * Implementations handle cell creation, styling, column merging,
 * and final workbook output.
 *
 * @see SheetGenerator
 * @see SpreadsheetSchema
 */
public interface SheetWriter extends AutoCloseable {

    SpreadsheetVersion getSpreadsheetVersion();

    void setSchema(SpreadsheetSchema schema);

    void setReference(CellAddress reference);

    void writeHeaders();

    void writeNumeric(double value);

    void writeString(String value);

    void writeBoolean(boolean value);

    void writeBlank();

    void adjustColumnWidth();

    void mergeScopedColumns(ColumnPointer pointer, int row, int size);

    void write(OutputStream out) throws IOException;

    @Override
    void close() throws IOException;

    boolean isDate1904();
}
