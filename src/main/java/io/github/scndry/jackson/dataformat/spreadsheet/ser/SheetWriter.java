package io.github.scndry.jackson.dataformat.spreadsheet.ser;

import io.github.scndry.jackson.dataformat.spreadsheet.schema.ColumnPointer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.CellAddress;

import java.io.IOException;
import java.io.OutputStream;

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
