package io.github.scndry.jackson.dataformat.spreadsheet.ser;

import java.io.IOException;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.CellAddress;

import io.github.scndry.jackson.dataformat.spreadsheet.schema.ColumnPointer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.internal.SpreadsheetSchemaImpl;

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

    void setSchema(SpreadsheetSchemaImpl schema);

    void setReference(CellAddress reference);

    void writeHeaders();

    void writeNumeric(double value);

    void writeString(String value);

    void writeBoolean(boolean value);

    void writeBlank();

    void adjustColumnWidth();

    void mergeScopedColumns(ColumnPointer pointer, int row, int size);

    default void ensureRowWindow(int requiredRows) {}

    default void restoreRowWindow() {}

    /**
     * Hook signaling that the generator entered a list scope. Streaming
     * writers can suspend output flushing until {@link #exitArrayScope()}
     * so that an outer field declared after the list can back-write into
     * the list's first element row.
     */
    default void enterArrayScope() {}

    /** Counterpart of {@link #enterArrayScope()}. */
    default void exitArrayScope() {}

    void write() throws IOException;

    @Override
    void close() throws IOException;

    boolean isDate1904();
}
