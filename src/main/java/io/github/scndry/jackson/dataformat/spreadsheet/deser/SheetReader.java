package io.github.scndry.jackson.dataformat.spreadsheet.deser;

import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.CellAddress;

/**
 * Low-level interface for reading spreadsheet data as a
 * sequence of {@link SheetToken}s. Implementations provide
 * row-by-row, cell-by-cell iteration over a single sheet.
 *
 * @see SheetToken
 * @see SheetParser
 */
public interface SheetReader extends AutoCloseable, Iterator<SheetToken> {

    SpreadsheetVersion getSpreadsheetVersion();

    boolean isDate1904();

    CellAddress getReference();

    CellValue getCellValue();

    int getRow();

    int getColumn();

    boolean isClosed();

    @Override
    void close() throws IOException;
}
