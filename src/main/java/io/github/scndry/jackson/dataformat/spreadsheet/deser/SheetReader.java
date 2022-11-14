package io.github.scndry.jackson.dataformat.spreadsheet.deser;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.CellAddress;

import java.io.IOException;
import java.util.Iterator;

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
