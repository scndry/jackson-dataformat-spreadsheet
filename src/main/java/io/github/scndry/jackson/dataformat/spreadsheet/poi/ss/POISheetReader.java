package io.github.scndry.jackson.dataformat.spreadsheet.poi.ss;

import io.github.scndry.jackson.dataformat.spreadsheet.deser.CellValue;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetReader;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetToken;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.format.CellFormat;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;

import java.io.IOException;
import java.util.Iterator;
import java.util.NoSuchElementException;

@Slf4j
public final class POISheetReader implements SheetReader {

    private final Sheet _sheet;
    private final DataFormatter _formatter = new DataFormatter();
    private final Iterator<Row> _rowIterator;
    private Iterator<Cell> _cellIterator;
    private SheetToken _next;
    private Cell _cell;
    private int _rowIndex = -1;
    private int _columnIndex = -1;
    private boolean _closed;

    public POISheetReader(final Sheet sheet) {
        _sheet = sheet;
        _rowIterator = _sheet.rowIterator();
        _next = SheetToken.SHEET_DATA_START;
    }

    @Override
    public SpreadsheetVersion getSpreadsheetVersion() {
        return _sheet.getWorkbook().getSpreadsheetVersion();
    }

    @Override
    public boolean isDate1904() {
        final Workbook workbook = _sheet.getWorkbook();
        return workbook instanceof Date1904Support && ((Date1904Support) workbook).isDate1904();
    }

    @Override
    public CellAddress getReference() {
        return _cell == null ? null : _cell.getAddress();
    }

    @Override
    public CellValue getCellValue() {
        if (_cell == null) return null;
        final CellType type = CellFormat.ultimateType(_cell);
        switch (type) {
            case NUMERIC:
                final double value = _cell.getNumericCellValue();
                return new CellValue(value, _formattedString(value, _cell.getCellStyle()));
            case STRING:
                return new CellValue(_cell.getStringCellValue());
            case BOOLEAN:
                return CellValue.valueOf(_cell.getBooleanCellValue());
            case ERROR:
                return CellValue.getError(_cell.getErrorCellValue());
            case BLANK:
                return null;
            case _NONE:
            case FORMULA:
        }
        throw new IllegalStateException("Unexpected cell value type: " + type);
    }

    @Override
    public int getRow() {
        return _rowIndex;
    }

    @Override
    public int getColumn() {
        return _columnIndex;
    }

    @Override
    public boolean isClosed() {
        return _closed;
    }

    @Override
    public void close() throws IOException {
        _sheet.getWorkbook().close();
        _closed = true;
    }

    @Override
    public boolean hasNext() {
        return _next != null;
    }

    @Override
    public SheetToken next() {
        if (_next == null) throw new NoSuchElementException();
        final SheetToken token = _next;
        switch (token) {
            case SHEET_DATA_START:
                _next = _rowIterator.hasNext() ? SheetToken.ROW_START : SheetToken.SHEET_DATA_END;
                break;
            case ROW_START:
                final Row row = _rowIterator.next();
                _rowIndex = row.getRowNum();
                _cellIterator = row.cellIterator();
                _next = _cellIterator.hasNext() ? SheetToken.CELL_VALUE : SheetToken.ROW_END;
                break;
            case CELL_VALUE:
                _cell = _cellIterator.next();
                _columnIndex = _cell.getColumnIndex();
                _next = _cellIterator.hasNext() ? SheetToken.CELL_VALUE : SheetToken.ROW_END;
                break;
            case ROW_END:
                _cell = null;
                _cellIterator = null;
                _next = _rowIterator.hasNext() ? SheetToken.ROW_START : SheetToken.SHEET_DATA_END;
                break;
            case SHEET_DATA_END:
                _next = null;
                break;
        }
        if (log.isTraceEnabled()) {
            if (token == SheetToken.CELL_VALUE) {
                log.trace("{} {} {}", token, getReference(), getCellValue());
            } else {
                log.trace("{}", token);
            }
        }
        return token;
    }

    private String _formattedString(final double value, final CellStyle style) {
        if (style != null && style.getDataFormatString() != null) {
            return _formatter.formatRawCellContents(value, style.getDataFormat(), style.getDataFormatString());
        }
        return null;
    }
}
