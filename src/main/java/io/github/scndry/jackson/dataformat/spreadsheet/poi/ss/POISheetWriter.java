package io.github.scndry.jackson.dataformat.spreadsheet.poi.ss;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.function.BiConsumer;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.SheetUtil;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.POICompat;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.ColumnPointer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Styles;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetWriter;

/**
 * POI cell-model {@link SheetWriter} implementation.
 * Writes values through {@link Cell} and delegates styles, merges, and column-width handling to POI APIs.
 *
 * @see io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.SSMLSheetWriter
 */
@Slf4j
public final class POISheetWriter implements SheetWriter {

    private static final int MAX_COLUMN_WIDTH = 255 * 256;

    private final Sheet _sheet;
    private final OutputStream _out;
    private SpreadsheetSchema _schema;
    private CellAddress _reference;
    private Styles _styles;
    private int _lastRow;
    private int _savedWindowSize = -1;

    public POISheetWriter(final Sheet sheet) {
        this(sheet, null);
    }

    public POISheetWriter(final Sheet sheet, final OutputStream out) {
        _sheet = sheet;
        _out = out;
    }

    @Override
    public SpreadsheetVersion getSpreadsheetVersion() {
        return _sheet.getWorkbook().getSpreadsheetVersion();
    }

    @Override
    public void setSchema(final SpreadsheetSchema schema) {
        _schema = schema;
        _styles = _schema.buildStyles(_sheet.getWorkbook());
    }

    @Override
    public void setReference(final CellAddress reference) {
        _reference = reference;
    }

    @Override
    public void writeHeaders() {
        final int row = _schema.getOriginRow();
        for (final Column column : _schema) {
            final int col = _schema.columnIndexOf(column);
            if (_sheet instanceof SXSSFSheet && column.getValue().isAutoSize()) {
                ((SXSSFSheet) _sheet).trackColumnForAutoSizing(col);
            }
            setReference(new CellAddress(row, col));
            writeString(column.getName());
        }
    }

    @Override
    public void writeNumeric(final double value) {
        _write(value, Cell::setCellValue);
    }

    @Override
    public void writeString(final String value) {
        _write(value, Cell::setCellValue);
    }

    @Override
    public void writeBoolean(final boolean value) {
        _write(value, Cell::setCellValue);
    }

    @Override
    public void writeBlank() {
        _write(null, (cell, o) -> cell.setBlank());
    }

    private <T> void _write(final T value, final BiConsumer<Cell, T> consumer) {
        final int row = _reference.getRow();
        final Cell cell;
        try {
            cell = CellUtil.getCell(CellUtil.getRow(row, _sheet), _reference.getColumn());
        } catch (IllegalArgumentException e) {
            if (_sheet instanceof SXSSFSheet) {
                final int window = ((SXSSFWorkbook) _sheet.getWorkbook())
                        .getRandomAccessWindowSize();
                throw new IllegalStateException(
                        "Cannot write to row " + row + " (already flushed)."
                        + " The nested list exceeded the SXSSF row window (" + window + ")."
                        + " Increase the window size via WorkbookProvider:"
                        + " new SpreadsheetFactory(() -> new SXSSFWorkbook(" + (row + 100) + "), ...)",
                        e);
            }
            throw e;
        }
        consumer.accept(cell, value);
        final Column column = _schema.findColumn(_reference);
        if (column != null) {
            final CellStyle style = _schema.getDataRow() > row
                    ? _styles.getHeaderStyle(column) : _styles.getStyle(column);
            if (style != null) {
                cell.setCellStyle(style);
            }
        }
        _lastRow = Math.max(_lastRow, row);
        if (log.isTraceEnabled()) {
            log.trace("{} {} {}", _reference, cell.getCellType(), cell);
        }
    }

    @Override
    public void ensureRowWindow(final int requiredRows) {
        if (!(_sheet instanceof SXSSFSheet)) return;
        final SXSSFSheet sxssf = (SXSSFSheet) _sheet;
        final int current = sxssf.getWorkbook().getRandomAccessWindowSize();
        if (current >= 0 && requiredRows > current) {
            try {
                sxssf.flushRows(1);
            } catch (java.io.IOException e) {
                throw new IllegalStateException(e);
            }
            _savedWindowSize = current;
            sxssf.setRandomAccessWindowSize(requiredRows + 1);
        }
    }

    @Override
    public void restoreRowWindow() {
        if (_savedWindowSize < 0) return;
        if (!(_sheet instanceof SXSSFSheet)) return;
        ((SXSSFSheet) _sheet).setRandomAccessWindowSize(_savedWindowSize);
        _savedWindowSize = -1;
    }

    @Override
    public void mergeScopedColumns(final ColumnPointer filter, final int row, final int size) {
        if (size <= 1) return;
        final List<Column> columns = _schema.getColumns(filter);
        for (final Column column : columns) {
            if (!column.isMerge()) continue;
            if (filter.relativize(column.getPointer()).contains(ColumnPointer.array())) continue;
            final int col = _schema.columnIndexOf(column);
            final CellRangeAddress region = new CellRangeAddress(row, row + size - 1, col, col);
            if (log.isTraceEnabled()) {
                log.trace(region.formatAsString());
            }
            _sheet.addMergedRegion(region);
        }
    }

    @Override
    public void write() throws IOException {
        if (_out != null) {
            _sheet.getWorkbook().write(_out);
        } else if (log.isDebugEnabled()) {
            log.debug("write() skipped — no OutputStream bound (direct Sheet mode)");
        }
    }

    @Override
    public void adjustColumnWidth() {
        for (final Column column : _schema) {
            final int col = _schema.columnIndexOf(column);
            final DataColumn.Value value = column.getValue();
            double width;
            if (value.isAutoSize()) {
                if (_sheet instanceof SXSSFSheet) {
                    _sheet.autoSizeColumn(col, true);
                    width = _sheet.getColumnWidth(col) / 256d;
                } else {
                    int firstRow = Math.max(_sheet.getFirstRowNum(), _lastRow - 100);
                    width = SheetUtil.getColumnWidth(_sheet, col, true, firstRow, _lastRow);
                }
                width = Math.max(width, value.getMinWidth());
                width = Math.min(width, value.getMaxWidth());
            } else {
                width = value.getWidth();
            }
            if (width <= 0) continue;
            width *= 256;
            width = Math.min(width, MAX_COLUMN_WIDTH);
            _sheet.setColumnWidth(col, (int) width);
        }
    }

    @Override
    public void close() throws IOException {
        final Workbook workbook = _sheet.getWorkbook();
        workbook.close();
        if (workbook instanceof SXSSFWorkbook) {
            ((SXSSFWorkbook) workbook).dispose();
        }
        if (_out != null) {
            _out.close();
        }
    }

    @Override
    public boolean isDate1904() {
        return POICompat.isDate1904(_sheet.getWorkbook());
    }
}
