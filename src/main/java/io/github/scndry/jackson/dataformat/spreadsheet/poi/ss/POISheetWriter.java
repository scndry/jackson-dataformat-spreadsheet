package io.github.scndry.jackson.dataformat.spreadsheet.poi.ss;

import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.BiConsumer;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.SheetUtil;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.POICompat;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.ColumnPointer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.HeaderLayoutVisitor;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Styles;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.internal.HeaderComments;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetWriter;

/**
 * POI User Model {@link SheetWriter} implementation.
 * Writes values through {@link Cell} and delegates styles, merges, and column-width handling to POI APIs.
 *
 * @see io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.SSMLSheetWriter
 */
@Slf4j
public final class POISheetWriter implements SheetWriter {

    private static final int MAX_COLUMN_WIDTH = 255 * 256;
    // AUTOSIZE_FULL_SCAN_ROWS must be >= AUTOSIZE_SAMPLE_PERIOD; otherwise rows between the
    // full-scan range and the first sample point would be silently skipped.
    private static final int AUTOSIZE_FULL_SCAN_ROWS = 100;
    private static final int AUTOSIZE_SAMPLE_PERIOD = 100;

    private final Sheet _sheet;
    private final OutputStream _out;
    private SpreadsheetSchema _schema;
    private CellAddress _reference;
    private Styles _styles;
    private int _lastRow;
    private int _savedWindowSize = -1;
    private final Map<Integer, Double> _columnMaxWidth = new HashMap<>();
    private DataFormatter _dataFormatter;
    private int _defaultCharWidth;

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
        HeaderComments.apply(_sheet, _schema);
        _dataFormatter = new DataFormatter();
        _defaultCharWidth = SheetUtil.getDefaultCharWidth(_sheet.getWorkbook());
    }

    @Override
    public void setReference(final CellAddress reference) {
        _reference = reference;
    }

    @Override
    public void writeHeaders() {
        _schema.forEachHeaderCell(new HeaderLayoutVisitor() {
            @Override
            public void visitColumnHeader(final int row, final int col, final Column column) {
                setReference(new CellAddress(row, col));
                writeString(column.getName());
            }

            @Override
            public void visitGroupCell(final int row, final int firstCol, final int lastCol,
                                       final DataColumnGroup.Value group) {
                setReference(new CellAddress(row, firstCol));
                writeString(group.getName());
                final CellStyle gs = _styles.resolve(group.getHeaderStyle(), null);
                final CellStyle anchorStyle;
                if (gs != null) {
                    CellUtil.getCell(CellUtil.getRow(row, _sheet), firstCol).setCellStyle(gs);
                    anchorStyle = gs;
                } else {
                    final Column column = _schema.findColumn(firstCol);
                    anchorStyle = column == null ? null
                            : _styles.resolve(column.getValue().getHeaderStyle(),
                                    column.getType().getRawClass());
                }
                if (firstCol < lastCol) {
                    _sheet.addMergedRegion(new CellRangeAddress(row, row, firstCol, lastCol));
                    _fillMergedInnerCellsHorizontal(row, firstCol, lastCol, anchorStyle);
                }
            }

            @Override
            public void visitVerticalMerge(final int firstRow, final int lastRow, final int col) {
                _sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, col, col));
                final Column column = _schema.findColumn(col);
                final CellStyle hs = column == null ? null
                        : _styles.resolve(column.getValue().getHeaderStyle(),
                                column.getType().getRawClass());
                _fillMergedInnerCellsVertical(firstRow, lastRow, col, hs);
            }
        });
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
            final String name = _schema.getDataRow() > row
                    ? column.getValue().getHeaderStyle()
                    : column.getValue().getStyle();
            final CellStyle style = _styles.resolve(name, column.getType().getRawClass());
            if (style != null) {
                cell.setCellStyle(style);
            }
            if (column.getValue().isAutoSize() && _shouldMeasureAutoSize(row)) {
                _accumulateAutoSizeWidth(_reference.getColumn(), cell);
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
            } catch (IOException e) {
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
            final CellStyle ds = _styles.resolve(column.getValue().getStyle(),
                    column.getType().getRawClass());
            _fillMergedInnerCellsVertical(row, row + size - 1, col, ds);
        }
    }

    /** Fill the inner cells of a vertical merge with the same style as the
     *  top cell so cell-by-cell viewers (LibreOffice / Numbers) render the
     *  merged region as a closed rectangle. Only triggers when the column
     *  has an explicit style — without one, the inner cells are left
     *  untouched (no cell entry created), matching SSML's writer. */
    private void _fillMergedInnerCellsVertical(final int firstRow, final int lastRow,
                                               final int col, final CellStyle style) {
        if (style == null) return;
        for (int r = firstRow + 1; r <= lastRow; r++) {
            CellUtil.getCell(CellUtil.getRow(r, _sheet), col).setCellStyle(style);
        }
    }

    /** Horizontal counterpart of
     *  {@link #_fillMergedInnerCellsVertical} for group-header merges. */
    private void _fillMergedInnerCellsHorizontal(final int row, final int firstCol,
                                                 final int lastCol, final CellStyle style) {
        if (style == null) return;
        for (int c = firstCol + 1; c <= lastCol; c++) {
            CellUtil.getCell(CellUtil.getRow(row, _sheet), c).setCellStyle(style);
        }
    }

    @Override
    public void write() throws IOException {
        _schema.configureSheet(_sheet, _styles, _lastRow);
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
                width = _columnMaxWidth.getOrDefault(col, 0.0);
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

    private boolean _shouldMeasureAutoSize(final int row) {
        final int dataStart = _schema.getDataRow();
        if (row < dataStart) return true;
        final int relRow = row - dataStart;
        return relRow < AUTOSIZE_FULL_SCAN_ROWS || relRow % AUTOSIZE_SAMPLE_PERIOD == 0;
    }

    private void _accumulateAutoSizeWidth(final int col, final Cell cell) {
        try {
            final double w = SheetUtil.getCellWidth(cell, _defaultCharWidth, _dataFormatter, false);
            if (w > 0) {
                _columnMaxWidth.merge(col, w, Math::max);
            }
        } catch (RuntimeException e) {
            if (log.isDebugEnabled()) {
                log.debug("autoSize measurement failed at col {}: {}", col, e.getMessage());
            }
        }
    }

    @Override
    public void close() throws IOException {
        final Workbook workbook = _sheet.getWorkbook();
        workbook.close();
        // Explicit dispose() — close→dispose became automatic only in recent POI (Bug 68183); safe across POI 4.1.1+.
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
