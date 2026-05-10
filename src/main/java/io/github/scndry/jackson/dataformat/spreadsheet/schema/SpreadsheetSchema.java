package io.github.scndry.jackson.dataformat.spreadsheet.schema;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;

import com.fasterxml.jackson.core.FormatSchema;

import com.fasterxml.jackson.databind.JavaType;

import io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.GridConfigurer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.internal.BackWriteProjection;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.style.StylesBuilder;

/**
 * {@link FormatSchema} implementation that defines the column
 * layout for spreadsheet reading and writing. Contains an
 * ordered list of {@link Column} definitions, a cell origin,
 * and style configuration. Do not mutate the builders or
 * columns after construction.
 *
 * @see Column
 * @see io.github.scndry.jackson.dataformat.spreadsheet.SchemaGenerator
 */
public final class SpreadsheetSchema implements FormatSchema, Iterable<Column> {

    public static final String SCHEMA_TYPE = "spreadsheet";

    public static final int FEATURE_USE_HEADER = 0x0001;
    public static final int FEATURE_COLUMN_REORDERING = 0x0002;

    public static final int DEFAULT_FEATURES = FEATURE_USE_HEADER;

    private static final int COMMENT_BOX_WIDTH_COLS = 2;
    private static final int COMMENT_BOX_HEIGHT_ROWS = 3;

    private final List<Column> _columns;
    private final CellAddress _origin;
    private final int _features;
    private final StylesBuilder _stylesBuilder;
    private final GridConfigurer _gridConfigurer;
    private volatile Boolean _backWriteRequired;   // lazy-memoized

    public SpreadsheetSchema(
            final List<Column> columns,
            final CellAddress origin,
            final int features,
            final StylesBuilder stylesBuilder,
            final GridConfigurer gridConfigurer) {
        _columns = columns;
        _origin = origin;
        _features = features;
        _stylesBuilder = stylesBuilder;
        _gridConfigurer = gridConfigurer;
    }

    @Override
    public String getSchemaType() {
        return SCHEMA_TYPE;
    }

    @Override
    public Iterator<Column> iterator() {
        return _columns.iterator();
    }

    public Column findColumn(final CellAddress reference) {
        if (_columns.isEmpty()) {
            return null;
        }
        final int idx = reference.getColumn() - getOriginColumn();
        if (idx < 0 || idx >= _columns.size()) {
            return null;
        }
        return _columns.get(idx);
    }

    public Column getColumn(final CellAddress reference) {
        return _columns.get(reference.getColumn() - getOriginColumn());
    }

    public int getDataRow() {
        return _origin.getRow() + (usesHeader() ? 1 : 0);
    }

    public boolean usesHeader() {
        return (_features & FEATURE_USE_HEADER) != 0;
    }

    public boolean reordersColumns() {
        return (_features & FEATURE_COLUMN_REORDERING) != 0;
    }

    public SpreadsheetSchema reorderColumns(final List<String> headers) {
        final List<Column> reordered = new ArrayList<>(headers.size());
        for (final String header : headers) {
            Column matched = null;
            if (header != null) {
                for (final Column col : _columns) {
                    if (col != null && col.matchesName(header)) {
                        matched = col;
                        break;
                    }
                }
            }
            reordered.add(matched);
        }
        return new SpreadsheetSchema(reordered, _origin, _features, _stylesBuilder, _gridConfigurer);
    }

    public int getOriginColumn() {
        return _origin.getColumn();
    }

    public int columnIndexOf(final ColumnPointer pointer) {
        for (int i = 0; i < _columns.size(); i++) {
            final Column col = _columns.get(i);
            if (col != null && col.matches(pointer)) {
                return i + getOriginColumn();
            }
        }
        return -1;
    }

    public int columnIndexOf(final Column column) {
        return columnIndexOf(column.getPointer());
    }

    public int getOriginRow() {
        return _origin.getRow();
    }

    public List<Column> getColumns(final ColumnPointer filter) {
        if (filter.isEmpty()) {
            return _columns;
        }
        return _columns
                .stream()
                .filter(c -> c
                .getPointer()
                .startsWith(filter)).collect(Collectors
                .toList());
    }

    // ----------------------------------------------------------------
    // Back-write safety — public utilities for cell-XML upper-bound
    // computation and the runtime limit. Per-row / per-list projection
    // and the back-write trigger detector live in
    // io.github.scndry.jackson.dataformat.spreadsheet.schema.internal
    // .BackWriteProjection (internal API).
    // ----------------------------------------------------------------

    /** True when this schema can trigger the SSML back-write code path
     *  (outer field declared after a {@code List<T>} in column order).
     *  Computed once and memoized — callers use this to gate the
     *  array-scope flush suspension so schemas where no back-write can
     *  happen avoid the buffer-accumulation overhead. */
    public boolean requiresBackWriteScope() {
        Boolean v = _backWriteRequired;
        if (v == null) {
            v = BackWriteProjection.hasOuterFieldAfterList(this);
            _backWriteRequired = v;
        }
        return v;
    }

    private static final int CELL_FIXED_TAGS_BYTES = 27;
    private static final int CELL_REF_MAX = 10;
    private static final int CELL_STYLE_MAX = 5;
    private static final int CELL_TYPE_MAX = 3;

    /** System property name for the back-write buffer limit (bytes).
     *  Resolved per call by {@link #backWriteBufferLimit()} so test code
     *  can vary it without JVM restart. */
    public static final String BACK_WRITE_BUFFER_BYTES_PROPERTY =
            "spreadsheet.backWriteBufferBytes";

    /** Back-write buffer limit (bytes). System property override:
     *  {@code -Dspreadsheet.backWriteBufferBytes=<n>}. Default =
     *  max(16 MB, heap/8) — heap/8 accounts for StringBuilder grow's
     *  2× peak plus headroom for other allocations. */
    public static long backWriteBufferLimit() {
        final String configured = System.getProperty(BACK_WRITE_BUFFER_BYTES_PROPERTY);
        if (configured != null) {
            try {
                final long v = Long.parseLong(configured);
                if (v > 0) return v;
            } catch (NumberFormatException ignored) {
                // fall through to default
            }
        }
        return Math.max(16L * 1024 * 1024, Runtime.getRuntime().maxMemory() / 8);
    }

    /** Human-readable byte size formatting for diagnostics. */
    public static String formatBytes(final long bytes) {
        if (bytes < 1024L) return bytes + " B";
        if (bytes < 1024L * 1024) return (bytes / 1024) + " KB";
        if (bytes < 1024L * 1024 * 1024) {
            return String.format("%.1f MB", bytes / (1024.0 * 1024));
        }
        return String.format("%.2f GB", bytes / (1024.0 * 1024 * 1024));
    }

    /** Upper-bound bytes for one cell XML of the given column.
     *
     *  Cell-XML output paths in this library:
     *    - All numeric overloads (writeNumber(int|long|float|double))
     *      route through SheetGenerator → _writer.writeNumeric(double),
     *      which calls StringBuilder.append(double) — so every numeric
     *      primitive emits a Double.toString-formatted value (e.g.
     *      "1.0", "-1.7976931348623157E308"), regardless of the source
     *      type. Worst-case value length = ~25 chars.
     *    - boolean emits "0" or "1" (1 char).
     *    - String / Enum / BigDecimal / BigInteger route through
     *      SheetGenerator → _writer.writeString → SharedStringsStore,
     *      leaving only a shared-string index (max 10 digits for
     *      Integer.MAX_VALUE) in the cell. Their content lives in
     *      sharedStrings.xml, a separate stream not gated by
     *      back-write buffer limits.
     *    - Date / java.time.* route through writeNumber(double serial),
     *      same 25-char bound.
     *
     *  Every supported Java type therefore has a bounded cell XML size. */
    public static int cellMaxBytes(final Column column) {
        final JavaType type = column.getType();
        final Class<?> raw = type.getRawClass();
        final int valueMax;
        if (raw == boolean.class || raw == Boolean.class) {
            valueMax = 1;
        } else if (raw == String.class || raw.isEnum()
                || raw == java.math.BigDecimal.class
                || raw == java.math.BigInteger.class) {
            valueMax = 10;                                    // shared string index
        } else {
            // All numeric primitives + Date/java.time.* emit
            // Double.toString-formatted values. Use the double worst case.
            valueMax = 25;
        }
        return CELL_FIXED_TAGS_BYTES + CELL_REF_MAX + CELL_STYLE_MAX + CELL_TYPE_MAX + valueMax;
    }


    public Styles buildStyles(final Workbook workbook) {
        return _stylesBuilder.build(workbook);
    }

    public void applyHeaderComments(final Sheet sheet) {
        if (!usesHeader()) return;
        final int row = getOriginRow();
        final CreationHelper factory = sheet.getWorkbook().getCreationHelper();
        Drawing<?> drawing = null;
        for (final Column column : _columns) {
            if (column == null) continue;
            final String text = column.getValue().getComment();
            if (text.isEmpty()) continue;
            if (drawing == null) {
                drawing = sheet.createDrawingPatriarch();
            }
            final int col = columnIndexOf(column);
            final ClientAnchor anchor = factory.createClientAnchor();
            anchor.setCol1(col);
            anchor.setRow1(row);
            anchor.setCol2(col + COMMENT_BOX_WIDTH_COLS);
            anchor.setRow2(row + COMMENT_BOX_HEIGHT_ROWS);
            final Comment comment = drawing.createCellComment(anchor);
            comment.setString(factory.createRichTextString(text));
            comment.setAddress(row, col);
        }
    }

    public void configureSheet(final Sheet sheet, final Styles styles, final int lastRow) {
        _gridConfigurer.apply(sheet, styles, this, lastRow);
    }

    public int columnIndexByName(final String name) {
        for (int i = 0; i < _columns.size(); i++) {
            final Column col = _columns.get(i);
            if (col != null && col.matchesName(name)) {
                return i + getOriginColumn();
            }
        }
        return -1;
    }

    public int columnCount() {
        return _columns.size();
    }

    public List<String> columnNames() {
        final List<String> names = new ArrayList<>();
        for (final Column col : _columns) {
            if (col != null) names.add(col.getName());
        }
        return names;
    }

    public boolean isInRowBounds(final int row) {
        return getDataRow() <= row;
    }

    public boolean isInColumnBounds(final int col) {
        return getOriginColumn() <= col && col < getOriginColumn() + _columns.size();
    }
}
