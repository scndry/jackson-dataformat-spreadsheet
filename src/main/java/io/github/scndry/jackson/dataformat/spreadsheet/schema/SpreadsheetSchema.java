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
import io.github.scndry.jackson.dataformat.spreadsheet.schema.style.StylesBuilder;
import lombok.extern.slf4j.Slf4j;

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
@Slf4j
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
    // Back-write safety projection — derived from "shared string ⇒ cell XML
    // upper bound is fixed by type" (BigDecimal/BigInteger excepted).
    //
    // Per-column max cell bytes:
    //   <c r="XFD1048576" s="32767" t="X"><v>VALUE</v></c>
    //   fixed tags = 27 bytes; cell ref ≤ 10; style index ≤ 5; type ≤ 3;
    //   value depends on JavaType (see _cellMaxBytes).
    //
    // Row tag overhead: <row r="N">...</row> = ~22 bytes.
    //
    // These projections let writeStartArray() fail-fast before _sb grows.
    // ----------------------------------------------------------------

    private static final int ROW_TAG_BYTES = 22;
    private static final int CELL_FIXED_TAGS_BYTES = 27;
    private static final int CELL_REF_MAX = 10;
    private static final int CELL_STYLE_MAX = 5;
    private static final int CELL_TYPE_MAX = 3;

    /** Upper-bound bytes for one cell XML of the given column.
     *  Returns -1 if the column's type is unbounded (BigDecimal/BigInteger).
     *  Caller can then decide between rejecting outright or using a worst-case
     *  estimate. */
    public static int cellMaxBytes(final Column column) {
        final JavaType type = column.getType();
        final Class<?> raw = type.getRawClass();
        final int valueMax;
        if (raw == int.class || raw == Integer.class
                || raw == short.class || raw == Short.class
                || raw == byte.class || raw == Byte.class) {
            valueMax = 12;                                    // -2147483648
        } else if (raw == long.class || raw == Long.class) {
            valueMax = 20;                                    // -9223372036854775808
        } else if (raw == float.class || raw == Float.class) {
            valueMax = 17;                                    // -3.4028235E38
        } else if (raw == double.class || raw == Double.class) {
            valueMax = 25;                                    // -1.7976931348623157E308
        } else if (raw == boolean.class || raw == Boolean.class) {
            valueMax = 1;
        } else if (raw == String.class || raw.isEnum()) {
            valueMax = 10;                                    // shared string index (Integer.MAX = 2147483647)
        } else if (java.util.Date.class.isAssignableFrom(raw)
                || raw.getName().startsWith("java.time.")) {
            valueMax = 25;                                    // Excel serial date double
        } else if (raw == java.math.BigDecimal.class || raw == java.math.BigInteger.class) {
            return -1;                                        // unbounded
        } else {
            valueMax = 25;                                    // unknown — assume double-like
        }
        return CELL_FIXED_TAGS_BYTES + CELL_REF_MAX + CELL_STYLE_MAX + CELL_TYPE_MAX + valueMax;
    }

    /** Upper-bound bytes for one inner row (all columns inside the given
     *  array pointer + the <row> tag). Returns -1 if any inner column type
     *  is unbounded. */
    public long innerRowMaxBytes(final ColumnPointer arrayPointer) {
        final List<Column> inners = getColumns(arrayPointer);
        long bytes = ROW_TAG_BYTES;
        for (final Column c : inners) {
            final int cellMax = cellMaxBytes(c);
            if (cellMax < 0) return -1;
            bytes += cellMax;
        }
        return bytes;
    }

    /** Projected back-write buffer max for a nested list of {@code listSize}
     *  elements rooted at {@code arrayPointer}. Returns -1 when unbounded. */
    public long projectBackWriteBuffer(final ColumnPointer arrayPointer, final int listSize) {
        if (listSize <= 0) return 0;
        final long perRow = innerRowMaxBytes(arrayPointer);
        if (perRow < 0) return -1;
        return (long) listSize * perRow;
    }

    /** Returns true when the column order has any outer (non-array-scope)
     *  column appearing after at least one inner (array-scope) column —
     *  the trigger for SSML back-write. Build-time detection used for an
     *  early warning to the user. */
    public boolean hasOuterFieldAfterList() {
        boolean seenInner = false;
        for (final Column c : _columns) {
            if (c == null) continue;
            final boolean inner = c.getPointer().contains(ColumnPointer.array());
            if (inner) {
                seenInner = true;
            } else if (seenInner) {
                return true;
            }
        }
        return false;
    }

    /** Returns true when any inner column (inside a nested array scope) has
     *  an unbounded type (BigDecimal/BigInteger). Such schemas cannot have
     *  their back-write buffer projected at writeStartArray — the safety
     *  layer falls back to the runtime _sb monitor. */
    public boolean hasUnboundedInnerType() {
        for (final Column c : _columns) {
            if (c == null) continue;
            if (!c.getPointer().contains(ColumnPointer.array())) continue;
            if (cellMaxBytes(c) < 0) return true;
        }
        return false;
    }

    /** Logs a warn once at schema build if an outer field is declared after
     *  a List<T> field — flags the back-write code path. Adds a second warn
     *  when any inner column has an unbounded type (BigDecimal/BigInteger)
     *  so the user knows the projected check cannot pre-detect, and only
     *  the runtime monitor will catch limit breaches. Call once after
     *  schema construction. */
    public void warnIfBackWriteScenario() {
        if (!hasOuterFieldAfterList()) return;
        log.warn(
                "Schema has an outer field declared after a List<T> field —"
                + " nested list contents are buffered in memory until the"
                + " outer field can be back-written. Large nested lists may"
                + " cause OOM in streaming mode."
                + " Consider USE_POI_USER_MODEL, splitting data, or"
                + " reordering fields so the outer field comes before the list."
                + " Columns: {}",
                _columns);
        if (hasUnboundedInnerType()) {
            log.warn(
                    "Schema has inner column(s) with unbounded type"
                    + " (BigDecimal/BigInteger). writeStartArray cannot"
                    + " pre-project buffer size; only the runtime _sb monitor"
                    + " will catch limit breaches.");
        }
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
