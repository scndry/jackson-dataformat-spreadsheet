package io.github.scndry.jackson.dataformat.spreadsheet.schema;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;

import com.fasterxml.jackson.core.FormatSchema;

/**
 * {@link FormatSchema} implementation that defines the column
 * layout for spreadsheet reading and writing. Contains an
 * ordered list of {@link Column} definitions, a cell origin,
 * and style configuration.
 *
 * @see Column
 * @see io.github.scndry.jackson.dataformat.spreadsheet.SchemaGenerator
 */
public final class SpreadsheetSchema implements FormatSchema, Iterable<Column> {

    public static final String SCHEMA_TYPE = "spreadsheet";

    public static final int FEATURE_USE_HEADER = 0x0001;
    public static final int FEATURE_COLUMN_REORDERING = 0x0002;

    public static final int DEFAULT_FEATURES = FEATURE_USE_HEADER;

    private final List<Column> _columns;
    private final Styles.Builder _stylesBuilder;
    private final CellAddress _origin;
    private final int _features;

    public SpreadsheetSchema(
            final List<Column> columns,
            final Styles.Builder stylesBuilder,
            final CellAddress origin,
            final int features) {
        _columns = columns;
        _stylesBuilder = stylesBuilder;
        _origin = origin;
        _features = features;
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
        return new SpreadsheetSchema(reordered, _stylesBuilder, _origin, _features);
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

    public Styles buildStyles(final Workbook workbook) {
        return _stylesBuilder.build(workbook);
    }

    public boolean isInRowBounds(final int row) {
        return getDataRow() <= row;
    }

    public boolean isInColumnBounds(final int col) {
        return getOriginColumn() <= col && col < getOriginColumn() + _columns.size();
    }
}
