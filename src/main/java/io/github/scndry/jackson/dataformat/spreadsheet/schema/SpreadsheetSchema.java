package io.github.scndry.jackson.dataformat.spreadsheet.schema;

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
 * @see SchemaGenerator
 */
public final class SpreadsheetSchema implements FormatSchema, Iterable<Column> {

    public static final String SCHEMA_TYPE = "spreadsheet";
    private final List<Column> _columns;
    private final Styles.Builder _stylesBuilder;
    private final CellAddress _origin;
    private final boolean _useHeader;

    public SpreadsheetSchema(
            final List<Column> columns,
            final Styles.Builder stylesBuilder,
            final CellAddress origin) {
        this(columns, stylesBuilder, origin, true);
    }

    public SpreadsheetSchema(
            final List<Column> columns,
            final Styles.Builder stylesBuilder,
            final CellAddress origin,
            final boolean useHeader) {
        _columns = columns;
        _stylesBuilder = stylesBuilder;
        _origin = origin;
        _useHeader = useHeader;
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
        return getColumn(reference);
    }

    public Column getColumn(final CellAddress reference) {
        return _columns.get(reference.getColumn() - getOriginColumn());
    }

    public int getDataRow() {
        return _origin.getRow() + (_useHeader ? 1 : 0);
    }

    public boolean usesHeader() {
        return _useHeader;
    }

    public SpreadsheetSchema withUseHeader(final boolean state) {
        return _useHeader == state ? this : new SpreadsheetSchema(
                _columns,
                _stylesBuilder,
                _origin,
                state);
    }

    public int getOriginColumn() {
        return _origin.getColumn();
    }

    public int columnIndexOf(final ColumnPointer pointer) {
        for (int i = 0; i < _columns.size(); i++) {
            if (_columns.get(i).matches(pointer)) {
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
