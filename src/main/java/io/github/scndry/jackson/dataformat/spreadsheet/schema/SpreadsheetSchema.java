package io.github.scndry.jackson.dataformat.spreadsheet.schema;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;

import com.fasterxml.jackson.core.FormatSchema;

import io.github.scndry.jackson.dataformat.spreadsheet.schema.feature.ConditionalFormattingConfigurer;

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
    private final ConditionalFormattingConfigurer _conditionalFormattings;
    private final CellAddress _origin;
    private final int _features;
    private final int _freezePaneColSplit;
    private final int _freezePaneRowSplit;
    private final boolean _autoFilter;

    public SpreadsheetSchema(
            final List<Column> columns,
            final Styles.Builder stylesBuilder,
            final ConditionalFormattingConfigurer conditionalFormattings,
            final CellAddress origin,
            final int features,
            final int freezePaneColSplit,
            final int freezePaneRowSplit,
            final boolean autoFilter) {
        _columns = columns;
        _stylesBuilder = stylesBuilder;
        _conditionalFormattings = conditionalFormattings;
        _origin = origin;
        _features = features;
        _freezePaneColSplit = freezePaneColSplit;
        _freezePaneRowSplit = freezePaneRowSplit;
        _autoFilter = autoFilter;
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
        return new SpreadsheetSchema(reordered, _stylesBuilder, _conditionalFormattings, _origin, _features,
                _freezePaneColSplit, _freezePaneRowSplit, _autoFilter);
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

    public void applyAutoFilter(final Sheet sheet, final int lastRow) {
        if (!_autoFilter || _columns.isEmpty()) return;
        final int firstCol = getOriginColumn();
        final int lastCol = firstCol + _columns.size() - 1;
        final int firstRow = getOriginRow();
        final int endRow = lastRow < 0
                ? sheet.getWorkbook().getSpreadsheetVersion().getMaxRows() - 1
                : lastRow;
        sheet.setAutoFilter(new CellRangeAddress(firstRow, endRow, firstCol, lastCol));
    }

    public void applyFreezePane(final Sheet sheet) {
        if (_freezePaneColSplit >= 0 || _freezePaneRowSplit >= 0) {
            sheet.createFreezePane(
                    Math.max(_freezePaneColSplit, 0),
                    Math.max(_freezePaneRowSplit, 0));
        }
    }

    public void applyConditionalFormattings(final Sheet sheet, final Styles styles, final int lastRow) {
        _conditionalFormattings.apply(sheet, styles, this, lastRow);
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
