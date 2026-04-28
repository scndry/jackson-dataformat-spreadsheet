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

import io.github.scndry.jackson.dataformat.spreadsheet.schema.sheet.SheetConfigurer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.style.StylesBuilder;

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

    private static final int COMMENT_BOX_WIDTH_COLS = 2;
    private static final int COMMENT_BOX_HEIGHT_ROWS = 3;

    private final List<Column> _columns;
    private final CellAddress _origin;
    private final int _features;
    private final StylesBuilder _stylesBuilder;
    private final SheetConfigurer _sheetConfigurer;

    public SpreadsheetSchema(
            final List<Column> columns,
            final CellAddress origin,
            final int features,
            final StylesBuilder stylesBuilder,
            final SheetConfigurer sheetConfigurer) {
        _columns = columns;
        _origin = origin;
        _features = features;
        _stylesBuilder = stylesBuilder;
        _sheetConfigurer = sheetConfigurer;
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
        return new SpreadsheetSchema(reordered, _origin, _features, _stylesBuilder, _sheetConfigurer);
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
        _sheetConfigurer.apply(sheet, styles, this, lastRow);
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
