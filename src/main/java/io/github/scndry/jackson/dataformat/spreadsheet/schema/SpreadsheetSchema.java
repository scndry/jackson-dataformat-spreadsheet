package io.github.scndry.jackson.dataformat.spreadsheet.schema;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;

import com.fasterxml.jackson.core.FormatSchema;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.GridConfigurer;
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

    private final List<Column> _columns;
    private final CellAddress _origin;
    private final int _features;
    private final StylesBuilder _stylesBuilder;
    private final GridConfigurer _gridConfigurer;
    private final int _headerRowCount;
    // Pre-built lookup tables for resolve / resolveArray on the write path's
    // per-cell currentPointer chain. Every cell's pointer is a prefix of
    // some schema column pointer, so all needed (parent, name) and
    // (parent, []) pairs can be enumerated at schema construction. Avoids
    // per-cell SegmentPointer / String[] allocation in
    // SheetStreamContext.ObjectContext/ArrayContext.currentPointer().
    private final Map<ColumnPointer, Map<String, ColumnPointer>> _resolveTable;
    private final Map<ColumnPointer, ColumnPointer> _resolveArrayTable;

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
        _headerRowCount = _computeHeaderRowCount(columns);
        _resolveTable = new HashMap<>();
        _resolveArrayTable = new HashMap<>();
        _buildResolveTables();
    }

    private void _buildResolveTables() {
        for (final Column col : _columns) {
            if (col == null) continue;
            ColumnPointer head = ColumnPointer.empty();
            for (final ColumnPointer seg : col.getPointer()) {
                ColumnPointer next;
                if (seg.equals(ColumnPointer.array())) {
                    next = _resolveArrayTable.get(head);
                    if (next == null) {
                        next = head.resolveArray();
                        _resolveArrayTable.put(head, next);
                    }
                } else {
                    final String name = seg.name();
                    final Map<String, ColumnPointer> nameMap =
                            _resolveTable.computeIfAbsent(head, k -> new HashMap<>());
                    next = nameMap.get(name);
                    if (next == null) {
                        next = head.resolve(name);
                        nameMap.put(name, next);
                    }
                }
                head = next;
            }
        }
    }

    /**
     * Internal hot-path cache for per-cell pointer chains; external
     * callers should use {@link ColumnPointer#resolve(String)}.
     */
    public ColumnPointer resolve(final ColumnPointer parent, final String name) {
        final Map<String, ColumnPointer> nameMap = _resolveTable.get(parent);
        if (nameMap != null) {
            final ColumnPointer cached = nameMap.get(name);
            if (cached != null) return cached;
        }
        return parent.resolve(name);
    }

    /**
     * Internal hot-path cache for per-cell pointer chains; external
     * callers should use {@link ColumnPointer#resolveArray()}.
     */
    public ColumnPointer resolveArray(final ColumnPointer parent) {
        final ColumnPointer cached = _resolveArrayTable.get(parent);
        if (cached != null) return cached;
        return parent.resolveArray();
    }

    private static int _computeHeaderRowCount(final List<Column> columns) {
        int max = 0;
        for (final Column col : columns) {
            if (col == null) continue;
            final int depth = col.getGroupHierarchy().depth();
            if (depth > max) max = depth;
        }
        return max + 1;
    }

    @Override
    public String getSchemaType() {
        return SCHEMA_TYPE;
    }

    /**
     * Iterates the schema's columns in declared order. May yield {@code null}
     * at sparse gap positions introduced by {@code @DataColumn(shift)} or
     * {@code @DataColumnGroup(shift)} — callers must null-check each element.
     */
    @Override
    public Iterator<Column> iterator() {
        return _columns.iterator();
    }

    public Column findColumn(final CellAddress reference) {
        return findColumn(reference.getColumn());
    }

    /**
     * Returns the column at the given sheet column index, or {@code null} if
     * the index falls outside the schema range OR lands on a sparse gap
     * position introduced by shift.
     */
    public Column findColumn(final int column) {
        if (_columns.isEmpty()) {
            return null;
        }
        final int idx = column - getOriginColumn();
        if (idx < 0 || idx >= _columns.size()) {
            return null;
        }
        return _columns.get(idx);
    }

    public Column getColumn(final CellAddress reference) {
        return getColumn(reference.getColumn());
    }

    /**
     * Returns the column at the given sheet column index. Throws
     * {@link IndexOutOfBoundsException} when out of range; returns {@code null}
     * at sparse gap positions introduced by shift. The {@code null}-on-gap
     * behavior diverges from typical {@code get} semantics — callers that
     * need a strict non-null contract should prefer explicit index checks.
     */
    public Column getColumn(final int column) {
        return _columns.get(column - getOriginColumn());
    }

    public int getDataRow() {
        return _origin.getRow() + (usesHeader() ? effectiveHeaderRowCount() : 0);
    }

    public int getHeaderRowCount() {
        return effectiveHeaderRowCount();
    }

    /** Row index of the leaf header row (the row carrying column names).
     *  Collapses to the origin row when {@link #usesHeader()} is disabled. */
    public int getLeafHeaderRow() {
        return _origin.getRow() + (usesHeader() ? effectiveHeaderRowCount() - 1 : 0);
    }

    /** Effective header rows: shift collapses the layout to a single leaf row
     *  (group label rows are skipped; cascade attributes still apply per-column). */
    private int effectiveHeaderRowCount() {
        return anyNestedShift() ? 1 : _headerRowCount;
    }

    private boolean anyNestedShift() {
        for (final Column column : _columns) {
            if (column == null) continue;
            if (column.getValue().getShift() > 0
                    && column.getGroupHierarchy().depth() > 0) return true;
        }
        return false;
    }

    /** Walks the header region (leaf headers, group cells, vertical merges)
     *  and invokes the visitor for each. Used by writers and the comment
     *  applicator to share the run-identification algorithm. */
    public void forEachHeaderCell(final HeaderLayoutVisitor visitor) {
        if (!usesHeader()) return;
        final int originRow = _origin.getRow();
        if (anyNestedShift()) {
            // Single leaf row: emit every column header (flat + grouped leaves)
            // on the origin row. Group label rows and vertical merges are skipped;
            // cascade attributes still apply per-column.
            for (final Column column : _columns) {
                if (column == null) continue;
                visitor.visitColumnHeader(originRow, columnIndexOf(column), column);
            }
            return;
        }
        for (int depth = 0; depth < _headerRowCount; depth++) {
            _visitHeaderRow(originRow + depth, depth, visitor);
        }
        // Vertical merges for shallow-hierarchy (incl. flat) columns.
        final int leafRow = originRow + _headerRowCount - 1;
        for (final Column column : _columns) {
            if (column == null) continue;
            final int h = column.getGroupHierarchy().depth();
            if (h >= _headerRowCount - 1) continue;
            final int colIdx = columnIndexOf(column);
            final int rowspanStart = originRow + h;
            if (leafRow > rowspanStart) {
                visitor.visitVerticalMerge(rowspanStart, leafRow, colIdx);
            }
        }
    }

    private void _visitHeaderRow(final int row, final int depth, final HeaderLayoutVisitor visitor) {
        int i = 0;
        while (i < _columns.size()) {
            final Column col = _columns.get(i);
            if (col == null) { i++; continue; }
            final DataColumnGroup.Hierarchy hierarchy = col.getGroupHierarchy();
            final int h = hierarchy.depth();

            if (h == depth) {
                // Column hierarchy ends at this depth — leaf header text.
                visitor.visitColumnHeader(row, columnIndexOf(col), col);
                i++;
                continue;
            }
            if (h < depth) {
                // Hierarchy ended above this row — vertical merge covers, skip.
                i++;
                continue;
            }

            // h > depth: group cell at this depth, possibly merged horizontally
            // across adjacent columns sharing parent path + group name.
            final DataColumnGroup.Hierarchy parentPath = hierarchy.parentPath(depth);
            final String groupName = hierarchy.at(depth).getName();

            int j = i + 1;
            while (j < _columns.size()) {
                final Column col2 = _columns.get(j);
                if (col2 == null) break;
                final DataColumnGroup.Hierarchy h2 = col2.getGroupHierarchy();
                if (h2.depth() <= depth) break;
                if (!h2.parentPath(depth).equals(parentPath)) break;
                if (!h2.at(depth).getName().equals(groupName)) break;
                j++;
            }

            final int firstColIdx = columnIndexOf(col);
            final int lastColIdx = columnIndexOf(_columns.get(j - 1));
            visitor.visitGroupCell(row, firstColIdx, lastColIdx, hierarchy.at(depth));
            i = j;
        }
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
                .filter(c -> c != null && c
                .getPointer()
                .startsWith(filter)).collect(Collectors
                .toList());
    }

    public Styles buildStyles(final Workbook workbook) {
        return _stylesBuilder.build(workbook);
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
