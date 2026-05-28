package io.github.scndry.jackson.dataformat.spreadsheet.schema.internal;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.ColumnPointer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.HeaderLayoutVisitor;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Styles;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.GridConfigurer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.style.StylesBuilder;

/**
 * Library-internal implementation of {@link SpreadsheetSchema}. Not part of
 * the public API; callers outside this library must not depend on it.
 */
public final class SpreadsheetSchemaImpl implements SpreadsheetSchema {

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

    public SpreadsheetSchemaImpl(
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

    public ColumnPointer resolve(final ColumnPointer parent, final String name) {
        final Map<String, ColumnPointer> nameMap = _resolveTable.get(parent);
        if (nameMap != null) {
            final ColumnPointer cached = nameMap.get(name);
            if (cached != null) return cached;
        }
        return parent.resolve(name);
    }

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

    @Override
    public Iterator<Column> iterator() {
        return _columns.iterator();
    }

    public Column findColumn(final CellAddress reference) {
        return findColumn(reference.getColumn());
    }

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

    @Override
    public Column column(final int index) {
        return _columns.get(index - getOriginColumn());
    }

    public Column column(final CellAddress reference) {
        return column(reference.getColumn());
    }

    @Override
    public Column column(final String name) {
        for (final Column col : _columns) {
            if (col != null && col.matchesName(name)) return col;
        }
        return null;
    }

    @Override
    public String columnName(final int index) {
        final Column col = column(index);
        return col == null ? null : col.getName();
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

    public SpreadsheetSchemaImpl reorderColumns(final List<String> headers) {
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
        return new SpreadsheetSchemaImpl(reordered, _origin, _features, _stylesBuilder, _gridConfigurer);
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

    @Override
    public int columnIndex(final String name) {
        for (int i = 0; i < _columns.size(); i++) {
            final Column col = _columns.get(i);
            if (col != null && col.matchesName(name)) {
                return i + getOriginColumn();
            }
        }
        return -1;
    }

    @Override
    public int size() {
        return _columns.size();
    }

    @Override
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
