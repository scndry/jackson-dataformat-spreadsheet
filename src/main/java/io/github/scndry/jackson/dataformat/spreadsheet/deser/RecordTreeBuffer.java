package io.github.scndry.jackson.dataformat.spreadsheet.deser;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.CellType;

import com.fasterxml.jackson.core.JsonToken;

import io.github.scndry.jackson.dataformat.spreadsheet.SheetStreamReadException;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.ColumnPointer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.internal.SpreadsheetSchemaImpl;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.internal.BackWriteProjection;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.internal.SchemaAnchorInspector;

/**
 * Record-tree buffer keyed by array scope. Each anchored record
 * (root or nested) carries its outer cells plus sibling-grouped
 * child records; close walks the tree depth-first to emit JSON in
 * {outer, child1[], child2[], ...} order. Sibling lists,
 * outer-after-list, and N-depth all fall out of the tree.
 *
 * <p>Buffer footprint per record: outer Cell refs + LinkedHashMap of
 * child lists. Cells hold Column/CellValue refs only — no token
 * serialization overhead.
 */
final class RecordTreeBuffer {

    private final SpreadsheetSchemaImpl _schema;
    private final List<ColumnPointer> _anchorScopesByDepth;
    private final Set<ColumnPointer> _leafArrayScopes;
    private final Map<ColumnPointer, ColumnPointer> _parentArrayScopeOf;
    private final Map<Column, ColumnPointer> _immediateScopeByColumn;
    private final Map<Column, Integer> _columnPositionByColumn;
    private final Map<Column, List<String>> _relativeSegmentsByColumn;
    private final long _bufferLimitBytes;
    private final boolean _blankRowAsNull;
    private final boolean _breakOnBlankRow;

    private final List<Cell> _rowBuffer = new ArrayList<>();
    private final Map<ColumnPointer, RecordNode> _openRecords = new HashMap<>();
    private final Map<ColumnPointer, Object> _lastAnchorByScope = new HashMap<>();
    private long _bufferedCells;

    RecordTreeBuffer(final SpreadsheetSchemaImpl schema) {
        this(schema, BackWriteProjection.backWriteBufferLimit(), true, false);
    }

    RecordTreeBuffer(final SpreadsheetSchemaImpl schema, final long bufferLimitBytes) {
        this(schema, bufferLimitBytes, true, false);
    }

    RecordTreeBuffer(final SpreadsheetSchemaImpl schema, final long bufferLimitBytes,
                  final boolean blankRowAsNull, final boolean breakOnBlankRow) {
        _schema = schema;
        _anchorScopesByDepth = _collectAnchorScopes(schema);
        _leafArrayScopes = _findLeafArrayScopes(schema);
        _parentArrayScopeOf = _computeParentScopeMap(schema);
        _immediateScopeByColumn = _computeImmediateScopeByColumn(schema);
        _columnPositionByColumn = _computeColumnPositionByColumn(schema);
        _relativeSegmentsByColumn = _computeRelativeSegmentsByColumn(schema);
        _bufferLimitBytes = bufferLimitBytes;
        _blankRowAsNull = blankRowAsNull;
        _breakOnBlankRow = breakOnBlankRow;
    }

    private static Map<Column, ColumnPointer> _computeImmediateScopeByColumn(
            final SpreadsheetSchemaImpl schema) {
        final Map<Column, ColumnPointer> result = new HashMap<>();
        for (final Column c : schema) {
            if (c == null) continue;
            result.put(c, SchemaAnchorInspector.immediateScope(c.getPointer()));
        }
        return result;
    }

    private static Map<Column, Integer> _computeColumnPositionByColumn(
            final SpreadsheetSchemaImpl schema) {
        final Map<Column, Integer> result = new HashMap<>();
        int i = 0;
        for (final Column c : schema) {
            if (c != null) result.put(c, i);
            i++;
        }
        return result;
    }

    private static Map<Column, List<String>> _computeRelativeSegmentsByColumn(
            final SpreadsheetSchemaImpl schema) {
        final Map<Column, List<String>> result = new HashMap<>();
        for (final Column c : schema) {
            if (c == null) continue;
            final ColumnPointer scope = SchemaAnchorInspector.immediateScope(c.getPointer());
            final List<String> segs = new ArrayList<>();
            for (final ColumnPointer seg : scope.relativize(c.getPointer())) {
                segs.add(seg.name());
            }
            result.put(c, java.util.Collections.unmodifiableList(segs));
        }
        return result;
    }

    void onSheetDataStart(final Emitter out) {
        out.token(JsonToken.START_ARRAY);
    }

    void onRowStart() {
        _rowBuffer.clear();
    }

    void onCellValue(final Column column, final CellValue value) throws SheetStreamReadException {
        if (!_isSupportedCellType(value.getCellType())) {
            throw SheetStreamReadException.unexpected(null, value.getCellType());
        }
        _rowBuffer.add(new Cell(column, value));
    }

    private static boolean _isSupportedCellType(final CellType type) {
        switch (type) {
            case NUMERIC:
            case STRING:
            case BLANK:
            case BOOLEAN:
                return true;
            default:
                return false;
        }
    }

    /** Returns false when iteration should stop (BREAK_ON_BLANK_ROW). */
    boolean onRowEnd(final Emitter out) throws SheetStreamReadException {
        if (_rowBuffer.isEmpty()) {
            if (_breakOnBlankRow) {
                _closeRecordsAtOrDeeperThan(ColumnPointer.empty(), out);
                return false;
            }
            if (_blankRowAsNull) {
                _closeRecordsAtOrDeeperThan(ColumnPointer.empty(), out);
                out.token(JsonToken.VALUE_NULL);
            }
            return true;
        }
        for (final ColumnPointer scope : _anchorScopesByDepth) {
            final Column anchorCol = SchemaAnchorInspector.findAnchorColumn(_schema, scope);
            if (anchorCol == null) continue;
            final Cell cell = _findCell(anchorCol);
            if (cell == null) continue;
            final CellValue v = cell.value;
            if (v == null || v.getCellType() == CellType.BLANK) continue;
            final Object newKey = _anchorKey(v);
            final Object lastKey = _lastAnchorByScope.get(scope);
            if (newKey.equals(lastKey)) continue;

            _closeRecordsAtOrDeeperThan(scope, out);

            final RecordNode openRecord = new RecordNode(scope);
            for (final Cell c : _rowBuffer) {
                if (c.value == null || c.value.getCellType() == CellType.BLANK) continue;
                if (_immediateScopeByColumn.get(c.column).equals(scope)) {
                    _addSortedByColumn(openRecord.outerCells, c);
                    _bufferedCells++;
                }
            }
            _openRecords.put(scope, openRecord);
            _lastAnchorByScope.put(scope, newKey);
        }

        for (final ColumnPointer leafScope : _leafArrayScopes) {
            final RecordNode leafRecord = new RecordNode(leafScope);
            for (final Cell c : _rowBuffer) {
                if (c.value == null || c.value.getCellType() == CellType.BLANK) continue;
                if (_immediateScopeByColumn.get(c.column).equals(leafScope)) {
                    _addSortedByColumn(leafRecord.outerCells, c);
                }
            }
            if (leafRecord.outerCells.isEmpty()) continue;
            _bufferedCells += leafRecord.outerCells.size();

            final ColumnPointer parentScope = _parentArrayScopeOf.get(leafScope);
            final RecordNode parentRecord = _openRecords.get(parentScope);
            parentRecord.childRecords
                    .computeIfAbsent(leafScope, k -> new ArrayList<>())
                    .add(leafRecord);
        }

        _checkBufferLimit();
        return true;
    }

    void onSheetDataEnd(final Emitter out) {
        _closeRecordsAtOrDeeperThan(ColumnPointer.empty(), out);
        out.token(JsonToken.END_ARRAY);
    }

    private void _checkBufferLimit() throws SheetStreamReadException {
        final long bufferedBytes = _bufferedCells * (long) BackWriteProjection.CELL_MEMORY_BYTES;
        if (bufferedBytes <= _bufferLimitBytes) return;
        throw new SheetStreamReadException(null,
                "Nested-list buffer (" + BackWriteProjection.formatBytes(bufferedBytes)
                + ") exceeds limit (" + BackWriteProjection.formatBytes(_bufferLimitBytes)
                + "). Use USE_POI_USER_MODEL to bypass back-write buffering,"
                + " or split the outer record into smaller chunks.");
    }

    private void _closeRecordsAtOrDeeperThan(final ColumnPointer scope, final Emitter out) {
        final List<ColumnPointer> toClose = new ArrayList<>();
        for (final ColumnPointer s : _openRecords.keySet()) {
            if (s.equals(scope) || s.startsWith(scope)) toClose.add(s);
        }
        toClose.sort(Comparator.comparingInt(RecordTreeBuffer::_scopeDepth).reversed());
        for (final ColumnPointer s : toClose) {
            final RecordNode record = _openRecords.remove(s);
            _lastAnchorByScope.remove(s);
            _emitOrAttach(record, out);
        }
    }

    private void _emitOrAttach(final RecordNode record, final Emitter out) {
        final ColumnPointer parentScope = _parentArrayScopeOf.get(record.scope);
        final RecordNode parentRecord = parentScope == null ? null : _openRecords.get(parentScope);
        if (parentRecord != null) {
            parentRecord.childRecords
                    .computeIfAbsent(record.scope, k -> new ArrayList<>())
                    .add(record);
        } else {
            _emitRecord(record, out);
            _bufferedCells = 0;
        }
    }

    private void _emitRecord(final RecordNode record, final Emitter out) {
        out.token(JsonToken.START_OBJECT);
        final List<String> openPath = new ArrayList<>();
        for (final Cell c : record.outerCells) {
            final List<String> path = _relativeSegmentsByColumn.get(c.column);
            final int common = _commonPrefixLength(openPath, path, path.size() - 1);
            for (int i = openPath.size() - common; i > 0; i--) {
                out.token(JsonToken.END_OBJECT);
            }
            for (int i = common; i < path.size() - 1; i++) {
                out.fieldName(path.get(i));
                out.token(JsonToken.START_OBJECT);
            }
            out.fieldName(path.get(path.size() - 1));
            out.scalar(c.value, _scalarToken(c.value));
            openPath.clear();
            openPath.addAll(path.subList(0, path.size() - 1));
        }
        for (int i = openPath.size(); i > 0; i--) {
            out.token(JsonToken.END_OBJECT);
        }
        for (final Map.Entry<ColumnPointer, List<RecordNode>> e : record.childRecords.entrySet()) {
            out.fieldName(e.getKey().getParent().name());
            out.token(JsonToken.START_ARRAY);
            for (final RecordNode child : e.getValue()) {
                _emitRecord(child, out);
            }
            out.token(JsonToken.END_ARRAY);
        }
        out.token(JsonToken.END_OBJECT);
    }

    private static int _commonPrefixLength(final List<String> a, final List<String> b, final int bLimit) {
        final int n = Math.min(a.size(), bLimit);
        int i = 0;
        while (i < n && a.get(i).equals(b.get(i))) i++;
        return i;
    }

    private void _addSortedByColumn(final List<Cell> cells, final Cell c) {
        final int pos = _columnPositionByColumn.get(c.column);
        int i = 0;
        while (i < cells.size()
                && _columnPositionByColumn.get(cells.get(i).column) < pos) {
            i++;
        }
        cells.add(i, c);
    }

    private Cell _findCell(final Column column) {
        for (final Cell c : _rowBuffer) {
            if (c.column == column) return c;
        }
        return null;
    }

    private static List<ColumnPointer> _collectAnchorScopes(final SpreadsheetSchemaImpl schema) {
        final List<ColumnPointer> scopes = new ArrayList<>();
        if (SchemaAnchorInspector.findAnchorColumn(schema, ColumnPointer.empty()) != null) {
            scopes.add(ColumnPointer.empty());
        }
        final List<ColumnPointer> arrayScopes =
                new ArrayList<>(SchemaAnchorInspector.allArrayScopes(schema));
        arrayScopes.sort(Comparator.comparingInt(RecordTreeBuffer::_scopeDepth));
        for (final ColumnPointer s : arrayScopes) {
            if (SchemaAnchorInspector.findAnchorColumn(schema, s) != null) scopes.add(s);
        }
        return scopes;
    }

    private static Set<ColumnPointer> _findLeafArrayScopes(final SpreadsheetSchemaImpl schema) {
        final Set<ColumnPointer> all = SchemaAnchorInspector.allArrayScopes(schema);
        final Set<ColumnPointer> leafs = new LinkedHashSet<>();
        for (final ColumnPointer scope : all) {
            boolean hasChild = false;
            for (final ColumnPointer other : all) {
                if (_isImmediateChildArrayScope(scope, other)) { hasChild = true; break; }
            }
            if (!hasChild) leafs.add(scope);
        }
        return leafs;
    }

    private static Map<ColumnPointer, ColumnPointer> _computeParentScopeMap(
            final SpreadsheetSchemaImpl schema) {
        final Map<ColumnPointer, ColumnPointer> result = new HashMap<>();
        for (final ColumnPointer scope : SchemaAnchorInspector.allArrayScopes(schema)) {
            result.put(scope, _parentArrayScopeOf(scope));
        }
        return result;
    }

    private static ColumnPointer _parentArrayScopeOf(final ColumnPointer scope) {
        ColumnPointer head = ColumnPointer.empty();
        ColumnPointer parent = ColumnPointer.empty();
        for (final ColumnPointer seg : scope) {
            head = head.resolve(seg);
            if (head.equals(scope)) break;
            if (seg.equals(ColumnPointer.array())) parent = head;
        }
        return parent;
    }

    private static boolean _isImmediateChildArrayScope(
            final ColumnPointer parent, final ColumnPointer child) {
        if (!child.startsWith(parent) || child.equals(parent)) return false;
        return _scopeDepth(child) == _scopeDepth(parent) + 1;
    }

    private static int _scopeDepth(final ColumnPointer scope) {
        int d = 0;
        for (final ColumnPointer seg : scope) {
            if (seg.equals(ColumnPointer.array())) d++;
        }
        return d;
    }

    private static JsonToken _scalarToken(final CellValue v) {
        switch (v.getCellType()) {
            case NUMERIC: return v.getNumberValue() % 1 == 0
                    ? JsonToken.VALUE_NUMBER_INT : JsonToken.VALUE_NUMBER_FLOAT;
            case STRING: return JsonToken.VALUE_STRING;
            case BLANK: return JsonToken.VALUE_NULL;
            case BOOLEAN: return v.getBooleanValue()
                    ? JsonToken.VALUE_TRUE : JsonToken.VALUE_FALSE;
            default: throw new IllegalStateException(
                    "Unsupported cell type leaked past onCellValue: " + v.getCellType());
        }
    }

    private static Object _anchorKey(final CellValue v) {
        switch (v.getCellType()) {
            case NUMERIC: return v.getNumberValue();
            case STRING: return v.getStringValue();
            case BOOLEAN: return v.getBooleanValue();
            default: return null;
        }
    }

    interface Emitter {
        void token(JsonToken token);
        void fieldName(String name);
        void scalar(CellValue value, JsonToken scalarToken);
    }

    private static final class Cell {
        final Column column;
        final CellValue value;
        Cell(final Column c, final CellValue v) { column = c; value = v; }
    }

    private static final class RecordNode {
        final ColumnPointer scope;
        final List<Cell> outerCells = new ArrayList<>();
        final Map<ColumnPointer, List<RecordNode>> childRecords = new LinkedHashMap<>();
        RecordNode(final ColumnPointer scope) { this.scope = scope; }
    }
}
