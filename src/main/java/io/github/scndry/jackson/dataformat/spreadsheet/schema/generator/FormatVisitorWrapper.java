package io.github.scndry.jackson.dataformat.spreadsheet.schema.generator;

import com.fasterxml.jackson.databind.JavaType;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.jsonFormatVisitors.JsonArrayFormatVisitor;
import com.fasterxml.jackson.databind.jsonFormatVisitors.JsonFormatVisitorWrapper;
import com.fasterxml.jackson.databind.jsonFormatVisitors.JsonMapFormatVisitor;
import com.fasterxml.jackson.databind.jsonFormatVisitors.JsonObjectFormatVisitor;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.ColumnPointer;
import lombok.extern.slf4j.Slf4j;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

@Slf4j
public final class FormatVisitorWrapper extends JsonFormatVisitorWrapper.Base implements Iterable<Column> {

    private final ColumnPointer _pointer;
    private final DataGrid.Value _grid;
    private final DataColumn.Value _column;
    private final List<Column> _columns;

    public FormatVisitorWrapper() {
        this(ColumnPointer.empty(), DataGrid.Value.empty(), DataColumn.Value.empty(), null);
    }

    FormatVisitorWrapper(final FormatVisitorWrapper base, final ColumnPointer pointer) {
        this(pointer, base._grid, base._column, base._provider);
    }

    FormatVisitorWrapper(final ColumnPointer pointer, final DataGrid.Value sheet, final DataColumn.Value column, final SerializerProvider provider) {
        super(provider);
        _pointer = pointer;
        _grid = sheet;
        _column = column;
        _columns = new ArrayList<>();
    }

    @Override
    public JsonObjectFormatVisitor expectObjectFormat(final JavaType type) throws JsonMappingException {
        return new ObjectFormatVisitor(this, _provider);
    }

    @Override
    public JsonArrayFormatVisitor expectArrayFormat(final JavaType type) throws JsonMappingException {
        return new ArrayFormatVisitor(this, type, _provider);
    }

    @Override
    public JsonMapFormatVisitor expectMapFormat(final JavaType type) throws JsonMappingException {
        throw JsonMappingException.from((SerializerProvider) null, "Unsupported type: " + type);
    }

    @Override
    public Iterator<Column> iterator() {
        return _columns.iterator();
    }

    ColumnPointer getPointer() {
        return _pointer;
    }

    DataGrid.Value getSheet() {
        return _grid;
    }

    DataColumn.Value getColumn() {
        return _column;
    }

    boolean isEmpty() {
        return _columns.isEmpty();
    }

    void add(final Column column) {
        _columns.add(column);
    }

    void addAll(final Iterable<Column> columns) {
        for (final Column column : columns) {
            add(column);
        }
    }
}
