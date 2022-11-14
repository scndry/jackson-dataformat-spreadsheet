package io.github.scndry.jackson.dataformat.spreadsheet.schema.generator;

import com.fasterxml.jackson.databind.JavaType;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.jsonFormatVisitors.JsonArrayFormatVisitor;
import com.fasterxml.jackson.databind.jsonFormatVisitors.JsonFormatTypes;
import com.fasterxml.jackson.databind.jsonFormatVisitors.JsonFormatVisitable;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.ColumnPointer;

final class ArrayFormatVisitor extends JsonArrayFormatVisitor.Base {

    private final FormatVisitorWrapper _wrapper;
    private final JavaType _type;

    ArrayFormatVisitor(final FormatVisitorWrapper wrapper, final JavaType type, final SerializerProvider provider) {
        super(provider);
        _wrapper = wrapper;
        _type = type;
    }

    @Override
    public void itemsFormat(final JsonFormatVisitable handler, final JavaType elementType) throws JsonMappingException {
        final ColumnPointer pointer = _wrapper.getPointer().resolveArray();
        final FormatVisitorWrapper visitor = new FormatVisitorWrapper(_wrapper, pointer);
        handler.acceptJsonFormatVisitor(visitor, elementType);
        if (visitor.isEmpty()) {
            _wrapper.add(new Column(pointer, _wrapper.getColumn(), _type));
        } else {
            _wrapper.addAll(visitor);
        }
    }

    @Override
    public void itemsFormat(final JsonFormatTypes format) throws JsonMappingException {
        final ColumnPointer pointer = _wrapper.getPointer().resolveArray();
        _wrapper.add(new Column(pointer, _wrapper.getColumn(), _type));
    }
}
