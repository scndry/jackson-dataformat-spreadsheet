package io.github.scndry.jackson.dataformat.spreadsheet.schema;

import com.fasterxml.jackson.databind.JavaType;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import lombok.EqualsAndHashCode;

import java.util.StringJoiner;

@EqualsAndHashCode
public final class Column {

    private final ColumnPointer _pointer;
    private final DataColumn.Value _value;
    private final JavaType _type;

    public Column(final ColumnPointer pointer, final DataColumn.Value value, final JavaType type) {
        this._pointer = pointer;
        this._value = value;
        this._type = type;
    }

    public ColumnPointer getPointer() {
        return _pointer;
    }

    public boolean matches(final ColumnPointer other) {
        if (_pointer.equals(other)) {
            return true;
        } else if (isArray()) {
            return _pointer.getParent().equals(other);
        }
        return false;
    }

    public String getName() {
        final String name = _value.getName();
        if (name.isEmpty()) {
            return _pointer.toString();
        }
        return name;
    }

    public DataColumn.Value getValue() {
        return _value;
    }

    public JavaType getType() {
        if (isArray()) {
            return _type.getContentType();
        }
        return _type;
    }

    private boolean isArray() {
        return _type.isArrayType() || _type.isCollectionLikeType();
    }

    @Override
    public String toString() {
        return new StringJoiner(", ", Column.class.getSimpleName() + "[", "]")
                .add("pointer=" + _pointer)
                .add("value=" + _value)
                .add("type=" + _type)
                .toString();
    }
}
