package io.github.scndry.jackson.dataformat.spreadsheet.schema;

import java.util.StringJoiner;

import lombok.EqualsAndHashCode;

import com.fasterxml.jackson.databind.JavaType;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;

/**
 * Metadata for a single spreadsheet column within a
 * {@link SpreadsheetSchema}. Holds a {@link ColumnPointer}
 * identifying the column's position in the JSON tree and the
 * {@link DataColumn.Value} annotation values.
 *
 * @see SpreadsheetSchema
 * @see ColumnPointer
 */
@EqualsAndHashCode
public final class Column {

    private static final String[] NO_ALIASES = new String[0];

    private final ColumnPointer _pointer;
    private final DataColumn.Value _value;
    private final JavaType _type;
    private final String[] _aliases;

    public Column(final ColumnPointer pointer, final DataColumn.Value value, final JavaType type) {
        this(pointer, value, type, NO_ALIASES);
    }

    public Column(final ColumnPointer pointer, final DataColumn.Value value,
                  final JavaType type, final String[] aliases) {
        this._pointer = pointer;
        this._value = value;
        this._type = type;
        this._aliases = aliases;
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

    public boolean matchesName(final String header) {
        if (getName().equals(header)) return true;
        for (final String alias : _aliases) {
            if (alias.equals(header)) return true;
        }
        return false;
    }

    public DataColumn.Value getValue() {
        return _value;
    }

    public boolean isMerge() {
        return _value.isMerge();
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
