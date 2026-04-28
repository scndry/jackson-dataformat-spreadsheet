package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import java.util.HashMap;
import java.util.Map;

/**
 * Lightweight representation of {@code ST_CellType} (ECMA-376 §18.18.11).
 * <p>
 * Replaces XMLBeans {@code STCellType.Enum} with a static enum mapped
 * directly from StAX attribute values.
 */
enum STCellType {
    BOOLEAN("b"),
    NUMBER("n"),
    ERROR("e"),
    SHARED_STRING("s"),
    FORMULA_STRING("str"),
    INLINE_STRING("inlineStr");

    private static final Map<String, STCellType> BY_CODE;

    static {
        BY_CODE = new HashMap<>();
        for (STCellType type : values()) {
            BY_CODE.put(type._code, type);
        }
    }

    private final String _code;

    STCellType(final String code) {
        _code = code;
    }

    /** Maps the {@code t} attribute value to an enum. Defaults to {@link #NUMBER} when absent. */
    static STCellType of(final String code) {
        if (code == null) return NUMBER;
        final STCellType type = BY_CODE.get(code);
        if (type == null) {
            throw new IllegalStateException("Unknown ST_CellType: " + code);
        }
        return type;
    }
}
