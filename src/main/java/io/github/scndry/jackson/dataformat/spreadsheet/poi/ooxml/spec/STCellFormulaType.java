package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.spec;

import java.util.HashMap;
import java.util.Map;

/**
 * Lightweight representation of {@code ST_CellFormulaType} (ECMA-376 §18.18.6).
 */
public enum STCellFormulaType {
    NORMAL("normal"),
    ARRAY("array"),
    DATA_TABLE("dataTable"),
    SHARED("shared");

    private static final Map<String, STCellFormulaType> BY_CODE;

    static {
        BY_CODE = new HashMap<>();
        for (STCellFormulaType type : values()) {
            BY_CODE.put(type._code, type);
        }
    }

    private final String _code;

    STCellFormulaType(final String code) {
        _code = code;
    }

    /**
     * Maps the {@code t} attribute value of {@code <f>}.
     * Defaults to {@link #NORMAL} when absent.
     */
    public static STCellFormulaType of(final String code) {
        if (code == null) return NORMAL;
        final STCellFormulaType type = BY_CODE.get(code);
        if (type == null) {
            throw new IllegalStateException("Unknown ST_CellFormulaType: " + code);
        }
        return type;
    }
}
