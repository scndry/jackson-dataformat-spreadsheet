package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

/**
 * Lightweight representation of {@code CT_Cell} (ECMA-376 §18.3.1.4).
 * <p>
 * Field names match the XSD attribute/element names:
 * <ul>
 *   <li>{@code r} — cell reference (e.g. "A1")</li>
 *   <li>{@code t} — cell type ({@link STCellType})</li>
 *   <li>{@code v} — cell value text</li>
 *   <li>{@code ft} — formula type from {@code <f t="...">} ({@link STCellFormulaType})</li>
 *   <li>{@code is} — inline string text from {@code <is>}</li>
 * </ul>
 */
final class CTCell {

    private final String _r;
    private final STCellType _t;
    private final String _v;
    private final STCellFormulaType _ft;
    private final String _is;

    CTCell(final String r, final STCellType t, final String v,
           final STCellFormulaType ft, final String is) {
        _r = r;
        _t = t;
        _v = v;
        _ft = ft;
        _is = is;
    }

    String getR() { return _r; }
    STCellType getT() { return _t; }
    String getV() { return _v; }
    STCellFormulaType getFt() { return _ft; }
    String getIs() { return _is; }
}
