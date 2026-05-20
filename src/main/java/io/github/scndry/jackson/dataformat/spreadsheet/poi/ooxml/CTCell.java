package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

/**
 * Lightweight representation of {@code CT_Cell} (ECMA-376 §18.3.1.4).
 * <p>
 * Field names match the XSD attribute/element names:
 * <ul>
 *   <li>{@code r} — cell reference (e.g. "A1")</li>
 *   <li>{@code s} — cell style (XF) index</li>
 *   <li>{@code t} — cell type ({@link STCellType})</li>
 *   <li>{@code v} — cell value text</li>
 *   <li>{@code ft} — formula type from {@code <f t="...">} ({@link STCellFormulaType})</li>
 *   <li>{@code is} — inline string text from {@code <is>}</li>
 * </ul>
 */
final class CTCell {

    private String _r;
    private STCellType _t;
    private int _s;
    private String _v;
    private STCellFormulaType _ft;
    private String _is;

    void set(final String r, final STCellType t, final int s, final String v,
             final STCellFormulaType ft, final String is) {
        _r = r;
        _t = t;
        _s = s;
        _v = v;
        _ft = ft;
        _is = is;
    }

    String getR() { return _r; }
    STCellType getT() { return _t; }
    int getS() { return _s; }
    String getV() { return _v; }
    STCellFormulaType getFt() { return _ft; }
    String getIs() { return _is; }
}
