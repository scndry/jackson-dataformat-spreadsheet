package io.github.scndry.jackson.dataformat.spreadsheet.deser;

import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.usermodel.CellType;

public final class CellValue {

    public static final CellValue BLANK = new CellValue(CellType.BLANK, 0.0, false, null, 0);
    public static final CellValue TRUE = new CellValue(CellType.BOOLEAN, 0.0, true, null, 0);
    public static final CellValue FALSE = new CellValue(CellType.BOOLEAN, 0.0, false, null, 0);

    private final CellType _cellType;
    private final double _numberValue;
    private final boolean _booleanValue;
    private final String _textValue;
    private final int _errorCode;

    private CellValue(final CellType cellType, final double numberValue, final boolean booleanValue, final String textValue, final int errorCode) {
        _cellType = cellType;
        _numberValue = numberValue;
        _booleanValue = booleanValue;
        _textValue = textValue;
        _errorCode = errorCode;
    }

    public CellValue(final double numberValue) {
        this(CellType.NUMERIC, numberValue, false, null, 0);
    }

    public CellValue(final double numberValue, final String stringValue) {
        this(CellType.NUMERIC, numberValue, false, stringValue, 0);
    }

    public CellValue(final String stringValue) {
        this(CellType.STRING, 0.0, false, stringValue, 0);
    }

    public static CellValue valueOf(final boolean booleanValue) {
        return booleanValue ? TRUE : FALSE;
    }

    public static CellValue getError(final int errorCode) {
        return new CellValue(CellType.ERROR, 0.0, false, null, errorCode);
    }

    public boolean getBooleanValue() {
        return _booleanValue;
    }

    public double getNumberValue() {
        return _numberValue;
    }

    public String getStringValue() {
        return _textValue;
    }

    public CellType getCellType() {
        return _cellType;
    }

    public byte getErrorValue() {
        return (byte) _errorCode;
    }

    public boolean noFractionalPart() {
        return getNumberValue() % 1 == 0;
    }

    public String toString() {
        return _cellType + " [" + formatAsString() + "]";
    }

    public String formatAsString() {
        switch (_cellType) {
            case BLANK:
                return "<blank>";
            case NUMERIC:
                return String.valueOf(_numberValue);
            case STRING:
                return '"' + _textValue + '"';
            case BOOLEAN:
                return _booleanValue ? "TRUE" : "FALSE";
            case ERROR:
                return ErrorEval.getText(_errorCode);
            default:
                return "<error unexpected cell type " + _cellType + ">";
        }

    }
}
