package io.github.scndry.jackson.dataformat.spreadsheet.poi.ss;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;

import io.github.scndry.jackson.dataformat.spreadsheet.deser.CellValue;

/**
 * {@link CellValue} for POI-mode NUMERIC cells that defers
 * {@link DataFormatter#formatRawCellContents} until {@link #getStringValue()}.
 *
 * @see POISheetReader
 */
public final class POICellValue extends CellValue {

    private final Cell _cell;
    private final DataFormatter _formatter;
    private String _lazyText;
    private boolean _computed;

    public POICellValue(final double numberValue, final Cell cell, final DataFormatter formatter) {
        super(numberValue);
        _cell = cell;
        _formatter = formatter;
    }

    @Override
    public String getStringValue() {
        if (!_computed) {
            _lazyText = _formattedString();
            _computed = true;
        }
        return _lazyText;
    }

    private String _formattedString() {
        final CellStyle style = _cell.getCellStyle();
        if (style == null || style.getDataFormatString() == null) return null;
        return _formatter.formatRawCellContents(
                getNumberValue(),
                style.getDataFormat(),
                style.getDataFormatString());
    }
}
